# -*- coding: utf-8 -*-
# Author: Opera-Jin
"""OOXML 引擎：.xlsx / .xlsm / .xltx / .xltm 以及新版 WPS .et。

设计原则
--------
1. 只做“删除”操作，绝不重写单元格、样式、图表、宏等任何 part；
2. 所有修改都以字符串区间的方式作用在 workbook.xml 原文上，
   未命中的字节 100% 原样保留（避免 ElementTree 重排命名空间前缀
   导致 mc:Ignorable 失效、Excel 报“内容有问题”）；
3. 删除名称前先扫描全表公式，被引用的名称一律保留，杜绝 #NAME?。
"""
from __future__ import annotations

import os
import re
import zipfile
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple

from .util import atomic_replace, human_size, make_temp_beside, unescape_full, xml_unescape

# ---------------------------------------------------------------- 常量 / 正则

W_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

DEFINED_NAME_RE = re.compile(r"<definedName\b([^>]*?)(?:/>|>(.*?)</definedName>)", re.S)
ATTR_RE = re.compile(r'([A-Za-z_][\w:.\-]*)\s*=\s*"([^"]*)"')
SHEET_RE = re.compile(r"<sheet\b([^>]*?)/>", re.S)
EXTERNAL_REF_RE = re.compile(r"<externalReference\b[^>]*?/>", re.S)
RELATIONSHIP_RE = re.compile(r"<Relationship\b[^>]*?/>", re.S)
OVERRIDE_RE = re.compile(r"<Override\b[^>]*?/>", re.S)

# 公式语料：只取可能承载“名称引用”的元素
FORMULA_PATTERNS = (
    re.compile(r"<f\b[^>]*>(.*?)</f>", re.S),
    re.compile(r"<formula(\d?)>(.*?)</formula\1>", re.S),
    re.compile(r"<c:f>(.*?)</c:f>", re.S),
    re.compile(r"<calculatedColumnFormula>(.*?)</calculatedColumnFormula>", re.S),
    re.compile(r"<totalsRowFormula>(.*?)</totalsRowFormula>", re.S),
    re.compile(r"<worksheetSource\b[^>]*?\bname=\"([^\"]*)\"", re.S),
)

CELL_RE = re.compile(r"<c\b([^>]*?)(?:/>|>(.*?)</c>)", re.S)
CELL_F_RE = re.compile(r"<f\b([^>]*?)(?:/>|>(.*?)</f>)", re.S)
CELL_V_RE = re.compile(r"<v>(.*?)</v>", re.S)
EXT_TOKEN_RE = re.compile(r"\[\d+\]")

# 标识符：名称允许中文、下划线、反斜杠、点
TOKEN_RE = re.compile(r"[A-Za-z_\\\u00c0-\uffff][A-Za-z0-9_.\\\u00c0-\uffff]*")

# 默认保留的系统名称（打印区域 / 打印标题 / 筛选范围）：属于“格式”而非数据
SYSTEM_KEEP_KEYS = {"print_area", "print_titles", "_filterdatabase"}

# 结果明细上限：30 万条全存下来没意义，只会拖慢界面与日志
DETAIL_CAP = 3000

FORMULA_PARTS = (
    "xl/worksheets/",
    "xl/charts/",
    "xl/tables/",
    "xl/pivotCache/",
    "xl/pivotTables/",
    "xl/workbook.xml",
)


def is_ooxml(path: str) -> bool:
    try:
        with open(path, "rb") as fh:
            return fh.read(2) == b"PK"
    except OSError:
        return False


# ---------------------------------------------------------------- 数据结构


@dataclass(slots=True)
class NameItem:
    name: str
    scope_index: Optional[int]
    scope_label: str
    refers_to: str
    hidden: bool
    is_system: bool
    is_invalid: bool
    is_external: bool
    is_referenced: bool
    action: str = "delete"
    reason: str = ""


@dataclass(slots=True)
class LinkItem:
    rid: str
    target: str
    part: str


@dataclass
class Analysis:
    names: List[NameItem] = field(default_factory=list)
    links: List[LinkItem] = field(default_factory=list)
    sheet_names: List[str] = field(default_factory=list)
    ext_formula_refs: int = 0
    size: int = 0

    @property
    def invalid_names(self) -> int:
        return sum(1 for n in self.names if n.is_invalid)

    @property
    def referenced_names(self) -> int:
        return sum(1 for n in self.names if n.is_referenced)


@dataclass
class OoxmlOutcome:
    ok: bool
    message: str = ""
    names_total: int = 0
    names_removed: int = 0
    names_kept: int = 0
    links_removed: int = 0
    converted_formulas: int = 0
    ext_formula_refs: int = 0
    size_before: int = 0
    size_after: int = 0
    removed_names: List[str] = field(default_factory=list)
    kept_names: List[Tuple[str, str]] = field(default_factory=list)
    backup_path: str = ""


# ---------------------------------------------------------------- 解析


def _attrs(raw: str) -> Dict[str, str]:
    return {k: v for k, v in ATTR_RE.findall(raw)}


def _sheet_names(workbook_xml: str) -> List[str]:
    out = []
    for m in SHEET_RE.finditer(workbook_xml):
        a = _attrs(m.group(1))
        out.append(unescape_full(a.get("name", "")))
    return out


def _collect_tokens(zf: zipfile.ZipFile, extra_texts: List[str]) -> set:
    """把所有公式中出现过的标识符收集成大写集合，用于判断名称是否被引用。"""
    tokens: set = set()

    def feed(text: str):
        for t in TOKEN_RE.findall(unescape_full(text)):
            tokens.add(t.upper())

    for name in zf.namelist():
        if not any(name.startswith(p) or name == p for p in FORMULA_PARTS):
            continue
        if name.endswith(".rels") or name.endswith(".vml"):
            continue
        try:
            blob = zf.read(name)
        except (KeyError, zipfile.BadZipFile):
            continue
        if not any(k in blob for k in (b"<f", b"formula", b"<c:f", b"worksheetSource")):
            continue
        text = blob.decode("utf-8", "replace")
        for pat in FORMULA_PATTERNS:
            for m in pat.finditer(text):
                feed(m.group(m.lastindex or 1))
    for text in extra_texts:
        feed(text)
    return tokens


def analyze(path: str, progress=None) -> Analysis:
    """只读分析：列出所有名称与外部链接，并判定处置建议。"""
    res = Analysis()
    res.size = os.path.getsize(path)
    with zipfile.ZipFile(path, "r") as zf:
        names = set(zf.namelist())
        if "xl/workbook.xml" not in names:
            raise ValueError("不是有效的 Excel 工作簿（缺少 xl/workbook.xml）")
        wb = zf.read("xl/workbook.xml").decode("utf-8", "replace")
        res.sheet_names = _sheet_names(wb)

        raw_names = list(DEFINED_NAME_RE.finditer(wb))
        _n_all = len(raw_names)
        refers_texts = []
        for m in raw_names:
            body = m.group(2) or ""
            refers_texts.append(unescape_full(body))

        tokens = _collect_tokens(zf, refers_texts)

        for _i, m in enumerate(raw_names):
            if progress and _i and _i % 20000 == 0:
                progress("分析名称", _i, _n_all)
            a = _attrs(m.group(1))
            nm = unescape_full(a.get("name", ""))
            body = unescape_full(m.group(2) or "")
            scope = a.get("localSheetId")
            scope_i = int(scope) if scope is not None and scope.isdigit() else None
            if scope_i is not None and 0 <= scope_i < len(res.sheet_names):
                scope_label = res.sheet_names[scope_i]
            elif scope_i is not None:
                scope_label = "工作表#%d" % scope_i
            else:
                scope_label = "全工作簿"
            low = nm.lower()
            item = NameItem(
                name=nm,
                scope_index=scope_i,
                scope_label=scope_label,
                refers_to=body,
                hidden=a.get("hidden") == "1",
                is_system=low.startswith("_xlnm."),
                is_invalid="#REF!" in body.upper(),
                is_external=bool(re.search(r"\[\d+\]", body)),
                is_referenced=nm.upper() in tokens,
            )
            res.names.append(item)

        # 外部链接
        rels_name = "xl/_rels/workbook.xml.rels"
        rels = zf.read(rels_name).decode("utf-8", "replace") if rels_name in names else ""
        rel_map = {}
        for m in RELATIONSHIP_RE.finditer(rels):
            a = _attrs(m.group(0))
            rel_map[a.get("Id", "")] = a.get("Target", "")
        for m in EXTERNAL_REF_RE.finditer(wb):
            a = _attrs(m.group(0))
            rid = a.get("r:id") or a.get("id") or ""
            target = rel_map.get(rid, "")
            part = _resolve_part("xl", target)
            res.links.append(LinkItem(rid=rid, target=target, part=part))

        # 统计引用外链的公式数量（用于风险提示）
        cnt = 0
        for name in zf.namelist():
            if not name.startswith("xl/worksheets/"):
                continue
            blob = zf.read(name)
            if b"[" not in blob:
                continue
            text = blob.decode("utf-8", "replace")
            for pat in FORMULA_PATTERNS[:1]:
                for m in pat.finditer(text):
                    if re.search(r"\[\d+\]", m.group(1)):
                        cnt += 1
        res.ext_formula_refs = cnt
    return res


def _resolve_part(base: str, target: str) -> str:
    if not target:
        return ""
    target = target.replace("\\", "/")
    if target.startswith("/"):
        return target.lstrip("/")
    parts = (base + "/" + target).split("/")
    out = []
    for p in parts:
        if p in ("", "."):
            continue
        if p == "..":
            if out:
                out.pop()
        else:
            out.append(p)
    return "/".join(out)


# ---------------------------------------------------------------- 决策


def decide(analysis: Analysis, remove_print_names: bool, remove_links: bool) -> None:
    """为每个名称标注 action / reason。规则顺序即安全优先级。"""
    for it in analysis.names:
        low = it.name.lower()
        syskey = low[6:] if low.startswith("_xlnm.") else ""
        if it.is_invalid:
            it.action, it.reason = "delete", "引用已失效（#REF!）"
        elif remove_links and it.is_external:
            it.action, it.reason = "delete", "引用外部工作簿（外链将被移除）"
        elif it.is_system and syskey in SYSTEM_KEEP_KEYS and not remove_print_names:
            it.action, it.reason = "keep", "打印/筛选设置，保留"
        elif it.is_referenced:
            it.action, it.reason = "keep", "被公式引用，保留"
        else:
            it.action = "delete"
            it.reason = "系统内部名称" if it.is_system else "未被任何公式引用"


# ---------------------------------------------------------------- 主流程


def process(
    path: str,
    remove_print_names: bool = False,
    remove_links: bool = True,
    do_backup: bool = True,
    log=None,
    progress=None,
) -> OoxmlOutcome:
    out = OoxmlOutcome(ok=False)
    out.size_before = os.path.getsize(path)

    def rep(stage, done=0, total=0):
        if progress:
            try:
                progress(stage, done, total)
            except Exception:  # noqa: BLE001
                pass

    try:
        rep("分析名称", 0, 0)
        ana = analyze(path, progress=progress)
    except Exception as exc:  # noqa: BLE001
        out.message = "无法解析文件：%s" % exc
        return out

    decide(ana, remove_print_names, remove_links)
    out.names_total = len(ana.names)
    out.ext_formula_refs = ana.ext_formula_refs

    overrides: Dict[str, bytes] = {}
    drops: List[str] = []
    drop_prefixes: List[str] = []

    # ---- 1. 名称
    with zipfile.ZipFile(path, "r") as zf:
        wb = zf.read("xl/workbook.xml").decode("utf-8", "replace")
    matches = list(DEFINED_NAME_RE.finditer(wb))
    spans = []
    n_kept = 0
    _n_all = len(matches)
    for _i, (m, item) in enumerate(zip(matches, ana.names)):
        if _i and _i % 20000 == 0:
            rep("删除名称", _i, _n_all)
        if item.action == "delete":
            spans.append((m.start(), m.end()))
            if len(out.removed_names) < DETAIL_CAP:
                out.removed_names.append(item.name)
        else:
            n_kept += 1
            if len(out.kept_names) < DETAIL_CAP:
                out.kept_names.append((item.name, item.reason))
    out.names_removed = len(spans)
    out.names_kept = n_kept
    new_wb = _cut(wb, spans)
    if spans:
        new_wb = re.sub(r"<definedNames\s*>\s*</definedNames>", "", new_wb)
        new_wb = re.sub(r"<definedNames\s*/>", "", new_wb)

    # ---- 2. 外部链接
    if remove_links and ana.links:
        with zipfile.ZipFile(path, "r") as zf:
            names = set(zf.namelist())
            rels_name = "xl/_rels/workbook.xml.rels"
            rels = zf.read(rels_name).decode("utf-8", "replace") if rels_name in names else ""
            ct_name = "[Content_Types].xml"
            ct = zf.read(ct_name).decode("utf-8", "replace") if ct_name in names else ""

            # 2a. 先把引用外链的公式换成它此刻的缓存值 ——
            #     与 Excel「断开链接」完全一致：留下数值，而不是 #REF!
            converted = 0
            for part in list(names):
                if not (part.startswith("xl/worksheets/") and part.endswith(".xml")):
                    continue
                raw = zf.read(part).decode("utf-8", "replace")
                if not EXT_TOKEN_RE.search(raw):
                    continue
                new_raw, cnt = _break_formulas(raw)
                if cnt:
                    overrides[part] = new_raw.encode("utf-8")
                    converted += cnt
            out.converted_formulas = converted

            # 公式变了，计算链必须丢掉（Excel 会自己重建），
            # 否则残留的 calcChain 会让文件被判定为「需要修复」。
            if converted:
                drops.append("xl/calcChain.xml")
                ct = _cut_matching(
                    ct, OVERRIDE_RE,
                    lambda a: a.get("PartName", "").lstrip("/") == "xl/calcChain.xml")
                rels = _cut_matching(
                    rels, RELATIONSHIP_RE,
                    lambda a: "calcChain" in a.get("Type", ""))

        rids = {l.rid for l in ana.links}
        new_rels = _cut_matching(rels, RELATIONSHIP_RE, lambda a: a.get("Id") in rids)
        new_wb = _cut_matching(new_wb, EXTERNAL_REF_RE, lambda a: True)
        new_wb = re.sub(r"<externalReferences\s*>\s*</externalReferences>", "", new_wb)
        new_wb = re.sub(r"<externalReferences\s*/>", "", new_wb)
        removed_parts = set()
        for l in ana.links:
            if l.part:
                removed_parts.add(l.part)
                d, base = os.path.split(l.part)
                removed_parts.add("%s/_rels/%s.rels" % (d, base))
        new_ct = _cut_matching(
            ct,
            OVERRIDE_RE,
            lambda a: a.get("PartName", "").lstrip("/") in removed_parts,
        )
        overrides["xl/workbook.xml"] = new_wb.encode("utf-8")
        if rels_name in names:
            overrides[rels_name] = new_rels.encode("utf-8")
        if ct_name in names:
            overrides[ct_name] = new_ct.encode("utf-8")
        drops.extend(sorted(p for p in removed_parts if p))
        drop_prefixes.append("xl/externalLinks/")
        out.links_removed = len(ana.links)
    elif spans:
        overrides["xl/workbook.xml"] = new_wb.encode("utf-8")

    if not overrides and not drops:
        out.ok = True
        out.size_after = out.size_before
        out.message = "没有需要清理的内容"
        out.names_kept = out.names_total
        out.names_removed = 0
        return out

    rep("写入文件", 0, 0)
    tmp = make_temp_beside(path, ".xlsx")
    try:
        _rewrite_zip(path, tmp, overrides, drops, drop_prefixes)
        _sanity_check(tmp)
    except Exception as exc:  # noqa: BLE001
        try:
            os.remove(tmp)
        except OSError:
            pass
        out.message = "写入失败：%s" % exc
        return out

    if do_backup:
        from .util import copy_backup

        bak = copy_backup(path)
        out.backup_path = bak
        if log:
            log("已备份原文件 → %s" % bak)
    atomic_replace(tmp, path, keep_stat_from=path)
    out.size_after = os.path.getsize(path)
    out.ok = True
    out.message = "完成"
    return out


def _cut(text: str, spans: List[Tuple[int, int]]) -> str:
    """一次扫描拼装，避免 30 万条区间逐个切片导致的 O(n^2)。"""
    if not spans:
        return text
    spans.sort()
    out = []
    prev = 0
    for start, end in spans:
        if start < prev:
            continue
        out.append(text[prev:start])
        prev = end
    out.append(text[prev:])
    return "".join(out)


def _cut_matching(text: str, pattern: re.Pattern, keep_pred) -> str:
    spans = []
    for m in pattern.finditer(text):
        if keep_pred(_attrs(m.group(0))):
            spans.append((m.start(), m.end()))
    return _cut(text, spans)


def _rewrite_zip(src, dst, overrides, drops, drop_prefixes):
    drop_set = set(drops)
    with zipfile.ZipFile(src, "r") as zin, zipfile.ZipFile(
        dst, "w", zipfile.ZIP_DEFLATED, allowZip64=True
    ) as zout:
        for info in zin.infolist():
            name = info.filename
            if name in drop_set:
                continue
            if any(name.startswith(p) for p in drop_prefixes):
                continue
            if name in overrides:
                data = overrides[name]
            elif info.is_dir():
                data = b""
            else:
                data = zin.read(name)
            zi = zipfile.ZipInfo(name, date_time=info.date_time)
            zi.compress_type = info.compress_type
            zi.external_attr = info.external_attr
            zi.internal_attr = info.internal_attr
            zi.create_system = info.create_system
            zout.writestr(zi, data)


def _sanity_check(path: str):
    """结构自检：能重新打开、workbook.xml 仍是良构 XML。"""
    import xml.etree.ElementTree as ET

    with zipfile.ZipFile(path, "r") as zf:
        bad = zf.testzip()
        if bad:
            raise ValueError("压缩包校验失败：%s" % bad)
        wb = zf.read("xl/workbook.xml")
    ET.fromstring(wb)


# ---------------------------------------------------------------- 断开外链

def _retag(attrs: str, new_t: str) -> str:
    if re.search(r'\bt\s*=\s*"[^"]*"', attrs):
        return re.sub(r'\bt\s*=\s*"[^"]*"', 't="%s"' % new_t, attrs, count=1)
    return attrs + ' t="%s"' % new_t


def _break_formulas(xml: str):
    """把引用外部工作簿的公式替换成它此刻的缓存值。

    与 Excel 的「断开链接」行为一致：结果是数值，不会出现 #REF!。
    共享公式（t="shared" si=n）按组处理，避免只改主单元格导致整组失效。
    """
    cells = []
    for m in CELL_RE.finditer(xml):
        attrs = m.group(1) or ""
        body = m.group(2)
        if body is None:
            continue
        fm = CELL_F_RE.search(body)
        if not fm:
            continue
        fattrs = fm.group(1) or ""
        ftext = fm.group(2) or ""
        si = None
        si_m = re.search(r'\bsi\s*=\s*"(\d+)"', fattrs)
        if si_m:
            si = si_m.group(1)
        cells.append({
            "start": m.start(), "end": m.end(), "attrs": attrs, "body": body,
            "f_start": fm.start(), "f_end": fm.end(),
            "f_attrs": fattrs, "f_text": ftext, "si": si,
            "has_v": bool(CELL_V_RE.search(body)),
            "is_shared": 't="shared"' in fattrs,
        })
    if not cells:
        return xml, 0

    # 共享公式组：只要组里有一条公式引用外链，整组都要转
    bad_si = set()
    for c in cells:
        if c["is_shared"] and c["f_text"] and EXT_TOKEN_RE.search(c["f_text"]) and c["si"]:
            bad_si.add(c["si"])
    for c in cells:
        if c["is_shared"] and c["si"] in bad_si:
            c["hit"] = True
        else:
            c["hit"] = bool(c["f_text"] and EXT_TOKEN_RE.search(c["f_text"]))
    # 组内只要有一个单元格没有缓存值，整组放弃（宁可不转，也不写出坏文件）
    for si in bad_si:
        group = [c for c in cells if c["si"] == si and c["is_shared"]]
        if any(not c["has_v"] for c in group):
            for c in group:
                c["hit"] = False

    spans = []
    for c in cells:
        if not c["hit"] or not c["has_v"]:
            continue
        body = c["body"][: c["f_start"]] + c["body"][c["f_end"]:]
        attrs = c["attrs"]
        if re.search(r'\bt\s*=\s*"str"', attrs):
            vm = CELL_V_RE.search(body)
            if vm:
                text = vm.group(1)
                space = (' xml:space="preserve"'
                         if text[:1].isspace() or text[-1:].isspace() else "")
                body = (body[: vm.start()]
                        + "<is><t%s>%s</t></is>" % (space, text)
                        + body[vm.end():])
                attrs = _retag(attrs, "inlineStr")
        spans.append((c["start"], c["end"], "<c%s>%s</c>" % (attrs, body)))

    if not spans:
        return xml, 0
    out = []
    prev = 0
    for start, end, rep in spans:
        out.append(xml[prev:start])
        out.append(rep)
        prev = end
    out.append(xml[prev:])
    return "".join(out), len(spans)
