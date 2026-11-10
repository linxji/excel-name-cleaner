# -*- coding: utf-8 -*-
# Author: Opera-Jin
"""格式识别与统一调度：把三种引擎收敛成同一个结果对象。"""
from __future__ import annotations

import os
import re
import time
import zipfile
from dataclasses import dataclass, field
from typing import Callable, List, Optional, Tuple

from . import biff, cfb, ooxml
from .util import check_writable

FMT_LABEL = {
    "xlsx": "Excel 工作簿",
    "xlsm": "Excel 启用宏工作簿",
    "xltx": "Excel 模板",
    "xltm": "Excel 启用宏模板",
    "xls": "Excel 97-2003",
    "et": "WPS 表格",
    "xlsb": "Excel 二进制工作簿",
}


@dataclass
class Options:
    keep_print_names: bool = True     # 保留打印区域 / 打印标题 / 筛选范围名称
    remove_links: bool = True         # 删除外部链接
    backup: bool = True               # 处理前备份
    recurse: bool = True              # 文件夹递归


@dataclass
class Outcome:
    path: str
    ok: bool = False
    skipped: bool = False
    fmt: str = ""
    message: str = ""
    names_total: int = 0
    names_removed: int = 0
    names_kept: int = 0
    links_removed: int = 0
    converted_formulas: int = 0
    ext_formula_refs: int = 0
    size_before: int = 0
    size_after: int = 0
    output_path: str = ""
    ts: str = ""
    seconds: float = 0.0
    removed_names: List[str] = field(default_factory=list)
    kept_names: List[Tuple[str, str]] = field(default_factory=list)
    backup_path: str = ""

    @property
    def saved(self) -> int:
        return max(0, self.size_before - self.size_after)


SUPPORTED_EXT = {".xlsx", ".xlsm", ".xltx", ".xltm", ".xls", ".xlsb", ".et", ".ett"}


def detect(path: str) -> str:
    """返回 'ooxml' / 'biff' / 'encrypted' / 'xlsb' / 'unknown'。"""
    try:
        with open(path, "rb") as fh:
            head = fh.read(8)
    except OSError:
        return "unknown"
    if head[:4] == b"PK\x03\x04":
        try:
            with zipfile.ZipFile(path) as zf:
                names = set(zf.namelist())
            if "xl/workbook.bin" in names:
                return "xlsb"
            if "xl/workbook.xml" in names:
                return "ooxml"
        except zipfile.BadZipFile:
            return "unknown"
        return "unknown"
    if head[:8] == cfb.SIGNATURE:
        try:
            cf = cfb.CompoundFile.load(path)
        except Exception:  # noqa: BLE001
            return "unknown"
        if cf.stream("EncryptedPackage") is not None:
            return "encrypted"
        if cf.stream("Workbook", "Book") is not None:
            return "biff"
        return "unknown"
    return "unknown"


def count_names(path: str):
    """快速统计文件里现有的定义名称数量（不做完整分析）。

    返回 int；无法统计时返回 None（例如 .xlsb 需要先转换）。
    """
    kind = detect(path)
    try:
        if kind == "ooxml":
            with zipfile.ZipFile(path) as zf:
                wb = zf.read("xl/workbook.xml")
            # 注意不能直接 count(b"<definedName")，会把 <definedNames> 也算进去
            return len(re.findall(rb"<definedName[ \t\r\n/>]", wb))
        if kind == "biff":
            from . import biff
            _cf, entry = biff._load_workbook(path)
            buf = entry.data
            first = biff._first_sheet_offset(buf)
            n = 0
            for off, rtype, rlen in biff._records(buf):
                if off >= first:
                    break
                if rtype == 0x0018:
                    n += 1
            return n
    except Exception:  # noqa: BLE001
        return None
    return None


def collect_files(paths: List[str], recurse: bool = True) -> List[str]:
    """展开文件夹，返回受支持的 Excel 文件列表。"""
    out: List[str] = []
    seen = set()
    for p in paths:
        p = os.path.abspath(p)
        if os.path.isdir(p):
            if recurse:
                for root, dirs, files in os.walk(p):
                    dirs[:] = [d for d in dirs if not d.startswith(".")]
                    for f in files:
                        if os.path.splitext(f)[1].lower() in SUPPORTED_EXT:
                            fp = os.path.join(root, f)
                            if fp.lower() not in seen:
                                seen.add(fp.lower())
                                out.append(fp)
            else:
                for f in os.listdir(p):
                    fp = os.path.join(p, f)
                    if os.path.isfile(fp) and os.path.splitext(f)[1].lower() in SUPPORTED_EXT:
                        if fp.lower() not in seen:
                            seen.add(fp.lower())
                            out.append(fp)
        elif os.path.isfile(p):
            if p.lower() not in seen:
                seen.add(p.lower())
                out.append(p)
    out.sort()
    return out


def process_file(path: str, opt: Options, log: Optional[Callable[[str], None]] = None,
                 progress: Optional[Callable[[str, int, int], None]] = None) -> Outcome:
    res = Outcome(path=path, fmt=os.path.splitext(path)[1].lower().lstrip("."))
    t0 = time.time()
    if not os.path.exists(path):
        res.message = "文件不存在"
        return res
    res.size_before = os.path.getsize(path)
    res.size_after = res.size_before

    reason = check_writable(path)
    if reason == "readonly":
        res.message = "文件是只读的，请右键属性里取消「只读」后再处理"
        res.skipped = True
        return res
    if reason == "locked":
        res.message = "文件正被占用（可能已在 Excel / WPS 中打开），请先关闭再处理"
        res.skipped = True
        return res

    kind = detect(path)
    if kind == "encrypted":
        res.message = "文件已加密，无法处理（请先取消密码）"
        return res
    if kind == "xlsb":
        out = _process_xlsb(path, opt, log, progress, res)
        out.seconds = time.time() - t0
        return out
    if kind == "unknown":
        res.message = "无法识别文件格式"
        return res

    try:
        if kind == "ooxml":
            r = ooxml.process(
                path,
                remove_print_names=not opt.keep_print_names,
                remove_links=opt.remove_links,
                do_backup=opt.backup,
                log=log,
                progress=progress,
            )
            res.backup_path = r.backup_path
            res.names_total = r.names_total
            res.names_removed = r.names_removed
            res.names_kept = r.names_kept
            res.links_removed = r.links_removed
            res.converted_formulas = r.converted_formulas
            res.ext_formula_refs = r.ext_formula_refs
            res.size_after = r.size_after
            res.removed_names = r.removed_names
            res.kept_names = r.kept_names
        else:
            r = biff.process(
                path,
                remove_print_names=not opt.keep_print_names,
                do_backup=opt.backup,
                log=log,
                progress=progress,
            )
            res.backup_path = r.backup_path
            res.names_total = r.names_total
            res.names_removed = r.names_removed
            res.names_kept = r.names_kept
            res.size_after = r.size_after
            res.removed_names = r.removed_names
            res.kept_names = r.kept_names
            if opt.remove_links and r.has_external:
                res.message = "旧版格式暂不支持断开外部链接，外链已保留"
        res.ok = r.ok
        res.message = r.message if not res.message else (
            r.message + "；" + res.message if r.ok else r.message
        )
    except Exception as exc:  # noqa: BLE001
        res.ok = False
        res.message = "处理异常：%s" % exc
    res.seconds = time.time() - t0
    return res


def _unique_xlsx(path: str) -> str:
    base = os.path.splitext(path)[0]
    cand = base + ".xlsx"
    idx = 1
    while os.path.exists(cand):
        cand = "%s(%d).xlsx" % (base, idx)
        idx += 1
    return cand


def _process_xlsb(path, opt, log, progress, res) -> Outcome:
    """先转成 .xlsx 再清理，最后直接留下 .xlsx（原 .xlsb 不动）。"""
    from . import xlsb

    dest = _unique_xlsx(path)
    tmp = dest + ".part"

    def rep(stage, done=0, total=0):
        if progress:
            try:
                progress(stage, done, total)
            except Exception:  # noqa: BLE001
                pass

    err = xlsb.convert(path, tmp, progress)
    if err:
        if os.path.exists(tmp):
            try:
                os.remove(tmp)
            except OSError:
                pass
        res.message = err
        res.skipped = True
        return res

    try:
        r = ooxml.process(
            tmp,
            remove_print_names=not opt.keep_print_names,
            remove_links=opt.remove_links,
            do_backup=False,
            log=log,
            progress=progress,
        )
        res.names_total = r.names_total
        res.names_removed = r.names_removed
        res.names_kept = r.names_kept
        res.links_removed = r.links_removed
        res.converted_formulas = r.converted_formulas
        res.ext_formula_refs = r.ext_formula_refs
        res.removed_names = r.removed_names
        res.kept_names = r.kept_names
        if not r.ok:
            res.message = "已转为 xlsx，但清理失败：%s" % r.message
            return res
        rep("保存 xlsx", 0, 0)
        os.replace(tmp, dest)
        res.output_path = dest
        res.size_after = os.path.getsize(dest)
        res.ok = True
        res.message = "已另存为 %s" % os.path.basename(dest)
    except Exception as exc:  # noqa: BLE001
        res.message = "转换后处理失败：%s" % exc
    finally:
        if not res.output_path and os.path.exists(tmp):
            try:
                os.remove(tmp)
            except OSError:
                pass
    return res
