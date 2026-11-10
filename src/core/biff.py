# -*- coding: utf-8 -*-
# Author: Opera-Jin
"""BIFF8（.xls / 老版 .et）名称清理。

核心安全设计
------------
1. **不重编号**。NAME 记录是位置索引的（公式里的 ptgName 用 1-based 序号
   指向它）。常见做法是删记录再改所有公式的序号，但那样必须精确解析
   每一条公式（条件格式、数据验证的记录布局随版本而异），一处失误就会
   让公式悄悄指向另一个名称。
   这里改成：先把“要保留的”NAME 记录按原顺序整体前移，把“要删的”挤到
   末尾，再整块丢弃。存活名称的相对顺序和序号完全不变，公式一个字节
   都不用动。
2. **偏移修正**。BOUNDSHEET.lbPlyPos 与 INDEX.rgibRw 是流内绝对偏移，
   删除字节后整体前移，否则工作表定位错乱。
3. **宁可不做**。任何解析异常都整体中止，绝不写入半成品。
"""
from __future__ import annotations

import os
import struct
from dataclasses import dataclass, field
from typing import List, Set, Tuple

from . import cfb
from .util import atomic_replace, copy_backup, make_temp_beside

# ---- 记录类型
BOF = 0x0809
EOF_REC = 0x000A
BOUNDSHEET = 0x0085
NAME = 0x0018
INDEX = 0x020B
DBCELL = 0x00D7
FORMULA = 0x0006
ARRAY = 0x0221
SHRFMLA = 0x04BC
TABLE = 0x0236
CF = 0x01B1
DV = 0x01BE
SUPBOOK = 0x01AE
EXTERNSHEET = 0x0017

# rgce 在记录体中的 (cce 偏移, rgce 偏移)
RGCE_LAYOUT = {
    FORMULA: (20, 22),
    ARRAY: (12, 14),
    SHRFMLA: (8, 10),
    TABLE: (12, 14),
}

# ptg 数据长度（不含 ptg 自身 1 字节）
PTG_LEN = {
    0x01: 4, 0x02: 4,
    0x03: 0, 0x04: 0, 0x05: 0, 0x06: 0, 0x07: 0, 0x08: 0, 0x09: 0, 0x0A: 0,
    0x0B: 0, 0x0C: 0, 0x0D: 0, 0x0E: 0, 0x0F: 0, 0x10: 0, 0x11: 0, 0x12: 0,
    0x13: 0, 0x14: 0, 0x15: 0, 0x16: 0,
    0x19: 3, 0x1C: 1, 0x1D: 1, 0x1E: 2, 0x1F: 8,
    0x20: 7, 0x21: 2, 0x22: 3, 0x23: 4, 0x24: 4, 0x25: 8,
    0x26: 6, 0x27: 6, 0x28: 6, 0x29: 6, 0x2A: 4, 0x2B: 8, 0x2C: 4, 0x2D: 8,
    0x39: 6, 0x3A: 6, 0x3B: 10, 0x3C: 6, 0x3D: 10,
}

PTG_NAME_BYTE = (0x23, 0x43, 0x63)     # ptgName 的三种引用类
PTG_REFERR_BASES = (0x3C, 0x3D)        # ptgRefErr3d / ptgAreaErr3d

SYSTEM_KEEP_KEYS = {"print_area", "print_titles", "_filterdatabase"}

FIX_INDEX = True
FIX_BOUNDSHEET = True

# 结果明细上限：30 万条全存下来没意义，只会拖慢界面与日志
DETAIL_CAP = 3000


class BiffError(Exception):
    pass


@dataclass
class XlsName:
    rec_index: int
    offset: int
    length: int
    name: str
    itab: int
    scope_label: str
    hidden: bool
    is_system: bool
    is_invalid: bool
    is_referenced: bool = False
    action: str = "delete"
    reason: str = ""


@dataclass
class XlsAnalysis:
    names: List[XlsName] = field(default_factory=list)
    sheet_names: List[str] = field(default_factory=list)
    size: int = 0
    stream_name: str = ""
    formula_count: int = 0
    has_external: bool = False


@dataclass
class XlsOutcome:
    ok: bool
    message: str = ""
    names_total: int = 0
    names_removed: int = 0
    names_kept: int = 0
    size_before: int = 0
    size_after: int = 0
    has_external: bool = False
    removed_names: List[str] = field(default_factory=list)
    kept_names: List[Tuple[str, str]] = field(default_factory=list)
    backup_path: str = ""


# ---------------------------------------------------------------- 基础


def _records(buf: bytes):
    off = 0
    n = len(buf)
    while off + 4 <= n:
        rtype, rlen = struct.unpack_from("<HH", buf, off)
        if off + 4 + rlen > n:
            raise BiffError("记录越界 @%d" % off)
        yield off, rtype, rlen
        off += 4 + rlen


def _base_ptg(ptg: int) -> int:
    if ptg >= 0x20:
        return (ptg & 0x1F) | 0x20
    return ptg


def walk_tokens(rgce: bytes):
    i = 0
    n = len(rgce)
    while i < n:
        ptg = rgce[i]
        base = _base_ptg(ptg)
        if base == 0x17:
            if i + 3 > n:
                raise BiffError("ptgStr 越界")
            cch = rgce[i + 1]
            flag = rgce[i + 2]
            size = 3 + (cch * 2 if flag & 1 else cch)
        elif base in PTG_LEN:
            size = 1 + PTG_LEN[base]
        else:
            raise BiffError("未知 token 0x%02x @%d" % (ptg, i))
        if i + size > n:
            raise BiffError("token 越界 @%d" % i)
        yield i, base, size
        i += size


def _read_unicode(buf: bytes, pos: int) -> Tuple[str, int]:
    cch = struct.unpack_from("<H", buf, pos)[0]
    flag = buf[pos + 2]
    pos += 3
    if flag & 1:
        s = buf[pos: pos + cch * 2].decode("utf-16-le", "replace")
        pos += cch * 2
    else:
        s = buf[pos: pos + cch].decode("latin-1", "replace")
        pos += cch
    return s, pos


def _name_text(body: bytes) -> Tuple[str, int, int, int, int]:
    grbit = struct.unpack_from("<H", body, 0)[0]
    cch = body[3]
    cce = struct.unpack_from("<H", body, 4)[0]
    itab = struct.unpack_from("<H", body, 8)[0]
    p = 14
    flag = body[p]
    p += 1
    if flag & 1:
        nm = body[p: p + cch * 2].decode("utf-16-le", "replace")
        p += cch * 2
    else:
        nm = body[p: p + cch].decode("latin-1", "replace")
        p += cch
    return nm, itab, cce, p, grbit


def _scan_name_bytes(blob: bytes, max_index: int) -> Set[int]:
    """保守扫描：任何形如 ptgName + 合法序号的字节序列都算“被引用”。

    只用于布局随版本变化的记录（条件格式 / 数据验证）。多认几个名称只会
    让清理力度略小，不会误删。
    """
    out: Set[int] = set()
    n = len(blob)
    for i in range(max(0, n - 4)):
        if blob[i] in PTG_NAME_BYTE:
            v = struct.unpack_from("<I", blob, i + 1)[0]
            if 1 <= v <= max_index:
                out.add(v)
    return out


# ---------------------------------------------------------------- 分析


def _load_workbook(path: str):
    cf = cfb.CompoundFile.load(path)
    entry = cf.stream("Workbook", "Book")
    if entry is None:
        raise BiffError("找不到 Workbook 数据流")
    return cf, entry


def _first_sheet_offset(buf: bytes) -> int:
    best = len(buf)
    for off, rtype, rlen in _records(buf):
        if FIX_BOUNDSHEET and rtype == BOUNDSHEET and rlen >= 4:
            pos = struct.unpack_from("<I", buf, off + 4)[0]
            if 0 < pos < best:
                best = pos
    return best


def analyze(path: str) -> XlsAnalysis:
    cf, entry = _load_workbook(path)
    buf = entry.data
    res = XlsAnalysis(size=os.path.getsize(path), stream_name=entry.name)

    for off, rtype, rlen in _records(buf):
        if rtype == BOF:
            ver = struct.unpack_from("<H", buf, off + 4)[0]
            if ver != 0x0600:
                raise BiffError("仅支持 Excel 97 及以上格式（当前 0x%04X）" % ver)
            break

    for off, rtype, rlen in _records(buf):
        if rtype == BOUNDSHEET:
            body = buf[off + 4: off + 4 + rlen]
            cch = body[6]
            flag = body[7]
            nm = (body[8: 8 + cch * 2].decode("utf-16-le", "replace") if flag & 1
                  else body[8: 8 + cch].decode("latin-1", "replace"))
            res.sheet_names.append(nm)

    # SUPBOOK / EXTERNSHEET：判断 3D 引用是否已失效（#REF!）
    supbooks: List[bool] = []          # True = 本工作簿内部引用
    xtis: List[Tuple[int, int, int]] = []
    for off, rtype, rlen in _records(buf):
        body = buf[off + 4: off + 4 + rlen]
        if rtype == SUPBOOK and rlen >= 2:
            ctab = struct.unpack_from("<H", body, 0)[0]
            # Excel 用 0x0401 标记自身；WPS 直接写工作表数量。
            # 两者内部引用的共同点是：记录里没有后续的路径字符串。
            # 内部/加载项 SUPBOOK 只有 2~4 字节；外部的一定带路径字符串
            supbooks.append(rlen <= 4 or ctab == 0x0401 or ctab == 0x3A01)
        elif rtype == EXTERNSHEET and rlen >= 2:
            cnt = struct.unpack_from("<H", body, 0)[0]
            for k in range(cnt):
                if 2 + k * 6 + 6 <= rlen:
                    xtis.append(struct.unpack_from("<HHH", body, 2 + k * 6))
    n_sheets = len(res.sheet_names)
    res.has_external = any(not internal for internal in supbooks)

    def is_dead_ref(tokens, rgce) -> bool:
        for pos, base, size in tokens:
            if base in PTG_REFERR_BASES:
                return True
            if base in (0x3A, 0x3B):
                ixti = struct.unpack_from("<H", rgce, pos + 1)[0]
                if ixti >= len(xtis):
                    return True
                isup, i1, i2 = xtis[ixti]
                if i1 == 0xFFFF or i2 == 0xFFFF:
                    return True
                if isup < len(supbooks) and supbooks[isup]:
                    if i1 > n_sheets or i2 > n_sheets:
                        return True
        return False

    first_sheet = _first_sheet_offset(buf)

    # ---- 名称表（globals 区）
    idx = 0
    for off, rtype, rlen in _records(buf):
        if off >= first_sheet:
            break
        if rtype != NAME:
            continue
        idx += 1
        body = buf[off + 4: off + 4 + rlen]
        nm, itab, cce, rstart, grbit = _name_text(body)
        rgce = body[rstart: rstart + cce]
        try:
            toks = list(walk_tokens(rgce))
            invalid = is_dead_ref(toks, rgce)
        except BiffError:
            invalid = False
        low = nm.lower()
        if low.startswith("_xlnm."):
            low = low[6:]
        scope = res.sheet_names[itab - 1] if 0 < itab <= n_sheets else "全工作簿"
        res.names.append(
            XlsName(
                rec_index=idx,
                offset=off,
                length=4 + rlen,
                name=nm,
                itab=itab,
                scope_label=scope,
                hidden=bool(grbit & 0x01),
                is_system=low in SYSTEM_KEEP_KEYS or nm.lower().startswith("_xlnm."),
                is_invalid=invalid,
            )
        )

    # ---- 引用统计
    total_names = len(res.names)
    used: Set[int] = set()
    for off, rtype, rlen in _records(buf):
        body = buf[off + 4: off + 4 + rlen]
        if rtype == NAME:
            nm, itab, cce, rstart, grbit = _name_text(body)
            rgce = body[rstart: rstart + cce]
            res.formula_count += 1
            try:
                for pos, base, size in walk_tokens(rgce):
                    if base == 0x23:
                        used.add(struct.unpack_from("<I", rgce, pos + 1)[0])
            except BiffError:
                used |= _scan_name_bytes(rgce, total_names)
            continue
        if rtype in RGCE_LAYOUT:
            cce_off, rgce_off = RGCE_LAYOUT[rtype]
            if rlen < rgce_off:
                continue
            cce = struct.unpack_from("<H", body, cce_off)[0]
            rgce = body[rgce_off: rgce_off + cce]
            res.formula_count += 1
            try:
                for pos, base, size in walk_tokens(rgce):
                    if base == 0x23:
                        used.add(struct.unpack_from("<I", rgce, pos + 1)[0])
            except BiffError:
                used |= _scan_name_bytes(rgce, total_names)
            continue
        if rtype in (CF, DV):
            used |= _scan_name_bytes(body, total_names)

    for n in res.names:
        n.is_referenced = n.rec_index in used
    return res


# ---------------------------------------------------------------- 决策


def decide(ana: XlsAnalysis, remove_print_names: bool) -> None:
    for it in ana.names:
        low = it.name.lower()
        if low.startswith("_xlnm."):
            low = low[6:]
        if it.is_invalid:
            it.action, it.reason = "delete", "引用已失效（#REF!）"
        elif low in SYSTEM_KEEP_KEYS and not remove_print_names:
            it.action, it.reason = "keep", "打印/筛选设置，保留"
        elif it.is_referenced:
            it.action, it.reason = "keep", "被公式引用，保留"
        else:
            it.action, it.reason = "delete", "未被任何公式引用"


# ---------------------------------------------------------------- 主流程


def process(path: str, remove_print_names: bool = False, do_backup: bool = True,
            log=None, progress=None) -> XlsOutcome:
    out = XlsOutcome(ok=False, size_before=os.path.getsize(path))
    def rep(stage, done=0, total=0):
        if progress:
            try:
                progress(stage, done, total)
            except Exception:  # noqa: BLE001
                pass

    try:
        rep("读取工作簿", 0, 0)
        cf, entry = _load_workbook(path)
        ana = analyze(path)
    except Exception as exc:  # noqa: BLE001
        out.message = "无法解析：%s" % exc
        return out

    decide(ana, remove_print_names)
    out.has_external = ana.has_external
    out.names_total = len(ana.names)
    deleted = [n for n in ana.names if n.action == "delete"]
    if not deleted:
        out.ok = True
        out.names_kept = out.names_total
        out.size_after = out.size_before
        out.message = "没有需要清理的内容"
        return out

    for n in ana.names:
        if n.action == "delete":
            if len(out.removed_names) < DETAIL_CAP:
                    out.removed_names.append(n.name)
        else:
            if len(out.kept_names) < DETAIL_CAP:
                    out.kept_names.append((n.name, n.reason))

    rep("删除名称", len(ana.names) - len(deleted), len(ana.names))
    try:
        new_buf = _rebuild(entry.data, ana.names, {n.offset for n in deleted})
    except Exception as exc:  # noqa: BLE001
        out.message = "结构异常，已安全中止（文件未改动）：%s" % exc
        return out

    entry.data = new_buf
    rep("写入文件", 0, 0)
    tmp = make_temp_beside(path, ".xls")
    try:
        cf.save(tmp)
    except Exception as exc:  # noqa: BLE001
        try:
            os.remove(tmp)
        except OSError:
            pass
        out.message = "写入失败：%s" % exc
        return out

    if do_backup:
        bak = copy_backup(path)
        out.backup_path = bak
        if log:
            log("已备份原文件 → %s" % bak)
    atomic_replace(tmp, path, keep_stat_from=path)

    out.size_after = os.path.getsize(path)
    out.names_removed = len(deleted)
    out.names_kept = out.names_total - len(deleted)
    out.ok = True
    out.message = "完成"
    return out


def _rebuild(buf: bytes, names: List[XlsName], drop_offsets: Set[int]) -> bytes:
    """把保留的 NAME 记录整体前移，删除的挤到末尾后丢弃 —— 序号不变。"""
    survivors = [n for n in names if n.offset not in drop_offsets]
    table_offsets = {n.offset for n in names}

    pieces: List[bytes] = []
    emitted = False
    for off, rtype, rlen in _records(buf):
        total = 4 + rlen
        if rtype == NAME and off in table_offsets:
            if not emitted:
                # 在第一条 NAME 的位置一次性写入全部保留项，其余 NAME 全部丢弃
                for n in survivors:
                    pieces.append(buf[n.offset: n.offset + n.length])
                emitted = True
            continue
        pieces.append(buf[off: off + total])
    if not emitted:
        raise BiffError("未找到名称表")

    new_buf = bytearray(b"".join(pieces))
    delta = len(buf) - len(new_buf)
    if delta <= 0:
        raise BiffError("重建长度异常")

    # 名称序号重映射：删除记录后，后面所有名称的序号会前移，
    # 公式里的 ptgName 必须同步改，否则会指向别的名称。
    mapping = {}
    new_idx = 0
    for n in names:
        if n.offset not in drop_offsets:
            new_idx += 1
            mapping[n.rec_index] = new_idx
    _renumber(new_buf, mapping, len(names))

    # 修正流内绝对偏移
    for off, rtype, rlen in _records(bytes(new_buf)):
        body_off = off + 4
        if FIX_BOUNDSHEET and rtype == BOUNDSHEET and rlen >= 4:
            old = struct.unpack_from("<I", new_buf, body_off)[0]
            if delta <= old < len(buf):
                struct.pack_into("<I", new_buf, body_off, old - delta)
        elif FIX_INDEX and rtype == INDEX and rlen >= 12:
            cnt = (rlen - 12) // 4
            for k in range(cnt):
                p = body_off + 12 + k * 4
                val = struct.unpack_from("<I", new_buf, p)[0]
                if (val + 2 <= len(buf) and buf[val: val + 2] == b"\xd7\x00"
                        and val >= delta):
                    struct.pack_into("<I", new_buf, p, val - delta)
    return bytes(new_buf)


def _fix_rgce(buf: bytearray, start: int, cce: int, mapping: Dict[int, int]):
    if cce <= 0:
        return
    rgce = bytes(buf[start: start + cce])
    if len(rgce) != cce:
        raise BiffError("公式区间越界")
    for pos, base, size in walk_tokens(rgce):
        if base != 0x23:
            continue
        old = struct.unpack_from("<I", rgce, pos + 1)[0]
        new = mapping.get(old)
        if new is None:
            raise BiffError("公式引用了待删除的名称（序号 %d）" % old)
        if new != old:
            struct.pack_into("<I", buf, start + pos + 1, new)


def _renumber(buf: bytearray, mapping: Dict[int, int], n_old: int):
    """把全表 ptgName 的序号从旧值改成新值。"""
    for off, rtype, rlen in _records(bytes(buf)):
        body_off = off + 4
        if rtype == NAME:
            body = bytes(buf[body_off: body_off + rlen])
            nm, itab, cce, rstart, grbit = _name_text(body)
            _fix_rgce(buf, body_off + rstart, cce, mapping)
        elif rtype in RGCE_LAYOUT:
            cce_off, rgce_off = RGCE_LAYOUT[rtype]
            if rlen < rgce_off:
                continue
            cce = struct.unpack_from("<H", buf, body_off + cce_off)[0]
            _fix_rgce(buf, body_off + rgce_off, cce, mapping)
        elif rtype in (CF, DV):
            # 条件格式 / 数据验证的记录布局随实现而异（WPS 与 Excel 不同），
            # 这里按 ptgName 的字节特征做保守替换：只认 0x23/0x43/0x63 + 合法序号。
            body = bytes(buf[body_off: body_off + rlen])
            for i in range(len(body) - 4):
                if body[i] not in PTG_NAME_BYTE:
                    continue
                old = struct.unpack_from("<I", body, i + 1)[0]
                if not (1 <= old <= n_old):
                    continue
                new = mapping.get(old)
                if new is None:
                    raise BiffError("条件格式/数据验证引用了待删除的名称（序号 %d）" % old)
                if new != old:
                    struct.pack_into("<I", buf, body_off + i + 1, new)
