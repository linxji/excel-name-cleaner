# -*- coding: utf-8 -*-
# Author: Opera-Jin
"""通用工具：大小格式化、XML 实体、ZIP 重写、原子替换。"""
from __future__ import annotations

import os
import re
import shutil
import tempfile
import time
import zipfile
from xml.sax.saxutils import unescape as _xml_unescape

_EXTRA_ENTITIES = {"&quot;": '"', "&apos;": "'"}


_NUM_ENT_RE = re.compile(r"&#(?:[xX]([0-9a-fA-F]+)|(\d+));")


def xml_unescape(text: str) -> str:
    """还原 XML 实体（属性值与公式文本里常见 &quot; &amp; 等）。"""
    if not text:
        return ""
    return _xml_unescape(text, _EXTRA_ENTITIES)


def _num_entity(m: "re.Match") -> str:
    try:
        code = int(m.group(1), 16) if m.group(1) else int(m.group(2))
        if 0 < code <= 0x10FFFF:
            return chr(code)
    except (TypeError, ValueError):
        pass
    return m.group(0)


def unescape_full(text: str) -> str:
    """同时还原命名实体与数字字符引用（&#20851; / &#x4E2D;）。

    有些导出工具会把中文名称写成数字实体，若不还原，公式里的名称
    就识别不出来，会被误判为“未被引用”而删掉，导致 #NAME?。
    """
    if not text:
        return ""
    if "&" not in text:
        return text
    text = xml_unescape(text)
    return _NUM_ENT_RE.sub(_num_entity, text)


def human_size(num: int) -> str:
    if num is None:
        return "-"
    if num < 1024:
        return "%d B" % num
    for unit in ("KB", "MB", "GB"):
        num /= 1024.0
        if num < 1024 or unit == "GB":
            return "%.1f %s" % (num, unit)
    return "%.1f GB" % num


def short_path(path: str, limit: int = 46) -> str:
    """过长路径中段省略，保留盘符与文件名。"""
    if len(path) <= limit:
        return path
    head, tail = os.path.split(path)
    if len(tail) >= limit - 8:
        return "…" + tail[-(limit - 1):]
    keep = limit - len(tail) - 4
    return head[:max(keep, 3)] + "…" + os.sep + tail


def check_writable(path: str):
    """返回 None 表示可写；否则返回 'readonly' / 'locked'。"""
    if not os.path.exists(path):
        return "missing"
    try:
        fd = os.open(path, os.O_RDWR)
    except (OSError, PermissionError):
        pass
    else:
        os.close(fd)
        return None
    try:
        if not (os.stat(path).st_mode & 0o200):
            return "readonly"
    except OSError:
        return "locked"
    return "locked"


def is_file_locked(path: str) -> bool:
    return check_writable(path) is not None


def read_zip(path: str):
    """返回 (有序条目列表, {名称: 字节})。"""
    entries = []
    data = {}
    with zipfile.ZipFile(path, "r") as zf:
        for info in zf.infolist():
            entries.append(info)
            if not info.is_dir():
                data[info.filename] = zf.read(info.filename)
    return entries, data


def write_zip(dst: str, entries, data, drop_names=(), drop_prefixes=()):
    """按原始顺序与压缩方式重建 ZIP，未修改的条目字节级原样保留。"""
    drop_names = set(drop_names)
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED, allowZip64=True) as zout:
        for info in entries:
            name = info.filename
            if name in drop_names:
                continue
            if drop_prefixes and any(name.startswith(p) for p in drop_prefixes):
                continue
            if name not in data:
                continue
            zi = zipfile.ZipInfo(name, date_time=info.date_time)
            zi.compress_type = info.compress_type
            zi.external_attr = info.external_attr
            zi.internal_attr = info.internal_attr
            zi.create_system = info.create_system
            zout.writestr(zi, data[name])


def make_temp_beside(path: str, suffix: str = ".tmp") -> str:
    folder = os.path.dirname(os.path.abspath(path)) or "."
    fd, tmp = tempfile.mkstemp(prefix="~clean_", suffix=suffix, dir=folder)
    os.close(fd)
    return tmp


def atomic_replace(tmp_path: str, target: str, keep_stat_from: str = None):
    """原子替换目标文件，并尽量保留原文件的时间戳与属性。"""
    src_stat = None
    if keep_stat_from and os.path.exists(keep_stat_from):
        try:
            src_stat = os.stat(keep_stat_from)
        except OSError:
            src_stat = None
    # 目标文件被 Excel/WPS 占用时 os.replace 会抛 PermissionError，
    # 这里绝不能"先删掉再替换"——那会把用户正在编辑的文件删掉。
    os.replace(tmp_path, target)
    if src_stat is not None:
        try:
            os.utime(target, (src_stat.st_atime, src_stat.st_mtime))
        except OSError:
            pass


def backup_path(path: str) -> str:
    """生成不覆盖既有文件的备份路径：原名.bak / 原名.bak2 …"""
    base = path + ".bak"
    if not os.path.exists(base):
        return base
    idx = 2
    while True:
        cand = "%s.bak%d" % (path, idx)
        if not os.path.exists(cand):
            return cand
        idx += 1


def copy_backup(path: str) -> str:
    dst = backup_path(path)
    shutil.copy2(path, dst)
    return dst


_TS_RE = re.compile(r"[^0-9]")


def stamp() -> str:
    return time.strftime("%Y-%m-%d %H:%M:%S")


def now_id() -> str:
    return time.strftime("%Y%m%d-%H%M%S")
