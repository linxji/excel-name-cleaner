# -*- coding: utf-8 -*-
# Author: Opera-Jin
"""极简 OLE2 复合文档（Compound File Binary）读写。

.xls / 老版 .et 都是 CFB 容器，里面装着名为 "Workbook" 的 BIFF 流。
要删掉名称就必须改写这个流，而流长度一变，就必须重建整个容器。

策略：完整保留原目录树（左右兄弟、子节点指针、颜色、CLSID、时间戳），
只更新各数据流的起始扇区与长度，因此对 Excel/WPS 而言文件结构与原来一致。
"""
from __future__ import annotations

import struct
from dataclasses import dataclass, field
from typing import Dict, List, Optional

SIGNATURE = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"

FREESECT = 0xFFFFFFFF
ENDOFCHAIN = 0xFFFFFFFE
FATSECT = 0xFFFFFFFD
DIFSECT = 0xFFFFFFFC

TYPE_EMPTY = 0
TYPE_STORAGE = 1
TYPE_STREAM = 2
TYPE_ROOT = 5


class CfbError(Exception):
    pass


@dataclass
class Entry:
    index: int
    raw: bytes
    name: str
    obj_type: int
    color: int
    left: int
    right: int
    child: int
    clsid: bytes
    state: int
    ctime: bytes
    mtime: bytes
    start: int
    size: int
    data: bytes = b""

    def rebuild(self, start: int, size: int) -> bytes:
        buf = bytearray(self.raw)
        struct.pack_into("<I", buf, 116, start & 0xFFFFFFFF)
        struct.pack_into("<Q", buf, 120, size & 0xFFFFFFFFFFFFFFFF)
        return bytes(buf)


@dataclass
class CompoundFile:
    major: int = 3
    sector_size: int = 512
    mini_size: int = 64
    mini_cutoff: int = 4096
    entries: List[Entry] = field(default_factory=list)
    root_index: int = 0

    # ---------------------------------------------------------------- 读

    @classmethod
    def load(cls, path: str) -> "CompoundFile":
        with open(path, "rb") as fh:
            blob = fh.read()
        if len(blob) < 512 or blob[:8] != SIGNATURE:
            raise CfbError("不是有效的 OLE 复合文档")
        cf = cls()
        cf.major = struct.unpack_from("<H", blob, 26)[0]
        shift = struct.unpack_from("<H", blob, 30)[0]
        mini_shift = struct.unpack_from("<H", blob, 32)[0]
        cf.sector_size = 1 << shift
        cf.mini_size = 1 << mini_shift
        num_fat = struct.unpack_from("<I", blob, 44)[0]
        dir_start = struct.unpack_from("<I", blob, 48)[0]
        cf.mini_cutoff = struct.unpack_from("<I", blob, 56)[0]
        minifat_start = struct.unpack_from("<I", blob, 60)[0]
        num_minifat = struct.unpack_from("<I", blob, 64)[0]
        difat_start = struct.unpack_from("<I", blob, 68)[0]
        num_difat = struct.unpack_from("<I", blob, 72)[0]

        def sector(n: int) -> bytes:
            off = 512 + n * cf.sector_size
            return blob[off: off + cf.sector_size]

        # DIFAT
        difat: List[int] = list(struct.unpack_from("<109I", blob, 76))
        nxt = difat_start
        guard = 0
        while nxt not in (FREESECT, ENDOFCHAIN) and guard < num_difat + 8:
            data = sector(nxt)
            vals = struct.unpack_from("<%dI" % (cf.sector_size // 4), data, 0)
            difat.extend(vals[:-1])
            nxt = vals[-1]
            guard += 1
        fat_sectors = [s for s in difat if s not in (FREESECT, ENDOFCHAIN)][:num_fat]

        fat: List[int] = []
        for s in fat_sectors:
            data = sector(s)
            fat.extend(struct.unpack_from("<%dI" % (cf.sector_size // 4), data, 0))

        def chain(start: int, table: List[int], limit: int = 1 << 22) -> List[int]:
            out = []
            cur = start
            while cur not in (ENDOFCHAIN, FREESECT) and cur < len(table) and len(out) < limit:
                out.append(cur)
                cur = table[cur]
            return out

        # 目录
        dir_bytes = b"".join(sector(s) for s in chain(dir_start, fat))
        entries: List[Entry] = []
        for i in range(0, len(dir_bytes) // 128):
            raw = dir_bytes[i * 128: (i + 1) * 128]
            name_len = struct.unpack_from("<H", raw, 64)[0]
            if name_len >= 2:
                name = raw[: name_len - 2].decode("utf-16-le", "replace")
            else:
                name = ""
            entries.append(
                Entry(
                    index=i,
                    raw=raw,
                    name=name,
                    obj_type=raw[66],
                    color=raw[67],
                    left=struct.unpack_from("<I", raw, 68)[0],
                    right=struct.unpack_from("<I", raw, 72)[0],
                    child=struct.unpack_from("<I", raw, 76)[0],
                    clsid=raw[80:96],
                    state=struct.unpack_from("<I", raw, 96)[0],
                    ctime=raw[100:108],
                    mtime=raw[108:116],
                    start=struct.unpack_from("<I", raw, 116)[0],
                    size=struct.unpack_from("<Q", raw, 120)[0],
                )
            )
        cf.entries = entries
        cf.root_index = 0
        root = entries[0]
        if root.obj_type != TYPE_ROOT:
            raise CfbError("根目录项异常")

        # 迷你流
        mini_stream = b"".join(sector(s) for s in chain(root.start, fat))
        minifat: List[int] = []
        if num_minifat and minifat_start not in (FREESECT, ENDOFCHAIN):
            mf = b"".join(sector(s) for s in chain(minifat_start, fat))
            minifat = list(struct.unpack_from("<%dI" % (len(mf) // 4), mf, 0))

        for e in entries:
            if e.obj_type != TYPE_STREAM or e.index == 0:
                continue
            if e.size == 0:
                e.data = b""
            elif e.size < cf.mini_cutoff:
                parts = []
                cur = e.start
                while cur not in (ENDOFCHAIN, FREESECT) and cur < len(minifat):
                    off = cur * cf.mini_size
                    parts.append(mini_stream[off: off + cf.mini_size])
                    cur = minifat[cur]
                e.data = b"".join(parts)[: e.size]
            else:
                e.data = b"".join(sector(s) for s in chain(e.start, fat))[: e.size]
        return cf

    # ---------------------------------------------------------------- 查

    def stream(self, *names: str) -> Optional[Entry]:
        wanted = {n.lower() for n in names}
        for e in self.entries:
            if e.obj_type == TYPE_STREAM and e.name.lower() in wanted:
                return e
        return None

    def set_stream(self, entry: Entry, data: bytes):
        entry.data = data

    # ---------------------------------------------------------------- 写

    def save(self, path: str):
        sec = 512
        mini = 64
        cutoff = self.mini_cutoff or 4096

        entries = [e for e in self.entries]
        # 补齐目录项到 4 的倍数
        while len(entries) % 4 != 0:
            entries.append(
                Entry(
                    index=len(entries),
                    raw=b"\x00" * 128,
                    name="",
                    obj_type=TYPE_EMPTY,
                    color=0,
                    left=FREESECT,
                    right=FREESECT,
                    child=FREESECT,
                    clsid=b"\x00" * 16,
                    state=0,
                    ctime=b"\x00" * 8,
                    mtime=b"\x00" * 8,
                    start=ENDOFCHAIN,
                    size=0,
                )
            )

        # ---- 迷你流
        mini_blob = bytearray()
        minifat: List[int] = []
        for e in entries:
            if e.obj_type != TYPE_STREAM or e.index == 0 or e.size >= cutoff:
                continue
            if e.size == 0:
                e.start, e.size = ENDOFCHAIN, 0
                continue
            n = (e.size + mini - 1) // mini
            base = len(mini_blob) // mini
            for k in range(n):
                minifat.append(base + k + 1 if k < n - 1 else ENDOFCHAIN)
            start = len(mini_blob) // mini
            chunk = e.data[: e.size]
            chunk += b"\x00" * (n * mini - len(chunk))
            mini_blob.extend(chunk)
            e.start, e.size = start, len(e.data[: e.size])
        # 迷你 FAT 补齐到整扇区
        entries_per_sector = sec // 4
        while len(minifat) % entries_per_sector:
            minifat.append(FREESECT)

        # ---- 扇区分配
        sectors: List[bytes] = []
        fat: List[int] = []

        def alloc_chain(data: bytes) -> int:
            if not data:
                return ENDOFCHAIN
            n = (len(data) + sec - 1) // sec
            start = len(sectors)
            padded = data + b"\x00" * (n * sec - len(data))
            for k in range(n):
                sectors.append(padded[k * sec: (k + 1) * sec])
                fat.append(start + k + 1 if k < n - 1 else ENDOFCHAIN)
            return start

        for e in entries:
            if e.obj_type == TYPE_STREAM and e.index != 0 and e.size >= cutoff:
                e.start = alloc_chain(e.data[: e.size])
                e.size = len(e.data[: e.size])

        root = entries[0]
        root.start = alloc_chain(bytes(mini_blob))
        root.size = len(mini_blob)

        minifat_start = ENDOFCHAIN
        if minifat:
            mf_bytes = struct.pack("<%dI" % len(minifat), *minifat)
            minifat_start = alloc_chain(mf_bytes)

        # ---- 目录
        dir_blob = b"".join(e.rebuild(e.start, e.size) for e in entries)
        dir_start = alloc_chain(dir_blob)

        # ---- FAT 扇区数量（迭代求解）
        base = len(sectors)
        n_fat = 1
        for _ in range(64):
            n_difat = 0 if n_fat <= 109 else (n_fat - 109 + 126) // 127
            total = base + n_fat + n_difat
            need = (total + entries_per_sector - 1) // entries_per_sector
            if need <= n_fat:
                break
            n_fat = need
        n_difat = 0 if n_fat <= 109 else (n_fat - 109 + 126) // 127

        fat_sector_ids = []
        for _ in range(n_fat):
            fat_sector_ids.append(len(sectors))
            sectors.append(b"\x00" * sec)
            fat.append(FATSECT)
        difat_sector_ids = []
        for _ in range(n_difat):
            difat_sector_ids.append(len(sectors))
            sectors.append(b"\x00" * sec)
            fat.append(DIFSECT)

        while len(fat) < len(sectors):
            fat.append(FREESECT)
        while len(fat) < n_fat * entries_per_sector:
            fat.append(FREESECT)

        # ---- 写 FAT
        for idx, sid in enumerate(fat_sector_ids):
            chunk = fat[idx * entries_per_sector: (idx + 1) * entries_per_sector]
            sectors[sid] = struct.pack("<%dI" % entries_per_sector, *chunk)

        # ---- 写 DIFAT
        remaining = fat_sector_ids[109:]
        difat_first = difat_sector_ids[0] if difat_sector_ids else ENDOFCHAIN
        for i, sid in enumerate(difat_sector_ids):
            part = remaining[i * 127: (i + 1) * 127]
            vals = list(part) + [FREESECT] * (127 - len(part))
            nxt = difat_sector_ids[i + 1] if i + 1 < len(difat_sector_ids) else ENDOFCHAIN
            vals.append(nxt)
            sectors[sid] = struct.pack("<128I", *vals)

        # ---- 头部
        header = bytearray(512)
        header[0:8] = SIGNATURE
        struct.pack_into("<H", header, 24, 0x003E)
        struct.pack_into("<H", header, 26, 3)
        struct.pack_into("<H", header, 28, 0xFFFE)
        struct.pack_into("<H", header, 30, 9)
        struct.pack_into("<H", header, 32, 6)
        struct.pack_into("<I", header, 40, 0)
        struct.pack_into("<I", header, 44, n_fat)
        struct.pack_into("<I", header, 48, dir_start)
        struct.pack_into("<I", header, 52, 0)
        struct.pack_into("<I", header, 56, cutoff)
        struct.pack_into("<I", header, 60, minifat_start)
        struct.pack_into("<I", header, 64, len(minifat) // entries_per_sector if minifat else 0)
        struct.pack_into("<I", header, 68, difat_first)
        struct.pack_into("<I", header, 72, n_difat)
        first = fat_sector_ids[:109] + [FREESECT] * (109 - min(109, len(fat_sector_ids)))
        struct.pack_into("<109I", header, 76, *first)

        with open(path, "wb") as fh:
            fh.write(bytes(header))
            for s in sectors:
                fh.write(s)
