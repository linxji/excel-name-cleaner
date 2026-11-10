# -*- coding: utf-8 -*-
# Author: Opera-Jin
"""历史日志：按天追加，UTF-8(BOM) 便于记事本直接打开。"""
from __future__ import annotations

import os
import time


def log_dir() -> str:
    base = os.environ.get("LOCALAPPDATA") or os.path.expanduser("~")
    path = os.path.join(base, "Excel名称清理工具", "logs")
    try:
        os.makedirs(path, exist_ok=True)
    except OSError:
        path = os.path.expanduser("~")
    return path


class RunLog:
    def __init__(self):
        self.dir = log_dir()
        self.path = os.path.join(self.dir, "clean-%s.log" % time.strftime("%Y%m%d"))
        self.session = []

    def _write(self, line: str):
        self.session.append(line)
        new = not os.path.exists(self.path)
        try:
            with open(self.path, "w" if new else "a",
                      encoding="utf-8-sig" if new else "utf-8") as fh:
                fh.write(line + "\n")
        except OSError:
            pass

    def header(self, text: str):
        self._write("")
        self._write("=" * 72)
        self._write("%s  %s" % (time.strftime("%Y-%m-%d %H:%M:%S"), text))
        self._write("=" * 72)

    def info(self, text: str):
        self._write("  " + text)

    def raw(self, text: str = ""):
        self._write(text)

    def outcome(self, r, opt=None):
        from .util import human_size

        self._write("")
        self._write("[%s] %s" % ("成功" if r.ok else "失败", r.path))
        if r.ts:
            self._write("  操作时间：%s   耗时：%.1f 秒" % (r.ts, r.seconds))
        if opt is not None:
            self._write("  选项：保留打印区域=%s  删除外链=%s  备份=%s"
                        % (opt.keep_print_names, opt.remove_links, opt.backup))
        self._write("  格式：%s   耗时：%.2f 秒" % (r.fmt or "?", r.seconds))
        self._write("  体积：%s → %s" % (human_size(r.size_before), human_size(r.size_after)))
        self._write("  名称：共 %d 个，删除 %d 个，保留 %d 个"
                    % (r.names_total, r.names_removed, r.names_kept))
        if r.links_removed:
            self._write("  外部链接：删除 %d 个，其中 %d 个引用公式已转为当前数值"
                        % (r.links_removed, r.converted_formulas))
        if r.output_path:
            self._write("  输出：%s（原 .xlsb 未改动）" % r.output_path)
        if r.message:
            self._write("  说明：%s" % r.message)
        if r.kept_names:
            self._write("  保留清单：")
            for name, reason in r.kept_names[:40]:
                self._write("    - %s（%s）" % (name, reason))
            if r.names_kept > 40:
                self._write("    ……共 %d 条，其余略" % r.names_kept)
        if r.removed_names:
            self._write("  删除样例（前 50 个）：")
            self._write("    " + "，".join(r.removed_names[:50]))
