# -*- coding: utf-8 -*-
# Author: Opera-Jin
"""运行历史：JSONL 存储，供界面随时回看。"""
from __future__ import annotations

import json
import os
import time

MAX_RECORDS = 5000


def _base_dir() -> str:
    base = os.environ.get("LOCALAPPDATA") or os.path.expanduser("~")
    path = os.path.join(base, "Excel名称清理工具")
    try:
        os.makedirs(path, exist_ok=True)
    except OSError:
        path = os.path.expanduser("~")
    return path


def history_path() -> str:
    return os.path.join(_base_dir(), "history.jsonl")


def append(outcome, opt=None):
    rec = {
        "ts": outcome.ts or time.strftime("%Y-%m-%d %H:%M"),
        "path": outcome.path,
        "name": os.path.basename(outcome.path),
        "ok": bool(outcome.ok),
        "skipped": bool(outcome.skipped),
        "fmt": outcome.fmt,
        "total": outcome.names_total,
        "removed": outcome.names_removed,
        "kept": outcome.names_kept,
        "links": outcome.links_removed,
        "size_before": outcome.size_before,
        "size_after": outcome.size_after,
        "seconds": round(outcome.seconds, 3),
        "message": outcome.message or "",
        "output": outcome.output_path or "",
        "backup": outcome.backup_path or "",
        "keep_print": bool(opt.keep_print_names) if opt else True,
        "remove_links": bool(opt.remove_links) if opt else False,
        "kept_names": [[n, r] for n, r in (outcome.kept_names or [])][:50],
        "removed_sample": (outcome.removed_names or [])[:80],
    }
    try:
        with open(history_path(), "a", encoding="utf-8") as fh:
            fh.write(json.dumps(rec, ensure_ascii=False) + "\n")
    except OSError:
        pass


def load(limit: int = MAX_RECORDS):
    path = history_path()
    if not os.path.exists(path):
        return []
    out = []
    try:
        with open(path, "r", encoding="utf-8", errors="replace") as fh:
            for line in fh:
                line = line.strip()
                if not line:
                    continue
                try:
                    out.append(json.loads(line))
                except ValueError:
                    continue
    except OSError:
        return []
    return out[-limit:][::-1]


def clear():
    try:
        os.remove(history_path())
    except OSError:
        pass
