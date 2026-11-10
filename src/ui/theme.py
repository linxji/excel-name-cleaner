# -*- coding: utf-8 -*-
# Author: Opera-Jin
"""配色与字号（浅色主题，8px 栅格）。"""
from __future__ import annotations

import tkinter.font as tkfont

BG = "#F4F6F9"
CARD = "#FFFFFF"
BORDER = "#E2E7EF"
BORDER_SOFT = "#EDF1F6"
TEXT = "#1A1F26"
SUBTEXT = "#6B7480"
MUTED = "#9AA3AF"

ACCENT = "#2563EB"
ACCENT_HOVER = "#1D4FD8"
ACCENT_PRESS = "#1A45BE"
ACCENT_SOFT = "#EFF4FF"

SUCCESS = "#15803D"
SUCCESS_SOFT = "#E9F7EE"
DANGER = "#B91C1C"
DANGER_SOFT = "#FDECEC"
WARN = "#B45309"
WARN_SOFT = "#FEF4E6"

DROP_BG = "#FBFCFE"
DROP_BORDER = "#CBD5E3"
TRACK = "#EDF1F6"

FAMILY = "Microsoft YaHei UI"
FAMILY_FALLBACK = "Microsoft YaHei"
MONO = "Consolas"


class Theme:
    """所有尺寸都经过 s() 缩放，保证 100%~200% DPI 下视觉一致。"""

    def __init__(self, root):
        try:
            dpi = root.winfo_fpixels("1i")
        except Exception:  # noqa: BLE001
            dpi = 96.0
        self.scale = max(1.0, round(dpi / 96.0, 2))
        fams = set(tkfont.families(root))
        family = FAMILY if FAMILY in fams else (FAMILY_FALLBACK if FAMILY_FALLBACK in fams else "Segoe UI")
        self.family = family

        self.f_h1 = (family, -self.s(19), "bold")
        self.f_h2 = (family, -self.s(14), "bold")
        self.f_body = (family, -self.s(13))
        self.f_body_b = (family, -self.s(13), "bold")
        self.f_small = (family, -self.s(12))
        self.f_small_b = (family, -self.s(12), "bold")
        self.f_tiny = (family, -self.s(11))
        self.f_btn = (family, -self.s(13), "bold")
        self.f_mono = (MONO, -self.s(12))

    def s(self, v):
        return int(round(v * self.scale))
