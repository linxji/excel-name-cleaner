# -*- coding: utf-8 -*-
# Author: Opera-Jin
"""自绘控件：圆角按钮、复选框、拖放区、进度条。"""
from __future__ import annotations

import tkinter as tk

from . import theme as T


def round_points(x1, y1, x2, y2, r):
    return [
        x1 + r, y1, x2 - r, y1, x2, y1, x2, y1 + r,
        x2, y2 - r, x2, y2, x2 - r, y2, x1 + r, y2,
        x1, y2, x1, y2 - r, x1, y1 + r, x1, y1,
    ]


class RoundedButton(tk.Canvas):
    """扁平圆角按钮，带 hover / press / disabled 三态。"""

    def __init__(self, master, text, command=None, th=None, kind="primary",
                 width=120, height=38, bg=None, icon=None):
        self.th = th
        self.kind = kind
        self.text = text
        self.icon = icon
        self.command = command
        self._enabled = True
        self._state = "normal"
        bg = bg or master["bg"]
        w = th.s(width) if isinstance(width, int) else width
        h = th.s(height) if isinstance(height, int) else height
        super().__init__(master, width=w, height=h, bg=bg, highlightthickness=0, bd=0)
        self.bind("<Enter>", self._on_enter)
        self.bind("<Leave>", self._on_leave)
        self.bind("<ButtonPress-1>", self._on_press)
        self.bind("<ButtonRelease-1>", self._on_release)
        self.bind("<Configure>", lambda e: self._draw())
        self._draw()

    # -- 配色
    def _colors(self):
        if self.kind == "primary":
            base = {"normal": T.ACCENT, "hover": T.ACCENT_HOVER,
                    "press": T.ACCENT_PRESS, "disabled": "#B9C6DC"}[self._state]
            return base, base, "#FFFFFF"
        if self.kind == "ghost":
            base = {"normal": T.CARD, "hover": T.ACCENT_SOFT,
                    "press": T.ACCENT_SOFT, "disabled": T.CARD}[self._state]
            border = {"normal": T.BORDER, "hover": "#BFD3F5",
                      "press": "#BFD3F5", "disabled": T.BORDER}[self._state]
            fg = {"normal": T.TEXT, "hover": T.ACCENT,
                  "press": T.ACCENT, "disabled": T.MUTED}[self._state]
            return base, border, fg
        base = {"normal": T.ACCENT_SOFT, "hover": "#E3EDFF",
                "press": "#D8E5FF", "disabled": "#F3F5F9"}[self._state]
        fg = {"normal": T.ACCENT, "hover": T.ACCENT_HOVER,
              "press": T.ACCENT_PRESS, "disabled": T.MUTED}[self._state]
        return base, base, fg

    def _draw(self):
        self.delete("all")
        w = self.winfo_width() or int(self["width"])
        h = self.winfo_height() or int(self["height"])
        r = self.th.s(8)
        fill, outline, fg = self._colors()
        self.create_polygon(round_points(1, 1, w - 1, h - 1, r), smooth=True,
                            splinesteps=16, fill=fill, outline=outline, width=1)
        tx = w / 2
        if self.icon:
            iw = self.th.s(15)
            gap = self.th.s(7)
            tw = self.th.s(len(self.text) * 13 * 0.62)
            start = tx - (iw + gap + tw) / 2
            self.create_text(start, h / 2, text=self.icon, anchor="w",
                             font=self.th.f_btn, fill=fg)
            self.create_text(start + iw + gap, h / 2, text=self.text, anchor="w",
                             font=self.th.f_btn, fill=fg)
        else:
            self.create_text(tx, h / 2, text=self.text, font=self.th.f_btn, fill=fg)

    # -- 事件
    def _on_enter(self, _):
        if self._enabled:
            self._state = "hover"
            self._draw()

    def _on_leave(self, _):
        if self._enabled:
            self._state = "normal"
            self._draw()

    def _on_press(self, _):
        if self._enabled:
            self._state = "press"
            self._draw()

    def _on_release(self, _):
        if not self._enabled:
            return
        self._state = "hover"
        self._draw()
        if self.command:
            self.command()

    def set_text(self, text: str):
        self.text = text
        self._draw()

    def set_enabled(self, on: bool):
        self._enabled = bool(on)
        self._state = "normal" if on else "disabled"
        self.configure(cursor="hand2" if on else "arrow")
        self._draw()


class CheckBox(tk.Frame):
    """自绘复选框：18px 圆角方框 + 勾。"""

    def __init__(self, master, text, value=False, command=None, th=None,
                 bg=None, note=None):
        bg = bg or master["bg"]
        super().__init__(master, bg=bg)
        self.th = th
        self.value = bool(value)
        self.command = command
        self._bg = bg
        self._note = note

        box = th.s(18)
        self.canvas = tk.Canvas(self, width=box, height=box, bg=bg,
                                highlightthickness=0, bd=0)
        self.canvas.grid(row=0, column=0, sticky="n", pady=(th.s(1), 0))
        self.label = tk.Label(self, text=text, bg=bg, fg=T.TEXT,
                              font=th.f_body, anchor="w", justify="left")
        self.label.grid(row=0, column=1, sticky="w", padx=(th.s(9), 0))
        self.note_label = None
        if note:
            self.note_label = tk.Label(self, text=note, bg=bg, fg=T.MUTED,
                                       font=th.f_tiny, anchor="w")
            self.note_label.grid(row=1, column=1, sticky="w", padx=(th.s(9), 0))

        for w in (self, self.canvas, self.label):
            w.bind("<Button-1>", self._toggle)
            try:
                w.configure(cursor="hand2")
            except tk.TclError:
                pass
        if self.note_label is not None:
            self.note_label.bind("<Button-1>", self._toggle)
        self._draw()

    def _draw(self):
        self.canvas.delete("all")
        b = self.th.s(18)
        r = self.th.s(5)
        if self.value:
            fill, outline = T.ACCENT, T.ACCENT
        else:
            fill, outline = "#FFFFFF", "#C6CEDA"
        self.canvas.create_polygon(round_points(1, 1, b - 1, b - 1, r),
                                   smooth=True, splinesteps=12,
                                   fill=fill, outline=outline, width=1)
        if self.value:
            self.canvas.create_line(b * 0.28, b * 0.52, b * 0.44, b * 0.68,
                                    b * 0.74, b * 0.33,
                                    fill="#FFFFFF", width=max(2, self.th.s(2)),
                                    capstyle="round", joinstyle="round")

    def _toggle(self, _=None):
        self.value = not self.value
        self._draw()
        if self.command:
            self.command(self.value)

    def set(self, v: bool):
        if self.value != bool(v):
            self.value = bool(v)
            self._draw()


class DropZone(tk.Canvas):
    """拖放 / 点击选择区域。"""

    def __init__(self, master, th, on_click, on_drop, bg=T.BG):
        self.th = th
        self.on_click = on_click
        self.on_drop = on_drop
        self._hover = False
        self._drag = False
        super().__init__(master, bg=bg, highlightthickness=0, bd=0, height=th.s(104))
        self.bind("<Configure>", lambda e: self._draw())
        self.bind("<Button-1>", lambda e: self.on_click())
        self.bind("<Enter>", self._enter)
        self.bind("<Leave>", self._leave)
        self.configure(cursor="hand2")

    def _enter(self, _):
        self._hover = True
        self._draw()

    def _leave(self, _):
        self._hover = False
        self._draw()

    def set_drag(self, on):
        self._drag = bool(on)
        self._draw()

    def _draw(self):
        self.delete("all")
        w = self.winfo_width()
        h = self.winfo_height()
        if w < 10 or h < 10:
            return
        r = self.th.s(10)
        fill = T.ACCENT_SOFT if (self._hover or self._drag) else T.DROP_BG
        outline = T.ACCENT if (self._hover or self._drag) else T.DROP_BORDER
        self.create_polygon(round_points(1, 1, w - 1, h - 1, r), smooth=True,
                            splinesteps=20, fill=fill, outline=outline, width=1)

        cx = w / 2
        top = h / 2 - self.th.s(26)
        # 简易上传图标
        s = self.th.s(1)
        aw = 11 * s
        ah = 13 * s
        col = T.ACCENT if (self._hover or self._drag) else "#8FA3BF"
        self.create_line(cx, top - ah / 2, cx, top + ah / 2 - 2 * s,
                         fill=col, width=max(2, 2 * s), capstyle="round")
        self.create_line(cx - aw / 2, top - 1 * s, cx, top - ah / 2,
                         cx + aw / 2, top - 1 * s,
                         fill=col, width=max(2, 2 * s), capstyle="round", joinstyle="round")
        self.create_line(cx - aw * 0.75, top + ah / 2, cx + aw * 0.75, top + ah / 2,
                         fill=col, width=max(2, 2 * s), capstyle="round")

        self.create_text(cx, h / 2 + self.th.s(6), text="拖入表格文件或文件夹",
                         font=self.th.f_body_b, fill=T.TEXT)
        self.create_text(cx, h / 2 + self.th.s(26),
                         text="支持 .xlsx  .xlsm  .xlsb  .xls  .et  .ett   ·   也可点击选择",
                         font=self.th.f_small, fill=T.SUBTEXT)


class ProgressBar(tk.Canvas):
    def __init__(self, master, th, bg=T.CARD):
        self.th = th
        super().__init__(master, bg=bg, highlightthickness=0, bd=0, height=th.s(6))
        self._v = 0.0
        self.bind("<Configure>", lambda e: self._draw())

    def set(self, v):
        self._v = max(0.0, min(1.0, float(v)))
        self._draw()

    def _draw(self):
        self.delete("all")
        w = self.winfo_width()
        h = self.winfo_height()
        if w < 4 or h < 2:
            return
        r = h / 2
        self.create_polygon(round_points(0, 0, w, h, r), smooth=True,
                            splinesteps=12, fill=T.TRACK, outline=T.TRACK)
        if self._v > 0:
            fw = max(h, w * self._v)
            self.create_polygon(round_points(0, 0, fw, h, r), smooth=True,
                                splinesteps=12, fill=T.ACCENT, outline=T.ACCENT)


class Card(tk.Frame):
    """白底 + 1px 描边的卡片。"""

    def __init__(self, master, th, padding=0, bg=T.CARD):
        super().__init__(master, bg=bg, highlightbackground=T.BORDER,
                         highlightcolor=T.BORDER, highlightthickness=1, bd=0)
        self.th = th
        self.inner = tk.Frame(self, bg=bg)
        self.inner.pack(fill="both", expand=True, padx=th.s(padding), pady=th.s(padding))
