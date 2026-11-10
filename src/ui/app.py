# -*- coding: utf-8 -*-
# Author: Opera-Jin
"""主界面。"""
from __future__ import annotations

import json
import os
import queue
import subprocess
import threading
import time
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from core import engine, history
from core.engine import Options
from core.runlog import RunLog, log_dir
from core.util import human_size
from ui import theme as T
from ui.dnd import DropTarget
from ui.widgets import (Card, CheckBox, DropZone, ProgressBar, RoundedButton,
                        round_points)

APP_NAME = "Excel 名称清理工具"
APP_VER = "v1.4.0"
APP_AUTHOR = "Opera-Jin"
FOOTER_TEXT = "本地处理 · 不联网 · 不上传任何文件"

# 列定义：(key, 标题, 默认宽, 对齐, 默认是否显示)
COLUMNS = (
    ("name",   "文件名", 240, "w", True),
    ("size",   "大小",    76, "w", True),
    ("names",  "名称",    66, "w", True),
    ("dur",    "耗时",    62, "w", True),
    ("status", "结果",   200, "w", True),
    ("time",   "时间",    56, "w", True),
)
COL_TITLES = {c[0]: c[1] for c in COLUMNS}
COL_ANCHOR = {c[0]: c[3] for c in COLUMNS}
COL_DEFAULT_W = {c[0]: c[2] for c in COLUMNS}
COL_ORDER = [c[0] for c in COLUMNS]


def ui_conf_path() -> str:
    base = os.environ.get("LOCALAPPDATA") or os.path.expanduser("~")
    return os.path.join(base, "Excel名称清理工具", "ui.json")

BELOW_NORMAL_PRIORITY_CLASS = 0x00004000
NORMAL_PRIORITY_CLASS = 0x00000020


def _set_process_low_priority(on: bool):
    try:
        import ctypes
        k32 = ctypes.windll.kernel32
        k32.SetPriorityClass(k32.GetCurrentProcess(),
                             BELOW_NORMAL_PRIORITY_CLASS if on else NORMAL_PRIORITY_CLASS)
    except Exception:  # noqa: BLE001
        pass


def _set_thread_low_priority():
    try:
        import ctypes
        k32 = ctypes.windll.kernel32
        k32.SetThreadPriority(k32.GetCurrentThread(), -1)  # THREAD_PRIORITY_BELOW_NORMAL
    except Exception:  # noqa: BLE001
        pass


def _fmt_sec(sec) -> str:
    """单文件耗时：精确到 0.1 秒。"""
    try:
        sec = max(0.0, float(sec))
    except (TypeError, ValueError):
        return "-"
    return "%.1f 秒" % sec if sec < 60 else "%.0f 秒" % sec


def _fmt_dur(sec: float) -> str:
    sec = max(0.0, float(sec))
    if sec < 10:
        return "%.1f 秒" % sec
    sec = int(sec)
    if sec < 60:
        return "%d 秒" % sec
    if sec < 3600:
        return "%d 分 %02d 秒" % (sec // 60, sec % 60)
    return "%d 时 %02d 分" % (sec // 3600, sec % 3600 // 60)


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.th = T.Theme(self)
        th = self.th

        self.title(APP_NAME)
        self.configure(bg=T.BG)
        w = th.s(960)
        h = min(th.s(720), int(self.winfo_screenheight() * 0.94))
        self.geometry("%dx%d" % (w, h))
        self.minsize(th.s(880), min(th.s(640), h))
        self._center(w, h)

        self.files = []
        self.results = {}
        self.running = False
        self.q = queue.Queue()
        self.log = RunLog()
        self._links_confirmed = False
        self._t0 = 0.0
        self._cur_opt = None
        self._cur_idx = 0
        self._cur_name = ""
        self._cur_stage = ""
        self._cur_done = 0
        self._cur_total = 0
        self._done_files = 0
        self._total_files = 0
        self._removed_sum = 0
        self._saved_sum = 0
        self._hist_recs = []
        self.rows = {}
        self.count_q = queue.Queue()
        self._counting = 0
        self._backups_removed = 0

        self._init_style()
        self._build()
        self._set_icon()

        self.after(200, self._poll_counts)
        self.drop = DropTarget(self, self._on_drop)
        if not self.drop.ok:
            # 拖放挂接失败时必须让用户看见，否则会以为"拖进去没反应"
            self.lbl_status.configure(
                text="拖放功能不可用（%s），请改用「添加文件」按钮"
                     % (self.drop.error or "未知原因"), fg=T.WARN)
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.bind("<Delete>", lambda e: self._remove_selected())

        import sys
        argv = [a for a in sys.argv[1:] if os.path.exists(a)]
        if argv:
            self.after(120, lambda: self._add(argv))

    # ------------------------------------------------------------ 基础
    def _center(self, w, h):
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        self.geometry("+%d+%d" % (max(0, (sw - w) // 2), max(0, (sh - h) // 3)))

    def _set_icon(self):
        import sys
        cands = []
        if hasattr(sys, "_MEIPASS"):
            cands.append(os.path.join(sys._MEIPASS, "app.ico"))
        here = os.path.dirname(os.path.abspath(__file__))
        cands.append(os.path.join(here, "..", "..", "assets", "app.ico"))
        for c in cands:
            try:
                if os.path.exists(c):
                    self.iconbitmap(c)
                    return
            except Exception:  # noqa: BLE001
                continue

    def _init_style(self):
        th = self.th
        st = ttk.Style(self)
        try:
            st.theme_use("clam")
        except tk.TclError:
            pass
        st.configure("Files.Treeview", background=T.CARD, fieldbackground=T.CARD,
                     foreground=T.TEXT, borderwidth=0, rowheight=th.s(31),
                     font=th.f_body)
        st.layout("Files.Treeview", [("Files.Treeview.treearea", {"sticky": "nswe"})])
        # 表头：浅灰底 + 细分隔线 + 悬停高亮，一眼能分清列
        # 表头布局：文字靠左、分隔线图片贴右边缘。
        # 必须保留 Treeheading.padding 元素，它负责绘制表头背景和边框。
        st.layout("Files.Treeview.Heading", [
            ("Treeheading.cell", {"sticky": "nswe", "children": [
                ("Treeheading.padding", {"sticky": "nswe", "children": [
                    ("Treeheading.text", {"sticky": "w", "side": "left"}),
                    ("Treeheading.image", {"sticky": "ns", "side": "right"}),
                ]}),
            ]}),
        ])
        st.configure("Files.Treeview.Heading",
                     background="#EFF3F9", foreground="#59636F",
                     relief="solid", borderwidth=1,
                     bordercolor="#D2DAE6", lightcolor="#EFF3F9",
                     darkcolor="#D2DAE6", font=th.f_small_b,
                     padding=(10, th.s(7), 0, th.s(7)))
        st.map("Files.Treeview.Heading",
               background=[("active", "#E3EDFD")],
               foreground=[("active", T.ACCENT)],
               lightcolor=[("active", "#E3EDFD")],
               darkcolor=[("active", "#D6E2F8")])
        st.map("Files.Treeview", background=[("selected", T.ACCENT_SOFT)],
               foreground=[("selected", T.TEXT)])
        st.configure("Thin.Vertical.TScrollbar", background="#D3DAE5", troughcolor=T.CARD,
                     bordercolor=T.CARD, arrowcolor=T.SUBTEXT, borderwidth=0,
                     arrowsize=th.s(12), width=th.s(8))
        st.map("Thin.Vertical.TScrollbar", background=[("active", "#BFC9D8")])
        st.layout("Thin.Vertical.TScrollbar", [
            ("Vertical.Scrollbar.trough", {
                "sticky": "ns",
                "children": [("Vertical.Scrollbar.thumb", {"expand": "1", "sticky": "nswe"})],
            })
        ])
        st.configure("Thin.Horizontal.TScrollbar", background="#D3DAE5",
                     troughcolor=T.CARD, bordercolor=T.CARD,
                     arrowcolor=T.SUBTEXT, borderwidth=0, arrowsize=th.s(12))
        st.map("Thin.Horizontal.TScrollbar", background=[("active", "#BFC9D8")])
        st.layout("Thin.Horizontal.TScrollbar", [
            ("Horizontal.Scrollbar.trough", {
                "sticky": "ew",
                "children": [("Horizontal.Scrollbar.thumb", {"expand": "1", "sticky": "nswe"})],
            })
        ])

    # ------------------------------------------------------------ 布局
    def _build(self):
        th = self.th
        root = tk.Frame(self, bg=T.BG)
        root.pack(fill="both", expand=True)

        self._build_header(root)
        self._build_footer(root)

        body = tk.Frame(root, bg=T.BG)
        body.pack(fill="both", expand=True, padx=th.s(22), pady=(th.s(16), th.s(16)))
        body.columnconfigure(0, weight=1)
        body.rowconfigure(2, weight=1)

        self.dropzone = DropZone(body, th, self._pick_files, self._on_drop, bg=T.BG)
        self.dropzone.grid(row=0, column=0, sticky="ew")

        bar = tk.Frame(body, bg=T.BG)
        bar.grid(row=1, column=0, sticky="ew", pady=(th.s(14), th.s(8)))
        self.lbl_files = tk.Label(bar, text="待处理文件", bg=T.BG, fg=T.TEXT,
                                  font=th.f_body_b)
        self.lbl_files.pack(side="left")

        RoundedButton(bar, "清空", self._clear, th, kind="ghost", width=64,
                      height=30, bg=T.BG).pack(side="right")
        self.btn_remove = RoundedButton(bar, "移除选中", self._remove_selected, th,
                                        kind="ghost", width=96, height=30, bg=T.BG)
        self.btn_remove.pack(side="right", padx=(0, th.s(8)))
        self.btn_remove.set_enabled(False)
        RoundedButton(bar, "添加文件夹", self._pick_folder, th, kind="ghost",
                      width=100, height=30, bg=T.BG).pack(side="right", padx=(0, th.s(8)))
        RoundedButton(bar, "添加文件", self._pick_files, th, kind="ghost",
                      width=88, height=30, bg=T.BG).pack(side="right", padx=(0, th.s(8)))

        card = Card(body, th, padding=1)
        card.grid(row=2, column=0, sticky="nsew")
        wrap = tk.Frame(card.inner, bg=T.CARD)
        wrap.pack(fill="both", expand=True)

        # 先建容器：Treeview 的 master 必须是它，否则 grid 会和外层 pack 冲突
        inner = tk.Frame(wrap, bg=T.CARD)
        # 不再留内边距：表头直接贴住卡片边框，避免"灰块浮在中间"的观感
        inner.pack(fill="both", expand=True, padx=0, pady=0)
        inner.rowconfigure(0, weight=1)
        inner.columnconfigure(0, weight=1)

        self.tree = ttk.Treeview(inner, columns=COL_ORDER, show="headings", height=6,
                                 style="Files.Treeview", selectmode="extended")
        self._sort_col = ""
        self._sort_rev = False
        self._col_cfg = {k: {"width": th.s(COL_DEFAULT_W[k]), "visible": True}
                         for k in COL_ORDER}
        self._col_order = list(COL_ORDER)
        # 列分隔线：ttk 的 clam 表头不画列间竖线，这里用一条 1px 图片补上
        self._vline = tk.PhotoImage(width=1, height=18)
        self._vline.put("#C6CFDC", to=(0, 0, 1, 18))
        self._head_press = None
        self._head_drag = False
        self._head_region = ""
        self._user_resized = False
        self._load_layout()
        self._refresh_headings()
        for k in self._col_order:
            self.tree.column(k, width=self._col_cfg[k]["width"],
                             minwidth=th.s(44), stretch=False,
                             anchor=COL_ANCHOR[k])
        self._apply_columns()
        self.tree.tag_configure("wait", foreground=T.SUBTEXT)
        self.tree.tag_configure("run", foreground=T.ACCENT)
        self.tree.tag_configure("ok", foreground=T.SUCCESS)
        self.tree.tag_configure("fail", foreground=T.DANGER)
        self.tree.tag_configure("skip", foreground=T.WARN)

        # 垂直/水平滚动条（水平条让用户能横向看完整列）
        sb = ttk.Scrollbar(inner, orient="vertical", command=self.tree.yview,
                           style="Thin.Vertical.TScrollbar")
        sbx = ttk.Scrollbar(inner, orient="horizontal", command=self.tree.xview,
                            style="Thin.Horizontal.TScrollbar")
        self.tree.configure(yscrollcommand=sb.set, xscrollcommand=sbx.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        sb.grid(row=0, column=1, sticky="ns", padx=(th.s(2), th.s(4)))
        sbx.grid(row=1, column=0, sticky="ew", pady=(th.s(2), 0))
        self.tree.bind("<Double-1>", self._on_double)
        self.tree.bind("<<TreeviewSelect>>", lambda e: self._sync_remove_btn())
        self.tree.bind("<Button-3>", self._on_right_click)
        self.tree.bind("<Configure>", self._fit_name_column, add="+")
        self.tree.bind("<Button-1>", self._on_head_press, add="+")
        self.tree.bind("<B1-Motion>", self._on_head_motion, add="+")
        self.tree.bind("<ButtonRelease-1>", self._on_head_release, add="+")
        self._build_menu()

        self.empty_hint = tk.Label(
            card.inner, text="把表格文件拖到上方，或点击「添加文件」",
            bg=T.CARD, fg=T.MUTED, font=th.f_body)
        self.empty_hint.place(relx=0.5, rely=0.52, anchor="center")

        opt_card = Card(body, th, padding=0)
        opt_card.grid(row=3, column=0, sticky="ew", pady=(th.s(14), 0))
        grid = tk.Frame(opt_card.inner, bg=T.CARD)
        grid.pack(fill="x", padx=th.s(16), pady=th.s(12))
        grid.columnconfigure(0, weight=1, uniform="o")
        grid.columnconfigure(1, weight=1, uniform="o")

        self.cb_print = CheckBox(grid, "保留打印区域 / 打印标题名称", True, None, th,
                                 bg=T.CARD, note="取消勾选清理更彻底，但打印设置会被重置")
        self.cb_print.grid(row=0, column=0, sticky="w", pady=(0, th.s(10)))
        self.cb_links = CheckBox(grid, "断开外部链接（减小体积）", True,
                                 self._on_links_toggle, th, bg=T.CARD,
                                 note="外链公式转成当前数值，与 Excel 手动断开一致")
        self.cb_links.grid(row=0, column=1, sticky="w", pady=(0, th.s(10)))
        self.cb_backup = CheckBox(grid, "处理前自动备份（.bak）", True, None, th,
                                  bg=T.CARD, note="成功后自动删除")
        self.cb_backup.grid(row=1, column=0, sticky="w")
        self.cb_bg = CheckBox(grid, "后台模式", True, None, th,
                              bg=T.CARD, note="降低性能占用，不影响前台工作")
        self.cb_bg.grid(row=1, column=1, sticky="w")

    def _build_header(self, root):
        th = self.th
        header = tk.Frame(root, bg=T.CARD, height=th.s(68))
        header.pack(fill="x")
        header.pack_propagate(False)

        left = tk.Frame(header, bg=T.CARD)
        # 标题 + 副标题实测需要 63px（DPI 1.25 下 header 为 85px）；
        # 原来上下各 14px 内边距只留 49px，会把副标题下缘裁掉。
        left.pack(side="left", padx=th.s(22), pady=th.s(2))

        icon = tk.Canvas(left, width=th.s(36), height=th.s(36), bg=T.CARD,
                         highlightthickness=0, bd=0)
        icon.pack(side="left")
        r = th.s(10)
        icon.create_polygon(round_points(0, 0, th.s(36), th.s(36), r), smooth=True,
                            splinesteps=18, fill=T.ACCENT, outline=T.ACCENT)
        s = th.s(36)
        icon.create_line(s * 0.28, s * 0.52, s * 0.44, s * 0.68, s * 0.74, s * 0.33,
                         fill="#FFFFFF", width=max(2, th.s(3)),
                         capstyle="round", joinstyle="round")

        txt = tk.Frame(left, bg=T.CARD)
        txt.pack(side="left", padx=(th.s(11), 0))
        tk.Label(txt, text=APP_NAME, bg=T.CARD, fg=T.TEXT,
                 font=th.f_h1).pack(anchor="w")
        tk.Label(txt, text="一键清除无效名称 · 不改动任何数据与公式",
                 bg=T.CARD, fg=T.SUBTEXT, font=th.f_small).pack(anchor="w")

        right = tk.Frame(header, bg=T.CARD)
        right.pack(side="right", padx=th.s(22))
        tk.Label(right, text=APP_VER, bg=T.CARD, fg=T.MUTED,
                 font=th.f_small).pack(side="right", padx=(th.s(12), 0))
        RoundedButton(right, "历史记录", self._open_history, th, kind="ghost",
                      width=88, height=32, bg=T.CARD).pack(side="right")

        tk.Frame(root, bg=T.BORDER, height=1).pack(fill="x")

    def _build_footer(self, root):
        th = self.th
        footer = tk.Frame(root, bg=T.CARD, height=th.s(112))
        footer.pack(fill="x", side="bottom")
        footer.pack_propagate(False)
        tk.Frame(root, bg=T.BORDER, height=1).pack(fill="x", side="bottom")

        inner = tk.Frame(footer, bg=T.CARD)
        inner.pack(fill="both", expand=True, padx=th.s(22), pady=(th.s(13), th.s(12)))

        self.progress = ProgressBar(inner, th, bg=T.CARD)
        self.progress.pack(fill="x")

        row = tk.Frame(inner, bg=T.CARD)
        row.pack(fill="x", pady=(th.s(11), 0))

        self.btn_start = RoundedButton(row, "开始清理", self._start, th,
                                       kind="primary", width=132, height=38, bg=T.CARD)
        self.btn_start.pack(side="right")

        self.btn_log = RoundedButton(row, "打开日志", self._open_log, th,
                                     kind="ghost", width=96, height=38, bg=T.CARD)
        self.btn_log.pack(side="right", padx=(0, th.s(10)))

        self.lbl_status = tk.Label(row, text="拖入文件后点击「开始清理」", bg=T.CARD,
                                   fg=T.SUBTEXT, font=th.f_small, anchor="w")
        self.lbl_status.pack(side="left", fill="x", expand=True)

        self.lbl_metrics = tk.Label(inner, text=FOOTER_TEXT,
                                    bg=T.CARD, fg=T.MUTED, font=th.f_tiny, anchor="w")
        self.lbl_metrics.pack(fill="x", pady=(th.s(7), 0))

    # ------------------------------------------------------------ 文件
    def _on_drop(self, paths):
        self._add(paths)

    def _pick_files(self):
        paths = filedialog.askopenfilenames(
            parent=self, title="选择 Excel / WPS 文件",
            filetypes=[("Excel / WPS 文件",
                        "*.xlsx *.xlsm *.xltx *.xltm *.xlsb *.xls *.et *.ett"),
                       ("所有文件", "*.*")])
        if paths:
            self._add(list(paths))

    def _pick_folder(self):
        path = filedialog.askdirectory(parent=self, title="选择文件夹")
        if path:
            self._add([path])

    def _add(self, paths):
        opt = self._opts()
        exts = engine.SUPPORTED_EXT

        # 明确拖进来但格式不对的文件，直接告诉用户
        bad = [p for p in paths
               if os.path.isfile(p) and os.path.splitext(p)[1].lower() not in exts]
        if bad:
            names = "\n".join("· " + os.path.basename(b) for b in bad[:6])
            more = "\n…… 共 %d 个" % len(bad) if len(bad) > 6 else ""
            messagebox.showwarning(
                "格式不支持",
                "以下文件不是受支持的 Excel / WPS 格式，已跳过：\n\n%s%s\n\n"
                "支持：.xlsx  .xlsm  .xlsb  .xls  .et" % (names, more),
                parent=self)

        found = engine.collect_files(paths, recurse=opt.recurse)
        added = []
        for p in found:
            if os.path.splitext(p)[1].lower() not in exts:
                continue
            if p in self.results or p in self.files:
                continue
            self.files.append(p)
            try:
                nbytes = os.path.getsize(p)
            except OSError:
                nbytes = 0
            self.rows[p] = {
                "size": human_size(nbytes),
                "size_bytes": nbytes,
                "names": "统计中…",
                "names_num": None,
                "dur": "",
                "dur_sec": None,
                "status": "等待处理",
                "time": "",
                "tag": "wait",
            }
            self.tree.insert("", "end", iid=p, values=(), tags=("wait",))
            self._render(p)
            added.append(p)
        self._refresh_count()
        if added:
            self.tree.see(added[-1])
            self._counting += len(added)
            self.lbl_status.configure(text="已添加 %d 个文件，正在统计名称数量…"
                                      % len(added), fg=T.SUBTEXT)
            threading.Thread(target=self._count_worker, args=(added,), daemon=True).start()
        elif not found and not bad:
            self.lbl_status.configure(text="没有找到可处理的 Excel / WPS 文件", fg=T.WARN)

    # ---- 行渲染 / 名称数统计
    def _render(self, path):
        d = self.rows.get(path)
        if not d or not self.tree.exists(path):
            return
        self.tree.item(path, values=(os.path.basename(path), d["size"], d["names"],
                                     d["dur"], d["status"], d["time"]),
                       tags=(d["tag"],))

    def _set_row(self, path, status=None, tag=None, tm=None, dur=None):
        d = self.rows.get(path)
        if not d:
            return
        if status is not None:
            d["status"] = status
        if tag is not None:
            d["tag"] = tag
        if tm is not None:
            d["time"] = tm
        if dur is not None:
            if isinstance(dur, str):
                d["dur"] = dur
            else:
                d["dur"] = _fmt_sec(dur)
                d["dur_sec"] = float(dur)
        self._render(path)

    def _count_worker(self, paths):
        for p in paths:
            if os.path.splitext(p)[1].lower() == ".xlsb":
                self.count_q.put((p, "需转换"))
                continue
            try:
                n = engine.count_names(p)
            except Exception:  # noqa: BLE001
                n = None
            self.count_q.put((p, "?" if n is None else "{:,}".format(n)))

    def _poll_counts(self):
        if getattr(self, "_closing", False):
            return
        try:
            while True:
                p, txt = self.count_q.get_nowait()
                d = self.rows.get(p)
                if d is not None:
                    d["names"] = txt
                    num = txt.replace(",", "")
                    d["names_num"] = int(num) if num.isdigit() else None
                    self._render(p)
                self._counting = max(0, self._counting - 1)
                if self._counting == 0 and not self.running:
                    self.lbl_status.configure(
                        text="已添加 %d 个文件" % len(self.files), fg=T.SUBTEXT)
        except queue.Empty:
            pass
        try:
            self.after(200, self._poll_counts)
        except tk.TclError:
            pass

    def _clear(self):
        if self.running:
            return
        self.files.clear()
        self.results.clear()
        self.rows.clear()
        for iid in self.tree.get_children():
            self.tree.delete(iid)
        self._refresh_count()
        self.progress.set(0)
        self.lbl_status.configure(text="拖入文件后点击「开始清理」", fg=T.SUBTEXT)
        self.lbl_metrics.configure(text=FOOTER_TEXT, fg=T.MUTED)

    def _remove_selected(self):
        """移除列表里选中的文件（可多选），不动磁盘上的任何文件。"""
        if self.running:
            return
        sel = list(self.tree.selection())
        if not sel:
            return
        for iid in sel:
            try:
                self.files.remove(iid)
            except ValueError:
                pass
            # results 也要清掉：否则再次添加同一个文件会被当成「已处理」挡掉
            self.results.pop(iid, None)
            self.rows.pop(iid, None)
            if self.tree.exists(iid):
                self.tree.delete(iid)
        self._refresh_count()
        self._sync_remove_btn()
        if self.files:
            self.lbl_status.configure(text="已移除 %d 个文件" % len(sel), fg=T.SUBTEXT)
        else:
            self.lbl_status.configure(text="拖入文件后点击「开始清理」", fg=T.SUBTEXT)

    def _sync_remove_btn(self):
        try:
            self.btn_remove.set_enabled(bool(self.tree.selection()) and not self.running)
        except Exception:  # noqa: BLE001
            pass

    # ---- 右键菜单
    def _build_menu(self):
        th = self.th
        self.menu = tk.Menu(self, tearoff=0, bg=T.CARD, fg=T.TEXT, bd=0,
                            activebackground=T.ACCENT_SOFT,
                            activeforeground=T.ACCENT,
                            font=th.f_body, activeborderwidth=0)
        self.menu.add_command(label="移除选中", command=self._remove_selected)
        self.menu.add_command(label="打开文件所在位置", command=self._open_in_explorer)
        self.menu.add_separator()
        self.menu.add_command(label="清空列表", command=self._clear)
        self._build_col_menu()

    # ---- 表头：列宽 / 列顺序 / 显示列
    def _build_col_menu(self):
        th = self.th
        self._col_menu = tk.Menu(self, tearoff=0, bg=T.CARD, fg=T.TEXT, bd=0,
                                 activebackground=T.ACCENT_SOFT,
                                 activeforeground=T.ACCENT,
                                 font=th.f_body, activeborderwidth=0)
        self._col_menu.add_command(label="将列调整为合适的大小",
                                   command=self._autosize_columns)
        self._col_menu.add_command(label="恢复默认列宽与顺序",
                                   command=self._reset_columns)
        self._col_menu.add_separator()
        self._col_vars = {}
        for k in COL_ORDER:
            var = tk.BooleanVar(value=self._col_cfg[k]["visible"])
            self._col_vars[k] = var
            self._col_menu.add_checkbutton(
                label=COL_TITLES[k], variable=var,
                command=lambda kk=k: self._toggle_column(kk))

    def _show_col_menu(self, event):
        for k, var in self._col_vars.items():
            var.set(self._col_cfg[k]["visible"])
        try:
            self._col_menu.tk_popup(event.x_root, event.y_root)
        finally:
            try:
                self._col_menu.grab_release()
            except Exception:  # noqa: BLE001
                pass

    def _col_at(self, x):
        """把鼠标 x 换算成当前可见列的 key。"""
        cid = self.tree.identify_column(x)
        if not cid:
            return None
        try:
            idx = int(cid[1:]) - 1
        except ValueError:
            return None
        disp = self.tree["displaycolumns"]
        if not disp or disp == ("#all",) or disp == "#all":
            disp = tuple(self._col_order)
        if 0 <= idx < len(disp):
            return disp[idx]
        return None

    def _refresh_headings(self):
        vis = [x for x in self._col_order if self._col_cfg[x]["visible"]]
        last = vis[-1] if vis else None
        for k in COL_ORDER:
            title = COL_TITLES[k]
            if self._sort_col == k:
                title += " ▼" if self._sort_rev else " ▲"
            # 每列右侧挂一条竖线；最后一列不挂，免得贴着表格右边缘
            if k == last:
                self.tree.heading(k, text=title, anchor=COL_ANCHOR[k])
            else:
                self.tree.heading(k, text=title, anchor=COL_ANCHOR[k],
                                  image=self._vline)

    def _apply_columns(self):
        vis = [k for k in self._col_order if self._col_cfg[k]["visible"]]
        if not vis:
            vis = ["name"]
        self.tree["displaycolumns"] = tuple(vis)
        self._refresh_headings()

    def _toggle_column(self, k):
        self._col_cfg[k]["visible"] = bool(self._col_vars[k].get())
        if not any(self._col_cfg[x]["visible"] for x in COL_ORDER):
            self._col_cfg[k]["visible"] = True   # 至少留一列
            self._col_vars[k].set(True)
        self._apply_columns()
        self._save_layout()

    def _autosize_columns(self):
        """按内容把每列调到刚好放得下（类似资源管理器）。"""
        import tkinter.font as tkfont
        fb = tkfont.Font(font=self.th.f_small_b)
        fr = tkfont.Font(font=self.th.f_body)
        pad = self.th.s(28)
        for k in self.tree["displaycolumns"]:
            w = fb.measure(COL_TITLES[k]) + pad
            for iid in self.tree.get_children():
                w = max(w, fr.measure(str(self.tree.set(iid, k))) + pad)
            self.tree.column(k, width=min(w, self.th.s(560)))
        self._user_resized = True
        self._save_layout()

    def _reset_columns(self):
        self._col_order = list(COL_ORDER)
        for k in COL_ORDER:
            self._col_cfg[k] = {"width": self.th.s(COL_DEFAULT_W[k]), "visible": True}
            self.tree.column(k, width=self.th.s(COL_DEFAULT_W[k]),
                             anchor=COL_ANCHOR[k])
            if k in self._col_vars:
                self._col_vars[k].set(True)
        self._user_resized = False
        self._apply_columns()
        self._save_layout()
        self.after_idle(self._fit_name_column)

    def _fit_name_column(self, _event=None):
        """未手动调整过列宽时，让「文件名」列吃掉剩余宽度，表头始终填满。

        用户一旦拖过列宽就停止干预 —— 否则拖完会被拉回去，横向滚动条也失去意义。
        """
        if getattr(self, "_user_resized", False):
            return
        try:
            total = self.tree.winfo_width()
            if total <= 20:
                return
            disp = self.tree["displaycolumns"]
            others = sum(self.tree.column(k, "width") for k in disp if k != "name")
            w = max(self.th.s(160), total - others - 2)
            if abs(self.tree.column("name", "width") - w) > 2:
                self.tree.column("name", width=w)
        except Exception:  # noqa: BLE001
            pass

    # ---- 表头拖动 / 点击排序
    def _on_head_press(self, event):
        region = self.tree.identify_region(event.x, event.y)
        self._head_region = region
        if region != "heading":
            self._head_press = None
            return
        self._head_press = (event.x, self._col_at(event.x))
        self._head_drag = False

    def _on_head_motion(self, event):
        if not self._head_press:
            return
        x0, col0 = self._head_press
        if col0 and abs(event.x - x0) > self.th.s(20):
            self._head_drag = True
            try:
                self.tree.configure(cursor="fleur")
            except Exception:  # noqa: BLE001
                pass

    def _on_head_release(self, event):
        try:
            self.tree.configure(cursor="")
        except Exception:  # noqa: BLE001
            pass
        if not self._head_press:
            # 拖动列宽分隔线：ttk 已改好宽度，这里只负责记住
            if getattr(self, "_head_region", "") == "separator":
                self._user_resized = True
                self.after(60, self._save_layout)
            return
        _x0, col0 = self._head_press
        was_drag = self._head_drag
        self._head_press = None
        self._head_drag = False

        if was_drag:
            col1 = self._col_at(event.x)
            if col0 and col1 and col0 != col1:
                vis = [k for k in self._col_order if self._col_cfg[k]["visible"]]
                i, j = vis.index(col0), vis.index(col1)
                vis.insert(j, vis.pop(i))
                rest = [k for k in self._col_order if k not in vis]
                self._col_order = vis + rest
                self._apply_columns()
                self._save_layout()
            return
        col = self._col_at(event.x)
        if col:
            self._sort_by(col)

    def _sort_key(self, iid):
        d = self.rows.get(iid, {})
        c = self._sort_col
        if c == "name":
            return os.path.basename(iid).lower()
        if c == "size":
            return d.get("size_bytes") or 0
        if c == "names":
            n = d.get("names_num")
            return -1 if n is None else n
        if c == "dur":
            v = d.get("dur_sec")
            return -1.0 if v is None else v
        if c == "status":
            return d.get("status") or ""
        if c == "time":
            return d.get("time") or ""
        return ""

    def _sort_by(self, col):
        if self._sort_col == col:
            self._sort_rev = not self._sort_rev
        else:
            self._sort_col, self._sort_rev = col, False
        self._refresh_headings()
        items = list(self.tree.get_children())
        if not items:
            return
        items.sort(key=self._sort_key, reverse=self._sort_rev)
        for idx, iid in enumerate(items):
            self.tree.move(iid, "", idx)

    # ---- 列布局持久化
    def _load_layout(self):
        try:
            with open(ui_conf_path(), "r", encoding="utf-8") as fh:
                data = json.load(fh)
        except Exception:  # noqa: BLE001
            return
        cols = data.get("columns") or {}
        for k, v in cols.items():
            if k in self._col_cfg and isinstance(v, dict):
                w = v.get("width")
                if isinstance(w, int) and w > 20:
                    self._col_cfg[k]["width"] = w
                if "visible" in v:
                    self._col_cfg[k]["visible"] = bool(v["visible"])
        if any(isinstance(v, dict) and isinstance(v.get("width"), int)
               for v in cols.values()):
            self._user_resized = True
        order = [k for k in (data.get("order") or []) if k in self._col_cfg]
        for k in COL_ORDER:
            if k not in order:
                order.append(k)
        self._col_order = order

    def _save_layout(self):
        try:
            path = ui_conf_path()
            os.makedirs(os.path.dirname(path), exist_ok=True)
            data = {
                "columns": {k: {"width": int(self.tree.column(k, "width")),
                                "visible": self._col_cfg[k]["visible"]}
                            for k in COL_ORDER},
                "order": list(self._col_order),
            }
            tmp = path + ".tmp"
            with open(tmp, "w", encoding="utf-8") as fh:
                json.dump(data, fh, ensure_ascii=False, indent=1)
            os.replace(tmp, path)
        except Exception:  # noqa: BLE001
            pass

    def _on_right_click(self, event):
        if self.tree.identify_region(event.x, event.y) == "heading":
            self._show_col_menu(event)
            return
        row = self.tree.identify_row(event.y)
        if not row:
            return
        if row not in self.tree.selection():
            self.tree.selection_set(row)
        self._sync_remove_btn()
        try:
            self.menu.tk_popup(event.x_root, event.y_root)
        finally:
            try:
                self.menu.grab_release()
            except Exception:  # noqa: BLE001
                pass

    def _open_in_explorer(self):
        sel = self.tree.selection()
        if not sel:
            return
        path = os.path.normpath(sel[0])
        try:
            subprocess.Popen(["explorer", "/select,", path])
        except Exception:  # noqa: BLE001
            try:
                os.startfile(os.path.dirname(path))  # noqa: S606
            except Exception:  # noqa: BLE001
                pass

    def _refresh_count(self):
        n = len(self.files)
        self.lbl_files.configure(text="待处理文件（%d）" % n if n else "待处理文件")
        if hasattr(self, "empty_hint"):
            if n:
                self.empty_hint.place_forget()
            else:
                self.empty_hint.place(relx=0.5, rely=0.52, anchor="center")

    # ------------------------------------------------------------ 运行
    def _opts(self) -> Options:
        return Options(keep_print_names=self.cb_print.value,
                       remove_links=self.cb_links.value,
                       backup=self.cb_backup.value,
                       recurse=True)

    def _on_links_toggle(self, on):
        if on and not self._links_confirmed:
            self.lbl_status.configure(
                text="提示：断开外链后，引用外部工作簿的公式会变成固定数值", fg=T.WARN)

    def _start(self):
        if self.running:
            return
        if not self.files:
            self.lbl_status.configure(text="请先添加要处理的文件", fg=T.WARN)
            return
        opt = self._opts()
        if opt.remove_links and not self._links_confirmed:
            ok = messagebox.askyesno(
                "确认断开外部链接",
                "将删除文件中的全部外部链接，效果与 Excel 里手动「断开链接」相同：\n\n"
                "· 引用外链的公式会替换成它此刻的数值，不会出现 #REF!\n"
                "· 这些单元格从此变成固定值，不再随源文件更新\n"
                "· 已勾选备份时，可随时用 .bak 还原\n\n是否继续？",
                parent=self)
            if not ok:
                return
            self._links_confirmed = True

        self.running = True
        self._sync_remove_btn()      # 运行中禁止移除
        self._t0 = time.time()
        self._cur_opt = opt
        self._run_ts = time.strftime("%Y-%m-%d %H:%M")
        self._run_hm = time.strftime("%H:%M")
        self._cur_idx = 0
        self._done_files = 0
        self._total_files = len(self.files)
        self._removed_sum = 0
        self._saved_sum = 0
        self._cur_stage = ""
        self._cur_done = 0
        self._cur_total = 0
        self._cur_name = ""
        self._backups_removed = 0
        self.btn_start.set_enabled(False)
        self.btn_start.set_text("处理中…")
        self.progress.set(0)
        self.lbl_status.configure(text="正在处理…", fg=T.ACCENT)
        for p in self.files:
            if p not in self.results:
                self._set_row(p, "排队中", "wait")
        self.log.header("开始批量处理：%d 个文件" % len(self.files))

        threading.Thread(target=self._worker, args=(list(self.files), opt),
                         daemon=True).start()
        self.after(60, self._poll)
        self.after(200, self._tick)

    def _worker(self, files, opt):
        bg = self.cb_bg.value
        if bg:
            _set_process_low_priority(True)
            _set_thread_low_priority()
        total = len(files)
        for i, p in enumerate(files):
            self.q.put(("begin", (p, i, total), None))

            def cb(stage, done, tot):
                self.q.put(("progress", (stage, done, tot), None))
                if bg:
                    time.sleep(0.002)

            try:
                r = engine.process_file(p, opt, progress=cb)
            except Exception as exc:  # noqa: BLE001
                r = engine.Outcome(path=p, ok=False, message="处理异常：%s" % exc)
            self.q.put(("done", r, None))
            time.sleep(0.01)
        self.q.put(("end", None, None))

    def _poll(self):
        try:
            while True:
                kind, payload, _ = self.q.get_nowait()
                if kind == "begin":
                    p, i, total = payload
                    self._cur_idx = i
                    self._cur_name = p
                    self._cur_stage = "准备中"
                    self._cur_done = 0
                    self._cur_total = 0
                    self._set_row(p, "处理中…", "run")
                    self.tree.see(p)
                elif kind == "progress":
                    self._cur_stage, self._cur_done, self._cur_total = payload
                elif kind == "done":
                    self._on_result(payload)
                elif kind == "end":
                    self._on_finished()
                    return
        except queue.Empty:
            pass
        if self.running:
            self.after(60, self._poll)

    def _tick(self):
        if not self.running:
            return
        elapsed = time.time() - self._t0
        sub = (self._cur_done / self._cur_total) if self._cur_total else 0.0
        frac = (self._done_files + sub) / max(1, self._total_files)
        frac = min(0.99, max(0.01, frac))
        eta = elapsed * (1 - frac) / frac
        self.progress.set(frac)

        if self._cur_total:
            stage = "%s（%s / %s）" % (self._cur_stage or "处理中",
                                      "{:,}".format(self._cur_done),
                                      "{:,}".format(self._cur_total))
        else:
            stage = self._cur_stage or "处理中"
        name = os.path.basename(self._cur_name) if self._cur_name else ""
        self.lbl_status.configure(
            text="(%d/%d) %s · %s" % (self._cur_idx + 1, self._total_files, name, stage),
            fg=T.ACCENT)
        self.lbl_metrics.configure(
            text="已用时 %s · 预计剩余 %s · 已删除 %s 个名称 · 已减少 %s"
                 % (_fmt_dur(elapsed), _fmt_dur(eta),
                    "{:,}".format(self._removed_sum), human_size(self._saved_sum)),
            fg=T.SUBTEXT)
        self.after(250, self._tick)

    def _on_result(self, r):
        self.results[r.path] = r
        self._done_files += 1
        self._removed_sum += r.names_removed
        self._saved_sum += r.saved

        # 处理成功且没有异常时，备份已经没有保留价值，直接删掉
        if r.ok and r.backup_path:
            try:
                os.remove(r.backup_path)
                r.backup_path = ""
                self._backups_removed += 1
            except OSError:
                pass

        history.append(r, self._cur_opt)
        if r.ok:
            if r.output_path:
                txt = "✓ 已转为 %s，删除 %s 个名称" % (
                    os.path.basename(r.output_path), "{:,}".format(r.names_removed))
            elif r.names_removed or r.links_removed:
                txt = "✓ 删除 %s 个名称，减少 %s" % (
                    "{:,}".format(r.names_removed), human_size(r.saved))
            else:
                txt = "✓ 无需处理"
            tag = "ok"
        elif r.skipped:
            txt = "! " + r.message
            tag = "skip"
        else:
            txt = "✕ " + (r.message or "失败")
            tag = "fail"
        r.ts = self._run_ts
        self._set_row(r.path, txt, tag, tm=self._run_hm, dur=r.seconds)
        self.log.outcome(r, self._cur_opt)

    def _on_finished(self):
        self.running = False
        self._sync_remove_btn()
        self.btn_start.set_enabled(True)
        self.btn_start.set_text("开始清理")
        self.progress.set(1.0)
        rs = [self.results[p] for p in self.files if p in self.results]
        ok = sum(1 for r in rs if r.ok)
        removed = sum(r.names_removed for r in rs)
        saved = sum(r.saved for r in rs if not r.output_path)
        n_conv = sum(1 for r in rs if r.output_path)
        secs = time.time() - self._t0
        parts = ["%d/%d 个文件成功" % (ok, len(rs)),
                 "共删除 %s 个名称" % "{:,}".format(removed)]
        if saved:
            parts.append("减少 %s" % human_size(saved))
        parts.append("用时 %s" % _fmt_dur(secs))
        if n_conv:
            parts.append("%d 个 .xlsb 已另存为 .xlsx" % n_conv)
        self.lbl_status.configure(
            text="完成：" + " · ".join(parts),
            fg=T.SUCCESS if ok == len(rs) else T.WARN)
        tail = ""
        if self._backups_removed:
            tail = " · 已自动删除 %d 个备份" % self._backups_removed
        self.lbl_metrics.configure(
            text="已用时 %s · 平均 %s/文件 · 结果已写入历史记录%s"
                 % (_fmt_dur(secs), _fmt_dur(secs / max(1, len(rs))), tail),
            fg=T.SUBTEXT)
        self.log.info("")
        self.log.info("本次汇总：成功 %d / %d，删除名称 %d 个，减少 %s，用时 %.1f 秒"
                      % (ok, len(rs), removed, human_size(saved), secs))
        _set_process_low_priority(False)

    # ------------------------------------------------------------ 详情 / 历史
    def _on_double(self, _):
        sel = self.tree.selection()
        if not sel:
            return
        r = self.results.get(sel[0])
        if r is None:
            messagebox.showinfo("尚未处理", "该文件还没有处理结果。", parent=self)
            return
        self._show_detail(r)

    def _text_window(self, title, w=780, h=580):
        th = self.th
        win = tk.Toplevel(self)
        win.title(title)
        win.configure(bg=T.CARD)
        win.geometry("%dx%d" % (th.s(w), th.s(h)))
        win.transient(self)
        txt = tk.Text(win, wrap="word", bg=T.CARD, fg=T.TEXT, bd=0,
                      highlightthickness=0, font=th.f_body, padx=th.s(20),
                      pady=th.s(16), spacing1=3, spacing3=3)
        sb = ttk.Scrollbar(win, orient="vertical", command=txt.yview,
                           style="Thin.Vertical.TScrollbar")
        txt.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        txt.pack(fill="both", expand=True)
        txt.tag_configure("h", font=th.f_h2, foreground=T.TEXT, spacing3=th.s(8))
        txt.tag_configure("sec", font=th.f_body_b, foreground=T.ACCENT,
                          spacing1=th.s(10), spacing3=th.s(4))
        txt.tag_configure("dim", foreground=T.SUBTEXT)
        txt.tag_configure("ok", foreground=T.SUCCESS)
        txt.tag_configure("warn", foreground=T.WARN)
        return win, txt

    def _show_detail(self, r):
        win, txt = self._text_window("处理详情 - " + os.path.basename(r.path))
        txt.insert("end", os.path.basename(r.path) + "\n", "h")
        txt.insert("end", r.path + "\n", "dim")
        txt.insert("end", "\n")
        txt.insert("end", "格式　　%s\n" % engine.FMT_LABEL.get(r.fmt, r.fmt or "未知"))
        txt.insert("end", "体积　　%s → %s（减少 %s）\n"
                   % (human_size(r.size_before), human_size(r.size_after),
                      human_size(r.saved)))
        txt.insert("end", "名称　　共 %s 个，删除 %s 个，保留 %s 个\n"
                   % ("{:,}".format(r.names_total), "{:,}".format(r.names_removed),
                      "{:,}".format(r.names_kept)))
        if r.links_removed:
            txt.insert("end", "外链　　删除 %d 个，其中 %d 个公式已转为数值\n"
                       % (r.links_removed, r.converted_formulas), "warn")
        txt.insert("end", "操作时间　%s\n" % (r.ts or "-"))
        txt.insert("end", "耗时　　%.1f 秒\n" % r.seconds)
        txt.insert("end", "结果　　%s\n" % (r.message or ("成功" if r.ok else "失败")),
                   "ok" if r.ok else "warn")
        if r.output_path:
            txt.insert("end", "输出　　%s\n" % r.output_path, "ok")
            txt.insert("end", "　　　　原 .xlsb 未改动，可直接使用这个 .xlsx\n", "dim")
        if r.backup_path:
            txt.insert("end", "备份　　%s\n" % r.backup_path, "dim")

        if r.kept_names:
            txt.insert("end", "\n保留的名称（%s 个）\n" % "{:,}".format(r.names_kept), "sec")
            for name, reason in r.kept_names[:60]:
                txt.insert("end", "  · %s" % name)
                txt.insert("end", "  —  %s\n" % reason, "dim")
            if r.names_kept > 60:
                txt.insert("end", "  …… 其余 %s 个略\n"
                           % "{:,}".format(r.names_kept - 60), "dim")
        if r.removed_names:
            txt.insert("end", "\n被删除的名称（共 %s 个，最多显示 3000 条）\n"
                       % "{:,}".format(r.names_removed), "sec")
            txt.insert("end", "  " + "，".join(r.removed_names) + "\n", "dim")
        txt.configure(state="disabled")

    def _open_history(self):
        th = self.th
        win = tk.Toplevel(self)
        win.title("历史记录")
        win.configure(bg=T.BG)
        win.geometry("%dx%d" % (th.s(980), th.s(620)))
        win.minsize(th.s(820), th.s(480))
        win.transient(self)

        head = tk.Frame(win, bg=T.CARD, height=th.s(64))
        head.pack(fill="x")
        head.pack_propagate(False)
        left = tk.Frame(head, bg=T.CARD)
        left.pack(side="left", padx=th.s(20), pady=th.s(12))
        tk.Label(left, text="历史记录", bg=T.CARD, fg=T.TEXT,
                 font=th.f_h2).pack(anchor="w")
        self.hist_info = tk.Label(left, text="", bg=T.CARD, fg=T.SUBTEXT,
                                  font=th.f_small)
        self.hist_info.pack(anchor="w", pady=(th.s(2), 0))

        btns = tk.Frame(head, bg=T.CARD)
        btns.pack(side="right", padx=th.s(18))
        RoundedButton(btns, "打开日志", self._open_log, th, kind="ghost",
                      width=96, height=32, bg=T.CARD).pack(side="right")
        RoundedButton(btns, "刷新", lambda: self._load_history(tree), th,
                      kind="ghost", width=64, height=32,
                      bg=T.CARD).pack(side="right", padx=(0, th.s(8)))
        tk.Frame(win, bg=T.BORDER, height=1).pack(fill="x")

        body = tk.Frame(win, bg=T.BG)
        body.pack(fill="both", expand=True, padx=th.s(16), pady=(th.s(14), th.s(6)))
        card = Card(body, th, padding=1)
        card.pack(fill="both", expand=True)
        wrap = tk.Frame(card.inner, bg=T.CARD)
        wrap.pack(fill="both", expand=True)

        cols = ("name", "ts", "names", "cost", "size", "result")
        tree = ttk.Treeview(wrap, columns=cols, show="headings", style="Files.Treeview")
        for c, t, w, a, st in (("name", "文件名", 240, "w", True),
                               ("ts", "清理时间", 150, "w", False),
                               ("names", "名称（删除/总数）", 140, "e", False),
                               ("cost", "用时", 80, "e", False),
                               ("size", "体积变化", 150, "e", False),
                               ("result", "结果", 70, "w", False)):
            tree.heading(c, text=t, anchor=a)
            tree.column(c, width=th.s(w), anchor=a, stretch=st)
        tree.tag_configure("ok", foreground=T.SUCCESS)
        tree.tag_configure("fail", foreground=T.DANGER)
        tree.tag_configure("skip", foreground=T.WARN)
        sb = ttk.Scrollbar(wrap, orient="vertical", command=tree.yview,
                           style="Thin.Vertical.TScrollbar")
        tree.configure(yscrollcommand=sb.set)
        tree.pack(side="left", fill="both", expand=True, padx=(th.s(8), 0), pady=th.s(4))
        sb.pack(side="right", fill="y", padx=(0, th.s(6)), pady=th.s(6))
        tree.bind("<Double-1>", lambda e: self._show_history_detail(tree))

        tip = tk.Frame(win, bg=T.BG)
        tip.pack(fill="x", padx=th.s(18), pady=(0, th.s(12)))
        tk.Label(tip, text="双击任意一行查看这一次的完整详情",
                 bg=T.BG, fg=T.MUTED, font=th.f_tiny).pack(side="left")
        tk.Label(tip, text="%s %s · 开发者 %s" % (APP_NAME, APP_VER, APP_AUTHOR),
                 bg=T.BG, fg=T.MUTED, font=th.f_tiny).pack(side="right")

        self._load_history(tree)

    def _load_history(self, tree):
        for iid in tree.get_children():
            tree.delete(iid)
        recs = history.load(2000)
        self._hist_recs = recs
        total_removed = 0
        for i, r in enumerate(recs):
            tag = "ok" if r.get("ok") else ("skip" if r.get("skipped") else "fail")
            sb, sa = r.get("size_before", 0), r.get("size_after", 0)
            total_removed += r.get("removed", 0) or 0
            name = r.get("name", "")
            if r.get("output"):
                name += "  →  " + os.path.basename(r["output"])
            tree.insert("", "end", iid=str(i), tags=(tag,), values=(
                name,
                r.get("ts", ""),
                "%s / %s" % ("{:,}".format(r.get("removed", 0)),
                             "{:,}".format(r.get("total", 0))),
                _fmt_sec(r.get("seconds", 0)),
                "%s → %s" % (human_size(sb), human_size(sa)),
                "成功" if r.get("ok") else ("跳过" if r.get("skipped") else "失败")))
        try:
            self.hist_info.configure(
                text="共 %d 次记录 · 累计删除 %s 个名称"
                     % (len(recs), "{:,}".format(total_removed)))
        except tk.TclError:
            pass

    def _show_history_detail(self, tree):
        sel = tree.selection()
        if not sel:
            return
        rec = self._hist_recs[int(sel[0])]
        win, txt = self._text_window("历史详情 - " + rec.get("name", ""), 760, 560)
        txt.insert("end", rec.get("name", "") + "\n", "h")
        txt.insert("end", rec.get("path", "") + "\n", "dim")
        txt.insert("end", "\n")
        txt.insert("end", "时间　　%s\n" % rec.get("ts", ""))
        txt.insert("end", "格式　　%s\n" % engine.FMT_LABEL.get(rec.get("fmt", ""),
                                                             rec.get("fmt", "")))
        txt.insert("end", "体积　　%s → %s\n"
                   % (human_size(rec.get("size_before", 0)),
                      human_size(rec.get("size_after", 0))))
        txt.insert("end", "名称　　共 %s 个，删除 %s 个，保留 %s 个\n"
                   % ("{:,}".format(rec.get("total", 0)),
                      "{:,}".format(rec.get("removed", 0)),
                      "{:,}".format(rec.get("kept", 0))))
        if rec.get("links"):
            txt.insert("end", "外链　　删除 %d 个\n" % rec["links"], "warn")
        txt.insert("end", "耗时　　%.1f 秒\n" % rec.get("seconds", 0))
        txt.insert("end", "结果　　%s\n" % (rec.get("message") or
                                         ("成功" if rec.get("ok") else "失败")),
                   "ok" if rec.get("ok") else "warn")
        if rec.get("output"):
            txt.insert("end", "输出　　%s\n" % rec["output"], "ok")
        if rec.get("backup"):
            txt.insert("end", "备份　　%s\n" % rec["backup"], "dim")
        txt.insert("end", "选项　　保留打印区域=%s · 删除外链=%s\n"
                   % (rec.get("keep_print"), rec.get("remove_links")), "dim")
        kept = rec.get("kept_names") or []
        if kept:
            txt.insert("end", "\n保留的名称\n", "sec")
            for name, reason in kept:
                txt.insert("end", "  · %s" % name)
                txt.insert("end", "  —  %s\n" % reason, "dim")
        sample = rec.get("removed_sample") or []
        if sample:
            txt.insert("end", "\n被删除的名称（样例）\n", "sec")
            txt.insert("end", "  " + "，".join(sample) + "\n", "dim")
        txt.configure(state="disabled")

    def _open_log(self):
        try:
            os.startfile(log_dir())  # noqa: S606
        except Exception:  # noqa: BLE001
            messagebox.showinfo("日志位置", log_dir(), parent=self)

    def _on_close(self):
        self._closing = True
        if self.running:
            if not messagebox.askyesno("正在处理", "任务尚未完成，确定要退出吗？", parent=self):
                return
        try:
            self._save_layout()
        except Exception:  # noqa: BLE001
            pass
        try:
            self.drop.close()
        except Exception:  # noqa: BLE001
            pass
        _set_process_low_priority(False)
        self.destroy()


def main():
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:  # noqa: BLE001
        pass
    app = App()
    app.mainloop()
