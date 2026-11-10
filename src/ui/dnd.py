# -*- coding: utf-8 -*-
# Author: Opera-Jin
"""Windows 原生文件拖放（WM_DROPFILES），纯 ctypes，无第三方依赖。

两条必须遵守的约束（都是踩过的坑）：

1. 窗口过程里**绝不能**调用 Tk/Tcl 的 API（after、createcommand 等）。
   WNDPROC 是 Tcl 的 notifier 在消息循环中回调进来的，此时重入 Tcl
   解释器会静默失败 —— 表现就是「文件拖进去毫无反应、也不报错」。
   路径必须先入队，由主循环的 after 轮询取走。
2. LRESULT / LONG_PTR 在 64 位下是 8 字节，返回类型必须用 c_ssize_t。
   用 c_int 会截断高 32 位，导致部分消息处理异常。
"""
from __future__ import annotations

import ctypes
import queue
from ctypes import wintypes

WM_DROPFILES = 0x0233
GWL_WNDPROC = -4

LRESULT = ctypes.c_ssize_t
WNDPROC = ctypes.WINFUNCTYPE(
    LRESULT,
    wintypes.HWND,
    ctypes.c_uint,
    wintypes.WPARAM,
    wintypes.LPARAM,
)


class DropTarget:
    """把窗口变成文件拖放目标（整个窗口任意位置都可以放下文件）。"""

    def __init__(self, root, on_files, poll_ms: int = 120):
        self.root = root
        self.on_files = on_files
        self.poll_ms = poll_ms
        self.queue: "queue.Queue[list]" = queue.Queue()
        self.ok = False
        self.error = ""
        self.hwnd = 0
        self._old = 0
        self._proc = None
        self._closed = False
        self._user32 = None
        self._shell32 = None
        try:
            self._install()
            self.ok = True
        except Exception as exc:  # noqa: BLE001
            self.error = "%s: %s" % (type(exc).__name__, exc)
        self._pump()

    # ---------------------------------------------------------------- 安装
    def _install(self):
        u = ctypes.WinDLL("user32", use_last_error=True)
        sh = ctypes.WinDLL("shell32", use_last_error=True)
        self._user32, self._shell32 = u, sh

        self.root.update_idletasks()
        inner = self.root.winfo_id()
        top = u.GetParent(inner) or inner
        self.hwnd = top

        # 只注册顶层窗口。子窗口若也注册，系统会把 WM_DROPFILES 投递给
        # 未子类化的 Tk 子窗口，消息直接丢失。
        sh.DragAcceptFiles.argtypes = [wintypes.HWND, wintypes.BOOL]
        sh.DragAcceptFiles.restype = None
        sh.DragAcceptFiles(wintypes.HWND(top), True)

        sh.DragQueryFileW.argtypes = [wintypes.HANDLE, wintypes.UINT,
                                      wintypes.LPWSTR, wintypes.UINT]
        sh.DragQueryFileW.restype = wintypes.UINT
        sh.DragFinish.argtypes = [wintypes.HANDLE]
        sh.DragFinish.restype = None

        setter = getattr(u, "SetWindowLongPtrW", None) or u.SetWindowLongW
        setter.restype = LRESULT
        setter.argtypes = [wintypes.HWND, ctypes.c_int, LRESULT]

        u.CallWindowProcW.restype = LRESULT
        u.CallWindowProcW.argtypes = [LRESULT, wintypes.HWND, ctypes.c_uint,
                                      wintypes.WPARAM, wintypes.LPARAM]

        self._proc = WNDPROC(self._wndproc)
        addr = ctypes.cast(self._proc, ctypes.c_void_p).value
        old = setter(wintypes.HWND(top), GWL_WNDPROC, addr)
        if not old:
            raise OSError("窗口过程挂接失败（错误码 %d）" % ctypes.get_last_error())
        self._old = old

    # ------------------------------------------------------------ 窗口过程
    def _wndproc(self, hwnd, msg, wparam, lparam):
        try:
            if msg == WM_DROPFILES:
                paths = []
                try:
                    paths = self._query(wparam)
                except Exception:  # noqa: BLE001
                    paths = []
                finally:
                    try:
                        self._shell32.DragFinish(wparam)
                    except Exception:  # noqa: BLE001
                        pass
                if paths:
                    # 只入队，绝不在这里碰 Tk
                    self.queue.put(paths)
                return 0
            if self._old:
                return self._user32.CallWindowProcW(self._old, hwnd, msg,
                                                    wparam, lparam)
            return 0
        except Exception:  # noqa: BLE001
            return 0

    def _query(self, hdrop):
        n = self._shell32.DragQueryFileW(hdrop, 0xFFFFFFFF, None, 0)
        out = []
        for i in range(n):
            buf = ctypes.create_unicode_buffer(32768)
            if self._shell32.DragQueryFileW(hdrop, i, buf, len(buf)):
                out.append(buf.value)
        return out

    # -------------------------------------------------------------- 取结果
    def _pump(self):
        if self._closed:
            return
        try:
            while True:
                paths = self.queue.get_nowait()
                try:
                    self.on_files(paths)
                except Exception:  # noqa: BLE001
                    pass
        except queue.Empty:
            pass
        try:
            self.root.after(self.poll_ms, self._pump)
        except Exception:  # noqa: BLE001
            pass

    def close(self):
        self._closed = True
        try:
            if self.hwnd:
                self._shell32.DragAcceptFiles(wintypes.HWND(self.hwnd), False)
        except Exception:  # noqa: BLE001
            pass
        # 恢复原窗口过程：否则 Tk 销毁窗口时可能回调到已释放的函数
        try:
            if self.hwnd and self._old and self._user32:
                u = self._user32
                setter = getattr(u, "SetWindowLongPtrW", None) or u.SetWindowLongW
                setter.restype = LRESULT
                setter.argtypes = [wintypes.HWND, ctypes.c_int, LRESULT]
                setter(wintypes.HWND(self.hwnd), GWL_WNDPROC, self._old)
        except Exception:  # noqa: BLE001
            pass
