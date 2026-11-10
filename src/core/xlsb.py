# -*- coding: utf-8 -*-
# Author: Opera-Jin
"""把 .xlsb 转成 .xlsx —— 借用本机 Excel / WPS 的另存能力（与旧版 VBA 思路一致）。

.xlsb 内部是二进制零件，纯 Python 解析风险极高；而用户的机器上只要能打开
.xlsb，就一定有 Excel（或新版 WPS）。所以这里只做一件事：
    用 COM 打开 → 另存为 .xlsx → 关掉。
之后的清理工作交给已经验证过的 OOXML 引擎，全程不碰原 .xlsb 文件。

注意：本模块是可选能力，任何一步失败都只返回错误信息，绝不半途写文件。
"""
from __future__ import annotations

import os

# xlOpenXMLWorkbook = 51
XL_XLSX = 51

# 依次尝试：微软 Excel → WPS 表格
PROGIDS = ("Excel.Application", "Ket.Application", "KET.Application")


class ConvertError(Exception):
    pass


def _com_init():
    try:
        import pythoncom
        pythoncom.CoInitialize()
        return pythoncom
    except Exception:  # noqa: BLE001
        return None


def available() -> bool:
    """探测本机有没有可用的 Excel / WPS（只尝试创建 COM 对象）。"""
    try:
        import win32com.client  # noqa: F401
    except Exception:  # noqa: BLE001
        return False
    com = _com_init()
    try:
        for prog in PROGIDS:
            app = None
            try:
                import win32com.client as win32
                app = win32.DispatchEx(prog)
                try:
                    app.Quit()
                except Exception:  # noqa: BLE001
                    pass
                return True
            except Exception:  # noqa: BLE001
                continue
        return False
    finally:
        if com is not None:
            try:
                com.CoUninitialize()
            except Exception:  # noqa: BLE001
                pass


def convert(src: str, dst: str, progress=None) -> str:
    """把 src(.xlsb) 另存为 dst(.xlsx)。成功返回空串，失败返回原因。"""
    if not os.path.exists(src):
        return "源文件不存在"
    try:
        import win32com.client as win32
    except Exception:  # noqa: BLE001
        return "缺少 Office 自动化组件"

    com = _com_init()
    last_err = ""
    try:
        for prog in PROGIDS:
            app = None
            wb = None
            try:
                app = win32.DispatchEx(prog)
            except Exception as exc:  # noqa: BLE001
                last_err = str(exc)
                continue
            try:
                try:
                    app.Visible = False
                except Exception:  # noqa: BLE001
                    pass
                try:
                    app.DisplayAlerts = False
                except Exception:  # noqa: BLE001
                    pass
                try:
                    app.AskToUpdateLinks = False
                except Exception:  # noqa: BLE001
                    pass
                try:
                    app.ScreenUpdating = False
                except Exception:  # noqa: BLE001
                    pass

                if progress:
                    progress("转换 xlsb", 0, 0)
                wb = app.Workbooks.Open(os.path.abspath(src), 0, True)
                if progress:
                    progress("另存为 xlsx", 0, 0)
                if os.path.exists(dst):
                    os.remove(dst)
                wb.SaveAs(os.path.abspath(dst), XL_XLSX)
                return ""
            except Exception as exc:  # noqa: BLE001
                last_err = str(exc)
                continue
            finally:
                try:
                    if wb is not None:
                        wb.Close(False)
                except Exception:  # noqa: BLE001
                    pass
                try:
                    if app is not None:
                        app.Quit()
                except Exception:  # noqa: BLE001
                    pass
        if not last_err:
            return "本机没有可用的 Excel / WPS，无法转换 .xlsb。请先用 Excel 打开并另存为 .xlsx"
        return "无法用 Excel / WPS 转换 .xlsb：%s" % last_err[:120]
    finally:
        if com is not None:
            try:
                com.CoUninitialize()
            except Exception:  # noqa: BLE001
                pass
