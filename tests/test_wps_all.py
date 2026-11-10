# -*- coding: utf-8 -*-
# Author: Opera-Jin
"""WPS 全格式回归：.et(新版) / .ett / .xlsx / .xls(老版)，逐格比对 + 专项检查。"""
import os
import shutil
import sys
import tempfile

sys.path.insert(0, "src")
sys.path.insert(0, "tests")
import win32com.client as win32
from core import engine
from core.engine import Options
from validate_wps import snap

SAMPLES = [
    ("WPS_新版.et", "WPS 新版 .et（OOXML）"),
    ("WPS_模板.ett", "WPS 模板 .ett"),
    ("WPS_兼容.xlsx", "WPS 兼容 .xlsx"),
    ("WPS_老版.xls", "WPS 老版 .xls（BIFF8）"),
]


def extra(app, path):
    """数组公式 / 数据验证 / 条件格式 专项检查"""
    wb = app.Workbooks.Open(os.path.abspath(path))
    out = {}
    try:
        ws = wb.Worksheets("报表")
        out["数组公式"] = bool(ws.Range("A1").HasArray)
        try:
            out["数据验证"] = str(ws.Range("B1").Validation.Formula1)
        except Exception:
            out["数据验证"] = ""
        try:
            out["条件格式"] = ws.Range("A1:A10").FormatConditions.Count
        except Exception:
            out["条件格式"] = -1
        ws2 = wb.Worksheets("数据")
        out["合并区"] = str(ws2.Range("E1").MergeArea.Address) if ws2.Range("E1").MergeCells else ""
        out["打印区域"] = str(ws2.PageSetup.PrintArea)
    finally:
        wb.Close(False)
    return out


def main():
    base = os.path.abspath("tests/samples/wps")
    tmp = tempfile.mkdtemp(prefix="wps_test_")
    app = win32.Dispatch("Ket.Application")
    try:
        app.Visible = False
    except Exception:
        pass
    app.DisplayAlerts = False
    try:
        app.AskToUpdateLinks = False
    except Exception:
        pass

    bad = 0
    for name, label in SAMPLES:
        src = os.path.join(base, name)
        work = os.path.join(tmp, name)
        shutil.copy2(src, work)
        r = engine.process_file(work, Options(remove_links=False, backup=True))
        note = ""
        diff = -1
        if r.ok:
            try:
                a = snap(app, src)
                b = snap(app, work)
                diff = 0
                for (na, ca), (nb, cb) in zip(a["sheets"], b["sheets"]):
                    for ra, rb in zip(ca, cb):
                        for xa, xb in zip(ra, rb):
                            if xa != xb:
                                diff += 1
                ea, eb = extra(app, src), extra(app, work)
                note = " | ".join("%s:%s→%s" % (k, ea[k], eb[k])
                                  for k in ea if ea[k] != eb[k])
                if ea != eb:
                    bad += 1
                    note = "专项不一致 " + note
            except Exception as exc:  # noqa: BLE001
                note = "校验异常 %s" % exc
        print("%-16s %-24s ok=%-5s 名称 %d→%d  差异=%s  %s"
              % (name, label, r.ok, r.names_total, r.names_kept, diff, note))
        if not r.ok or diff != 0:
            bad += 1

    try:
        app.Quit()
    except Exception:
        pass
    shutil.rmtree(tmp, ignore_errors=True)
    print("\n结果：", "全部通过" if bad == 0 else "有 %d 项异常" % bad)
    return 0 if bad == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
