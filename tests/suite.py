# -*- coding: utf-8 -*-
# Author: Opera-Jin
"""全量回归：各类文件 + 边界情况。"""
import os, shutil, sys, tempfile, zipfile
sys.path.insert(0, "src")
sys.path.insert(0, "tests")
from core import engine
from core.engine import Options
from core.util import human_size

import win32com.client as win32
from validate_wps import snap

SAMPLES = [
    ("tests/samples/样本_xlsx.xlsx", "xlsx 多表+外链+打印区"),
    ("tests/samples/极限_xlsx.xlsx", "xlsx 数组/验证/条件格式"),
    ("tests/samples/样本_xls.xls", "xls 多表+外链"),
    ("tests/samples/极限_xls.xls", "xls 数组/验证/条件格式"),
    ("tests/samples/新版_wps.et", "et (OOXML)"),
    ("tests/samples/样本_wps.et", "et (BIFF)"),
]

def make_extra(tmp):
    # 空工作簿
    from openpyxl import Workbook
    p = os.path.join(tmp, "空表.xlsx")
    wb = Workbook(); wb.active["A1"] = 1; wb.save(p)
    SAMPLES.append((p, "xlsx 无名称"))
    # 非 Excel
    p2 = os.path.join(tmp, "说明.txt")
    open(p2, "w").write("hello")
    # 损坏
    p3 = os.path.join(tmp, "损坏.xlsx")
    with open("tests/samples/样本_xlsx.xlsx", "rb") as f:
        data = f.read()
    open(p3, "wb").write(data[: len(data) // 3])

def main():
    tmp = tempfile.mkdtemp(prefix="clean_suite_")
    make_extra(tmp)
    app = win32.Dispatch("Ket.Application")
    try: app.Visible = False
    except Exception: pass
    app.DisplayAlerts = False
    try: app.AskToUpdateLinks = False
    except Exception: pass

    opt = Options(keep_print_names=True, remove_links=False, backup=True)
    rows = []
    for src, label in SAMPLES:
        work = os.path.join(tmp, os.path.basename(src))
        shutil.copy2(src, work)
        r = engine.process_file(work, opt)
        note = ""
        if r.ok and r.fmt in ("xlsx", "xlsm", "xltx", "xltm", "xls", "et"):
            try:
                a = snap(app, src); b = snap(app, work)
                diff = 0
                for (na, ca), (nb, cb) in zip(a["sheets"], b["sheets"]):
                    for ra, rb in zip(ca, cb):
                        for xa, xb in zip(ra, rb):
                            if xa != xb:
                                diff += 1
                note = "单元格差异 %d" % diff
                if diff:
                    r.ok = False
            except Exception as e:
                note = "校验失败: %s" % e
        rows.append((label, r.ok, r.names_total, r.names_removed, r.names_kept,
                     r.size_before, r.size_after, r.message, note))
        print('  .. %-26s %s %s' % (label, 'OK' if r.ok else 'FAIL', note), flush=True)
    # 边界
    for name in ("说明.txt", "损坏.xlsx"):
        p = os.path.join(tmp, name)
        r = engine.process_file(p, opt)
        rows.append((name, r.ok, 0, 0, 0, r.size_before, r.size_after, r.message, ""))
    try: app.Quit()
    except Exception: pass

    print("%-26s %-6s %8s %8s %8s %-22s %s" % ("样本", "结果", "名称", "删除", "保留", "体积", "备注"))
    print("-" * 118)
    bad = 0
    for label, ok, tot, rem, kept, sb, sa, msg, note in rows:
        print("%-26s %-6s %8d %8d %8d %-22s %s %s" % (
            label, "通过" if ok else "失败", tot, rem, kept,
            "%s→%s" % (human_size(sb), human_size(sa)), note, "" if ok else msg[:40]))
        if not ok:
            bad += 1
    print("-" * 118)
    print("失败项:", bad)
    return 0 if bad == 0 else 1

if __name__ == "__main__":
    sys.exit(main())
