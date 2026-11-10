# -*- coding: utf-8 -*-
# Author: Opera-Jin
"""用真实 Excel 打开原始文件与清理后文件，逐项比对值/公式/格式/名称。"""
import os, sys, shutil
import win32com.client as win32

def open_wb(app, path):
    return app.Workbooks.Open(os.path.abspath(path), UpdateLinks=0, ReadOnly=True)

def snapshot(app, path):
    wb = open_wb(app, path)
    data = {"sheets": [], "names": [], "merged": {}, "widths": {}, "fills": {}}
    try:
        app.CalculateFullRebuild()
        for ws in wb.Worksheets:
            rows, cols = ws.UsedRange.Rows.Count, ws.UsedRange.Columns.Count
            cells = []
            for r in range(1, min(rows, 60) + 1):
                row = []
                for c in range(1, min(cols, 10) + 1):
                    cell = ws.Cells(r, c)
                    row.append((cell.Formula, cell.Text))
                cells.append(tuple(row))
            data["sheets"].append((ws.Name, tuple(cells)))
            data["merged"][ws.Name] = tuple(
                str(a.Address) for a in ws.Range("A1:Z60").MergeArea if a.MergeCells
            ) if False else tuple(sorted(str(m.Address) for m in ws.UsedRange.MergeArea)) if ws.UsedRange.MergeCells else ()
            data["widths"][ws.Name] = tuple(round(ws.Columns(c).ColumnWidth, 2) for c in range(1, 8))
            data["fills"][ws.Name] = tuple(
                int(ws.Cells(1, c).Interior.Color or 0) for c in range(1, 5)
            )
        for i in range(1, wb.Names.Count + 1):
            n = wb.Names(i)
            data["names"].append((n.Name, n.RefersTo))
    finally:
        wb.Close(False)
    return data

def main():
    orig, cleaned = sys.argv[1], sys.argv[2]
    app = win32.gencache.EnsureDispatch("Excel.Application")
    app.Visible = False
    app.DisplayAlerts = False
    app.AskToUpdateLinks = False
    try:
        a = snapshot(app, orig)
        b = snapshot(app, cleaned)
    finally:
        app.Quit()
    ok = True
    if len(a["sheets"]) != len(b["sheets"]):
        print("FAIL 工作表数量不同", len(a["sheets"]), len(b["sheets"])); ok = False
    for (na, ca), (nb, cb) in zip(a["sheets"], b["sheets"]):
        if na != nb:
            print("FAIL 表名不同", na, nb); ok = False
        if ca != cb:
            ok = False
            print("FAIL 单元格内容不同:", na)
            for i, (ra, rb) in enumerate(zip(ca, cb)):
                if ra != rb:
                    for j, (xa, xb) in enumerate(zip(ra, rb)):
                        if xa != xb:
                            print("   行%d 列%d: %r != %r" % (i+1, j+1, xa, xb))
                    break
    for key in ("merged", "widths", "fills"):
        if a[key] != b[key]:
            print("FAIL 格式不同:", key, a[key], "!=", b[key]); ok = False
    print("原始名称数:", len(a["names"]), " 清理后:", len(b["names"]))
    kept_a = {n for n, _ in a["names"]}
    kept_b = {n for n, _ in b["names"]}
    print("被删除的名称样例:", sorted(kept_a - kept_b)[:3])
    print("保留的名称:", sorted(kept_b))
    if not kept_b.issubset(kept_a):
        print("FAIL 出现了原文件没有的名称"); ok = False
    print("=== 校验", "通过 ===" if ok else "未通过 ===")
    return 0 if ok else 1

if __name__ == "__main__":
    sys.exit(main())
