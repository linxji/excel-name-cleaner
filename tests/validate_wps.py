# -*- coding: utf-8 -*-
# Author: Opera-Jin
"""用 WPS 打开原始/清理后文件，逐项比对数据、公式、格式、名称。"""
import os, sys, re
import win32com.client as win32


def snap(app, path):
    wb = app.Workbooks.Open(os.path.abspath(path))
    out = {"sheets": [], "names": [], "merged": {}, "width": {}, "color": {}, "fonts": {}}
    try:
        try:
            app.Calculate()
        except Exception:
            pass
        for ws in wb.Worksheets:
            ur = ws.UsedRange
            rows = min(ur.Rows.Count, 60)
            cols = min(ur.Columns.Count, 8)
            cells = []
            for r in range(1, rows + 1):
                row = []
                for c in range(1, cols + 1):
                    cell = ws.Cells(r, c)
                    row.append((str(cell.Formula), str(cell.Text)))
                cells.append(tuple(row))
            out["sheets"].append((ws.Name, tuple(cells)))
            out["merged"][ws.Name] = tuple(
                str(ws.Cells(1, 1).MergeArea.Address) if ws.Cells(1, 1).MergeCells else ""
            )
            out["width"][ws.Name] = tuple(round(float(ws.Columns(c).ColumnWidth), 2) for c in range(1, 6))
            out["color"][ws.Name] = tuple(int(ws.Cells(1, c).Interior.Color or 0) for c in range(1, 4))
            out["fonts"][ws.Name] = tuple(bool(ws.Cells(1, c).Font.Bold) for c in range(1, 4))
        for i in range(1, wb.Names.Count + 1):
            n = wb.Names(i)
            out["names"].append((str(n.Name), str(n.RefersTo)))
    finally:
        wb.Close(False)
    return out


def main():
    orig, cleaned = sys.argv[1], sys.argv[2]
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
    try:
        a = snap(app, orig)
        b = snap(app, cleaned)
    finally:
        try:
            app.Quit()
        except Exception:
            pass

    ok = True
    if len(a["sheets"]) != len(b["sheets"]):
        print("FAIL 工作表数不同", len(a["sheets"]), len(b["sheets"])); ok = False
    diffs = 0
    for (na, ca), (nb, cb) in zip(a["sheets"], b["sheets"]):
        if na != nb:
            print("FAIL 表名", na, nb); ok = False
        for i, (ra, rb) in enumerate(zip(ca, cb)):
            for j, (xa, xb) in enumerate(zip(ra, rb)):
                if xa != xb:
                    diffs += 1
                    if diffs <= 8:
                        print("  差异 %s R%dC%d: %r -> %r" % (na, i + 1, j + 1, xa, xb))
    for key in ("merged", "width", "color", "fonts"):
        if a[key] != b[key]:
            for k in a[key]:
                if a[key][k] != b[key].get(k):
                    print("FAIL 格式不同 %s %s: %r -> %r" % (key, k, a[key][k], b[key].get(k)))
            ok = False
    print("名称数 %d -> %d" % (len(a["names"]), len(b["names"])))
    na = {n for n, _ in a["names"]}
    nb = {n for n, _ in b["names"]}
    print("保留:", sorted(nb))
    if not nb.issubset(na):
        print("FAIL 出现新名称"); ok = False
    print("单元格差异数:", diffs)
    print("=== 校验", "通过 ===" if ok else "未通过 ===")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
