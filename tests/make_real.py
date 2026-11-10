# -*- coding: utf-8 -*-
# Author: Opera-Jin
"""用 WPS COM 生成真实场景样本：大量名称、外链、公式、格式。"""
import os, sys, time
import win32com.client as win32

OUT = os.path.abspath("tests/samples")
os.makedirs(OUT, exist_ok=True)

app = win32.Dispatch("Ket.Application")
try:
    app.Visible = False
except Exception:
    pass
app.DisplayAlerts = False
app.AskToUpdateLinks = False

# ---- 源工作簿（供外链引用）
src_path = os.path.join(OUT, "外部源.xlsx")
wb_src = app.Workbooks.Add()
ws_src = wb_src.Worksheets(1)
for r in range(1, 11):
    ws_src.Cells(r, 1).Value = r * 100
wb_src.SaveAs(src_path)
wb_src.Close(False)

# ---- 主工作簿
wb = app.Workbooks.Add()
ws = wb.Worksheets(1)
ws.Name = "数据"
for r in range(1, 51):
    ws.Cells(r, 1).Value = "行%d" % r
    ws.Cells(r, 2).Value = r * 1.5
    ws.Cells(r, 3).Formula = "=B%d*2" % r
ws.Range("A1:C1").Font.Bold = True
ws.Range("A1:C1").Interior.Color = 0x4A7C59
ws.Range("A1:C1").Font.Color = 0xFFFFFF
ws.Columns(1).ColumnWidth = 22
ws.Range("G1:H2").Merge()
ws.Range("G1").Value = "合并单元格"
ws.Range("A1:E50").Select()
app.ActiveWindow.DisplayGridlines = True

wb.Names.Add("关键数据", "=数据!$B$1")
wb.Names.Add("有效区域", "=数据!$B$1:$B$50")
ws.Range("E1").Formula = "=SUM(B1:B50)"
ws.Range("E2").Formula = "=关键数据*2"
ws.Range("E3").Formula = "=SUM(有效区域)"
ws.Range("E4").Formula = "=A1&\"测试\""
# 外链公式
ws.Range("E5").Formula = "=[外部源.xlsx]Sheet1!$A$1"
ws.Range("E6").Formula = "=SUM([外部源.xlsx]Sheet1!$A$1:$A$10)"

# 打印区域（会生成 _xlnm.Print_Area）
ws.PageSetup.PrintArea = "$A$1:$E$50"
ws.PageSetup.PrintTitleRows = "$1:$1"

# 复制工作表 30 次，模拟“反复复制产生大量名称”
for i in range(30):
    ws.Copy(None, wb.Worksheets(wb.Worksheets.Count))

# 垃圾名称 300 个
for i in range(300):
    wb.Names.Add("垃圾名称%d" % i, "=数据!$A$1:$B$3")

# 失效名称：先建表再删表 → #REF!
tmp = wb.Worksheets.Add()
tmp.Name = "临时表"
tmp.Range("A1").Value = 1
for i in range(40):
    wb.Names.Add("失效名称%d" % i, "=临时表!$A$1:$B$3")
tmp.Delete()

xlsx = os.path.join(OUT, "样本_xlsx.xlsx")
wb.SaveAs(xlsx)
xls = os.path.join(OUT, "样本_xls.xls")
try:
    wb.SaveAs(xls, 56)
except Exception as e:
    print("save xls fail", e)
et = os.path.join(OUT, "样本_wps.et")
try:
    wb.SaveAs(et)
except Exception as e:
    print("save et fail", e)
wb.Close(False)
app.Quit()

for p in (xlsx, xls, et, src_path):
    if os.path.exists(p):
        with open(p, "rb") as f:
            print(os.path.basename(p), os.path.getsize(p), f.read(4).hex())
