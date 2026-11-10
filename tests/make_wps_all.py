# -*- coding: utf-8 -*-
# Author: Opera-Jin
"""用 WPS 生成各种格式的复杂样本：.et(新版) / .ett / .xlsx / .xls / .et(老版)"""
import os
import win32com.client as win32

OUT = os.path.abspath("tests/samples/wps")
os.makedirs(OUT, exist_ok=True)

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

wb = app.Workbooks.Add()
ws = wb.Worksheets(1)
ws.Name = "数据"
ws2 = wb.Worksheets.Add()
ws2.Name = "报表"

for r in range(1, 31):
    ws.Cells(r, 1).Value = r
    ws.Cells(r, 2).Value = "项目%d" % r
    ws.Cells(r, 3).Value = r * 3.7

# 名称
wb.Names.Add("数据区", "=数据!$A$1:$C$30")
wb.Names.Add("下拉列表", "=数据!$B$1:$B$10")
wb.Names.Add("条件名", "=数据!$C$1")
wb.Names.Add("报表值", "=报表!$A$1")

# 数组公式 + 复杂函数
ws2.Range("A1").FormulaArray = "=SUM(数据区*2)"
ws2.Range("A2").Formula = "=VLOOKUP(5,数据区,2,FALSE)"
ws2.Range("A3").Formula = '=IF(条件名>10,"大","小")'
ws2.Range("A4").Formula = "=SUMPRODUCT((数据!A1:A30>5)*(数据!C1:C30))"
ws2.Range("A5").Formula = "=报表值+1"

# 数据验证 + 条件格式
try:
    ws2.Range("B1:B5").Validation.Delete()
    ws2.Range("B1:B5").Validation.Add(3, 1, 1, "=下拉列表")
except Exception as e:
    print("DV:", e)
try:
    ws2.Range("A1:A10").FormatConditions.Delete()
    fc = ws2.Range("A1:A10").FormatConditions.Add(2, 1, "=A1>条件名")
    fc.Interior.Color = 0x99CCFF
except Exception as e:
    print("CF:", e)

# 合并 + 打印区域 + 格式
ws.Range("E1:F2").Merge()
ws.Range("E1").Value = "合并"
ws.Range("A1:C1").Font.Bold = True
ws.Range("A1:C1").Interior.Color = 0x4A7C59
ws.Range("A1:C1").Font.Color = 0xFFFFFF
ws.Columns(1).ColumnWidth = 22
ws.PageSetup.PrintArea = "$A$1:$E$30"
ws.PageSetup.PrintTitleRows = "$1:$1"

# 大量垃圾名称
for i in range(200):
    wb.Names.Add("冗余%d" % i, "=数据!$A$1:$B$2")
# 失效名称
t = wb.Worksheets.Add()
t.Name = "临时X"
t.Range("A1").Value = 1
for i in range(20):
    wb.Names.Add("坏名字%d" % i, "=临时X!$A$1:$B$2")
t.Delete()

results = []
for name, fmt in (("WPS_新版.et", None), ("WPS_模板.ett", None),
                  ("WPS_兼容.xlsx", None), ("WPS_老版.xls", 56)):
    path = os.path.join(OUT, name)
    try:
        if fmt:
            wb.SaveAs(path, fmt)
        else:
            wb.SaveAs(path)
        results.append((name, os.path.getsize(path), open(path, "rb").read(4).hex()))
    except Exception as e:
        results.append((name, -1, str(e)[:60]))

wb.Close(False)
app.Quit()
for r in results:
    print(r)
