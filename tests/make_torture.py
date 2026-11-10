# -*- coding: utf-8 -*-
# Author: Opera-Jin
"""极限样本：数组公式、数据验证引用名称、条件格式、复杂函数。"""
import os
import win32com.client as win32

OUT = os.path.abspath("tests/samples")
app = win32.Dispatch("Ket.Application")
try: app.Visible = False
except Exception: pass
app.DisplayAlerts = False
app.AskToUpdateLinks = False

wb = app.Workbooks.Add()
ws = wb.Worksheets(1); ws.Name = "数据"
ws2 = wb.Worksheets.Add(); ws2.Name = "报表"

for r in range(1, 31):
    ws.Cells(r, 1).Value = r
    ws.Cells(r, 2).Value = "项目%d" % r
    ws.Cells(r, 3).Value = r * 3.7

wb.Names.Add("数据区", "=数据!$A$1:$C$30")
wb.Names.Add("下拉列表", "=数据!$B$1:$B$10")
wb.Names.Add("条件名", "=数据!$C$1")
wb.Names.Add("报表值", "=报表!$A$1")

# 数组公式
ws2.Range("A1").FormulaArray = "=SUM(数据区*2)"
# 复杂函数
ws2.Range("A2").Formula = "=VLOOKUP(5,数据区,2,FALSE)"
ws2.Range("A3").Formula = "=IF(条件名>10,\"大\"&\"值\",\"小\")"
ws2.Range("A4").Formula = "=SUMPRODUCT((数据!A1:A30>5)*(数据!C1:C30))"
ws2.Range("A5").Formula = "=INDEX(数据区,2,3)+COUNTIF(数据!B1:B30,\"项目*\")"
ws2.Range("A6").Formula = "=报表值+1"
ws2.Range("A7").Formula = "=SUM(数据!A1:A5,报表!A1:A2)"
ws2.Range("A8").Formula = "=\"文本\"&CHAR(34)&数据!B1"
# 数据验证：引用名称
try:
    ws2.Range("B1:B5").Validation.Delete()
    ws2.Range("B1:B5").Validation.Add(3, 1, 1, "=下拉列表")
except Exception as e:
    print("DV fail", e)
# 条件格式：公式引用名称
try:
    ws2.Range("A1:A10").FormatConditions.Delete()
    fc = ws2.Range("A1:A10").FormatConditions.Add(2, 1, "=A1>条件名")
    fc.Interior.Color = 0x99CCFF
except Exception as e:
    print("CF fail", e)

# 大量垃圾名称
for i in range(150):
    wb.Names.Add("冗余%d" % i, "=数据!$A$1:$B$2")
# 失效名称
t = wb.Worksheets.Add(); t.Name = "临时X"; t.Range("A1").Value = 1
for i in range(20):
    wb.Names.Add("坏名字%d" % i, "=临时X!$A$1:$B$2")
t.Delete()

for path, fmt in ((os.path.join(OUT, "极限_xlsx.xlsx"), None),
                  (os.path.join(OUT, "极限_xls.xls"), 56)):
    if fmt: wb.SaveAs(path, fmt)
    else: wb.SaveAs(path)
wb.Close(False)
app.Quit()
for f in ("极限_xlsx.xlsx", "极限_xls.xls"):
    p = os.path.join(OUT, f)
    print(f, os.path.getsize(p), open(p, "rb").read(4).hex())
