# -*- coding: utf-8 -*-
# Author: Opera-Jin
"""生成测试样本：大量名称 + 公式引用 + 格式 + 打印区域。"""
import os, sys, random
from openpyxl import Workbook
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

def build(path):
    wb = Workbook()
    ws = wb.active
    ws.title = "数据"
    for r in range(1, 51):
        ws.cell(r, 1, "行%d" % r)
        ws.cell(r, 2, r * 1.5)
        ws.cell(r, 3, "=B%d*2" % r)
    ws["E1"] = "=SUM(B1:B50)"
    ws["E2"] = "=关键数据*2"          # 引用名称
    ws["E3"] = "=SUM(有效区域)"        # 引用名称
    ws["E4"] = "=A1&\"测试\""
    # 样式
    for c in range(1, 4):
        cell = ws.cell(1, c)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="4472C4")
        cell.alignment = Alignment(horizontal="center")
    ws.column_dimensions["A"].width = 22
    ws.merge_cells("G1:H2")
    ws["G1"] = "合并单元格"
    ws.print_area = "A1:E50"

    ws2 = wb.create_sheet("第二表")
    ws2["A1"] = "=数据!B1+1"

    # 有效名称
    wb.defined_names.add(DefinedName("关键数据", attr_text="数据!$B$1"))
    wb.defined_names.add(DefinedName("有效区域", attr_text="数据!$B$1:$B$50"))
    # 大量垃圾名称（模拟反复复制粘贴产生）
    for i in range(200):
        wb.defined_names.add(DefinedName("Sheet1!Print_Area_%d" % i, attr_text="数据!$A$1:$C$%d" % (i+1)))
    for i in range(200):
        wb.defined_names.add(DefinedName("_xlnm.Print_Titles_%d" % i, attr_text="数据!$1:$2"))
    for i in range(100):
        wb.defined_names.add(DefinedName("名称%d" % i, attr_text="第二表!$A$1:$B$5"))
    # 无效名称（#REF!）
    for i in range(50):
        wb.defined_names.add(DefinedName("失效名称%d" % i, attr_text="#REF!$A$1:$B$5"))
    wb.save(path)
    return path

if __name__ == "__main__":
    out = sys.argv[1] if len(sys.argv) > 1 else "tests/fixture.xlsx"
    os.makedirs(os.path.dirname(out) or ".", exist_ok=True)
    build(out)
    print("built", out, os.path.getsize(out))
