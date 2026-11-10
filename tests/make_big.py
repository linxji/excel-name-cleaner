# -*- coding: utf-8 -*-
# Author: Opera-Jin
"""构造含 30 万个名称的大文件，用于压测。"""
import os, re, sys, zipfile, shutil, time

SRC = "tests/samples/样本_xlsx.xlsx"
OUT = "tests/big.xlsx"
N = int(sys.argv[1]) if len(sys.argv) > 1 else 300000

t0 = time.time()
zin = zipfile.ZipFile(SRC)
wb = zin.read("xl/workbook.xml").decode("utf-8")
parts = []
for i in range(N):
    parts.append('<definedName name="_xlnm.Print_Area_%d">数据!$A$1:$C$%d</definedName>' % (i, i % 50 + 1))
for i in range(N // 3):
    parts.append('<definedName name="冗余名称_%d">#REF!$A$1:$B$5</definedName>' % i)
block = "<definedNames>" + "".join(parts) + "</definedNames>"
if "<definedNames>" in wb:
    wb = re.sub(r"<definedNames>.*?</definedNames>", block, wb, flags=re.S)
else:
    wb = wb.replace("</workbook>", block + "</workbook>")

with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as zout:
    for info in zin.infolist():
        data = zin.read(info.filename)
        if info.filename == "xl/workbook.xml":
            data = wb.encode("utf-8")
        zout.writestr(info, data)
zin.close()
print("built", OUT, os.path.getsize(OUT), "names", N + N // 3, "%.1fs" % (time.time() - t0))
