# -*- coding: utf-8 -*-
# Author: Opera-Jin
import sys, os, shutil
sys.path.insert(0, "src")
from core import ooxml
from core.util import human_size

src = "tests/fixture.xlsx"
work = "tests/work.xlsx"
shutil.copy2(src, work)

ana = ooxml.analyze(work)
print("名称总数:", len(ana.names), "无效:", ana.invalid_names, "被引用:", ana.referenced_names)
print("外链:", len(ana.links), "外链公式:", ana.ext_formula_refs)
for n in ana.names[:3]:
    print("  ", n.name, "|", n.scope_label, "|", n.refers_to[:30], "| ref=", n.is_referenced)

out = ooxml.process(work, remove_print_names=False, remove_links=True, do_backup=True)
print("结果:", out.ok, out.message)
print("  名称 %d -> 删除 %d / 保留 %d" % (out.names_total, out.names_removed, out.names_kept))
print("  体积 %s -> %s" % (human_size(out.size_before), human_size(out.size_after)))
print("  保留原因:", out.kept_names[:5])

ana2 = ooxml.analyze(work)
print("处理后名称数:", len(ana2.names))
for n in ana2.names:
    print("   keep:", n.name, "|", n.refers_to[:40])
