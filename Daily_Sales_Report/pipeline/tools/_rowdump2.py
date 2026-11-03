# -*- coding: utf-8 -*-
"""临时诊断：欧盟表 mergeCells / 列定义 / CF 落点"""
import zipfile, re, os

P = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))),
                 "1.每日销售数据表", "欧盟英国库存销售表.xlsx")
z = zipfile.ZipFile(P)
for tgt, label in (("xl/worksheets/sheet9.xml", "英国"), ("xl/worksheets/sheet10.xml", "欧洲")):
    x = z.read(tgt).decode()
    print("=====", label)
    m = re.search(r"<mergeCells.*?</mergeCells>", x, re.S)
    print("mergeCells:", (m.group(0)[:800] if m else None))
    m = re.search(r"<cols>.*?</cols>", x, re.S)
    print("cols:", (m.group(0)[:500] if m else None))
    for mm in re.finditer(r'<conditionalFormatting[^>]*sqref="([^"]+)"[^>]*>(.*?)</conditionalFormatting>', x, re.S):
        for r2 in re.finditer(r"<cfRule([^>]*?)(?:/>|>(.*?)</cfRule>)", mm.group(2), re.S):
            f = re.search(r"<formula>(.*?)</formula>", r2.group(2) or "", re.S)
            prio = re.search(r'priority="(\d+)"', r2.group(1))
            dx = re.search(r'dxfId="(\d+)"', r2.group(1))
            print("  CF", mm.group(1), "p" + (prio.group(1) if prio else "?"), "dxf" + (dx.group(1) if dx else "-"),
                  (f.group(1)[:110] if f else ""))
