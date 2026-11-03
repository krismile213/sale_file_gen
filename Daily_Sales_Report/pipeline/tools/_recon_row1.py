# -*- coding: utf-8 -*-
"""看若干表第 1 行到底有什么内容（判断合并范围是否安全）。"""
import os
import re
import sys
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.abspath(os.path.join(HERE, "..", "core")))
import config as C
from _recon_cf_all import sheet_map, sst_list, cell_text, row_cells, colnum, _col

BOOKS = os.path.join(C.DSR, "1.每日销售数据表")
TARGETS = [
    ("北美加拿大库存销售表.xlsx", "美国-每日库存分析表"),
    ("北美加拿大库存销售表.xlsx", "墨西哥-每日库存分析表"),
    ("沙特库存销售表.xlsx", "沙特阿拉伯-每日库存分析表"),
    ("Interflowd沙特库存销售表.xlsx", "沙特-每日库存分析表"),
    ("欧盟英国库存销售表.xlsx", "英国-每日库存分析表"),
]

for fname, sn in TARGETS:
    p = os.path.join(BOOKS, fname)
    z = zipfile.ZipFile(p)
    ss = sst_list(z)
    sp = dict(sheet_map(z))[sn]
    xml = z.read(sp).decode("utf-8")
    rc = row_cells(xml, 1)
    print("=" * 78)
    print(fname, "::", sn, " 行1 共 %d 个 cell" % len(rc))
    items = sorted(((colnum(k), k) for k in rc))
    for n, k in items[:30]:
        a, b = rc[k]
        t = cell_text(b, a, ss)
        print("   %-4s %-22r %s" % (k, t[:20], "MERGED-OWNER?" if False else ""))
    # 行2 也看一眼
    rc2 = row_cells(xml, 2)
    if rc2:
        print("   行2:", [(k, cell_text(v[1], v[0], ss)[:10])
                        for k, v in sorted(rc2.items(), key=lambda x: colnum(x[0]))[:12]])
    z.close()
