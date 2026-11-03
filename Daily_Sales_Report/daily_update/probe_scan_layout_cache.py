# -*- coding: utf-8 -*-
"""probe_scan_layout_cache.py — ① trunk 次表 sheet2 块布局对照 ② 两副本公式缓存扫描"""
import os
import re
import zipfile
import xml.etree.ElementTree as ET

import openpyxl

REPO = r"D:\Desktop\work-main\sale_file_gen"
NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
TRUNK_SUB = os.path.join(REPO, "trunk", "result", "东南亚销量表", "202609")
DIST_SUB = os.path.join(REPO, "东南亚次表分发_202609")

print("=" * 78)
print("[A] trunk 202609 次表 sheet2 块布局（对照分发目录）")
for name, key in [("Philippines_202609_sales.xlsx", "Philippines"),
                  ("Indonesia_202609_sales.xlsx", "Indonesia"),
                  ("Thailand_202609_sales.xlsx", "Thailand")]:
    for base, tag in [(TRUNK_SUB, "trunk"), (DIST_SUB, "分发")]:
        p = os.path.join(base, name)
        if not os.path.exists(p):
            print(f"  [{tag}] {name}: 不存在")
            continue
        wb = openpyxl.load_workbook(p, data_only=True)
        s2 = None
        for sn in wb.sheetnames:
            if key in sn:
                s2 = wb[sn]
                break
        # 找含「9月」的列与块首列
        col9 = None
        blk_first = None
        for c in range(3, s2.max_column + 1):
            h = s2.cell(1, c).value
            if h and "9月" in str(h):
                col9 = c
        # 块首 = 9月总销量列 - (块宽-1)；块宽从相邻月标签推
        hdrs = [(c, str(s2.cell(1, c).value)) for c in range(3, s2.max_column + 1) if s2.cell(1, c).value]
        months = [(c, h) for c, h in hdrs if re.search(r"(\d+)月", h)]
        first_month = months[0][1] if months else "?"
        print(f"  [{tag}] {key}: max_col={s2.max_column} 首块月份={first_month!r} "
              f"9月总销量列={col9} 行2值={s2.cell(2, col9).value if col9 else None}")
        wb.close()

print("=" * 78)
print("[B] 公式缓存扫描（缺缓存/零缓存的公式格）")
SCAN = [
    ("主表副本", os.path.join(REPO, "菲律宾重算_202609", "2026年东南亚销量表Southeast Asia Sales table.xlsx")),
    ("月销副本", os.path.join(REPO, "菲律宾重算_202609", "月销售总表.xlsx")),
]
for tag, path in SCAN:
    print(f"  --- {tag}: {os.path.basename(path)}")
    with zipfile.ZipFile(path) as z:
        wbxml = z.read("xl/workbook.xml").decode("utf-8")
        rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
        relmap = {rel.get("Id"): rel.get("Target") for rel in ET.fromstring(rels)}
        sheets = re.findall(r'<sheet name="([^"]+)"[^>]*r:id="(rId\d+)"', wbxml)
        for nm, rid in sheets:
            t = relmap[rid].lstrip("/")
            t = t if t.startswith("xl/") else "xl/" + t
            root = ET.fromstring(z.read(t).decode("utf-8"))
            n_f = n_nocache = 0
            nocache_samples = []
            for c in root.iter(f"{{{NS}}}c"):
                f = c.find(f"{{{NS}}}f")
                if f is None:
                    continue
                n_f += 1
                v = c.find(f"{{{NS}}}v")
                if v is None or v.text in (None, "", "0"):
                    n_nocache += 1
                    if len(nocache_samples) < 6:
                        nocache_samples.append(c.get("r"))
            if n_f:
                print(f"    {nm!r}: 公式格 {n_f}，缺/零缓存 {n_nocache}"
                      + (f"  样例 {nocache_samples}" if nocache_samples else ""))
