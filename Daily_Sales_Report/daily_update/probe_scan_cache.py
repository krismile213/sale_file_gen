# -*- coding: utf-8 -*-
"""probe_scan_cache.py — 两个副本的公式缓存扫描（缺缓存/零缓存格）"""
import os
import re
import zipfile
import xml.etree.ElementTree as ET

REPO = r"D:\Desktop\work-main\sale_file_gen"
NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

SCAN = [
    ("主表副本", os.path.join(REPO, "菲律宾重算_202609", "2026年东南亚销量表Southeast Asia Sales table.xlsx")),
    ("月销副本", os.path.join(REPO, "菲律宾重算_202609", "月销售总表.xlsx")),
]
for tag, path in SCAN:
    print("=" * 78)
    print(f"{tag}: {os.path.basename(path)}")
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
                    if len(nocache_samples) < 8:
                        nocache_samples.append(c.get("r"))
            if n_f:
                print(f"  {nm!r}: 公式格 {n_f}，缺/零缓存 {n_nocache}"
                      + (f"  样例 {nocache_samples}" if nocache_samples else ""))
