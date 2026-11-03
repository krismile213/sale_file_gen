# -*- coding: utf-8 -*-
"""临时诊断：英国表 460~472 行的列完整度明细（判断残缺是否早已存在）"""
import os, sys, zipfile, re

BOOK = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))),
                    "1.每日销售数据表", "欧盟英国库存销售表.xlsx")


def sheet_map(path):
    z = zipfile.ZipFile(path)
    wbxml = z.read("xl/workbook.xml").decode()
    rels = z.read("xl/_rels/workbook.xml.rels").decode()
    rmap = {}
    for tag in re.findall(r"<Relationship\b[^>]*/?>", rels):
        rid = re.search(r'\bId="([^"]+)"', tag)
        tgt = re.search(r'\bTarget="([^"]+)"', tag)
        if rid and tgt:
            rmap[rid.group(1)] = tgt.group(1)
    out = {}
    for m in re.finditer(r'<sheet[^>]*name="([^"]+)"[^>]*r:id="([^"]+)"', wbxml):
        t = rmap.get(m.group(2), "").lstrip("/")
        if not t.startswith("xl/"):
            t = "xl/" + t
        out[m.group(1)] = t
    return z, out


def ci(c):
    n = 0
    for ch in c:
        n = n * 26 + (ord(ch) - 64)
    return n


z, sm = sheet_map(BOOK)
for sn, lo, hi in (("英国-每日库存分析表", 455, 475), ("欧洲-每日库存分析表", 455, 472)):
    x = z.read(sm[sn]).decode()
    rows = {}
    for m in re.finditer(r'<row r="(\d+)"[^>]*>(.*?)</row>', x, re.S):
        rn = int(m.group(1))
        cols = []
        sku = ""
        for mc in re.finditer(r'<c r="([A-Z]+)%d"([^>/]*)(?:/>|>(.*?)</c>)' % rn, m.group(2), re.S):
            cols.append(mc.group(1))
            if mc.group(1) == "A":
                body = mc.group(3) or ""
                t = re.search(r"<t[^>]*>(.*?)</t>", body, re.S)
                sku = t.group(1) if t else ""
        rows[rn] = (cols, sku)
    print("=" * 60)
    print(sn)
    for rn in range(lo, hi + 1):
        if rn not in rows:
            continue
        cols, sku = rows[rn]
        mx = max((ci(c) for c in cols), default=0)
        print("  r%-4d n=%-4d maxCol=%-4s sku=%-10s 末列=%s" % (
            rn, len(cols), mx, sku[:10], ",".join(cols[-3:])))
