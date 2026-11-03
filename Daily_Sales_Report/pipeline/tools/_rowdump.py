# -*- coding: utf-8 -*-
"""临时诊断：看欧盟表某几行的列覆盖与关键公式（不写文件）"""
import zipfile, re, sys, os

HERE = os.path.dirname(os.path.abspath(__file__))
P = os.path.join(os.path.dirname(os.path.dirname(HERE)),
                 "1.每日销售数据表", "欧盟英国库存销售表.xlsx")
z = zipfile.ZipFile(P)
x = z.read("xl/worksheets/sheet10.xml").decode()


def rowcells(rn):
    m = re.search(r'<row r="%d"[^>]*>(.*?)</row>' % rn, x, re.S)
    if not m:
        return None
    out = []
    for cr, attr, body in re.findall(r'<c r="([A-Z]+%d)"([^>]*)>(.*?)</c>' % rn, m.group(1), re.S):
        col = re.match(r"[A-Z]+", cr).group(0)
        f = re.search(r"<f[^>]*>(.*?)</f>", body, re.S)
        t = re.search(r"<t[^>]*>(.*?)</t>", body, re.S)
        v = re.search(r"<v>(.*?)</v>", body, re.S)
        out.append((col, (f.group(1) if f else ""), (t.group(1) if t else ""), (v.group(1) if v else "")))
    return out


for rn in (7, 100, 462, 465, 467, 468):
    cs = rowcells(rn)
    if cs is None:
        print(rn, "no row")
        continue
    print("=== row", rn, "cells", len(cs))
    print("   cols:", ",".join(c[0] for c in cs[:6]), "...", ",".join(c[0] for c in cs[-6:]))
    for col in ("A", "B", "C", "D", "E", "F", "G", "H"):
        for c in cs:
            if c[0] == col:
                print("   %-2s F=%s | T=%s | V=%s" % (col, c[1][:70], c[2][:24], c[3][:10]))
