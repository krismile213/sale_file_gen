# -*- coding: utf-8 -*-
"""找英国分析表 9 月：日销有销量但表里无行的 SKU"""
import os, re, zipfile
from collections import defaultdict
import openpyxl
from openpyxl.utils import column_index_from_string as CI

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.abspath(os.path.join(HERE, "..", "..", ".."))
XLSX = os.path.join(ROOT, "Daily_Sales_Report", "1.每日销售数据表", "欧盟英国库存销售表.xlsx")
ARCH9 = os.path.join(ROOT, "Daily_Sales_Report", "1.每日销售数据表", "2026年每日数据",
                     "2026年9月", "2026年9月-欧盟英国库存销售表.xlsx")

# 1) 日销 9 月 UK
wb = openpyxl.load_workbook(ARCH9, read_only=True, data_only=True)
ws = wb["每日销售数据"]
uk = defaultdict(int)
for r in ws.iter_rows(min_row=3, max_col=44, values_only=True):
    if not r[0] or r[42] != 1: continue
    if str(r[2])[:7] != "2026-09": continue
    if str(r[6]) != "Amazon.co.uk": continue
    if not r[11] or not isinstance(r[14], (int, float)): continue
    uk[str(r[11]).strip()] += int(r[14])
wb.close()
print("日销 UK 9月: %d 件 / %d SKU" % (sum(uk.values()), len(uk)))

# 2) 表内 SKU
z = zipfile.ZipFile(XLSX)
rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
rid2t = dict(re.findall(r'Id="(rId\d+)"[^>]*Target="([^"]+)"', rels))
wbx = z.read("xl/workbook.xml").decode("utf-8")
sm = {n: "xl/" + rid2t[r].lstrip("/") for n, r in
      re.findall(r'<sheet [^>]*name="([^"]+)"[^>]*r:id="(rId\d+)"', wbx)}
sst = []
if "xl/sharedStrings.xml" in z.namelist():
    sx = z.read("xl/sharedStrings.xml").decode("utf-8")
    sst = ["".join(re.findall(r"<t[^>]*>(.*?)</t>", si, re.S))
           for si in re.findall(r"<si>(.*?)</si>", sx, re.S)]
x = z.read(sm["英国-每日库存分析表"]).decode("utf-8")
z.close()

inb = {}
for m in re.finditer(r'<row r="(\d+)"', x):
    r = int(m.group(1))
    if r < 7: continue
    c = re.search(r'<c r="A%d"(?![0-9])([^>]*?)(?:/>|>(.*?)</c>)' % r, x, re.S)
    if not c: continue
    at, b = c.group(1), c.group(2) or ""
    if 't="s"' in at:
        v = re.search(r"<v>(\d+)</v>", b)
        s = sst[int(v.group(1))] if v and int(v.group(1)) < len(sst) else ""
    elif 't="inlineStr"' in at:
        s = "".join(re.findall(r"<t[^>]*>(.*?)</t>", b, re.S))
    else:
        s = b
    if str(s).strip(): inb[str(s).strip()] = r
print("表内 SKU: %d" % len(inb))

miss = {k: v for k, v in uk.items() if k not in inb}
print("\n缺行 SKU: %d 个 / %d 件" % (len(miss), sum(miss.values())))
for k, v in sorted(miss.items(), key=lambda kv: -kv[1]):
    print("   %-16s %d 件" % (k, v))
extra = {k: v for k, v in inb.items() if k not in uk}
print("\n表里有行但 9 月无销量: %d 个" % len(extra))
