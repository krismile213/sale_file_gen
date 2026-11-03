# -*- coding: utf-8 -*-
"""probe_inject_au_v3.py — 东南亚销量表年度 sheet AU 列公式缓存注入（修正 rels 正则）"""
import os
import re
import zipfile

SEA = r"D:\Desktop\work-main\sale_file_gen\菲律宾重算_202609\2026年东南亚销量表Southeast Asia Sales table.xlsx"
AU_CACHE = {"AU3": 3485, "AU4": 3238, "AU5": 4026, "AU6": 174, "AU7": 15,
            "AU8": 11, "AU9": 361, "AU10": 0, "AU2": 11310}

with zipfile.ZipFile(SEA) as z:
    wbxml = z.read("xl/workbook.xml").decode("utf-8")
    rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")

m = re.search(r'<sheet[^>]*name="2023-2026年销量情况 Sales"[^>]*r:id="(rId\d+)"', wbxml)
assert m, "找不到年度 sheet"
rid = m.group(1)
print("rid =", rid)
rel_m = re.search(r'<Relationship[^>]*Id="%s"[^>]*/>' % rid, rels)
assert rel_m, f"rels 里找不到 {rid}"
target = re.search(r'Target="([^"]+)"', rel_m.group(0)).group(1).lstrip("/")
if not target.startswith("xl/"):
    target = "xl/" + target
print("target =", target)

with zipfile.ZipFile(SEA) as z:
    sheet_xml = z.read(target).decode("utf-8")

injected = 0
for ref, val in AU_CACHE.items():
    pat = re.compile(r'(<c r="%s"[^>]*>)(.*?)(</c>)' % ref, re.S)
    mm = pat.search(sheet_xml)
    if not mm:
        print(f"  ⚠️ 未找到 {ref}")
        continue
    inner = mm.group(2)
    if "<v>" in inner:
        new_inner = re.sub(r"<v>.*?</v>", "<v>%s</v>" % val, inner, flags=re.S)
    else:
        new_inner = inner + "<v>%s</v>" % val
    sheet_xml = sheet_xml[:mm.start()] + mm.group(1) + new_inner + mm.group(3) + sheet_xml[mm.end():]
    injected += 1
print(f"注入 {injected}/{len(AU_CACHE)} 个缓存值")

tmp = SEA + ".tmp"
with zipfile.ZipFile(SEA) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
    for item in zin.infolist():
        data = zin.read(item.filename)
        if item.filename == target:
            data = sheet_xml.encode("utf-8")
        zout.writestr(item, data)
os.replace(tmp, SEA)
print("已重打包")

# 复核
from openpyxl import load_workbook


def n0(v):
    if v is None:
        return 0
    if isinstance(v, str):
        s = v.strip()
        if s in ("00", ""):
            return 0
        try:
            return float(s)
        except ValueError:
            return 0
    return v if isinstance(v, (int, float)) else 0


wv = load_workbook(SEA, data_only=True)
yrv = None
for sn in wv.sheetnames:
    if "销量情况" in sn:
        yrv = wv[sn]
        break
ok = True
for ref, exp in AU_CACHE.items():
    v = n0(yrv[ref].value)
    flag = "✓" if v == exp else "✗"
    if v != exp:
        ok = False
    print(f"  {flag} {ref} = {v}（应 {exp}）")
wv.close()
print("复核：" + ("全部通过 ✓" if ok else "存在失败项 ✗"))
