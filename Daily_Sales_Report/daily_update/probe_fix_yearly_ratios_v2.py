# -*- coding: utf-8 -*-
"""probe_fix_yearly_ratios_v2.py — 修正 I10/AC10/I14 公式错误 + J14/AD14 派生传播 + AU 缓存注入

已确认的定义（从全表值反推，多行精确验证）：
  全球组：H = G/F-1（同比，今年/去年同月）｜I = G/G上行-1（环比，今年当月/今年上月）
          行9 I=0.0234569=71861/70214-1 ✓；行8 I=12532/12951 系东南亚组误看，全球组行8 I=0.0922251=68808/62998-1 ✓
  东南亚组：AB = AA/Z-1（同比）｜AC = AA/AA上行-1（环比）
          行8 AC=-0.0323527=12532/12951-1 ✓
  汇总行：I14 = G14/F14-1（原值 -0.2882244=595569/836737-1 ✓）｜AC14 = AA14/Z14-1（原值 -0.35271=117017/180780-1 ✓）
  J14 = SUM(组内年份列) 含 G14 → +2；AD14 = SUM(V14:AA14) 含 AA14 → +2（原 AD14=832308=旧 AA14=117017 时之和 ✓）
"""
import os
import re
import zipfile

from openpyxl import load_workbook

SEA = r"D:\Desktop\work-main\sale_file_gen\菲律宾重算_202609\2026年东南亚销量表Southeast Asia Sales table.xlsx"
MON = r"D:\Desktop\work-main\sale_file_gen\菲律宾重算_202609\月销售总表.xlsx"

# ================================================================ ① 月销年度 sheet 修正
print("=" * 84)
print("① 月销售总表年度 sheet 比率修正 + 派生传播")
print("=" * 84)
wm = load_workbook(MON, data_only=False)
yr = None
for sn in wm.sheetnames:
    if "销量情况" in sn:
        yr = wm[sn]
        break


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


g9, g10 = n0(yr["G9"].value), n0(yr["G10"].value)
aa9, aa10 = n0(yr["AA9"].value), n0(yr["AA10"].value)
f14, z14 = n0(yr["F14"].value), n0(yr["Z14"].value)
print(f"  基数: G9(8月全球)={g9} G10={g10}｜AA9(8月东南亚)={aa9} AA10={aa10}｜F14={f14} Z14={z14}")
assert g10 == 63199 and aa10 == 11310

fixes = [
    ("I10", g10 / g9 - 1, "全球环比 = G10/G9-1"),
    ("AC10", aa10 / aa9 - 1, "东南亚环比 = AA10/AA9-1"),
    ("I14", n0(yr["G14"].value) / f14 - 1, "全球汇总环比 = G14/F14-1"),
    ("J14", int(n0(yr["E14"].value) + n0(yr["F14"].value) + n0(yr["G14"].value)
              + sum(n0(yr.cell(14, c).value) for c in (2, 3, 4))), "全球历年总计 = SUM(组内年份列)"),
    ("AD14", int(sum(n0(yr.cell(14, c).value) for c in range(22, 28))), "东南亚历年总计 = SUM(V14:AA14)"),
]
for cell, new, why in fixes:
    old = yr[cell].value
    yr[cell] = new
    print(f"  {cell}: {old} → {new}   （{why}）")
wm.save(MON)
wm.close()
print("  已保存")

# ================================================================ ② 东南亚表 AU 缓存注入 + calcPr
print("\n" + "=" * 84)
print("② 东南亚销量表年度 sheet AU 列公式缓存注入 + calcPr 检查")
print("=" * 84)
AU_CACHE = {"AU3": 3485, "AU4": 3238, "AU5": 4026, "AU6": 174, "AU7": 15,
            "AU8": 11, "AU9": 361, "AU10": 0, "AU2": 11310}

with zipfile.ZipFile(SEA) as z:
    wbxml = z.read("xl/workbook.xml").decode("utf-8")
    rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
m = re.search(r'<sheet[^>]*name="2023-2026年销量情况 Sales"[^>]*r:id="(rId\d+)"', wbxml)
rid = m.group(1)
m2 = re.search(r'<Relationship[^>]*Id="%s"[^>]*Target="([^"]+)"' % rid, rels)
target = m2.group(1).lstrip("/")
if not target.startswith("xl/"):
    target = "xl/" + target
print(f"  年度 sheet → {target}")

with zipfile.ZipFile(SEA) as z:
    sheet_xml = z.read(target).decode("utf-8")
    # calcPr 检查
    mcalc = re.search(r"<calcPr[^>]*/?>", wbxml)
    calc_old = mcalc.group(0) if mcalc else "(无)"
print(f"  当前 calcPr: {calc_old}")

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
print(f"  注入 {injected}/{len(AU_CACHE)} 个缓存值")

# calcPr：若无 fullCalcOnLoad 则补上（openpyxl 重存丢了其他公式缓存，靠打开时重算补偿）
if "fullCalcOnLoad" not in calc_old:
    if mcalc:
        new_calc = calc_old.replace("/>", ' fullCalcOnLoad="1"/>') if calc_old.endswith("/>") \
            else calc_old.replace(">", ' fullCalcOnLoad="1">')
        wbxml_new = wbxml.replace(calc_old, new_calc)
    else:
        wbxml_new = wbxml.replace("</workbook>", '<calcPr calcId="124519" fullCalcOnLoad="1"/></workbook>')
    print(f"  calcPr 补 fullCalcOnLoad: {calc_old} → 已加")
else:
    wbxml_new = wbxml
    print("  calcPr 已含 fullCalcOnLoad，无需补")

tmp = SEA + ".tmp"
with zipfile.ZipFile(SEA) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
    for item in zin.infolist():
        data = zin.read(item.filename)
        if item.filename == target:
            data = sheet_xml.encode("utf-8")
        elif item.filename == "xl/workbook.xml":
            data = wbxml_new.encode("utf-8")
        zout.writestr(item, data)
os.replace(tmp, SEA)
print("  已重打包")

# ================================================================ ③ 立即复核
print("\n" + "=" * 84)
print("③ 复核：重读缓存值")
print("=" * 84)
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
wv2 = load_workbook(MON, data_only=True)
yr2 = None
for sn in wv2.sheetnames:
    if "销量情况" in sn:
        yr2 = wv2[sn]
        break
for cell, exp in [("G10", 63199), ("AA10", 11310), ("G14", 595571), ("AA14", 117019),
                  ("J14", 4188317), ("AD14", 832310)]:
    v = n0(yr2[cell].value)
    flag = "✓" if v == exp else "✗"
    if v != exp:
        ok = False
    print(f"  {flag} 月销年度 {cell} = {v}（应 {exp}）")
wv2.close()
print("\n复核：" + ("全部通过 ✓" if ok else "存在失败项 ✗"))
