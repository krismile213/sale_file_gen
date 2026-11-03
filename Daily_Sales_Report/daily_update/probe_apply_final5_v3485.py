# -*- coding: utf-8 -*-
"""probe_apply_final5_v3485.py — 3,483 → 3,485：补回 G554×2（9/12 Shopee，WMS+旧报告双记载）

用户 2026-10-16 拍板"听你的，就按你的算" = 推荐口径 Y+（3,485）。
传播链（全部只改 202609 的 5 个重算副本，原件/trunk 不动）：
  ① 东南亚销量表 Philippines!G554 行19 BA(SHP)+2/BD(合计)+2；行2 BA2/BD2 +2
  ② 东南亚销量表 SEA Total!G554 行19 CI(9月东南亚)/CJ(9月菲) +2；行2 CI2/CJ2 +2
  ③ 东南亚销量表年度 sheet AU2/AU3~AU10：公式保留，XML 注入缓存（+2 / 新值）
  ④ 月销售总表 2026年!G554 行520 EM(143菲)/HS(227)/KA(287)/KJ(296) +2
  ⑤ 月销售总表年度 sheet G10/AA10/G14/AA14 +2，同比环比 H10/I10/AB10/AC10/I14/AC14 重算
  ⑥ 菲次表 Sheet1 G554 行 col7/col10 +2；sheet2 G554 行 col5/col8 +2、行2 col5/col8 +2
  ⑦ 印尼/泰次表不动
"""
import os
import re
import shutil
import zipfile

from openpyxl import load_workbook

REPO = r"D:\Desktop\work-main\sale_file_gen"
DIR_RECALC = os.path.join(REPO, "菲律宾重算_202609")
DIR_DIST = os.path.join(REPO, "东南亚次表分发_202609")
DIR_BAK = os.path.join(DIR_RECALC, "备份_3,483口径_20261016")
SEA = os.path.join(DIR_RECALC, "2026年东南亚销量表Southeast Asia Sales table.xlsx")
MON = os.path.join(DIR_RECALC, "月销售总表.xlsx")
PH_SUB = os.path.join(DIR_DIST, "Philippines_202609_sales.xlsx")

changes = []


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


def set_cell(ws, row, col, new, tag):
    old = ws.cell(row, col).value
    ws.cell(row, col).value = new
    changes.append((tag, old, new))
    print(f"  {tag}: {old!r} → {new!r}")


# ================================================================ Step 0 备份
print("=" * 84)
print("Step 0 · 备份 3,483 口径 5 表")
print("=" * 84)
os.makedirs(DIR_BAK, exist_ok=True)
for src in (SEA, MON, PH_SUB,
            os.path.join(DIR_DIST, "Indonesia_202609_sales.xlsx"),
            os.path.join(DIR_DIST, "Thailand_202609_sales.xlsx")):
    dst = os.path.join(DIR_BAK, os.path.basename(src))
    if not os.path.exists(dst):
        shutil.copy2(src, dst)
        print(f"  备份 {os.path.basename(src)}")

# ================================================================ Step 1 东南亚销量表
print("\n" + "=" * 84)
print("Step 1 · 东南亚销量表（Philippines + SEA Total，静态值）")
print("=" * 84)
wb = load_workbook(SEA, data_only=False)
ph = wb["Philippines 菲律宾销量表"]
st = wb["SEA Total 东南亚总销量表"]
g554_ph = g554_st = None
for r in range(3, ph.max_row + 1):
    if str(ph.cell(r, 1).value or "").strip().upper() == "G554":
        g554_ph = r
        break
for r in range(3, st.max_row + 1):
    if str(st.cell(r, 1).value or "").strip().upper() == "G554":
        g554_st = r
        break
print(f"  Philippines!G554 行={g554_ph}｜SEA Total!G554 行={g554_st}")
assert g554_ph and g554_st
assert n0(ph.cell(g554_ph, 53).value) == 42, f"BA 现值异常: {ph.cell(g554_ph, 53).value}"
assert n0(ph.cell(g554_ph, 56).value) == 115, f"BD 现值异常: {ph.cell(g554_ph, 56).value}"
assert n0(ph.cell(2, 53).value) == 1809 and n0(ph.cell(2, 56).value) == 3483
assert n0(st.cell(g554_st, 87).value) == 217 and n0(st.cell(g554_st, 88).value) == 115
assert n0(st.cell(2, 87).value) == 11308 and n0(st.cell(2, 88).value) == 3483
set_cell(ph, g554_ph, 53, 44, "Philippines!BA%d G554 Shopee9月" % g554_ph)
set_cell(ph, g554_ph, 56, 117, "Philippines!BD%d G554 9月合计" % g554_ph)
set_cell(ph, 2, 53, 1811, "Philippines!BA2 Shopee9月合计行")
set_cell(ph, 2, 56, 3485, "Philippines!BD2 9月合计行")
set_cell(st, g554_st, 87, 219, "SEA Total!CI%d G554 9月东南亚" % g554_st)
set_cell(st, g554_st, 88, 117, "SEA Total!CJ%d G554 9月菲律宾" % g554_st)
set_cell(st, 2, 87, 11310, "SEA Total!CI2 9月东南亚合计行")
set_cell(st, 2, 88, 3485, "SEA Total!CJ2 9月菲律宾合计行")
# 年度 sheet AU 列公式保留（XLOOKUP/SUM），保存后 XML 注入缓存
wb.save(SEA)
print("  已保存（年度 sheet AU 列公式保留，稍后 XML 注入缓存）")
wb.close()

# ================================================================ Step 2 月销售总表
print("\n" + "=" * 84)
print("Step 2 · 月销售总表（2026年 G554 行 + 年度 sheet 静态传播）")
print("=" * 84)
wm = load_workbook(MON, data_only=False)
ws = wm["2026年"]
g554_m = None
for r in range(4, ws.max_row + 1):
    if str(ws.cell(r, 1).value or "").strip().upper() == "G554":
        g554_m = r
        break
print(f"  2026年!G554 行 = {g554_m}")
assert g554_m
em_old = n0(ws.cell(g554_m, 143).value)
print(f"  现值: EM(143)={em_old} HS(227)={ws.cell(g554_m, 227).value} "
      f"KA(287)={ws.cell(g554_m, 287).value} KJ(296)={ws.cell(g554_m, 296).value}")
assert em_old == 115, f"EM 现值异常: {em_old}"
assert n0(ws.cell(g554_m, 227).value) == 217
assert n0(ws.cell(g554_m, 287).value) == 235
assert n0(ws.cell(g554_m, 296).value) == 728
set_cell(ws, g554_m, 143, 117, "2026年!EM%d G554 菲9月" % g554_m)
set_cell(ws, g554_m, 227, 219, "2026年!HS%d G554 东南亚9月" % g554_m)
set_cell(ws, g554_m, 287, 237, "2026年!KA%d G554 总计9月" % g554_m)
set_cell(ws, g554_m, 296, 730, "2026年!KJ%d G554 当年总销量(差值传播)" % g554_m)

yr = None
for sn in wm.sheetnames:
    if "销量情况" in sn:
        yr = wm[sn]
        break
print(f"  年度 sheet: {yr.title!r}")
assert n0(yr["G10"].value) == 63197 and n0(yr["AA10"].value) == 11308
assert n0(yr["G14"].value) == 595569 and n0(yr["AA14"].value) == 117017
set_cell(yr, 10, 7, 63199, "年度!G10 全球9月")
set_cell(yr, 10, 27, 11310, "年度!AA10 东南亚9月")
set_cell(yr, 14, 7, 595571, "年度!G14 全球年计")
set_cell(yr, 14, 27, 117019, "年度!AA14 东南亚年计")
# 同比/环比重算（静态值，公式=上月/去年-1）
e10, f10 = n0(yr["E10"].value), n0(yr["F10"].value)
y10, z10 = n0(yr["Y10"].value), n0(yr["Z10"].value)
y14, z14 = n0(yr["Y14"].value), n0(yr["Z14"].value)
print(f"  基数: E10(8月全球)={e10} F10(去年9月全球)={f10}｜Y10(8月东南亚)={y10} Z10(去年9月东南亚)={z10}")
print(f"        Y14(去年东南亚)={y14} Z14(前年东南亚)={z14}")
set_cell(yr, 10, 8, 63199 / f10 - 1, "年度!H10 全球同比")
set_cell(yr, 10, 9, 63199 / e10 - 1, "年度!I10 全球环比")
set_cell(yr, 10, 28, 11310 / z10 - 1, "年度!AB10 东南亚同比")
set_cell(yr, 10, 29, 11310 / y10 - 1, "年度!AC10 东南亚环比")
set_cell(yr, 14, 9, 117019 / z14 - 1, "年度!I14 全球汇总环比")
set_cell(yr, 14, 29, 117019 / y14 - 1, "年度!AC14 东南亚汇总环比")
wm.save(MON)
print("  已保存")
wm.close()

# ================================================================ Step 3 菲次表
print("\n" + "=" * 84)
print("Step 3 · 菲律宾次表（Sheet1 + sheet2）")
print("=" * 84)
wp = load_workbook(PH_SUB, data_only=False)
s1 = wp["Sheet1"]
g554_s1 = None
for r in range(2, s1.max_row + 1):
    if str(s1.cell(r, 1).value or "").strip().upper() == "G554":
        g554_s1 = r
        break
print(f"  Sheet1!G554 行 = {g554_s1}")
assert g554_s1
assert n0(s1.cell(g554_s1, 7).value) == 42, f"Sheet1 Shopee 现值异常: {s1.cell(g554_s1, 7).value}"
assert n0(s1.cell(g554_s1, 10).value) == 115, f"Sheet1 Total 现值异常: {s1.cell(g554_s1, 10).value}"
set_cell(s1, g554_s1, 7, 44, "菲次表Sheet1!G%d G554 Shopee" % g554_s1)
set_cell(s1, g554_s1, 10, 117, "菲次表Sheet1!J%d G554 Total" % g554_s1)

s2 = None
for sn in wp.sheetnames:
    if "Philippines" in sn:
        s2 = wp[sn]
        break
g554_s2 = None
for r in range(3, s2.max_row + 1):
    if str(s2.cell(r, 1).value or "").strip().upper() == "G554":
        g554_s2 = r
        break
print(f"  sheet2({s2.title!r})!G554 行 = {g554_s2}")
assert g554_s2
assert n0(s2.cell(g554_s2, 5).value) == 42, f"sheet2 Shopee 现值异常: {s2.cell(g554_s2, 5).value}"
assert n0(s2.cell(g554_s2, 8).value) == 115, f"sheet2 Total 现值异常: {s2.cell(g554_s2, 8).value}"
assert n0(s2.cell(2, 5).value) == 1809 and n0(s2.cell(2, 8).value) == 3483
set_cell(s2, g554_s2, 5, 44, "菲次表sheet2!E%d G554 Shopee" % g554_s2)
set_cell(s2, g554_s2, 8, 117, "菲次表sheet2!H%d G554 Total" % g554_s2)
set_cell(s2, 2, 5, 1811, "菲次表sheet2!E2 Shopee合计行")
set_cell(s2, 2, 8, 3485, "菲次表sheet2!H2 Total合计行")
wp.save(PH_SUB)
print("  已保存")
wp.close()

# ================================================================ Step 4 XML 注入年度 sheet AU 缓存
print("\n" + "=" * 84)
print("Step 4 · 东南亚销量表年度 sheet AU 列公式缓存注入（openpyxl 保存后缓存丢失）")
print("=" * 84)
AU_CACHE = {"AU3": 3485, "AU4": 3238, "AU5": 4026, "AU6": 174, "AU7": 15,
            "AU8": 11, "AU9": 361, "AU10": 0, "AU2": 11310}

with zipfile.ZipFile(SEA) as z:
    wbxml = z.read("xl/workbook.xml").decode("utf-8")
    rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    names = z.namelist()
# 找年度 sheet 的 sheetN.xml
m = re.search(r'<sheet[^>]*name="2023-2026年销量情况 Sales"[^>]*r:id="(rId\d+)"', wbxml)
rid = m.group(1)
m2 = re.search(r'<Relationship[^>]*Id="%s"[^>]*Target="([^"]+)"' % rid, rels)
target = m2.group(1).lstrip("/")
if not target.startswith("xl/"):
    target = "xl/" + target
print(f"  年度 sheet → {target}")
sheet_xml = None
with zipfile.ZipFile(SEA) as z:
    sheet_xml = z.read(target).decode("utf-8")

injected = 0
for ref, val in AU_CACHE.items():
    # 形如 <c r="AU3" s="12"><f>...</f></c> → 插入 <v>
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

# 重打包（只替换该 sheetN.xml）
tmp = SEA + ".tmp"
with zipfile.ZipFile(SEA) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
    for item in zin.infolist():
        data = zin.read(item.filename)
        if item.filename == target:
            data = sheet_xml.encode("utf-8")
        zout.writestr(item, data)
os.replace(tmp, SEA)
print("  已重打包：", os.path.basename(SEA))

print("\n" + "=" * 84)
print(f"完成：共修改 {len(changes)} 处静态值 + {injected} 处公式缓存")
print("=" * 84)
for tag, old, new in changes:
    print(f"  {tag}: {old} → {new}")
