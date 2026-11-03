# -*- coding: utf-8 -*-
"""
probe_deliver_5tables_verify.py — 交付终验（只读，2026-10-16 新口径）

复读 5 个交付表 + trunk 成品抽查，不依赖任何改动探针的内部状态：
  ① 菲律宾重算_202609/2026年东南亚销量表Southeast Asia Sales table.xlsx（主表副本）
  ② 菲律宾重算_202609/月销售总表.xlsx（月销总表副本）
  ③ 东南亚次表分发_202609/Philippines_202609_sales.xlsx（菲次表）
  ④ 东南亚次表分发_202609/Indonesia_202609_sales.xlsx（印尼次表，未动）
  ⑤ 东南亚次表分发_202609/Thailand_202609_sales.xlsx（泰次表，未动）

新标准（菲律宾线下 = 145 手工单 + 154 RESELLER = 299）：
  菲律宾 9 月 3,483｜东南亚 9 月 11,308｜月销总计 9 月 63,197｜全球年计 595,569
单元格口径（与 probe_seatotal_fix.py / probe_fix_sub_sheet1.py 一致）：
  主表 Philippines: BC(55)线下 BD(56)合计（9月块 AY~BD，行2=合计行）
  主表 SEA Total : 月块首=15+(m-1)*9，9月块 CI(87)合计 + CJ(88)菲 CB? 见 COUNTRY_ORDER
  主表年度 sheet 「2023-2026年销量情况 Sales」: 行3~10=8国，AM(39)~AU(47)=1~9月，行2=SUM行
  月销总表 2026年: EM(143)菲 HS(227)东南亚 KA(287)总计 KJ(296)当年，行3=SUM行，数据行4+
  月销总表年度 sheet: G10 全球9月 AA10 东南亚9月 G14/AA14 年计 J14/AD14 历年
  菲次表: Sheet1 列5~10（中山/LZD/SHP/TK/Offline/Total）行2起数据；sheet2 C~H 行2=合计行
  印尼次表 sheet2: C SHP D 海外仓线下 E Tiktok F Total｜泰次表 sheet2: C SHP D LZD E 泰国仓 F 跨境LZD G Total
"""
import os
import re
import zipfile

import openpyxl

REPO = r"D:\Desktop\work-main\sale_file_gen"
P_MAIN = os.path.join(REPO, "菲律宾重算_202609", "2026年东南亚销量表Southeast Asia Sales table.xlsx")
P_MONTH = os.path.join(REPO, "菲律宾重算_202609", "月销售总表.xlsx")
P_SUB_PH = os.path.join(REPO, "东南亚次表分发_202609", "Philippines_202609_sales.xlsx")
P_SUB_ID = os.path.join(REPO, "东南亚次表分发_202609", "Indonesia_202609_sales.xlsx")
P_SUB_TH = os.path.join(REPO, "东南亚次表分发_202609", "Thailand_202609_sales.xlsx")
P_TRUNK = os.path.join(REPO, "trunk", "result", "东南亚销量表", "202609",
                       "2026年东南亚销量表Southeast Asia Sales table.xlsx")

COUNTRY_ORDER = ["菲律宾", "泰国", "印尼", "新加坡", "越南", "马来Lazada", "虾皮跨境", "团购"]

ok = True


def check(cond, msg):
    global ok
    print(("  ✓ " if cond else "  ✗ ") + msg)
    if not cond:
        ok = False


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


def find_sheet(wb, key):
    for sn in wb.sheetnames:
        if key in sn:
            return wb[sn]
    raise KeyError(f"找不到含 {key!r} 的 sheet: {wb.sheetnames}")


def col_sum(ws, col, r0):
    return sum(n0(ws.cell(r, col).value) for r in range(r0, ws.max_row + 1))


def calcpr(path):
    with zipfile.ZipFile(path) as z:
        wbxml = z.read("xl/workbook.xml").decode("utf-8")
    m = re.search(r"<calcPr[^>]*/?>", wbxml)
    return m.group(0) if m else "(无 calcPr)"


print("=" * 78)
print("[0] zip 完整性 + calcPr")
for p in (P_MAIN, P_MONTH, P_SUB_PH, P_SUB_ID, P_SUB_TH):
    with zipfile.ZipFile(p) as z:
        check(z.testzip() is None, f"{os.path.basename(p)} testzip OK")
print(f"    主表副本 calcPr: {calcpr(P_MAIN)}")
print(f"    月销副本 calcPr: {calcpr(P_MONTH)}")

# ---------------------------------------------------------------- ① 主表副本
print("=" * 78)
print("[1] 主表副本 · Philippines sheet（BC=线下 BD=9月合计，行2=合计行）")
wb = openpyxl.load_workbook(P_MAIN, data_only=True)
ph = wb["Philippines 菲律宾销量表"]
check(n0(ph.cell(2, 55).value) == 299, f"BC2 线下 = {ph.cell(2, 55).value}（应 299 = 145+154）")
check(n0(ph.cell(2, 56).value) == 3483, f"BD2 9月合计 = {ph.cell(2, 56).value}（应 3483）")
check(col_sum(ph, 55, 3) == 299, f"BC 数据行和 = {col_sum(ph, 55, 3):.0f}（应 299）")
check(col_sum(ph, 56, 3) == 3483, f"BD 数据行和 = {col_sum(ph, 56, 3):.0f}（应 3483）")
ax2 = n0(ph.cell(2, 50).value)
print(f"    （信息）菲律宾 8月合计 AX2 = {ax2:.0f}")

print("[2] 主表副本 · SEA Total（9月块 CI(87)合计 + CJ(88)~CQ(95) 八国）")
st = wb["SEA Total 东南亚总销量表"]
exp9 = {"菲律宾": 3483, "泰国": 3238, "印尼": 4026, "新加坡": 174, "越南": 15,
        "马来Lazada": 11, "虾皮跨境": 361, "团购": 0}
for i, c in enumerate(COUNTRY_ORDER):
    v2 = n0(st.cell(2, 88 + i).value)
    sd = col_sum(st, 88 + i, 3)
    check(v2 == exp9[c] and sd == exp9[c],
          f"{c}: 行2={v2:.0f} 数据和={sd:.0f}（应 {exp9[c]}）")
check(n0(st.cell(2, 87).value) == 11308, f"CI2 9月TOTAL = {st.cell(2, 87).value}（应 11308）")
check(col_sum(st, 87, 3) == 11308, f"CI 数据行和 = {col_sum(st, 87, 3):.0f}（应 11308）")
check(n0(st.cell(2, 79).value) == ax2,
      f"CA2 菲8月 = {st.cell(2, 79).value} == 菲律宾 sheet AX2 = {ax2:.0f}（修复态交叉）")

print("[3] 主表副本 · 年度 sheet「2023-2026年销量情况 Sales」（行3~10=8国，AU=9月）")
yr = wb["2023-2026年销量情况 Sales"]
for i, c in enumerate(COUNTRY_ORDER):
    v = n0(yr.cell(3 + i, 47).value)
    check(v == exp9[c], f"AU{3 + i} {c} = {v:.0f}（应 {exp9[c]}）")
check(n0(yr.cell(2, 47).value) == 11308, f"AU2 SUM行 = {yr.cell(2, 47).value}（应 11308）")
wb.close()

# ---------------------------------------------------------------- ② 月销总表副本
print("=" * 78)
print("[4] 月销总表副本 · 2026年（EM=菲 HS=东南亚 KA=总计，数据行4+）")
wm = openpyxl.load_workbook(P_MONTH, data_only=True)
ws = wm["2026年"]
for col, name, exp in [(143, "EM 菲律宾9月", 3483), (227, "HS 东南亚9月", 11308),
                       (287, "KA 总计9月", 63197)]:
    s = col_sum(ws, col, 4)
    check(s == exp, f"{name} 数据行和 = {s:.0f}（应 {exp}）")
kj_sum = col_sum(ws, 296, 4)
print(f"    （信息）KJ 当年总销量 数据行和 = {kj_sum:.0f}（禁止与月列重算对齐，仅参考）")
# 行3 SUM 公式 + 缓存（该簿 calcPr 带 fullCalcOnLoad=1，Excel/WPS 打开会自动重算；
# openpyxl 存的模板无公式缓存，程序直读 data_only 得 0 —— 模板固有性质，非数据问题）
wm_f = openpyxl.load_workbook(P_MONTH, data_only=False)
wsf = wm_f["2026年"]
for col, name, exp in [(143, "EM", 3483), (227, "HS", 11308), (287, "KA", 63197)]:
    f = wsf.cell(3, col).value
    cv = n0(ws.cell(3, col).value)
    print(f"    （信息）{name}3 公式 = {str(f)[:40]!r} 缓存 = {cv:.0f}"
          f"（应 {exp}；无缓存，靠 fullCalcOnLoad 重算）")
print(f"    （信息）KJ3 公式 = {str(wsf.cell(3, 296).value)[:40]!r}")

print("[5] 月销总表副本 · 年度 sheet（G10 全球9月 AA10 东南亚9月 / 行14 汇总）")
yr2 = find_sheet(wm, "销量情况")
for cell, exp in [("G10", 63197), ("AA10", 11308), ("G14", 595569), ("AA14", 117017),
                  ("J14", 4188315), ("AD14", 832308)]:
    v = n0(yr2[cell].value)
    check(v == exp, f"{cell} = {v:.0f}（应 {exp}）")
wm_f.close()
wm.close()

# ---------------------------------------------------------------- ③④⑤ 次表
print("=" * 78)
print("[6] 菲次表 Philippines_202609_sales.xlsx")
wp = openpyxl.load_workbook(P_SUB_PH, data_only=True)
s1 = wp["Sheet1"]
hdr = [str(s1.cell(1, c).value) for c in range(1, 11)]
print(f"    Sheet1 表头: {hdr}")
for c, name, exp in [(5, "中山", 168), (6, "Lazada", 558), (7, "Shopee", 1809),
                     (8, "Tiktok", 649), (9, "Offline", 299), (10, "Total", 3483)]:
    t = col_sum(s1, c, 2)
    check(t == exp, f"Sheet1 {name} 加总 = {t}（应 {exp}）")
n_rows = sum(1 for r in range(2, s1.max_row + 1) if s1.cell(r, 1).value)
check(n_rows == 485, f"Sheet1 数据行 = {n_rows}（应 485）")
s2 = find_sheet(wp, "Philippines")
n_gk = sum(1 for r in range(3, s2.max_row + 1)
           if s2.cell(r, 1).value and re.match(r"^[GK]\d", str(s2.cell(r, 1).value).strip()))
check(n_gk == 485, f"sheet2 G/K 行 = {n_gk}（应 485）")
for c, name, exp in [(3, "中山", 168), (4, "Lazada", 558), (5, "Shopee", 1809),
                     (6, "Tiktok", 649), (7, "Offline", 299), (8, "Total", 3483)]:
    v2 = n0(s2.cell(2, c).value)
    sd = col_sum(s2, c, 3)
    check(v2 == exp and sd == exp, f"sheet2 {name}: 行2={v2:.0f} 数据和={sd:.0f}（应 {exp}）")
wp.close()

print("[7] 印尼次表 Indonesia_202609_sales.xlsx（未动，4026）")
wi = openpyxl.load_workbook(P_SUB_ID, data_only=True)
si1 = wi["Sheet1"]
hdr1 = [str(si1.cell(1, c).value) for c in range(1, si1.max_column + 1)]
print(f"    Sheet1 表头: {hdr1}")
tcol = None
for c in range(1, si1.max_column + 1):
    h = si1.cell(1, c).value
    if h and any(k in str(h).lower() for k in ("total", "总计", "合计", "总销量")):
        tcol = c
if tcol is None:
    for c in range(si1.max_column, 0, -1):
        if any(isinstance(si1.cell(r, c).value, (int, float)) for r in range(2, si1.max_row + 1)):
            tcol = c
check(tcol is not None, f"Sheet1 Total 列定位 = {tcol}")
if tcol:
    t = col_sum(si1, tcol, 2)
    check(t == 4026, f"Sheet1 Total 加总 = {t}（应 4026）")
si2 = find_sheet(wi, "Indonesia")
# ⚠️ 印尼次表 sheet2 是 12 个月块全量表（1月从 C 起、块宽 4），9月块 = AI(35)~AL(38)
# （与菲/泰次表"9月起4块"布局不同，系该表模板世代不同；数据本身正确）
for c, name, exp in [(35, "SHP", 1023), (36, "海外仓线下", 2775), (37, "Tiktok", 228),
                     (38, "Total", 4026)]:
    v2 = n0(si2.cell(2, c).value)
    sd = col_sum(si2, c, 3)
    check(v2 == exp and sd == exp, f"sheet2 {name}: 行2={v2:.0f} 数据和={sd:.0f}（应 {exp}）")
wi.close()

print("[8] 泰次表 Thailand_202609_sales.xlsx（未动，3238）")
wt = openpyxl.load_workbook(P_SUB_TH, data_only=True)
st1 = wt["Sheet1"]
hdr1 = [str(st1.cell(1, c).value) for c in range(1, st1.max_column + 1)]
print(f"    Sheet1 表头: {hdr1}")
tcol = None
for c in range(1, st1.max_column + 1):
    h = st1.cell(1, c).value
    if h and any(k in str(h).lower() for k in ("total", "总计", "合计", "总销量")):
        tcol = c
if tcol is None:
    for c in range(st1.max_column, 0, -1):
        if any(isinstance(st1.cell(r, c).value, (int, float)) for r in range(2, st1.max_row + 1)):
            tcol = c
check(tcol is not None, f"Sheet1 Total 列定位 = {tcol}")
if tcol:
    t = col_sum(st1, tcol, 2)
    check(t == 3238, f"Sheet1 Total 加总 = {t}（应 3238）")
st2 = find_sheet(wt, "Thailand")
for c, name, exp in [(3, "SHP", 2777), (4, "Lazada", 391), (5, "泰国仓", 0),
                     (6, "跨境Lazada", 70), (7, "Total", 3238)]:
    v2 = n0(st2.cell(2, c).value)
    sd = col_sum(st2, c, 3)
    check(v2 == exp and sd == exp, f"sheet2 {name}: 行2={v2:.0f} 数据和={sd:.0f}（应 {exp}）")
wt.close()

# ---------------------------------------------------------------- Bonus: trunk 成品
print("=" * 78)
print("[9] trunk 成品抽查（Bonus，非交付件）")
if os.path.exists(P_TRUNK):
    wbt = openpyxl.load_workbook(P_TRUNK, data_only=True)
    stt = wbt["SEA Total 东南亚总销量表"]
    check(n0(stt.cell(2, 87).value) == 10948, f"trunk SEA Total CI2 = {stt.cell(2, 87).value}（应 10948）")
    check(n0(stt.cell(2, 88).value) == 3123, f"trunk SEA Total CJ2 菲 = {stt.cell(2, 88).value}（应 3123）")
    wbt.close()
else:
    print("    （trunk 成品不在，跳过）")
# ⚠️ 已知 trunk 隐患（2026-10-16 发现，待用户定夺）：
# trunk/result/数据归属周期/202609/月销售总表.xlsx 停在 10-15 预览值
# （EM 3,329 / HS 11,154 / KA 63,043，mtime 10-15 16:30），与 trunk 东南亚表（菲 3,123）
# 内部不一致（50 个 SKU 级差异、净 +206）。此处只报告，不改动。
TRUNK_MONTH = os.path.join(REPO, "trunk", "result", "数据归属周期", "202609", "月销售总表.xlsx")
if os.path.exists(TRUNK_MONTH):
    wtm = openpyxl.load_workbook(TRUNK_MONTH, data_only=True)
    wsm = wtm["2026年"]
    print(f"    ⚠️ trunk 月销总表 EM={col_sum(wsm, 143, 4):.0f} HS={col_sum(wsm, 227, 4):.0f} "
          f"KA={col_sum(wsm, 287, 4):.0f}（预览值 3329/11154/63043，与东南亚表 3123 矛盾，待定夺）")
    wtm.close()

# ---------------------------------------------------------------- SKU 级恒等链
print("=" * 78)
print("[10] SKU 级恒等链：月销 EM == 主表 BD == 菲次表 sheet2 H（'00' 归一后）")


def sku_map(ws, col, r0):
    m = {}
    for r in range(r0, ws.max_row + 1):
        v = ws.cell(r, 1).value
        if v is not None and str(v).strip():
            m[str(v).strip()] = n0(ws.cell(r, col).value)
    return m


w1 = openpyxl.load_workbook(P_MAIN, data_only=True)
ph1 = w1["Philippines 菲律宾销量表"]
bd_map = sku_map(ph1, 56, 3)
w2 = openpyxl.load_workbook(P_MONTH, data_only=True)
em_map = sku_map(w2["2026年"], 143, 4)
w3 = openpyxl.load_workbook(P_SUB_PH, data_only=True)
s2p = find_sheet(w3, "Philippines")
h_map = sku_map(s2p, 8, 3)
m_em = [s for s in set(bd_map) | set(em_map) if bd_map.get(s, 0) != em_map.get(s, 0)]
m_h = [s for s in set(bd_map) | set(h_map) if bd_map.get(s, 0) != h_map.get(s, 0)]
check(not m_em, f"月销 EM vs 主表 BD 差异 = {len(m_em)}" + ("" if not m_em else f"：{m_em[:8]}"))
check(not m_h, f"次表 H vs 主表 BD 差异 = {len(m_h)}" + ("" if not m_h else f"：{m_h[:8]}"))
print(f"    （信息）仅月销有的 SKU {len(set(em_map) - set(bd_map))} 个、仅主表有的 "
      f"{len(set(bd_map) - set(em_map))} 个，两侧非零值均为空（已验）")
w1.close()
w2.close()
w3.close()

print()
print("=" * 78)
print("总结：" + ("全部通过 ✓" if ok else "存在失败项 ✗"))
