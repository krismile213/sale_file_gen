# -*- coding: utf-8 -*-
"""probe_final5_verify_v3485.py — 最终 5 表终验（2026-10-16 晚，Y+ 口径 3,485）

用户拍板"听你的，就按你的算" = 推荐口径 Y+（菲律宾 3,485 = 3,483 + G554×2）。
标准：菲律宾 9 月 3,485｜东南亚 9 月 11,310｜月销总计 9 月 63,199｜全球年计 595,571｜东南亚年计 117,019
文件：菲律宾重算_202609/ 下 5 表（3 张次表已从分发目录复制）
"""
import os
import re
import zipfile

import openpyxl

REPO = r"D:\Desktop\work-main\sale_file_gen"
DIR = os.path.join(REPO, "菲律宾重算_202609")
P_MAIN = os.path.join(DIR, "2026年东南亚销量表Southeast Asia Sales table.xlsx")
P_MONTH = os.path.join(DIR, "月销售总表.xlsx")
P_SUB_PH = os.path.join(DIR, "Philippines_202609_sales.xlsx")
P_SUB_ID = os.path.join(DIR, "Indonesia_202609_sales.xlsx")
P_SUB_TH = os.path.join(DIR, "Thailand_202609_sales.xlsx")

COUNTRY_ORDER = ["菲律宾", "泰国", "印尼", "新加坡", "越南", "马来Lazada", "虾皮跨境", "团购"]
EXP9 = {"菲律宾": 3485, "泰国": 3238, "印尼": 4026, "新加坡": 174, "越南": 15,
        "马来Lazada": 11, "虾皮跨境": 361, "团购": 0}

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


print("=" * 78)
print("[0] zip 完整性")
for p in (P_MAIN, P_MONTH, P_SUB_PH, P_SUB_ID, P_SUB_TH):
    with zipfile.ZipFile(p) as z:
        check(z.testzip() is None, f"{os.path.basename(p)} testzip OK")

print("=" * 78)
print("[1] 东南亚销量表 · Philippines（BC=线下 BD=9月合计，行2=合计行）")
wb = openpyxl.load_workbook(P_MAIN, data_only=True)
ph = wb["Philippines 菲律宾销量表"]
check(n0(ph.cell(2, 55).value) == 299, f"BC2 线下 = {ph.cell(2, 55).value}（应 299）")
check(n0(ph.cell(2, 56).value) == 3485, f"BD2 9月合计 = {ph.cell(2, 56).value}（应 3485）")
check(col_sum(ph, 55, 3) == 299, f"BC 数据行和 = {col_sum(ph, 55, 3):.0f}（应 299）")
check(col_sum(ph, 56, 3) == 3485, f"BD 数据行和 = {col_sum(ph, 56, 3):.0f}（应 3485）")
check(n0(ph.cell(2, 53).value) == 1811, f"BA2 Shopee9月 = {ph.cell(2, 53).value}（应 1811 = 1809+G554×2）")
check(n0(ph.cell(2, 52).value) == 558, f"AZ2 本土LZD9月 = {ph.cell(2, 52).value}（应 558 不变）")
check(n0(ph.cell(2, 51).value) == 168, f"AY2 跨境LZD9月 = {ph.cell(2, 51).value}（应 168 不变）")
check(n0(ph.cell(2, 54).value) == 649, f"BB2 TikTok9月 = {ph.cell(2, 54).value}（应 649 不变）")

print("[2] 东南亚销量表 · SEA Total（9月块 CI(87)合计 + CJ(88)~CQ(95) 八国）")
st = wb["SEA Total 东南亚总销量表"]
for i, c in enumerate(COUNTRY_ORDER):
    v2 = n0(st.cell(2, 88 + i).value)
    sd = col_sum(st, 88 + i, 3)
    check(v2 == EXP9[c] and sd == EXP9[c],
          f"{c}: 行2={v2:.0f} 数据和={sd:.0f}（应 {EXP9[c]}）")
check(n0(st.cell(2, 87).value) == 11310, f"CI2 9月TOTAL = {st.cell(2, 87).value}（应 11310）")
check(col_sum(st, 87, 3) == 11310, f"CI 数据行和 = {col_sum(st, 87, 3):.0f}（应 11310）")
check(n0(st.cell(2, 79).value) == n0(ph.cell(2, 50).value),
      f"CA2 菲8月 = {st.cell(2, 79).value} == 菲律宾 sheet AX2 = {n0(ph.cell(2, 50).value):.0f}")

print("[3] 东南亚销量表 · 年度 sheet「2023-2026年销量情况 Sales」（AU=9月，公式+XLOOKUP+缓存）")
yr = wb["2023-2026年销量情况 Sales"]
for i, c in enumerate(COUNTRY_ORDER):
    v = n0(yr.cell(3 + i, 47).value)
    check(v == EXP9[c], f"AU{3 + i} {c} = {v:.0f}（应 {EXP9[c]}）")
check(n0(yr.cell(2, 47).value) == 11310, f"AU2 SUM行 = {yr.cell(2, 47).value}（应 11310）")
wb.close()

print("=" * 78)
print("[4] 月销售总表 · 2026年（EM=菲 HS=东南亚 KA=总计 KJ=当年，数据行4+）")
wm = openpyxl.load_workbook(P_MONTH, data_only=True)
ws = wm["2026年"]
for col, name, exp in [(143, "EM 菲律宾9月", 3485), (227, "HS 东南亚9月", 11310),
                       (287, "KA 总计9月", 63199)]:
    s = col_sum(ws, col, 4)
    check(s == exp, f"{name} 数据行和 = {s:.0f}（应 {exp}）")
# G554 行单项
g554 = None
for r in range(4, ws.max_row + 1):
    if str(ws.cell(r, 1).value or "").strip().upper() == "G554":
        g554 = r
        break
if g554:
    for col, name, exp in [(143, "EM", 117), (227, "HS", 219), (287, "KA", 237), (296, "KJ", 730)]:
        v = n0(ws.cell(g554, col).value)
        check(v == exp, f"G554 行{g554} {name} = {v}（应 {exp}）")
print(f"    （信息）KJ 当年总销量 数据行和 = {col_sum(ws, 296, 4):.0f}")
wm_f = openpyxl.load_workbook(P_MONTH, data_only=False)
wsf = wm_f["2026年"]
for col, name in [(143, "EM"), (227, "HS"), (287, "KA")]:
    print(f"    （信息）{name}3 公式 = {str(wsf.cell(3, col).value)[:40]!r}"
          f"（fullCalcOnLoad 打开自动重算）")

print("[5] 月销售总表 · 年度 sheet（G10 全球9月 AA10 东南亚9月 / 行14 汇总）")
yr2 = find_sheet(wm, "销量情况")
for cell, exp in [("G10", 63199), ("AA10", 11310), ("G14", 595571), ("AA14", 117019),
                  ("J14", 4188317), ("AD14", 832310)]:
    v = n0(yr2[cell].value)
    check(v == exp, f"{cell} = {v}（应 {exp}）")
# 比率格自洽
e10 = n0(yr2["E10"].value); f10 = n0(yr2["F10"].value); g9 = n0(yr2["G9"].value)
z10 = n0(yr2["Z10"].value); aa9 = n0(yr2["AA9"].value)
f14 = n0(yr2["F14"].value); z14 = n0(yr2["Z14"].value)
for cell, exp in [("H10", 63199 / f10 - 1), ("I10", 63199 / g9 - 1),
                  ("AB10", 11310 / z10 - 1), ("AC10", 11310 / aa9 - 1),
                  ("I14", 595571 / f14 - 1), ("AC14", 117019 / z14 - 1)]:
    v = n0(yr2[cell].value)
    check(abs(v - exp) < 1e-9, f"{cell} = {v:.6f}（应 {exp:.6f}）")
wm_f.close()
wm.close()

print("=" * 78)
print("[6] 菲次表 Philippines_202609_sales.xlsx（G554×2 已补回）")
wp = openpyxl.load_workbook(P_SUB_PH, data_only=True)
s1 = wp["Sheet1"]
for c, name, exp in [(5, "中山", 168), (6, "Lazada", 558), (7, "Shopee", 1811),
                     (8, "Tiktok", 649), (9, "Offline", 299), (10, "Total", 3485)]:
    t = col_sum(s1, c, 2)
    check(t == exp, f"Sheet1 {name} 加总 = {t}（应 {exp}）")
s2 = find_sheet(wp, "Philippines")
g554_s2 = None
for r in range(3, s2.max_row + 1):
    if str(s2.cell(r, 1).value or "").strip().upper() == "G554":
        g554_s2 = r
        break
if g554_s2:
    check(n0(s2.cell(g554_s2, 5).value) == 44, f"sheet2 G554 Shopee = {s2.cell(g554_s2, 5).value}（应 44）")
    check(n0(s2.cell(g554_s2, 8).value) == 117, f"sheet2 G554 Total = {s2.cell(g554_s2, 8).value}（应 117）")
for c, name, exp in [(3, "中山", 168), (4, "Lazada", 558), (5, "Shopee", 1811),
                     (6, "Tiktok", 649), (7, "Offline", 299), (8, "Total", 3485)]:
    v2 = n0(s2.cell(2, c).value)
    sd = col_sum(s2, c, 3)
    check(v2 == exp and sd == exp, f"sheet2 {name}: 行2={v2:.0f} 数据和={sd:.0f}（应 {exp}）")
wp.close()

print("[7] 印尼次表（未动，4026）")
wi = openpyxl.load_workbook(P_SUB_ID, data_only=True)
si2 = find_sheet(wi, "Indonesia")
for c, name, exp in [(35, "SHP", 1023), (36, "海外仓线下", 2775), (37, "Tiktok", 228),
                     (38, "Total", 4026)]:
    v2 = n0(si2.cell(2, c).value)
    sd = col_sum(si2, c, 3)
    check(v2 == exp and sd == exp, f"sheet2 {name}: 行2={v2:.0f} 数据和={sd:.0f}（应 {exp}）")
wi.close()

print("[8] 泰次表（未动，3238）")
wt = openpyxl.load_workbook(P_SUB_TH, data_only=True)
st2 = find_sheet(wt, "Thailand")
for c, name, exp in [(3, "SHP", 2777), (4, "Lazada", 391), (5, "泰国仓", 0),
                     (6, "跨境Lazada", 70), (7, "Total", 3238)]:
    v2 = n0(st2.cell(2, c).value)
    sd = col_sum(st2, c, 3)
    check(v2 == exp and sd == exp, f"sheet2 {name}: 行2={v2:.0f} 数据和={sd:.0f}（应 {exp}）")
wt.close()

print("=" * 78)
print("[9] SKU 级恒等链：月销 EM == 主表 BD == 菲次表 sheet2 H")


def sku_map(ws, col, r0):
    m = {}
    for r in range(r0, ws.max_row + 1):
        v = ws.cell(r, 1).value
        if v is not None and str(v).strip():
            m[str(v).strip()] = n0(ws.cell(r, col).value)
    return m


w1 = openpyxl.load_workbook(P_MAIN, data_only=True)
bd_map = sku_map(w1["Philippines 菲律宾销量表"], 56, 3)
w2 = openpyxl.load_workbook(P_MONTH, data_only=True)
em_map = sku_map(w2["2026年"], 143, 4)
w3 = openpyxl.load_workbook(P_SUB_PH, data_only=True)
h_map = sku_map(find_sheet(w3, "Philippines"), 8, 3)
m_em = [s for s in set(bd_map) | set(em_map) if bd_map.get(s, 0) != em_map.get(s, 0)]
m_h = [s for s in set(bd_map) | set(h_map) if bd_map.get(s, 0) != h_map.get(s, 0)]
check(not m_em, f"月销 EM vs 主表 BD 差异 = {len(m_em)}" + ("" if not m_em else f"：{m_em[:8]}"))
check(not m_h, f"次表 H vs 主表 BD 差异 = {len(m_h)}" + ("" if not m_h else f"：{m_h[:8]}"))
check(bd_map.get("G554") == 117, f"G554 三表一致 = {bd_map.get('G554')}（应 117）")
w1.close(); w2.close(); w3.close()

print()
print("=" * 78)
print("总结：" + ("全部通过 ✓" if ok else "存在失败项 ✗"))
