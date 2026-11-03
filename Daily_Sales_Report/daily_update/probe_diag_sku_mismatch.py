# -*- coding: utf-8 -*-
"""probe_diag_sku_mismatch.py v2 — 修复 "00" 文本格导致的 str-int 运算崩溃"""
import os
import datetime

import openpyxl

REPO = r"D:\Desktop\work-main\sale_file_gen"
MAIN_COPY = os.path.join(REPO, "菲律宾重算_202609", "2026年东南亚销量表Southeast Asia Sales table.xlsx")
MONTH_COPY = os.path.join(REPO, "菲律宾重算_202609", "月销售总表.xlsx")
SUB_PH = os.path.join(REPO, "东南亚次表分发_202609", "Philippines_202609_sales.xlsx")
TRUNK_SEA = os.path.join(REPO, "trunk", "result", "东南亚销量表", "202609",
                         "2026年东南亚销量表Southeast Asia Sales table.xlsx")
TRUNK_MONTH = os.path.join(REPO, "trunk", "result", "数据归属周期", "202609", "月销售总表.xlsx")


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


def sku_map(ws, col, r0):
    m = {}
    for r in range(r0, ws.max_row + 1):
        v = ws.cell(r, 1).value
        if v is not None and str(v).strip():
            m[str(v).strip()] = n0(ws.cell(r, col).value)
    return m


def col_sum(ws, col, r0):
    return sum(n0(ws.cell(r, col).value) for r in range(r0, ws.max_row + 1))


print("=" * 78)
print("[1] trunk 月销总表 vs trunk 东南亚表（内部一致性）")
wt = openpyxl.load_workbook(TRUNK_SEA, data_only=True)
pht = wt["Philippines 菲律宾销量表"]
bd_trunk = sku_map(pht, 56, 3)
wmt = openpyxl.load_workbook(TRUNK_MONTH, data_only=True)
wst = wmt["2026年"]
em_trunk = sku_map(wst, 143, 4)
print(f"    trunk 东南亚表 菲 BD 行2 = {pht.cell(2, 56).value}，数据行和 = {sum(bd_trunk.values()):.0f}")
print(f"    trunk 月销 EM 数据行和 = {col_sum(wst, 143, 4):.0f}，HS = {col_sum(wst, 227, 4):.0f}，"
      f"KA = {col_sum(wst, 287, 4):.0f}，KJ = {col_sum(wst, 296, 4):.0f}")
mism_t = [(s, bd_trunk.get(s, 0), em_trunk.get(s, 0)) for s in set(bd_trunk) | set(em_trunk)
          if bd_trunk.get(s, 0) != em_trunk.get(s, 0)]
net_t = sum(em_trunk.get(s, 0) - bd_trunk.get(s, 0) for s in set(bd_trunk) | set(em_trunk))
print(f"    trunk 内部 SKU 级差异: {len(mism_t)} 个，净差 = {net_t:+.0f}")
for s, a, b in mism_t[:8]:
    print(f"      {s}: 东南亚表BD={a:.0f} 月销EM={b:.0f} 差={b - a:+.0f}")
wt.close()
wmt.close()

print()
print("=" * 78)
print("[2] 交付副本逐 SKU 三方对照（主表 BD / 次表 H / 月销 EM）")
wm = openpyxl.load_workbook(MAIN_COPY, data_only=True)
ph = wm["Philippines 菲律宾销量表"]
bd = sku_map(ph, 56, 3)
wmc = openpyxl.load_workbook(MONTH_COPY, data_only=True)
wsc = wmc["2026年"]
em = sku_map(wsc, 143, 4)
wp = openpyxl.load_workbook(SUB_PH, data_only=True)
s2 = None
for sn in wp.sheetnames:
    if "Philippines" in sn:
        s2 = wp[sn]
h_sub = sku_map(s2, 8, 3)
print(f"    主表 BD 数据行和 = {sum(bd.values()):.0f}；次表 H 数据行和 = {sum(h_sub.values()):.0f}；"
      f"月销 EM 数据行和 = {sum(em.values()):.0f}")
m1 = [(s, bd.get(s, 0), h_sub.get(s, 0)) for s in set(bd) | set(h_sub)
      if bd.get(s, 0) != h_sub.get(s, 0)]
print(f"    次表 H vs 主表 BD 差异: {len(m1)} 个")
for s, a, b in m1[:10]:
    print(f"      {s}: 主表BD={a:.0f} 次表H={b:.0f} 差={b - a:+.0f}")
m2 = [(s, bd.get(s, 0), em.get(s, 0)) for s in set(bd) | set(em)
      if bd.get(s, 0) != em.get(s, 0)]
net2 = sum(em.get(s, 0) - bd.get(s, 0) for s in set(bd) | set(em))
print(f"    月销 EM vs 主表 BD 差异: {len(m2)} 个，净差 = {net2:+.0f}")
for s, a, b in m2[:25]:
    print(f"      {s}: 主表BD={a:.0f} 月销EM={b:.0f} 差={b - a:+.0f}")
only_em = sorted(set(em) - set(bd))
only_bd = sorted(set(bd) - set(em))
print(f"    只在月销有的 SKU: {only_em[:12]}（{len(only_em)} 个）")
print(f"    只在主表明的 SKU: {only_bd[:12]}（{len(only_bd)} 个）")
wm.close()
wmc.close()
wp.close()

print()
print("=" * 78)
print("[3] 差异 SKU 的渠道明细（主表 9 月块 AY~BD = 51~56）")
if m2:
    wm2 = openpyxl.load_workbook(MAIN_COPY, data_only=True)
    ws = wm2["Philippines 菲律宾销量表"]
    row_of = {}
    for r in range(3, ws.max_row + 1):
        v = ws.cell(r, 1).value
        if v is not None and str(v).strip():
            row_of[str(v).strip()] = r
    for s, a, b in m2[:15]:
        r = row_of.get(s)
        if r:
            vals = [n0(ws.cell(r, c).value) for c in range(51, 57)]
            print(f"      {s} 行{r}: 中山={vals[0]:.0f} LZD={vals[1]:.0f} SHP={vals[2]:.0f} "
                  f"TK={vals[3]:.0f} 线下={vals[4]:.0f} 合计={vals[5]:.0f}（月销EM={b:.0f}）")
    wm2.close()
