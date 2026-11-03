# -*- coding: utf-8 -*-
"""
probe_explain_154.py — 回应用户"我想知道你在算什么"（2026-10-16）

把菲律宾 9 月 3,483 的构成逐层拆开，全部只读复验：
  A. 用户截图 154 单（17 SKU，9 SEP 2026）逐 SKU 落库验证：
     主表 Philippines BC(线下) / 菲次表 sheet2 G(Offline)，与 145 手工单合并后 = 299
  B. 交付值 vs trunk 原值：菲律宾 9 月各渠道增量 → 总增量 +360 的分解
  C. 恒等链：菲律宾 BD2 == SEA Total CJ2 == 月销 EM；KA 62,837 → 63,197
  D. 未动部分确认：印尼/泰次表、trunk 菲律宾 3,123、trunk 月销 KA 62,837
"""
import os

import openpyxl

REPO = r"D:\Desktop\work-main\sale_file_gen"
MAIN_COPY = os.path.join(REPO, "菲律宾重算_202609", "2026年东南亚销量表Southeast Asia Sales table.xlsx")
MONTH_COPY = os.path.join(REPO, "菲律宾重算_202609", "月销售总表.xlsx")
SUB_PH = os.path.join(REPO, "东南亚次表分发_202609", "Philippines_202609_sales.xlsx")
TRUNK_SEA = os.path.join(REPO, "trunk", "result", "东南亚销量表", "202609",
                         "2026年东南亚销量表Southeast Asia Sales table.xlsx")
TRUNK_MONTH = os.path.join(REPO, "trunk", "result", "数据归属周期", "202609", "月销售总表.xlsx")

# 用户截图（9 SEP 2026，RESELLER 154 单，17 SKU）
RESELLER_154 = {"G564": 20, "G565": 50, "G177": 7, "G300": 3, "G398": 8, "G400": 3,
                "G461": 3, "G343": 2, "G510": 5, "G484": 5, "G426": 7, "G356": 10,
                "G551": 11, "G513": 1, "G374": 1, "G445": 8, "G443": 10}
# 9/24 手工大单（运单 YFDZ123456789，订单 403990300315398144，145 件，15 SKU）
MANUAL_145 = {"G376": 9, "G397": 10, "G401": 7, "G418": 15, "G427": 2, "G443": 9,
              "G445": 5, "G474": 12, "G484": 20, "G497": 8, "G530": 10, "G557": 7,
              "G567": 10, "G568": 10, "G569": 11}
EXP_OFFLINE = {s: MANUAL_145.get(s, 0) + RESELLER_154.get(s, 0)
               for s in set(MANUAL_145) | set(RESELLER_154)}

ok = True


def check(cond, msg):
    global ok
    print(("  ✓ " if cond else "  ✗ ") + msg)
    if not cond:
        ok = False


def sku_map(ws, col, r0):
    m = {}
    for r in range(r0, ws.max_row + 1):
        v = ws.cell(r, 1).value
        if v is not None and str(v).strip():
            m[str(v).strip()] = ws.cell(r, col).value or 0
    return m


print("=" * 78)
print("[A] 154 单逐 SKU 落库验证（截图口径）")
print(f"    截图 17 SKU 合计 = {sum(RESELLER_154.values())}（应 154）")
check(sum(RESELLER_154.values()) == 154, "截图 17 行数量合计 = 154")

wm = openpyxl.load_workbook(MAIN_COPY, data_only=True)
ph = wm["Philippines 菲律宾销量表"]
off_main = sku_map(ph, 55, 3)      # BC 线下
shp_main = sku_map(ph, 53, 3)      # BA Shopee
tot_main = sku_map(ph, 56, 3)      # BD 9月合计

wp = openpyxl.load_workbook(SUB_PH, data_only=True)
s2 = None
for sn in wp.sheetnames:
    if "Philippines" in sn:
        s2 = wp[sn]
off_sub = sku_map(s2, 7, 3)        # G Offline
tot_sub = sku_map(s2, 8, 3)        # H Total

bad = []
for s, q in EXP_OFFLINE.items():
    vm, vs = off_main.get(s), off_sub.get(s)
    if vm != q or vs != q:
        bad.append((s, q, vm, vs))
check(not bad, f"29 个线下 SKU（145+154 合并）主表 BC 与次表 G 逐格一致"
      + ("" if not bad else f"，异常: {bad[:6]}"))
# 线下列里不允许有第 30 个非零 SKU
extra = {s: v for s, v in off_sub.items() if s not in EXP_OFFLINE and v}
check(not extra, f"次表线下列无额外非零 SKU" + ("" if not extra else f": {extra}"))
check(sum(off_sub.values()) == 299, f"次表线下列总和 = {sum(off_sub.values())}（应 299 = 145+154）")
check(sum(off_main.values()) == 299, f"主表 BC 列总和 = {sum(off_main.values())}（应 299）")
# 154 单本身（17 SKU）在主线列的份额
check(sum(RESELLER_154.values()) == 154, "154 单数量合计自证")
overlap = sorted(set(MANUAL_145) & set(RESELLER_154))
print(f"    145 与 154 重叠 SKU: {overlap}"
      f"（合并后 {', '.join(f'{s}={EXP_OFFLINE[s]}' for s in overlap)}）")

print()
print("[B] 交付值 vs trunk 原值（菲律宾 9 月，行2 合计行）")
wt = openpyxl.load_workbook(TRUNK_SEA, data_only=True)
pht = wt["Philippines 菲律宾销量表"]
CH = [("中山仓库(跨境Lazada)", 51), ("菲律宾仓Lazada", 52), ("菲律宾仓Shopee", 53),
      ("TikTok", 54), ("线下订单", 55), ("9月合计", 56)]
print(f"    {'渠道':<22}{'trunk':>8}{'交付':>8}{'增量':>8}")
tot_d = 0
for name, c in CH:
    tv = pht.cell(2, c).value or 0
    dv = ph.cell(2, c).value or 0
    tot_d += dv - tv
    print(f"    {name:<22}{tv:>8}{dv:>8}{dv - tv:>+8}")
check(tot_d == 360, f"各渠道增量合计 = {tot_d}（应 +360）")
print(f"    → trunk 3,123 + 61（旧报告漏单：监控表9/1~9/12 多18 + 新报告9/13~9/30 多43）"
      f" + 299（线下：145手工 + 154 RESELLER） = 3,483")

print()
print("[C] 恒等链与月销总表")
st = wm["SEA Total 东南亚总销量表"]
check(ph.cell(2, 56).value == 3483, f"菲律宾 BD2 = {ph.cell(2, 56).value}")
check(st.cell(2, 88).value == 3483, f"SEA Total CJ2(菲9月) = {st.cell(2, 88).value}")
check(st.cell(2, 87).value == 11308, f"SEA Total CI2(东南亚9月) = {st.cell(2, 87).value}")
wmc = openpyxl.load_workbook(MONTH_COPY, data_only=True)
wsc = wmc["2026年"]


def col_sum(ws, col, r0):
    return sum(v for v in (ws.cell(r, col).value for r in range(r0, ws.max_row + 1))
               if isinstance(v, (int, float)))


check(col_sum(wsc, 143, 4) == 3483, f"月销 EM(菲9月) 数据行和 = {col_sum(wsc, 143, 4)}")
check(col_sum(wsc, 227, 4) == 11308, f"月销 HS(东南亚9月) 数据行和 = {col_sum(wsc, 227, 4)}")
ka_c = col_sum(wsc, 287, 4)
check(ka_c == 63197, f"月销 KA(总计9月) 数据行和 = {ka_c}")
# SKU 级恒等：月销 EM == 次表 H == 主表 BD（G/K SKU）
em_map = sku_map(wsc, 143, 4)
mism = [s for s in tot_sub if s in em_map and em_map[s] != tot_sub[s]]
check(not mism, f"SKU 级恒等 月销EM == 次表H == 主表BD"
      + ("" if not mism else f"，差异 {mism[:6]}"))
wt.close()
wm.close()
wp.close()

print()
print("[D] 未动部分确认")
wmt = openpyxl.load_workbook(TRUNK_MONTH, data_only=True)
wst = wmt["2026年"]
ka_t = col_sum(wst, 287, 4)
check(ka_t == 62837, f"trunk 月销 KA(9月) = {ka_t}（未动，原值）")
wmt.close()
wi = openpyxl.load_workbook(os.path.join(REPO, "东南亚次表分发_202609", "Indonesia_202609_sales.xlsx"), data_only=True)
si2 = None
for sn in wi.sheetnames:
    if "Indonesia" in sn:
        si2 = wi[sn]
check(si2.cell(2, 38).value == 4026, f"印尼次表 sheet2 9月块 AL2 = {si2.cell(2, 38).value}（未动）")
wi.close()
wth = openpyxl.load_workbook(os.path.join(REPO, "东南亚次表分发_202609", "Thailand_202609_sales.xlsx"), data_only=True)
st2 = None
for sn in wth.sheetnames:
    if "Thailand" in sn:
        st2 = wth[sn]
check(st2.cell(2, 7).value == 3238, f"泰次表 sheet2 9月块 G2 = {st2.cell(2, 7).value}（未动）")
wth.close()

print()
print("=" * 78)
print("总结：" + ("全部通过 ✓" if ok else "存在失败项 ✗"))
print(f"月销 KA: trunk {ka_t} → 交付 {ka_c}（Δ+{ka_c - ka_t} = 61 漏单 + 299 线下）")
