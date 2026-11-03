# -*- coding: utf-8 -*-
"""probe_final_caliber_v1.py — 回答用户：「按这份新报告 + 299 件手工单，最终菲/总计多少」

两种算法都算，并逐项验证为什么不能只用新报告：
  X. 字面口径：新报告全月三平台(有效) + 跨境168 + 线下299
  Y. 两段式口径：监控表段A(9/1~9/12) + 新报告段B(≥9/13) + 跨境168 + 线下299
  Y+. 两段式并集：段A 监控表∪新报告 + 段B新报告 + 跨境168 + 线下299
验证点：
  1. 新报告全月/分段三平台有效件数（手工单已排除，防与 299 双计）
  2. 段A 新报告 vs 监控表：哪些监控表有新报告没有（=字面口径会丢的件）
  3. G554×2 的双记载确认（新报告+旧报告有、监控表无）
  4. 月销总计线性换算：Δ菲 = Δ东南亚 = Δ总计（trunk 基线 3,123 / 10,948 / 62,837）
"""
import os
import re
import pandas as pd
from openpyxl import load_workbook

SNAP = r"D:\Desktop\work-main\sale_file_gen\trunk\data\数据归属周期\202609\快照日期\20261006\东南亚"
MON = r"C:\Users\dell\Desktop\DAILY MONITORING - SEPTEMBER 1-13(1).xlsx"
NEWRPT = os.path.join(SNAP, "菲律宾202609订单报告.xlsx")
D0, D1, D2 = "2026-09-01", "2026-09-12", "2026-09-13"
SEP_HI = "2026-09-30"

SHOPEE_MAP = {
    "Completed": "Completed", "Order Received": "Completed",
    "Shipping": "Shipping", "To ship": "Shipping", "Delivered": "Delivered",
    "Cancelled": "Cancelled", "Unpaid": "Cancelled", "Cancellation requested": "Cancelled",
}
RW_PREFIX = "Order Received, but buyer can raise return/refund"
LZD_KEEP = {"confirmed", "shipped", "delivered", "packed", "pending",
            "ready_to_ship", "Lost by 3PL", "Package scrapped"}
TK_MAP = {
    "Completed": "Completed", "已完成": "Completed", "Selesai": "Completed",
    "Shipped": "Shipped", "已发货": "Shipped", "Dikirim": "Shipped",
    "To ship": "To ship", "待发货": "To ship", "Perlu dikirim": "To ship",
    "Canceled": "Canceled", "Cancelled": "Cancelled", "已取消": "Canceled",
    "Dibatalkan": "Canceled", "Unpaid": "Canceled", "Belum dibayar": "Canceled",
}
TK_KEEP = {"Shipped", "Completed", "To ship"}
PLATMAP = {"Shopee": "Shopee", "Lazada": "Lazada", "TikTok": "Tiktok", "Tiktok shop": "Tiktok"}


def nsku(v):
    s = str(v or "").strip().upper()
    s = re.sub(r"^\d+-", "", s)
    return "" if s in ("", "NAN", "NONE") else s


def nid(v):
    if v is None:
        return ""
    if isinstance(v, float) and v == int(v):
        v = int(v)
    s = str(v).strip().upper()
    return "" if s in ("", "NAN", "NONE") else s


def q10(v):
    try:
        return float(v)
    except Exception:
        return 0.0


def d10(v):
    if v is None:
        return ""
    if hasattr(v, "strftime"):
        return v.strftime("%Y-%m-%d")
    s = str(v).strip()
    m = re.match(r"(\d{4}-\d{2}-\d{2})", s)
    return m.group(1) if m else ""


# ---------------------------------------------------------------- 新订单报告（全集）
def read_new():
    raw = pd.read_excel(NEWRPT, dtype=str)
    _line = ["订单ID", "Sku", "数量", "多属性Id", "平台数量", "sku图片", "sku中文名称"]
    _parent = [c for c in raw.columns if c not in _line]
    raw[_parent] = raw[_parent].ffill()
    v = raw[raw["订单状态"].isin(["已交运", "已发货"])].copy()
    v["sku"] = v["Sku"].map(nsku)
    v["d"] = v["创建时间"].map(d10)
    v["q"] = pd.to_numeric(v["数量"], errors="coerce").fillna(0)
    v["plat"] = v["平台类型"].fillna("").astype(str).str.strip()
    v["otype"] = v["订单类型"].fillna("").astype(str).str.strip()
    v["oid"] = v["交易号"].map(nid)
    v["is_man"] = v["otype"] == "手工订单"
    return v


# ---------------------------------------------------------------- 监控表（9 月）
def read_monitor():
    out = {"Shopee": [], "Lazada": [], "Tiktok": []}
    wb = load_workbook(MON, read_only=True, data_only=True)
    ws = wb["SHP Data"]
    for r in ws.iter_rows(min_row=2, values_only=True):
        d = d10(r[0])
        if not (D0 <= d <= SEP_HI):
            continue
        st = str(r[2] or "").strip()
        if st.startswith(RW_PREFIX):
            st = "Completed"
        st2 = SHOPEE_MAP.get(st)
        if st2 in (None, "Cancelled"):
            continue
        out["Shopee"].append((nid(r[1]), nsku(r[4]), q10(r[6]), d))
    ws = wb["LZD Data"]
    for r in ws.iter_rows(min_row=2, values_only=True):
        d = d10(r[2])
        if not (D0 <= d <= SEP_HI):
            continue
        if str(r[5] or "").strip() not in LZD_KEEP:
            continue
        out["Lazada"].append((nid(r[0]), nsku(r[1]), 1.0, d))
    ws = wb["Tiktok Data"]
    for r in ws.iter_rows(min_row=2, values_only=True):
        d = d10(r[6])
        if not (D0 <= d <= SEP_HI):
            continue
        st2 = TK_MAP.get(str(r[1] or "").strip())
        if st2 not in TK_KEEP:
            continue
        out["Tiktok"].append((nid(r[0]), nsku(r[2]), q10(r[5]), d))
    wb.close()
    return out


NEW = read_new()
MONR = read_monitor()

# 新报告三平台桶（非手工、有效状态）
NEW3 = NEW[(~NEW["is_man"]) & (NEW["plat"].map(lambda x: PLATMAP.get(x, x)).notna())].copy()
NEW3["p"] = NEW3["plat"].map(lambda x: PLATMAP.get(x, x))
segA_new = NEW3[(NEW3["d"] >= D0) & (NEW3["d"] <= D1)]
segB_new = NEW3[(NEW3["d"] >= D2) & (NEW3["d"] <= SEP_HI)]

print("=" * 84)
print("① 新报告自身结构（有效单 = 已交运/已发货）")
print("=" * 84)
print(f"  全月三平台(非手工): {NEW3['q'].sum():.0f} 件｜手工单: {NEW[NEW['is_man']]['q'].sum():.0f} 件")
print(f"  分段A 9/1~9/12: {segA_new['q'].sum():.0f} 件（其中手工 {segA_new['is_man'].sum()} 行）")
print(f"  分段B ≥9/13   : {segB_new['q'].sum():.0f} 件")
print(f"  段A明细: " + " / ".join(
    f"{p} {g['q'].sum():.0f}" for p, g in segA_new.groupby("p")))
print(f"  段B明细: " + " / ".join(
    f"{p} {g['q'].sum():.0f}" for p, g in segB_new.groupby("p")))

# 段A 新报告 vs 监控表 逐键重叠
print("\n" + "=" * 84)
print("② 段A（9/1~9/12）新报告 vs 监控表：逐键重叠（字面口径会丢多少）")
print("=" * 84)
a_new_keys, a_mon_keys = set(), set()
for _, r in segA_new.iterrows():
    if r["oid"] and r["sku"]:
        a_new_keys.add((r["oid"], r["sku"]))
for p in ("Shopee", "Lazada", "Tiktok"):
    for oid, sku, q, d in MONR[p]:
        if D0 <= d <= D1 and oid and sku:
            a_mon_keys.add((oid, sku))
only_new = a_new_keys - a_mon_keys
only_mon = a_mon_keys - a_new_keys
print(f"  新报告段A键 {len(a_new_keys)} 个｜监控表段A键 {len(a_mon_keys)} 个")
print(f"  交集 {len(a_new_keys & a_mon_keys)}｜仅新报告有 {len(only_new)}｜仅监控表有 {len(only_mon)}")
# 仅新报告有的件数
q_only_new = 0.0
for _, r in segA_new.iterrows():
    if (r["oid"], r["sku"]) in only_new:
        q_only_new += r["q"]
print(f"  仅新报告有的件数 = {q_only_new:.0f}（新报告+旧报告双记载、监控表无 → 即 G554×2）")
for k in sorted(only_new):
    for _, r in segA_new.iterrows():
        if (r["oid"], r["sku"]) == k:
            print(f"    {k[0]:<24} {k[1]:<8} {r['q']:.0f} 件 {r['d']} {r['p']}")
            break
# 仅监控表有的件数（按平台）
q_only_mon = 0.0
for p in ("Shopee", "Lazada", "Tiktok"):
    qp = sum(q for oid, sku, q, d in MONR[p]
             if D0 <= d <= D1 and (oid, sku) in only_mon)
    if qp:
        print(f"  仅监控表有 {p}: {qp:.0f} 件")
    q_only_mon += qp
print(f"  ⇒ 仅监控表有合计 {q_only_mon:.0f} 件 = 字面口径(只用新报告)会丢失的段A订单")

print("\n" + "=" * 84)
print("③ 最终口径计算")
print("=" * 84)
monA = {p: sum(q for oid, sku, q, d in MONR[p] if D0 <= d <= D1) for p in ("Shopee", "Lazada", "Tiktok")}
monA_tot = sum(monA.values())
newB = segB_new["q"].sum()
CROSS = 168.0   # Lazada 跨境菲律宾（快照/东南亚/跨境/ 单独文件，新旧口径都不动）
OFFLINE = 299.0  # 手工单 145（新报告内）+ RESELLER 154（监控表）

x_lit = NEW3["q"].sum() + CROSS + OFFLINE          # 字面：只用新报告
y_2seg = monA_tot + newB + CROSS + OFFLINE         # 两段式：监控表A + 新报告B
y_plus = monA_tot + q_only_new + newB + CROSS + OFFLINE  # 两段式并集（补回 G554×2）

print(f"  段A 监控表 = {monA_tot:.0f}（SHP {monA['Shopee']:.0f} / LZD {monA['Lazada']:.0f} / TK {monA['Tiktok']:.0f}）")
print(f"  段B 新报告 = {newB:.0f}")
print(f"  新报告全月三平台 = {NEW3['q'].sum():.0f}（段A {segA_new['q'].sum():.0f} + 段B {newB:.0f}）")
print(f"  跨境(不动) = {CROSS:.0f}｜线下手工单 = {OFFLINE:.0f}（145 手工 + 154 RESELLER）")
print()
print(f"  X 字面口径（只按这份报告+299）      : 菲 = {x_lit:.0f}")
print(f"  Y 两段式（监控表A + 新报告B + 299）  : 菲 = {y_2seg:.0f}  ← 已交付口径")
print(f"  Y+ 两段式并集（再补回 G554×2）       : 菲 = {y_plus:.0f}")

print("\n" + "=" * 84)
print("④ 月销总计线性换算（trunk 基线：菲 3,123｜东南亚 10,948｜总计 KA 62,837）")
print("=" * 84)
for tag, ph in (("X 字面", x_lit), ("Y 两段式(已交付)", y_2seg), ("Y+ 并集", y_plus)):
    d = ph - 3123.0
    print(f"  {tag:<18} 菲 {ph:>7.0f} → 东南亚 {10948 + d:>8.0f} → 总计 {62837 + d:>8.0f}")

print("\n" + "=" * 84)
print("⑤ 自证：Y 必须等于 3,483（交付值）")
print("=" * 84)
print(f"  Y = {y_2seg:.0f}｜与交付值 3,483 差 {y_2seg - 3483:+.0f}")
print(f"  字面口径丢失 = 3,483 − X = {3483 - x_lit:.0f} 件（应 = 段A仅监控表有的 {q_only_mon:.0f} − 仅新报告有的 {q_only_new:.0f} = {q_only_mon - q_only_new:.0f}）")
