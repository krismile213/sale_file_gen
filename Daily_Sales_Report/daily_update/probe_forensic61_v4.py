# -*- coding: utf-8 -*-
"""probe_forensic61_v4.py — 用「运单号」复核段B Shopee 互有键（只读）

v3 结论：35 个旧独有键按 (交易号,SKU) 精确匹配全不中；其中 30 个有「同SKU同日期」的
新报告行，但候选单号明显是**别的订单」（同一天同款很正常），不能算同一笔。
《核查-菲律宾新订单报告能否一敌多.md》的判据是**运单号**（Tracking Number）——本轮补上。

同时反查 50 个新独有键的运单号在旧报告有没有（防"同一笔单两个单号"双向虚增）。
"""
import os
import re
import pandas as pd

SNAP = r"D:\Desktop\work-main\sale_file_gen\trunk\data\数据归属周期\202609\快照日期\20261006\东南亚"
NEWRPT = os.path.join(SNAP, "菲律宾202609订单报告.xlsx")
OLD_SHOPEE = os.path.join(SNAP, "Shopee菲律宾站点202609订单报告.xlsx")
D2, SEP_HI = "2026-09-13", "2026-09-30"

SHOPEE_MAP = {
    "Completed": "Completed", "Order Received": "Completed",
    "Shipping": "Shipping", "To ship": "Shipping", "Delivered": "Delivered",
    "Cancelled": "Cancelled", "Unpaid": "Cancelled", "Cancellation requested": "Cancelled",
}
RW_PREFIX = "Order Received, but buyer can raise return/refund"


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


def ntrk(v):
    s = str(v or "").strip().upper()
    return "" if s in ("", "NAN", "NONE", "-") else s


def d10(v):
    if v is None:
        return ""
    if hasattr(v, "strftime"):
        return v.strftime("%Y-%m-%d")
    s = str(v).strip()
    m = re.match(r"(\d{4}-\d{2}-\d{2})", s)
    return m.group(1) if m else ""


# ---------- 旧 Shopee ----------
od = pd.read_excel(OLD_SHOPEE, dtype=str)
od["订单状态"] = od["Order Status"].astype(str).str.strip()
od.loc[od["订单状态"].str.startswith(RW_PREFIX), "订单状态"] = "Completed"
od["st"] = od["订单状态"].map(SHOPEE_MAP)
od = od[(od["st"].notna()) & (od["st"] != "Cancelled")].copy()
od["oid"] = od["Order ID"].map(nid)
od["sku"] = od["SKU Reference No."].map(nsku)
od["d"] = od["Order Creation Date"].map(d10)
od["q"] = pd.to_numeric(od["Quantity"], errors="coerce").fillna(0)
od["trk"] = od["Tracking Number*"].map(ntrk)
print(f"旧Shopee 有运单号的行: {int((od['trk'] != '').sum())} / {len(od)}")

# ---------- 新报告 ----------
raw = pd.read_excel(NEWRPT, dtype=str)
_line = ["订单ID", "Sku", "数量", "多属性Id", "平台数量", "sku图片", "sku中文名称"]
_parent = [c for c in raw.columns if c not in _line]
raw[_parent] = raw[_parent].ffill()
nd = raw[raw["订单状态"].isin(["已交运", "已发货"])].copy()
nd["oid"] = nd["交易号"].map(nid)
nd["sku"] = nd["Sku"].map(nsku)
nd["d"] = nd["创建时间"].map(d10)
nd["q"] = pd.to_numeric(nd["数量"], errors="coerce").fillna(0)
nd["otype"] = nd["订单类型"].fillna("").astype(str).str.strip()
nd["plat"] = nd["平台类型"].fillna("").astype(str).str.strip()
nd["trk"] = nd["运单号"].map(ntrk)
shp_new = nd[nd["plat"] == "Shopee"]
print(f"新报告Shopee 有运单号的行: {int((shp_new['trk'] != '').sum())} / {len(shp_new)}")

# 运单号索引
new_by_trk = {}
for _, r in shp_new.iterrows():
    if r["trk"]:
        new_by_trk.setdefault(r["trk"], []).append(
            (r["oid"], r["sku"], r["d"], r["q"], r["otype"], r["订单状态"]))
old_by_trk = {}
for _, r in od.iterrows():
    if r["trk"]:
        old_by_trk.setdefault(r["trk"], []).append((r["oid"], r["sku"], r["d"], r["q"]))

# ---------- 段B 键集 ----------
odB = od[(od["d"] >= D2) & (od["d"] <= SEP_HI)]
oldB_keys = {}
for _, r in odB.iterrows():
    if r["oid"] and r["sku"]:
        oldB_keys.setdefault((r["oid"], r["sku"]), []).append((r["q"], r["trk"], r["d"]))
ndB = shp_new[(shp_new["d"] >= D2) & (shp_new["otype"] != "手工订单")]
newB_keys = {}
for _, r in ndB.iterrows():
    if r["oid"] and r["sku"]:
        newB_keys.setdefault((r["oid"], r["sku"]), []).append((r["q"], r["trk"], r["d"]))
only_o = {k: v for k, v in oldB_keys.items() if k not in newB_keys}
only_n = {k: v for k, v in newB_keys.items() if k not in oldB_keys}
print(f"段B Shopee：旧独有 {len(only_o)} 键；新独有 {len(only_n)} 键")

print("\n===== 35 旧独有键 → 按运单号在新报告找同一笔 =====")
hit, no_trk, miss = 0, 0, 0
for k, rows in sorted(only_o.items()):
    trks = {t for _, t, _ in rows if t}
    found = []
    for t in trks:
        found.extend(new_by_trk.get(t, []))
    if found:
        hit += 1
        print(f"  [运单号命中] {k[0]} {k[1]} → 新报告: {found[:2]}")
    elif not trks:
        no_trk += 1
        print(f"  [旧报告无运单号] {k[0]} {k[1]} {rows[0][2]}")
    else:
        miss += 1
        print(f"  [运单号也不在] {k[0]} {k[1]} {rows[0][2]} trk={trks}")
print(f"  ⇒ 命中 {hit} / 无运单号 {no_trk} / 真不在 {miss}")

print("\n===== 50 新独有键 → 按运单号在旧报告找同一笔 =====")
hit2, no_trk2, miss2 = 0, 0, 0
aas_hits = 0
for k, rows in sorted(only_n.items()):
    trks = {t for _, t, _ in rows if t}
    found = []
    for t in trks:
        found.extend(old_by_trk.get(t, []))
    is_aas = "AAS" in k[0]
    if found:
        hit2 += 1
        if is_aas:
            aas_hits += 1
        print(f"  [运单号命中] {k[0]} {k[1]} → 旧报告: {found[:2]}")
    elif not trks:
        no_trk2 += 1
    else:
        miss2 += 1
print(f"  ⇒ 命中 {hit2}（其中 AAS 长单号 {aas_hits}）/ 无运单号 {no_trk2} / 真不在 {miss2}")

# AAS/17位长单号本身有多少件
long_rows = shp_new[shp_new["oid"].str.len() > 14]
print(f"\n新报告 Shopee 长单号(>14位)行数={len(long_rows)} 件数={long_rows['q'].sum():.0f}"
      f"（AAS={int(long_rows['oid'].str.contains('AAS').sum())} 行）")
print(f"  长单号行里有运单号的: {int((long_rows['trk'] != '').sum())} 行")

# ---------- 净差复核 ----------
print("\n===== 汇总 =====")
print(f"  段B Shopee 旧独有 35 键中：运单号能在新报告找到同一笔 = {hit} 键；找不到 = {no_trk + miss} 键")
print(f"  段B Shopee 新独有 50 键中：运单号能在旧报告找到同一笔 = {hit2} 键；找不到 = {no_trk2 + miss2} 键")
