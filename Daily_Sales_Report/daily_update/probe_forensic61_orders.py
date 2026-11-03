# -*- coding: utf-8 -*-
"""probe_forensic61_orders.py — 61 件订单级取证（只读，回答"61 来源哪里/原统计是否出错"）

段A（9/1~9/12）：监控表 SHP/LZD/TK vs 旧三报告同段；旧报告另存全月键，查"时间移位"。
段B（≥9/13）：新订单报告（非手工，已交运/已发货）vs 旧三报告同段。
按 (平台, 订单号, SKU) 匹配，分类：
  · 真实漏单——旧报告全月都没有这笔
  · 时间移位——旧报告有这笔但在其他日期
  · 旧报告独有——监控表/新报告没有
  · 数量不一致——两边都有但件数不同
"""
import os
import re
import datetime as dt

import pandas as pd
from openpyxl import load_workbook

SNAP = r"D:\Desktop\work-main\sale_file_gen\trunk\data\数据归属周期\202609\快照日期\20261006\东南亚"
MON = r"C:\Users\dell\Desktop\DAILY MONITORING - SEPTEMBER 1-13(1).xlsx"
NEWRPT = os.path.join(SNAP, "菲律宾202609订单报告.xlsx")
D0, D1, D2 = "2026-09-01", "2026-09-12", "2026-09-13"
SEP_LO, SEP_HI = "2026-09-01", "2026-09-30"

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
    """→ 'YYYY-MM-DD'（支持 datetime / '2026-09-01 00:05' / '2026-09-01'）"""
    if v is None:
        return ""
    if hasattr(v, "strftime"):
        return v.strftime("%Y-%m-%d")
    s = str(v).strip()
    m = re.match(r"(\d{4}-\d{2}-\d{2})", s)
    return m.group(1) if m else ""


# ---------------------------------------------------------------- 读旧三报告（全月，订单级）
def read_old():
    out = {"Shopee": [], "Lazada": [], "Tiktok": []}   # (oid, sku, qty, date)
    d = pd.read_excel(os.path.join(SNAP, "Shopee菲律宾站点202609订单报告.xlsx"), dtype=str)
    d["订单状态"] = d["Order Status"].astype(str).str.strip()
    d.loc[d["订单状态"].str.startswith(RW_PREFIX), "订单状态"] = "Completed"
    d["st"] = d["订单状态"].map(SHOPEE_MAP)
    d = d[(d["st"].notna()) & (d["st"] != "Cancelled")]
    for _, r in d.iterrows():
        out["Shopee"].append((nid(r["Order ID"]), nsku(r["SKU Reference No."]),
                              q10(r["Quantity"]), d10(r["Order Creation Date"])))
    d = pd.read_excel(os.path.join(SNAP, "Lazada菲律宾站点202609订单报告.xlsx"), dtype=str)
    d["st"] = d["status"].astype(str).str.strip()
    d = d[d["st"].isin(LZD_KEEP)]
    ct = pd.to_datetime(d["createTime"], errors="coerce")
    for (_, r), c in zip(d.iterrows(), ct):
        out["Lazada"].append((nid(r["orderNumber"]), nsku(r["sellerSku"]), 1.0,
                              c.strftime("%Y-%m-%d") if pd.notna(c) else ""))
    d = pd.read_csv(os.path.join(SNAP, "Tiktok菲律宾站点202609订单报告.csv"), dtype=str, low_memory=False)
    d["st"] = d["Order Status"].astype(str).str.strip().map(TK_MAP)
    d = d[d["st"].isin(TK_KEEP)]
    ct = pd.to_datetime(d["Created Time"].astype(str).str.strip(), errors="coerce")
    for (_, r), c in zip(d.iterrows(), ct):
        out["Tiktok"].append((nid(r["Order ID"]), nsku(r["Seller SKU"]),
                              q10(r["Quantity"]), c.strftime("%Y-%m-%d") if pd.notna(c) else ""))
    return out


# ---------------------------------------------------------------- 读监控表三平台（9 月，订单级）
def read_monitor():
    out = {"Shopee": [], "Lazada": [], "Tiktok": []}
    wb = load_workbook(MON, read_only=True, data_only=True)
    ws = wb["SHP Data"]
    for r in ws.iter_rows(min_row=2, values_only=True):
        d = d10(r[0])
        if not (SEP_LO <= d <= SEP_HI):
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
        if not (SEP_LO <= d <= SEP_HI):
            continue
        if str(r[5] or "").strip() not in LZD_KEEP:
            continue
        out["Lazada"].append((nid(r[0]), nsku(r[1]), 1.0, d))
    ws = wb["Tiktok Data"]
    for r in ws.iter_rows(min_row=2, values_only=True):
        d = d10(r[6])
        if not (SEP_LO <= d <= SEP_HI):
            continue
        st2 = TK_MAP.get(str(r[1] or "").strip())
        if st2 not in TK_KEEP:
            continue
        out["Tiktok"].append((nid(r[0]), nsku(r[2]), q10(r[5]), d))
    wb.close()
    return out


# ---------------------------------------------------------------- 读新报告（≥9/13 非手工，订单级）
def read_new():
    raw = pd.read_excel(NEWRPT, dtype=str)
    _line = ["订单ID", "Sku", "数量", "多属性Id", "平台数量", "sku图片", "sku中文名称"]
    _parent = [c for c in raw.columns if c not in _line]
    raw[_parent] = raw[_parent].ffill()
    v = raw[raw["订单状态"].isin(["已交运", "已发货"])].copy()
    v["sku"] = v["Sku"].map(nsku)
    v["d"] = v["创建时间"].map(d10)
    v["q"] = pd.to_numeric(v["数量"], errors="coerce").fillna(0)
    v["plat"] = v["平台类型"].fillna("手工订单").astype(str).str.strip()
    v["oid"] = v["交易号"].map(nid)
    return v


def key(rows):
    """(oid, sku) -> [qty...]（同键多行合并）"""
    m = {}
    for oid, sku, q, d in rows:
        if not oid or not sku:
            continue
        m.setdefault((oid, sku), []).append((q, d))
    return m


def merge(m):
    return {k: sum(q for q, _ in v) for k, v in m.items()}


def datemap(m):
    return {k: sorted({d for _, d in v}) for k, v in m.items()}


def seg(rows, lo, hi):
    return [r for r in rows if lo <= r[3] <= hi]


print("=" * 78)
print("读原始数据……")
OLD = read_old()
MONR = read_monitor()
NEW = read_new()
for p in ("Shopee", "Lazada", "Tiktok"):
    print(f"  旧报告 {p:<7} 全月行数={len(OLD[p]):>6}  件数={sum(r[2] for r in OLD[p]):>8.0f}"
          f" ｜ 监控表 9月行数={len(MONR[p]):>6}  件数={sum(r[2] for r in MONR[p]):>8.0f}")
segB_new = NEW[(NEW["d"] >= D2) & (NEW["plat"] != "手工订单")]
print(f"  新报告 ≥9/13 非手工行数={len(segB_new)}  件数={segB_new['q'].sum():.0f}"
      f"（其中手工单另行={int((NEW[(NEW['d'] >= D2) & (NEW['plat'] == '手工订单')].shape[0]))} 行）")

# 新报告按平台拆（平台类型 → 旧报告平台）
PLATMAP = {"Shopee": "Shopee", "Lazada": "Lazada", "TikTok": "Tiktok", "Tiktok shop": "Tiktok"}
print(f"  新报告平台类型分布（≥9/13 非手工）: {segB_new['plat'].value_counts().to_dict()}")
print(f"  新报告 交易号 样例: {segB_new['oid'].head(3).tolist()}")

print()
print("=" * 78)
print("段A（9/1~9/12）：监控表 vs 旧三报告同段")
print("=" * 78)
for p in ("Shopee", "Lazada", "Tiktok"):
    mA = merge(key(seg(MONR[p], D0, D1)))
    oA = merge(key(seg(OLD[p], D0, D1)))
    oAll = key(OLD[p])           # 旧报告全月（查时间移位）
    oAllD = datemap(oAll)
    mAllD = datemap(key(MONR[p]))
    only_m = {k: v for k, v in mA.items() if k not in oA}
    only_o = {k: v for k, v in oA.items() if k not in mA}
    qdiff = {k: (mA[k], oA[k]) for k in set(mA) & set(oA) if abs(mA[k] - oA[k]) > 1e-9}
    net = sum(mA.values()) - sum(oA.values())
    print(f"\n--- {p}: 监控表 {sum(mA.values()):.0f} 件 vs 旧报告 {sum(oA.values()):.0f} 件，净差 {net:+.0f}")
    # 监控表独有：旧报告全月有没有？
    miss, shift = [], []
    for k, v in only_m.items():
        if k in oAll:
            shift.append((k, v, oAllD[k]))
        else:
            miss.append((k, v))
    print(f"    监控表独有 {len(only_m)} 键 / {sum(only_m.values()):.0f} 件"
          f" → 旧报告全月也没有(真实漏单) {len(miss)} 键 / {sum(v for _, v in miss):.0f} 件；"
          f"旧报告其他日期有(时间移位) {len(shift)} 键")
    for k, v, ds in shift[:8]:
        print(f"      移位 {k}: 监控表 {v:.0f} 件，旧报告日期 {ds}")
    for k, v in miss[:12]:
        print(f"      漏单 {k}: {v:.0f} 件")
    if len(miss) > 12:
        print(f"      ……漏单还有 {len(miss) - 12} 键")
    # 旧报告独有：监控表全月有没有？
    o_miss, o_shift = [], []
    for k, v in only_o.items():
        if k in mAllD:
            o_shift.append((k, v, mAllD[k]))
        else:
            o_miss.append((k, v))
    print(f"    旧报告独有 {len(only_o)} 键 / {sum(only_o.values()):.0f} 件"
          f" → 监控表全月也没有 {len(o_miss)} 键 / {sum(v for _, v in o_miss):.0f} 件；"
          f"监控表其他日期有 {len(o_shift)} 键")
    for k, v in o_miss[:8]:
        print(f"      旧独有 {k}: {v:.0f} 件")
    if qdiff:
        print(f"    数量不一致 {len(qdiff)} 键: "
              + "; ".join(f"{k}:监{mA[k]:.0f}/旧{oA[k]:.0f}" for k in list(qdiff)[:6]))

print()
print("=" * 78)
print("段B（≥9/13）：新订单报告(非手工) vs 旧三报告同段")
print("=" * 78)
# 新报告按平台分组（oid 形态决定匹配键）
new_by_plat = {}
for plat, g in segB_new.groupby("plat"):
    new_by_plat[plat] = g
for plat, g in sorted(new_by_plat.items()):
    p = PLATMAP.get(plat)
    if p is None:
        print(f"--- 新报告平台 {plat!r}: 无对应旧报告平台，跳过匹配（{len(g)} 行 / {g['q'].sum():.0f} 件）")
        continue
    nB = {}
    for _, r in g.iterrows():
        if r["oid"] and r["sku"]:
            nB[(r["oid"], r["sku"])] = nB.get((r["oid"], r["sku"]), 0) + r["q"]
    oB = merge(key(seg(OLD[p], D2, SEP_HI)))
    oAll = key(OLD[p])
    oAllD = datemap(oAll)
    only_n = {k: v for k, v in nB.items() if k not in oB}
    only_o = {k: v for k, v in oB.items() if k not in nB}
    qdiff = {k: (nB[k], oB[k]) for k in set(nB) & set(oB) if abs(nB[k] - oB[k]) > 1e-9}
    net = sum(nB.values()) - sum(oB.values())
    print(f"\n--- {plat}→{p}: 新报告 {sum(nB.values()):.0f} 件 vs 旧报告 {sum(oB.values()):.0f} 件，净差 {net:+.0f}")
    miss, shift = [], []
    for k, v in only_n.items():
        if k in oAll:
            shift.append((k, v, oAllD[k]))
        else:
            miss.append((k, v))
    print(f"    新报告独有 {len(only_n)} 键 / {sum(only_n.values()):.0f} 件"
          f" → 旧报告全月也没有(真实漏单) {len(miss)} 键 / {sum(v for _, v in miss):.0f} 件；"
          f"旧报告其他日期有(时间移位) {len(shift)} 键")
    for k, v, ds in shift[:8]:
        print(f"      移位 {k}: 新报告 {v:.0f} 件，旧报告日期 {ds}")
    for k, v in miss[:15]:
        print(f"      漏单 {k}: {v:.0f} 件")
    if len(miss) > 15:
        print(f"      ……漏单还有 {len(miss) - 15} 键")
    print(f"    旧报告独有 {len(only_o)} 键 / {sum(only_o.values()):.0f} 件")
    for k, v in list(only_o.items())[:8]:
        print(f"      旧独有 {k}: {v:.0f} 件")
    if qdiff:
        print(f"    数量不一致 {len(qdiff)} 键: "
              + "; ".join(f"{k}:新{nB[k]:.0f}/旧{oB[k]:.0f}" for k in list(qdiff)[:6]))

# 手工单（新报告 ≥9/13）— 确认它们旧报告确实没有
man = NEW[(NEW["d"] >= D2) & (NEW["plat"] == "手工订单")]
print(f"\n--- 手工单（新报告 ≥9/13）: {len(man)} 行 / {man['q'].sum():.0f} 件"
      f"（这批 10-16 已归「线下订单」列，与 +61 无关）")
