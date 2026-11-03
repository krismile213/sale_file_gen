# -*- coding: utf-8 -*-
"""probe_newreport_scope_v3.py — 补丁：① P4 下标修正（qty 在 r[2] 不是 r[0]）
② Shopee泰国站点报告用正确单号列（泰语列名，宽探测曾误选 SKU 列）补测交集"""
import os
import re
import pandas as pd

SNAP = r"D:\Desktop\work-main\sale_file_gen\trunk\data\数据归属周期\202609\快照日期\20261006\东南亚"
NEWRPT = os.path.join(SNAP, "菲律宾202609订单报告.xlsx")

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


def ntrk(v):
    s = str(v or "").strip().upper()
    return "" if s in ("", "NAN", "NONE", "-") else s


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


def read_old():
    out = {"Shopee": [], "Lazada": [], "Tiktok": []}
    d = pd.read_excel(os.path.join(SNAP, "Shopee菲律宾站点202609订单报告.xlsx"), dtype=str)
    d["订单状态"] = d["Order Status"].astype(str).str.strip()
    d.loc[d["订单状态"].str.startswith(RW_PREFIX), "订单状态"] = "Completed"
    d["st"] = d["订单状态"].map(SHOPEE_MAP)
    trk_col = "Tracking Number*" if "Tracking Number*" in d.columns else None
    for _, r in d[(d["st"].notna()) & (d["st"] != "Cancelled")].iterrows():
        out["Shopee"].append((nid(r["Order ID"]), nsku(r["SKU Reference No."]),
                              q10(r["Quantity"]), d10(r["Order Creation Date"]),
                              ntrk(r[trk_col]) if trk_col else ""))
    d = pd.read_excel(os.path.join(SNAP, "Lazada菲律宾站点202609订单报告.xlsx"), dtype=str)
    d["st"] = d["status"].astype(str).str.strip()
    keep = d[d["st"].isin(LZD_KEEP)]
    trk_col = next((c for c in d.columns if "racking" in c or "racking" in c.lower()
                    or "awb" in c.lower() or "运单" in c), None)
    for (_, r), c in zip(keep.iterrows(),
                         pd.to_datetime(keep["createTime"], errors="coerce")):
        out["Lazada"].append((nid(r["orderNumber"]), nsku(r["sellerSku"]), 1.0,
                              c.strftime("%Y-%m-%d") if pd.notna(c) else "",
                              ntrk(r[trk_col]) if trk_col else ""))
    d = pd.read_csv(os.path.join(SNAP, "Tiktok菲律宾站点202609订单报告.csv"), dtype=str, low_memory=False)
    d["st"] = d["Order Status"].astype(str).str.strip().map(TK_MAP)
    keep = d[d["st"].isin(TK_KEEP)]
    trk_col = next((c for c in d.columns if "racking" in c or "ackage" in c
                    or "运单" in c), None)
    for (_, r), c in zip(keep.iterrows(),
                         pd.to_datetime(keep["Created Time"].astype(str).str.strip(), errors="coerce")):
        out["Tiktok"].append((nid(r["Order ID"]), nsku(r["Seller SKU"]),
                              q10(r["Quantity"]), c.strftime("%Y-%m-%d") if pd.notna(c) else "",
                              ntrk(r[trk_col]) if trk_col else ""))
    return out


def read_new_full():
    raw = pd.read_excel(NEWRPT, dtype=str)
    _line = ["订单ID", "Sku", "数量", "多属性Id", "平台数量", "sku图片", "sku中文名称"]
    _parent = [c for c in raw.columns if c not in _line]
    raw[_parent] = raw[_parent].ffill()
    raw["d"] = raw["创建时间"].map(d10)
    raw["q"] = pd.to_numeric(raw["数量"], errors="coerce").fillna(0)
    raw["plat"] = raw["平台类型"].fillna("").astype(str).str.strip()
    raw["oid"] = raw["交易号"].map(nid)
    return raw


OLD = read_old()
RAW = read_new_full()
V = RAW[RAW["订单状态"].isin(["已交运", "已发货"])].copy()

print("=" * 84)
print("P4-fix 9/13~9/14 逐平台新旧对比（有效单；旧报告侧按 v5 同款状态过滤）")
print("=" * 84)
for d in ("2026-09-13", "2026-09-14"):
    for p in ("Shopee", "Lazada", "Tiktok"):
        n = V[(V["d"] == d) & (V["plat"].map(lambda x: PLATMAP.get(x, x)) == p)]["q"].sum()
        o = sum(r[2] for r in OLD[p] if r[3] == d)
        print(f"  {d} {p:<7}: 新报告 {n:>6.0f} vs 旧报告 {o:>6.0f}  ({n - o:+.0f})")

print("\n" + "=" * 84)
print("补测：Shopee泰国站点报告（泰语列名，取正确单号列）∩ 新报告Shopee桶")
print("=" * 84)
th = os.path.join(SNAP, "Shopee泰国站点202609订单报告.xlsx")
d = pd.read_excel(th, dtype=str)
cand = [c for c in d.columns if "คำสั่งซื้อ" in str(c) and "สถานะ" not in str(c)
        and "ยกเลิก" not in str(c) and "คืน" not in str(c)]
print("候选单号列:", cand)
col = cand[0]
ids = {nid(v) for v in d[col]} - {""}
bucket = set(RAW.loc[RAW["plat"].map(lambda x: PLATMAP.get(x, x)) == "Shopee", "oid"]) - {""}
inter = ids & bucket
print(f"单号样例: {[str(x)[:24] for x in d[col].dropna().head(3).tolist()]}")
print(f"泰国站点: {len(d)} 行｜单号 {len(ids)} 个｜∩新报告Shopee桶 {len(inter)}"
      + (f" ⚠️ {sorted(inter)[:5]}" if inter else "（无重叠）"))
