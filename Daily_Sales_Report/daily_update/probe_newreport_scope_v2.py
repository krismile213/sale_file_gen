# -*- coding: utf-8 -*-
"""probe_newreport_scope_v2.py — v1 续：补 P3b（跳过文件的宽探测）、P5（新报告段A）、P4 加固

v1 已确认（不再重复论证，只复跑保证可复现）：
  - 新报告日期覆盖 2026-09-01~09-30 全月（9/1~9/12 零星 158 行，9/13 起密集）
  - 段A 净 +18、段B 净 +43（复现 v5）；找回 64 键/68 件、旧独有 7 键/7 件
  - 跨境/跨站点双计 = 0（Lazada跨境菲律宾 198 单号 ∩ 新报告 = 0）
  - 段A 旧独有 2 件（G554×2，9/12）在新报告全集按订单号/运单号均命中 1 行
  - 段B 旧独有 5 件（9/14 Shopee）在新报告全集 0 命中，旧报告状态均 Completed

v2 新增：
  P3b 对 v1 跳过的 3 个文件（Shopee印尼/Shopee泰国/Shopee跨境越南）宽探测单号列后再测交集
  P5  新报告 9/1~9/12 段（非手工,已交运/已发货）vs 监控表同段：WMS 切量前期的零星单
      是否被监控表全覆盖（若新报告段A有监控表没有的单 = 段A还有漏口）
  P4  包 try/except，崩了也看得到 traceback
"""
import os
import re
import glob
import traceback
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
        out["Shopee"].append((nid(r[1]), nsku(r[4]), q10(r[6]), d, ""))
    ws = wb["LZD Data"]
    for r in ws.iter_rows(min_row=2, values_only=True):
        d = d10(r[2])
        if not (SEP_LO <= d <= SEP_HI):
            continue
        if str(r[5] or "").strip() not in LZD_KEEP:
            continue
        out["Lazada"].append((nid(r[0]), nsku(r[1]), 1.0, d, ""))
    ws = wb["Tiktok Data"]
    for r in ws.iter_rows(min_row=2, values_only=True):
        d = d10(r[6])
        if not (SEP_LO <= d <= SEP_HI):
            continue
        st2 = TK_MAP.get(str(r[1] or "").strip())
        if st2 not in TK_KEEP:
            continue
        out["Tiktok"].append((nid(r[0]), nsku(r[2]), q10(r[5]), d, ""))
    wb.close()
    return out


def read_new_full():
    raw = pd.read_excel(NEWRPT, dtype=str)
    _line = ["订单ID", "Sku", "数量", "多属性Id", "平台数量", "sku图片", "sku中文名称"]
    _parent = [c for c in raw.columns if c not in _line]
    raw[_parent] = raw[_parent].ffill()
    raw["sku"] = raw["Sku"].map(nsku)
    raw["d"] = raw["创建时间"].map(d10)
    raw["q"] = pd.to_numeric(raw["数量"], errors="coerce").fillna(0)
    raw["plat"] = raw["平台类型"].fillna("").astype(str).str.strip()
    raw["otype"] = raw["订单类型"].fillna("").astype(str).str.strip()
    raw["oid"] = raw["交易号"].map(nid)
    raw["trk"] = raw["运单号"].map(ntrk)
    raw["is_man"] = raw["otype"] == "手工订单"
    raw["ts"] = raw["创建时间"].fillna("").astype(str).str.strip()
    return raw


def keymap(rows):
    m = {}
    for oid, sku, q, d, t in rows:
        if oid and sku:
            m.setdefault((oid, sku), []).append((q, d, t))
    return m


def oidmap(rows):
    m = {}
    for oid, sku, q, d, t in rows:
        if oid:
            m.setdefault(oid, []).append((sku, d, t, q))
    return m


def trkmap(rows):
    m = {}
    for oid, sku, q, d, t in rows:
        if t:
            m.setdefault(t, []).append((oid, sku, d, q))
    return m


def seg(rows, lo, hi):
    return [r for r in rows if lo <= r[3] <= hi]


def classify(only_keys, src_side, old_rows, src_rows, tag):
    o_oid, o_trk = oidmap(old_rows), trkmap(old_rows)
    s_oid, s_trk = oidmap(src_rows), trkmap(src_rows)
    real, tshift, split, details = 0.0, 0.0, 0.0, []
    for k, rows in only_keys.items():
        q = sum(r[0] for r in rows)
        trks = {r[2] for r in rows if r[2]}
        if src_side == "new":
            hits = o_oid.get(k[0], [])
        else:
            hits = s_oid.get(k[0], [])
        if not hits:
            t_hits = []
            for t in trks:
                t_hits.extend((o_trk if src_side == "new" else s_trk).get(t, []))
            if t_hits:
                split += q
                details.append(("运单号同笔(单号格式差异)", k, q, t_hits[:2]))
            else:
                real += q
                details.append((tag, k, q, []))
        else:
            same_sku = [h for h in hits if h[0] == k[1]]
            if same_sku:
                tshift += q
                details.append(("时间移位", k, q, same_sku[:2]))
            else:
                split += q
                details.append(("命名分裂(同单号不同SKU)", k, q, hits[:2]))
    return real, tshift, split, details


def read_any(path):
    if path.lower().endswith(".csv"):
        return pd.read_csv(path, dtype=str, low_memory=False)
    return pd.read_excel(path, dtype=str)


def find_id_col(d):
    for c in d.columns:
        if str(c).strip().lower() in ("order id", "orderid", "ordernumber", "order number",
                                      "订单号", "平台订单号", "交易号"):
            return c
    for c in d.columns:
        lc = str(c).lower()
        if "order" in lc and ("id" in lc or "number" in lc or "号" in str(c)):
            return c
    return None


def wide_id_col(d):
    """宽探测：列名含 order/id/单/号/no/package/track 的第一列；否则第一列"""
    keys = ("order", "id", "单", "号", "no.", "package", "track")
    for c in d.columns:
        lc = str(c).lower()
        if any(k in lc for k in keys):
            return c
    return d.columns[0]


RAW = read_new_full()
V = RAW[RAW["订单状态"].isin(["已交运", "已发货"])].copy()
OLD = read_old()
MONR = read_monitor()

print("=" * 84)
print("P3b v1 跳过的文件：宽探测单号列后补测交集（防双计漏网）")
print("=" * 84)
new_buckets = {}
for p in ("Shopee", "Lazada", "Tiktok"):
    new_buckets[p] = set(RAW.loc[RAW["plat"].map(lambda x: PLATMAP.get(x, x)) == p, "oid"]) - {""}
SKIP = ["Shopee印尼站点202609订单报告.xlsx", "Shopee泰国站点202609订单报告.xlsx",
        "Shopee跨境越南站点202609订单报告.xlsx"]
for name in SKIP:
    path = os.path.join(SNAP, name)
    if not os.path.exists(path):
        path = os.path.join(SNAP, "跨境", name)
    if not os.path.exists(path):
        print(f"  {name}: 文件不存在")
        continue
    d = read_any(path)
    col = wide_id_col(d)
    m = re.match(r"(Shopee|Lazada|Tiktok)(.+?)站点", name)
    plat = PLATMAP.get(m.group(1), m.group(1)) if m else None
    ids = {nid(v) for v in d[col]} - {""}
    bucket = new_buckets.get(plat, set())
    inter = ids & bucket
    print(f"  {name}: {len(d)} 行｜全部列: {list(d.columns)[:14]}")
    print(f"      宽探测选中列: {col!r}｜样例值: {[str(x)[:24] for x in d[col].dropna().head(5).tolist()]}")
    print(f"      单号 {len(ids)} 个｜∩新报告{plat}桶 {len(inter)}" + (f" ⚠️ {sorted(inter)[:5]}" if inter else ""))

print("\n" + "=" * 84)
print("P5 新报告 9/1~9/12 段（WMS 切量前期零星单，非手工,已交运/已发货）vs 监控表同段")
print("=" * 84)
nA = V[(V["d"] <= D1) & (~V["is_man"])]
print("新报告段A按平台:", nA.groupby("plat")["q"].agg(rows="size", qty="sum").to_dict("index"))
for plat, g in sorted(nA.groupby("plat")):
    p = PLATMAP.get(plat)
    if p is None:
        print(f"--- {plat!r}: 无对应旧平台（{len(g)} 行/{g['q'].sum():.0f} 件）")
        continue
    nAm = {}
    for _, r in g.iterrows():
        if r["oid"] and r["sku"]:
            nAm.setdefault((r["oid"], r["sku"]), []).append((r["q"], r["d"], r["trk"]))
    mAm = keymap(seg(MONR[p], D0, D1))
    only_n = {k: v for k, v in nAm.items() if k not in mAm}
    src_rows = [(r["oid"], r["sku"], r["q"], r["d"], r["trk"]) for _, r in g.iterrows()]
    r1, t1, s1, d1 = classify(only_n, "new", MONR[p], src_rows, "真实漏单(新报告段A有/监控表无)")
    print(f"--- {p}: 新报告段A {sum(sum(r[0] for r in v) for v in nAm.values()):.0f} 件 vs 监控表段A "
          f"{sum(sum(r[0] for r in v) for v in mAm.values()):.0f} 件｜新报告独有 {len(only_n)} 键/"
          f"{sum(sum(r[0] for r in v) for v in only_n.values()):.0f} 件：真实漏单 {r1:.0f}｜移位 {t1:.0f}｜分裂 {s1:.0f}")
    for tag, k, q, hits in d1:
        print(f"      [{tag}] {k[0][:30]:<32} {k[1]:<10} {q:.0f} 件")
print("⚠️ 监控表无运单号列 ⇒ 本段 AAS 长单号无法用运单号兜底，独有键需人工看单号格式判断")

try:
    print("\n" + "=" * 84)
    print("P4 “13号部分”：9/13~9/14 逐平台新旧对比")
    print("=" * 84)
    for d in ("2026-09-13", "2026-09-14"):
        for p in ("Shopee", "Lazada", "Tiktok"):
            n = V[(V["d"] == d) & (V["plat"].map(lambda x: PLATMAP.get(x, x)) == p)]["q"].sum()
            o = sum(r[0] for r in OLD[p] if r[3] == d)
            print(f"  {d} {p:<7}: 新报告 {n:>6.0f} vs 旧报告 {o:>6.0f}  ({n - o:+.0f})")
except Exception:
    print("P4 崩了，traceback 如下：")
    print(traceback.format_exc())

print("\n" + "=" * 84)
print("附：新报告段A（9/1~9/12）逐日件数（全状态）")
print("=" * 84)
print(RAW[RAW["d"] <= D1].groupby("d")["q"].agg(rows="size", qty="sum").to_string())
