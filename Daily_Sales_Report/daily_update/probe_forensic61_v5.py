# -*- coding: utf-8 -*-
"""probe_forensic61_v5.py — 61 件取证终版（修正 v2 的 oid 级误判 + 纳入运单号/命名分裂，只读）

v2 的暗坑：`oAll = key(OLD[p])` 的键是 (oid, sku) 元组，而判断写的是 `k[0] in oAll`
（k[0] 是 oid 字符串）⇒ **永远 False**，「旧报告全月无此订单号」这一分支实际没验证过。
v4 已证实：段B Shopee 50 个"新独有"里有 30 个的运单号能在旧报告找到（AAS 长单号 vs 短单号，同一笔）。

v5 做正确的三层判定（对每个"源独有"键）：
  L1 订单号 oid 在不在旧报告全月（任意 SKU/日期）→ 不在 = 真实漏单/真实缺失
  L2 在、同 SKU 不同日期 → 时间移位
  L3 在、不同 SKU → 命名分裂（再用运单号确认是否同一笔）
段A 的"新源"= 监控表；段B 的"新源"= 新订单报告（非手工）。旧 = 对应平台旧报告。
"""
import os
import re
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


# ---------------------------------------------------------------- 旧三报告（本土，全月，订单级，带运单号）
def read_old():
    """returns rows[平台] = [(oid, sku, qty, date, trk)]"""
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
    print("Lazada 报告列:", list(d.columns)[:20], "…运单列:", trk_col)
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
    print("Tiktok 报告列(前25):", list(d.columns)[:25], "…运单列:", trk_col)
    for (_, r), c in zip(keep.iterrows(),
                         pd.to_datetime(keep["Created Time"].astype(str).str.strip(), errors="coerce")):
        out["Tiktok"].append((nid(r["Order ID"]), nsku(r["Seller SKU"]),
                              q10(r["Quantity"]), c.strftime("%Y-%m-%d") if pd.notna(c) else "",
                              ntrk(r[trk_col]) if trk_col else ""))
    return out


# ---------------------------------------------------------------- 监控表（9 月，订单级）
def read_monitor():
    out = {"Shopee": [], "Lazada": [], "Tiktok": []}
    wb = load_workbook(MON, read_only=True, data_only=True)
    hdr = {}
    for name in ("SHP Data", "LZD Data", "Tiktok Data"):
        ws = wb[name]
        row1 = next(ws.iter_rows(min_row=1, max_row=1, values_only=True))
        hdr[name] = [(i, v) for i, v in enumerate(row1) if v]
        print(f"监控表 {name} 表头:", hdr[name])
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


# ---------------------------------------------------------------- 新订单报告（全集，非手工）
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
    v["trk"] = v["运单号"].map(ntrk)
    v["is_man"] = v["otype"] == "手工订单"
    return v


def keymap(rows):
    """(oid, sku) -> [(q, d, trk)]"""
    m = {}
    for oid, sku, q, d, t in rows:
        if oid and sku:
            m.setdefault((oid, sku), []).append((q, d, t))
    return m


def oidmap(rows):
    """oid -> [(sku, d, trk, q)]"""
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
    """src_side: 'new' 表示 only_keys 来自新源（查旧报告）；'old' 反之。
    返回 (真实漏单件数, 时间移位件数, 命名分裂件数, 明细)"""
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
            # oid 不在对面 → 再用运单号兜底（防单号格式差异）
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


OLD = read_old()
MONR = read_monitor()
NEW = read_new()

PLATMAP = {"Shopee": "Shopee", "Lazada": "Lazada", "TikTok": "Tiktok", "Tiktok shop": "Tiktok"}
grand = {"real": 0.0, "tshift": 0.0, "split": 0.0}

print("\n" + "=" * 84)
print("段A（9/1~9/12）：监控表独有 → 查旧报告（三层判定）")
print("=" * 84)
segA_net = 0.0
for p in ("Shopee", "Lazada", "Tiktok"):
    mA = keymap(seg(MONR[p], D0, D1))
    oA = keymap(seg(OLD[p], D0, D1))
    only_m = {k: v for k, v in mA.items() if k not in oA}
    only_o = {k: v for k, v in oA.items() if k not in mA}
    net = sum(sum(r[0] for r in v) for v in mA.values()) - sum(sum(r[0] for r in v) for v in oA.values())
    segA_net += net
    r1, t1, s1, d1 = classify(only_m, "new", OLD[p], MONR[p], "真实漏单(监控表有/旧报告无)")
    r2, t2, s2, d2 = classify(only_o, "old", OLD[p], MONR[p], "真实缺失(旧报告有/监控表无)")
    grand["real"] += r1; grand["tshift"] += t1; grand["split"] += s1
    grand["real"] -= r2; grand["tshift"] -= t2; grand["split"] -= s2
    print(f"\n--- {p}: 净差 {net:+.0f}（监 {sum(sum(r[0] for r in v) for v in mA.values()):.0f} vs 旧 {sum(sum(r[0] for r in v) for v in oA.values()):.0f}）")
    print(f"    监控表独有 {len(only_m)} 键：真实漏单 {r1:.0f} 件｜时间移位 {t1:.0f}｜命名/格式分裂 {s1:.0f}")
    for tag, k, q, hits in d1:
        if tag != "真实漏单(监控表有/旧报告无)":
            print(f"      [{tag}] {k[0][:24]:<26} {k[1]:<10} {q:.0f} 件 → {hits}")
    print(f"    旧报告独有 {len(only_o)} 键：真实缺失 {r2:.0f} 件｜时间移位 {t2:.0f}｜命名/格式分裂 {s2:.0f}")
    for tag, k, q, hits in d2:
        print(f"      [{tag}] {k[0][:24]:<26} {k[1]:<10} {q:.0f} 件 → {hits}")

print("\n" + "=" * 84)
print("段B（≥9/13）：新订单报告(非手工) 独有 → 查旧报告（三层判定）")
print("=" * 84)
segB_net = 0.0
for plat, g in sorted(NEW[(NEW["d"] >= D2) & (~NEW["is_man"])].groupby("plat")):
    p = PLATMAP.get(plat)
    if p is None:
        print(f"--- 新报告平台 {plat!r}: 无对应旧报告平台（{len(g)} 行/{g['q'].sum():.0f} 件）")
        continue
    nB = {}
    for _, r in g.iterrows():
        if r["oid"] and r["sku"]:
            nB.setdefault((r["oid"], r["sku"]), []).append((r["q"], r["d"], r["trk"]))
    oB = keymap(seg(OLD[p], D2, SEP_HI))
    only_n = {k: v for k, v in nB.items() if k not in oB}
    only_o = {k: v for k, v in oB.items() if k not in nB}
    net = sum(sum(r[0] for r in v) for v in nB.values()) - sum(sum(r[0] for r in v) for v in oB.values())
    segB_net += net
    # 新报告侧 rows 需带 trk：从 NEW 重建该平台段B行
    src_rows = [(r["oid"], r["sku"], r["q"], r["d"], r["trk"]) for _, r in g.iterrows()]
    r1, t1, s1, d1 = classify(only_n, "new", OLD[p], src_rows, "真实漏单(新报告有/旧报告无)")
    r2, t2, s2, d2 = classify(only_o, "old", OLD[p], src_rows, "真实缺失(旧报告有/新报告无)")
    grand["real"] += r1; grand["tshift"] += t1; grand["split"] += s1
    grand["real"] -= r2; grand["tshift"] -= t2; grand["split"] -= s2
    print(f"\n--- {plat}→{p}: 净差 {net:+.0f}（新 {sum(sum(r[0] for r in v) for v in nB.values()):.0f} vs 旧 {sum(sum(r[0] for r in v) for v in oB.values()):.0f}）")
    print(f"    新报告独有 {len(only_n)} 键：真实漏单 {r1:.0f} 件｜时间移位 {t1:.0f}｜单号格式/命名分裂 {s1:.0f}")
    for tag, k, q, hits in d1:
        if tag != "真实漏单(新报告有/旧报告无)":
            print(f"      [{tag}] {k[0][:26]:<28} {k[1]:<10} {q:.0f} 件 → {hits}")
    print(f"    旧报告独有 {len(only_o)} 键：真实缺失 {r2:.0f} 件｜时间移位 {t2:.0f}｜单号格式/命名分裂 {s2:.0f}")
    for tag, k, q, hits in d2:
        if tag != "真实缺失(旧报告有/新报告无)":
            print(f"      [{tag}] {k[0][:26]:<28} {k[1]:<10} {q:.0f} 件 → {hits}")
    for tag, k, q, hits in d2:
        if tag == "真实缺失(旧报告有/新报告无)":
            print(f"      [真实缺失] {k[0][:26]:<28} {k[1]:<10} {q:.0f} 件")

print("\n" + "=" * 84)
print("终版总账")
print("=" * 84)
print(f"  段A 净差 = {segA_net:+.0f}；段B 净差 = {segB_net:+.0f}；合计 = {segA_net + segB_net:+.0f}")
print(f"  其中：找回真实漏单 {grand['real']:+.0f} 件｜时间移位 {grand['tshift']:+.0f}｜单号格式/命名分裂 {grand['split']:+.0f}")
print(f"  校验：真实漏单 − 真实缺失 − 时间移位 ± 分裂 = {grand['real'] + grand['tshift'] + grand['split']:+.0f}（应=段A+段B净差）")
print(f"  ⇒ 新口径 = trunk 3,123 + {segA_net + segB_net:.0f} + 手工145 + RESELLER154 = {3123 + segA_net + segB_net + 145 + 154:.0f}")
print(f"  ⇒ 若业务要'并集口径'（把真实缺失的旧报告订单也捞回）应再加 {-grand['real'] if grand['real'] < 0 else 0:.0f} 件中属于缺失的部分")
