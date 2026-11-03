# -*- coding: utf-8 -*-
"""probe_forensic61_v2.py — 61 件订单级取证 v2（修正 v1 手工单识别 bug，只读）

v1 bug 回放：
  v1 用 `平台类型 == "手工订单"` 识别手工单，但「平台类型」是合并单元格父列、ffill 后
  手工单块继承上一行 = "Shopee"，从不为 "手工订单" → 145 件手工单被算进
  「新报告独有(漏单)」，段B 净差从 +43 虚报成 +188，总差从 +61 虚报成 +206。
v2 改用「订单类型」列（ffill 后 == '手工订单'）正确剥离，重算段A/段B 并逐键取证。

另对齐基线：trunk 菲 9 月 3,123 = 本土三报告 2,955 + Lazada跨境菲律宾 168
（跨境在 快照/东南亚/跨境/ 单独文件，新旧口径都不动，与 61 无关）。
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


# ---------------------------------------------------------------- 旧三报告（本土，全月，订单级）
def read_old():
    out = {"Shopee": [], "Lazada": [], "Tiktok": []}   # (oid, sku, qty, date)
    stats = {}
    d = pd.read_excel(os.path.join(SNAP, "Shopee菲律宾站点202609订单报告.xlsx"), dtype=str)
    d["订单状态"] = d["Order Status"].astype(str).str.strip()
    d.loc[d["订单状态"].str.startswith(RW_PREFIX), "订单状态"] = "Completed"
    d["st"] = d["订单状态"].map(SHOPEE_MAP)
    bad_date = 0
    for _, r in d[(d["st"].notna()) & (d["st"] != "Cancelled")].iterrows():
        dd = d10(r["Order Creation Date"])
        if not dd:
            bad_date += 1
        out["Shopee"].append((nid(r["Order ID"]), nsku(r["SKU Reference No."]),
                              q10(r["Quantity"]), dd))
    stats["Shopee"] = (len(d), bad_date)
    d = pd.read_excel(os.path.join(SNAP, "Lazada菲律宾站点202609订单报告.xlsx"), dtype=str)
    d["st"] = d["status"].astype(str).str.strip()
    bad_date = 0
    for (_, r), c in zip(d[d["st"].isin(LZD_KEEP)].iterrows(),
                         pd.to_datetime(d.loc[d["st"].isin(LZD_KEEP), "createTime"], errors="coerce")):
        dd = c.strftime("%Y-%m-%d") if pd.notna(c) else ""
        if not dd:
            bad_date += 1
        out["Lazada"].append((nid(r["orderNumber"]), nsku(r["sellerSku"]), 1.0, dd))
    stats["Lazada"] = (len(d), bad_date)
    d = pd.read_csv(os.path.join(SNAP, "Tiktok菲律宾站点202609订单报告.csv"), dtype=str, low_memory=False)
    d["st"] = d["Order Status"].astype(str).str.strip().map(TK_MAP)
    keep = d[d["st"].isin(TK_KEEP)]
    bad_date = 0
    for (_, r), c in zip(keep.iterrows(),
                         pd.to_datetime(keep["Created Time"].astype(str).str.strip(), errors="coerce")):
        dd = c.strftime("%Y-%m-%d") if pd.notna(c) else ""
        if not dd:
            bad_date += 1
        out["Tiktok"].append((nid(r["Order ID"]), nsku(r["Seller SKU"]),
                              q10(r["Quantity"]), dd))
    stats["Tiktok"] = (len(d), bad_date)
    return out, stats


# ---------------------------------------------------------------- 监控表三平台（9 月，订单级）
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


# ---------------------------------------------------------------- 新订单报告（ffill + 订单类型识别手工单）
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


def key(rows):
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


def oid_index(m):
    """oid -> [(sku, date), ...]（旧报告全月，查"订单在不在、SKU/日期是否不同"）"""
    idx = {}
    for (oid, sku), v in m.items():
        idx.setdefault(oid, []).extend((sku, d) for _, d in v)
    return idx


print("=" * 80)
print("基线对齐：trunk 菲 9 月 3,123 = 本土三报告 + Lazada跨境菲律宾（单独文件，与 61 无关）")
print("=" * 80)
OLD, OSTAT = read_old()
MONR = read_monitor()
NEW = read_new()
for p in ("Shopee", "Lazada", "Tiktok"):
    rows = OLD[p]
    ds = sorted({r[3] for r in rows if r[3]})
    print(f"  旧[{p:<7}] 全月 {len(rows):>5} 行 / {sum(r[2] for r in rows):>7.0f} 件"
          f"｜日期范围 {ds[0]}~{ds[-1]}｜解析失败 {OSTAT[p][1]} 行")
print(f"  本土三报告合计 = {sum(sum(r[2] for r in OLD[p]) for p in OLD):.0f} 件"
      f"（+ Lazada跨境 168 = trunk 3,123）")
#  Lazada跨境件数核对
try:
    cx = pd.read_excel(os.path.join(SNAP, "跨境", "Lazada跨境菲律宾站点202609订单报告.xlsx"), dtype=str)
    cxk = cx[cx["status"].astype(str).str.strip().isin(LZD_KEEP)]
    print(f"   Lazada跨境菲律宾 = {len(cxk)} 行 / {len(cxk):.0f} 件（status 白名单后）")
except Exception as e:
    print(f"   Lazada跨境菲律宾读取失败: {e}")

print()
print("=" * 80)
print("新报告手工单块核实（v1 bug 的根源）")
print("=" * 80)
man = NEW[NEW["is_man"]]
print(f"  订单类型=='手工订单'（ffill 后）= {len(man)} 行 / {man['q'].sum():.0f} 件")
print(f"  其 平台类型（ffill 继承）= {man['plat'].unique().tolist()} ← v1 按此列筛手工单，永远筛不中")
print(f"  订单号={man['oid'].unique().tolist()}  日期={sorted(man['d'].unique())}")
print(f"  v1 若按 plat!='手工订单' 放行 → 这 {man['q'].sum():.0f} 件全被算进「新报告独有漏单」")

print()
print("=" * 80)
print("段A（9/1~9/12）：监控表 vs 旧三报告同段 —— 逐键取证")
print("=" * 80)
segA_net = 0.0
for p in ("Shopee", "Lazada", "Tiktok"):
    mA = merge(key(seg(MONR[p], D0, D1)))
    oA = merge(key(seg(OLD[p], D0, D1)))
    oAll = key(OLD[p])
    oAllD = datemap(oAll)
    mAll = key(MONR[p])
    only_m = {k: v for k, v in mA.items() if k not in oA}
    only_o = {k: v for k, v in oA.items() if k not in mA}
    net = sum(mA.values()) - sum(oA.values())
    segA_net += net
    print(f"\n--- {p}: 监控表 {sum(mA.values()):.0f} vs 旧 {sum(oA.values()):.0f}，净差 {net:+.0f}"
          f"（键数 监{len(mA)} / 旧{len(oA)}）")
    # 监控表独有 → 旧报告全月有没有这笔订单（换 SKU/日期也算"在"）
    absent, sku_mismatch, shift = [], [], []
    for k, v in only_m.items():
        if k[0] in oAll:
            alts = [(s, d) for (s, d) in oid_index(oAll)[k[0]] if s != k[1]]
            if alts:
                sku_mismatch.append((k, v, alts[:3]))
            else:
                shift.append((k, v, oAllD.get(k, [])))
        else:
            absent.append((k, v))
    print(f"    监控表独有 {len(only_m)} 键/{sum(only_m.values()):.0f} 件：")
    print(f"      · 旧报告全月无此订单号（真实漏单）{len(absent)} 键/{sum(v for _, v in absent):.0f} 件")
    for k, v in sorted(absent, key=lambda x: -x[1]):
        print(f"        漏 {k[0][:22]:<24} {k[1]:<10} {v:.0f} 件")
    if sku_mismatch:
        print(f"      · 订单号在旧报告但 SKU 不同（命名分裂？）{len(sku_mismatch)} 键：")
        for k, v, alts in sku_mismatch[:6]:
            print(f"        名 {k[0][:22]:<24} 监{k[1]}({v:.0f}) 旧={alts}")
    if shift:
        print(f"      · 订单+SKU 在旧报告但日期不同 {len(shift)} 键：{shift[:4]}")
    # 旧报告独有 → 监控表全月有没有
    o_absent, o_shift = [], []
    for k, v in only_o.items():
        if k in mAll:
            o_shift.append((k, v))
        else:
            o_absent.append((k, v))
    print(f"    旧报告独有 {len(only_o)} 键/{sum(only_o.values()):.0f} 件"
          f"（监控表全月也无此(单,SKU) {len(o_absent)} 键/{sum(v for _, v in o_absent):.0f} 件）")
    for k, v in sorted(o_absent, key=lambda x: -x[1])[:10]:
        print(f"        旧独 {k[0][:22]:<24} {k[1]:<10} {v:.0f} 件")
print(f"\n  ⇒ 段A 净差合计 = {segA_net:+.0f} 件")

print()
print("=" * 80)
print("段B（≥9/13）：新订单报告(非手工) vs 旧三报告同段 —— 逐键取证")
print("=" * 80)
segB_new = NEW[(NEW["d"] >= D2) & (~NEW["is_man"])]
# 新报告全集（任意状态/日期/类型）键，用于查"旧独有键在新报告里到底有没有"
new_all = NEW.copy()
new_all_oid = {}
for _, r in new_all.iterrows():
    if r["oid"]:
        new_all_oid.setdefault(r["oid"], []).append(
            (r["sku"], r["d"], r["订单状态"], r["plat"], r["otype"]))

PLATMAP = {"Shopee": "Shopee", "Lazada": "Lazada", "TikTok": "Tiktok", "Tiktok shop": "Tiktok"}
segB_net = 0.0
for plat, g in sorted(segB_new.groupby("plat")):
    p = PLATMAP.get(plat)
    if p is None:
        print(f"--- 新报告平台 {plat!r}: 无对应旧报告平台（{len(g)} 行/{g['q'].sum():.0f} 件）")
        continue
    nB = {}
    for _, r in g.iterrows():
        if r["oid"] and r["sku"]:
            nB[(r["oid"], r["sku"])] = nB.get((r["oid"], r["sku"]), 0) + r["q"]
    oB = merge(key(seg(OLD[p], D2, SEP_HI)))
    oAll = key(OLD[p])
    only_n = {k: v for k, v in nB.items() if k not in oB}
    only_o = {k: v for k, v in oB.items() if k not in nB}
    net = sum(nB.values()) - sum(oB.values())
    segB_net += net
    print(f"\n--- {plat}→{p}: 新(非手工) {sum(nB.values()):.0f} vs 旧 {sum(oB.values()):.0f}，净差 {net:+.0f}"
          f"（键数 新{len(nB)} / 旧{len(oB)}）")
    absent, sku_mismatch, shift = [], [], []
    for k, v in only_n.items():
        if k[0] in oAll:
            alts = [(s, d) for (s, d) in oid_index(oAll)[k[0]] if s != k[1]]
            if alts:
                sku_mismatch.append((k, v, alts[:3]))
            else:
                shift.append((k, v))
        else:
            absent.append((k, v))
    print(f"    新报告独有 {len(only_n)} 键/{sum(only_n.values()):.0f} 件：")
    print(f"      · 旧报告全月无此订单号（真实漏单）{len(absent)} 键/{sum(v for _, v in absent):.0f} 件")
    for k, v in sorted(absent, key=lambda x: -x[1])[:20]:
        print(f"        漏 {k[0][:22]:<24} {k[1]:<10} {v:.0f} 件")
    if len(absent) > 20:
        print(f"        ……还有 {len(absent)-20} 键")
    if sku_mismatch:
        print(f"      · 订单号在旧报告但 SKU 不同 {len(sku_mismatch)} 键：")
        for k, v, alts in sku_mismatch[:6]:
            print(f"        名 {k[0][:22]:<24} 新{k[1]}({v:.0f}) 旧={alts}")
    if shift:
        print(f"      · 订单+SKU 在旧报告但日期不同 {len(shift)} 键")
    # 旧报告独有 → 去新报告全集里查（任意状态/日期）
    n_absent, n_cancel, n_other = [], [], []
    for k, v in only_o.items():
        hits = new_all_oid.get(k[0], [])
        if not hits:
            n_absent.append((k, v))
        elif all(h[2] == "取消" for h in hits):
            n_cancel.append((k, v, hits[:2]))
        else:
            n_other.append((k, v, hits[:2]))
    print(f"    旧报告独有 {len(only_o)} 键/{sum(only_o.values()):.0f} 件：")
    print(f"      · 新报告全集也无此订单号 {len(n_absent)} 键/{sum(v for _, v in n_absent):.0f} 件")
    for k, v in sorted(n_absent, key=lambda x: -x[1])[:10]:
        print(f"        新无 {k[0][:22]:<24} {k[1]:<10} {v:.0f} 件")
    if n_cancel:
        print(f"      · 新报告里是「取消」状态 {len(n_cancel)} 键/{sum(v for _, v, _ in n_cancel):.0f} 件")
    if n_other:
        print(f"      · 新报告里有但状态/日期/SKU 不同 {len(n_other)} 键/{sum(v for _, v, _ in n_other):.0f} 件")
        for k, v, hits in n_other[:6]:
            print(f"        异 {k[0][:22]:<24} 旧{k[1]}({v:.0f}) 新={hits}")
print(f"\n  ⇒ 段B 非手工净差合计 = {segB_net:+.0f} 件")

print()
print("=" * 80)
print("总对账")
print("=" * 80)
print(f"  段A 净差 = {segA_net:+.0f}")
print(f"  段B 非手工净差 = {segB_net:+.0f}")
print(f"  ⇒ 平台真实漏单合计 = {segA_net + segB_net:+.0f} 件（= 文档里的「61」？）")
print(f"  手工单（新报告 ≥9/13）= {man['q'].sum():.0f} 件 → 已归「线下订单」列")
print(f"  RESELLER 页（监控表）= 154 件 → 已归「线下订单」列")
print(f"  新口径 = trunk 3,123 + {segA_net + segB_net:.0f} + {man['q'].sum():.0f} + 154"
      f" = {3123 + segA_net + segB_net + man['q'].sum() + 154:.0f}（交付值 3,483）")
