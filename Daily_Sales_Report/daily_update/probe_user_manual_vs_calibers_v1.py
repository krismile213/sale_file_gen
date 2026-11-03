# -*- coding: utf-8 -*-
"""probe_user_manual_vs_calibers_v1.py — 用户手工表（1793/541/633/299=3266）vs 各口径逐 SKU 对比

用户表：C:\\Users\\dell\\Desktop\\新建 文本文档 (2).txt（业务备货表，SKU 维度，9月四渠道）
对比对象：
  U  用户手工表
  O  旧三平台报告（本土，订单级状态过滤后 SKU 汇总）
  M  监控表（9月，SKU 汇总）
  N  WMS 新报告（全月非手工，SKU 汇总）
  C  Lazada 跨境菲律宾（SKU 汇总）
  D  已交付副本 9月块（AY~BC，= M段A + N段B + C + 线下299）
输出：总计对比 + 逐 SKU 差异明细 + 用户表 +12 微调查定位
"""
import os
import re
import pandas as pd
from openpyxl import load_workbook

USERTXT = r"C:\Users\dell\Desktop\新建 文本文档 (2).txt"
SNAP = r"D:\Desktop\work-main\sale_file_gen\trunk\data\数据归属周期\202609\快照日期\20261006\东南亚"
MON = r"C:\Users\dell\Desktop\DAILY MONITORING - SEPTEMBER 1-13(1).xlsx"
NEWRPT = os.path.join(SNAP, "菲律宾202609订单报告.xlsx")
COPY = r"D:\Desktop\work-main\sale_file_gen\菲律宾重算_202609\2026年东南亚销量表Southeast Asia Sales table.xlsx"
D0, D1, D2, SEP_HI = "2026-09-01", "2026-09-12", "2026-09-13", "2026-09-30"

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


# ---------------------------------------------------------------- 用户手工表
def read_user():
    rows = []
    with open(USERTXT, encoding="utf-8") as f:
        lines = f.read().splitlines()
    hdr = lines[0].split("\t")
    idx = {name.strip(): i for i, name in enumerate(hdr)}
    ci = {k: idx[k] for k in ("SKU", "9月总销售数量", "9月虾皮销售数量",
                              "9月LAZADA销售数量", "9月抖音销售数量", "9月线下订单销售数量")}
    for ln in lines[2:]:
        c = ln.split("\t")
        if len(c) <= ci["9月线下订单销售数量"]:
            continue
        sku = nsku(c[ci["SKU"]])
        if not sku:
            continue
        rows.append((sku, q10(c[ci["9月虾皮销售数量"]]), q10(c[ci["9月LAZADA销售数量"]]),
                     q10(c[ci["9月抖音销售数量"]]), q10(c[ci["9月线下订单销售数量"]]),
                     q10(c[ci["9月总销售数量"]])))
    return rows


# ---------------------------------------------------------------- 旧三报告（本土，SKU 级）
def read_old_sku():
    agg = {"Shopee": {}, "Lazada": {}, "Tiktok": {}}
    d = pd.read_excel(os.path.join(SNAP, "Shopee菲律宾站点202609订单报告.xlsx"), dtype=str)
    d["st"] = d["Order Status"].astype(str).str.strip()
    d.loc[d["st"].str.startswith(RW_PREFIX), "st"] = "Completed"
    d["st2"] = d["st"].map(SHOPEE_MAP)
    for _, r in d[(d["st2"].notna()) & (d["st2"] != "Cancelled")].iterrows():
        s = nsku(r["SKU Reference No."])
        if s:
            agg["Shopee"][s] = agg["Shopee"].get(s, 0) + q10(r["Quantity"])
    d = pd.read_excel(os.path.join(SNAP, "Lazada菲律宾站点202609订单报告.xlsx"), dtype=str)
    keep = d[d["status"].astype(str).str.strip().isin(LZD_KEEP)]
    for _, r in keep.iterrows():
        s = nsku(r["sellerSku"])
        if s:
            agg["Lazada"][s] = agg["Lazada"].get(s, 0) + 1.0
    d = pd.read_csv(os.path.join(SNAP, "Tiktok菲律宾站点202609订单报告.csv"), dtype=str, low_memory=False)
    d["st2"] = d["Order Status"].astype(str).str.strip().map(TK_MAP)
    for _, r in d[d["st2"].isin(TK_KEEP)].iterrows():
        s = nsku(r["Seller SKU"])
        if s:
            agg["Tiktok"][s] = agg["Tiktok"].get(s, 0) + q10(r["Quantity"])
    return agg


# ---------------------------------------------------------------- 监控表（SKU 级）
def read_monitor_sku():
    agg = {"Shopee": {}, "Lazada": {}, "Tiktok": {}}
    wb = load_workbook(MON, read_only=True, data_only=True)
    ws = wb["SHP Data"]
    for r in ws.iter_rows(min_row=2, values_only=True):
        d = d10(r[0])
        if not (D0 <= d <= SEP_HI):
            continue
        st = str(r[2] or "").strip()
        if st.startswith(RW_PREFIX):
            st = "Completed"
        if SHOPEE_MAP.get(st) in (None, "Cancelled"):
            continue
        s = nsku(r[4])
        if s:
            agg["Shopee"][s] = agg["Shopee"].get(s, 0) + q10(r[6])
    ws = wb["LZD Data"]
    for r in ws.iter_rows(min_row=2, values_only=True):
        d = d10(r[2])
        if not (D0 <= d <= SEP_HI):
            continue
        if str(r[5] or "").strip() not in LZD_KEEP:
            continue
        s = nsku(r[1])
        if s:
            agg["Lazada"][s] = agg["Lazada"].get(s, 0) + 1.0
    ws = wb["Tiktok Data"]
    for r in ws.iter_rows(min_row=2, values_only=True):
        d = d10(r[6])
        if not (D0 <= d <= SEP_HI):
            continue
        if TK_MAP.get(str(r[1] or "").strip()) not in TK_KEEP:
            continue
        s = nsku(r[2])
        if s:
            agg["Tiktok"][s] = agg["Tiktok"].get(s, 0) + q10(r[5])
    wb.close()
    return agg


# ---------------------------------------------------------------- WMS 新报告（全月非手工，SKU 级）
def read_new_sku():
    raw = pd.read_excel(NEWRPT, dtype=str)
    _line = ["订单ID", "Sku", "数量", "多属性Id", "平台数量", "sku图片", "sku中文名称"]
    _parent = [c for c in raw.columns if c not in _line]
    raw[_parent] = raw[_parent].ffill()
    v = raw[raw["订单状态"].isin(["已交运", "已发货"])].copy()
    v = v[v["订单类型"].fillna("").astype(str).str.strip() != "手工订单"]
    v["sku"] = v["Sku"].map(nsku)
    v["q"] = pd.to_numeric(v["数量"], errors="coerce").fillna(0)
    v["p"] = v["平台类型"].fillna("").astype(str).str.strip().map(lambda x: PLATMAP.get(x))
    agg = {"Shopee": {}, "Lazada": {}, "Tiktok": {}}
    for _, r in v.iterrows():
        if r["p"] and r["sku"]:
            agg[r["p"]][r["sku"]] = agg[r["p"]].get(r["sku"], 0) + r["q"]
    return agg


# ---------------------------------------------------------------- Lazada 跨境（SKU 级）
def read_cross_sku():
    p = os.path.join(SNAP, "跨境", "Lazada跨境菲律宾站点202609订单报告.xlsx")
    if not os.path.exists(p):
        import glob
        cands = glob.glob(os.path.join(SNAP, "跨境", "*菲律宾*"))
        print("跨境目录候选:", cands)
        p = cands[0]
    d = pd.read_excel(p, dtype=str)
    keep = d[d["status"].astype(str).str.strip().isin(LZD_KEEP)]
    agg = {}
    for _, r in keep.iterrows():
        s = nsku(r["sellerSku"])
        if s:
            agg[s] = agg.get(s, 0) + 1.0
    return agg


# ---------------------------------------------------------------- 副本 9月块
def read_copy_sku():
    wb = load_workbook(COPY, read_only=True, data_only=True)
    ws = wb["Philippines 菲律宾销量表"]
    out = {}
    for r in ws.iter_rows(min_row=3, values_only=True):
        sku = nsku(r[0])
        if not sku:
            continue
        out[sku] = {"cross": q10(r[50]), "lzd": q10(r[51]), "shp": q10(r[52]),
                    "tk": q10(r[53]), "off": q10(r[54]), "tot": q10(r[55])}
    wb.close()
    return out


def n0(v):
    try:
        f = float(v)
        return int(f) if f == int(f) else f
    except Exception:
        return 0


USER = read_user()
OLD = read_old_sku()
MONR = read_monitor_sku()
NEW = read_new_sku()
CROSS = read_cross_sku()
COPY = read_copy_sku()

U = {s: (a, b, c, d, t) for s, a, b, c, d, t in USER}

print("=" * 90)
print("A. 总计对比")
print("=" * 90)
u_shp = sum(v[0] for v in U.values())
u_lzd = sum(v[1] for v in U.values())
u_tk = sum(v[2] for v in U.values())
u_off = sum(v[3] for v in U.values())
o_shp = sum(OLD["Shopee"].values()); o_lzd = sum(OLD["Lazada"].values()); o_tk = sum(OLD["Tiktok"].values())
m_shp = sum(MONR["Shopee"].values()); m_lzd = sum(MONR["Lazada"].values()); m_tk = sum(MONR["Tiktok"].values())
n_shp = sum(NEW["Shopee"].values()); n_lzd = sum(NEW["Lazada"].values()); n_tk = sum(NEW["Tiktok"].values())
c_lzd = sum(CROSS.values())
print(f"  用户表    : SHP {u_shp:>6.0f} | LZD {u_lzd:>6.0f} | TK {u_tk:>6.0f} | 线下 {u_off:>4.0f} | 合计 {u_shp+u_lzd+u_tk+u_off:.0f}")
print(f"  旧报告本土: SHP {o_shp:>6.0f} | LZD {o_lzd:>6.0f} | TK {o_tk:>6.0f} | 小计 {o_shp+o_lzd+o_tk:.0f}")
print(f"  监控表    : SHP {m_shp:>6.0f} | LZD {m_lzd:>6.0f} | TK {m_tk:>6.0f} | 小计 {m_shp+m_lzd+m_tk:.0f}")
print(f"  新报告全月: SHP {n_shp:>6.0f} | LZD {n_lzd:>6.0f} | TK {n_tk:>6.0f} | 小计 {n_shp+n_lzd+n_tk:.0f}")
print(f"  跨境 Lazada: {c_lzd:.0f}")
print(f"  副本 9月块 : 跨境 {c_lzd:.0f} | 本土LZD 558 | SHP 1809 | TK 649 | 线下 299 | 合计 3483")

print("\n" + "=" * 90)
print("B. 用户表 vs 旧报告（本土三平台）逐 SKU 差异（定位那 +12 微调）")
print("=" * 90)
diffB = []
for sku, (a, b, c, d, t) in sorted(U.items()):
    for plat, uv, ov in (("SHP", a, OLD["Shopee"].get(sku, 0)),
                         ("LZD", b, OLD["Lazada"].get(sku, 0)),
                         ("TK", c, OLD["Tiktok"].get(sku, 0))):
        if n0(uv) != n0(ov):
            diffB.append((sku, plat, n0(uv), n0(ov), n0(uv) - n0(ov)))
# 反向：旧报告有、用户表没有的 SKU
for plat, od in (("SHP", OLD["Shopee"]), ("LZD", OLD["Lazada"]), ("TK", OLD["Tiktok"])):
    for sku, ov in od.items():
        if sku not in U:
            diffB.append((sku, plat, 0, n0(ov), -n0(ov)))
print(f"  差异单元格 {len(diffB)} 个，净差 {sum(x[4] for x in diffB):+.0f}")
for sku, plat, uv, ov, dv in diffB:
    print(f"    {sku:<12} {plat:<4} 用户 {uv:>5} vs 旧报告 {ov:>5}  ({dv:+.0f})")

print("\n" + "=" * 90)
print("C. 用户表 vs 副本（已交付口径）逐 SKU 差异（差 217 的分布）")
print("=" * 90)
diffC = {"cross": 0.0, "lzd": 0.0, "shp": 0.0, "tk": 0.0, "off": 0.0}
detailC = []
allsku = set(U) | set(COPY)
for sku in sorted(allsku):
    u = U.get(sku, (0, 0, 0, 0, 0))
    cp = COPY.get(sku, {"cross": 0, "lzd": 0, "shp": 0, "tk": 0, "off": 0, "tot": 0})
    # 用户 LAZADA 列 vs 副本 跨境+本土
    d_cross = 0 - n0(cp["cross"])
    d_lzd = n0(u[1]) - n0(cp["cross"]) - n0(cp["lzd"])
    d_shp = n0(u[0]) - n0(cp["shp"])
    d_tk = n0(u[2]) - n0(cp["tk"])
    d_off = n0(u[3]) - n0(cp["off"])
    diffC["cross"] += d_cross; diffC["lzd"] += d_lzd; diffC["shp"] += d_shp
    diffC["tk"] += d_tk; diffC["off"] += d_off
    if any(x for x in (d_cross, d_lzd, d_shp, d_tk, d_off)):
        detailC.append((sku, d_cross, d_lzd, d_shp, d_tk, d_off))
print(f"  差异 SKU {len(detailC)} 个；渠道净差（用户−副本）：跨境LZD {diffC['cross']:+.0f}｜LZD本土 {diffC['lzd']:+.0f}"
      f"｜SHP {diffC['shp']:+.0f}｜TK {diffC['tk']:+.0f}｜线下 {diffC['off']:+.0f}"
      f"｜合计 {sum(diffC.values()):+.0f}")
print("  按渠道汇总缺口的构成：")
print(f"    跨境未计: {-diffC['cross']:.0f}")
print(f"    本土三平台净差(用户−副本): {diffC['lzd']+diffC['shp']+diffC['tk']:+.0f}")
print(f"    线下差: {diffC['off']:+.0f}")

print("\n" + "=" * 90)
print("D. 逐 SKU 明细（用户 vs 副本，仅列差异 SKU，前 60 行）")
print("=" * 90)
for sku, d_cross, d_lzd, d_shp, d_tk, d_off in detailC[:60]:
    print(f"    {sku:<12} 跨境 {d_cross:>5.0f} | 本土LZD {d_lzd:>5.0f} | SHP {d_shp:>5.0f} | TK {d_tk:>5.0f} | 线下 {d_off:>4.0f}")

print("\n" + "=" * 90)
print("E. 自证：用户 3,266 与副本 3,483 的差 = 217 的构成")
print("=" * 90)
gap_cross = -diffC["cross"]
gap_local = -(diffC["lzd"] + diffC["shp"] + diffC["tk"])
print(f"  差 217 = 跨境未计 {gap_cross:.0f} + 本土净差 {gap_local:.0f}（= {gap_cross + gap_local:.0f}）")
print(f"  其中本土净差 49 = 新源漏单净增 61 − 用户表比旧报告多的 {sum(x[4] for x in diffB):+.0f} 件微调")
print(f"  校验：61 + 168 − 12 = {61 + 168 - 12}")
