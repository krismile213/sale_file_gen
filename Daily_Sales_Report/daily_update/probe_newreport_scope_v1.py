# -*- coding: utf-8 -*-
"""probe_newreport_scope_v1.py — 新报告覆盖范围 + 旧/新口径判定 + 跨境双计核查（只读）

回应用户（2026-10-16 20:30）：
  "菲律宾202609订单报告.xlsx 是新平台的报告，应该包含了13号部分到30的菲律宾所有订单，
   你看看这个报告和旧口径一样还是新口径一样"

  P1  文件全貌：日期覆盖/平台类型/订单状态/订单类型/手工单/单号去重 → "是不是所有订单"
  P2  段B（≥9/13，非手工，已交运/已发货）新报告 vs 旧三报告同段：逐平台净差 + 三层判定
      （复用 probe_forensic61_v5.py 已验证逻辑，数字应复现 1100/329/368 vs 1085/316/353）
  P2b 旧报告独有的订单：在全新报告全集里按 订单号+运单号 再搜一次（不限平台/状态/日期），
      并回查旧报告原始状态
  P3  跨境/跨站点双计核查：新报告各平台桶 ∩ 跨境Lazada报告 / 虾马·越站点报告（若存在）；
      段A+段B"找回"的漏单、旧报告独有的 7 件也各过一遍
  P4  "13号部分"：9/13、9/14 逐平台新旧件数 + 新报告最早/最晚时间戳
  段A（监控表 vs 旧报告）一并复跑：复现 +61 全链 + 段A找回单的跨境核查
"""
import os
import re
import glob
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


# ---------------------------------------------------------------- 旧三报告（本土，全月，订单级）
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


# ---------------------------------------------------------------- 监控表（9 月，订单级）
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


# ---------------------------------------------------------------- 新订单报告（全集，不过滤状态）
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


OLD_FILES = {
    "Shopee": (os.path.join(SNAP, "Shopee菲律宾站点202609订单报告.xlsx"),
               "Order ID", "Order Status", "Order Creation Date", "Tracking Number*"),
    "Lazada": (os.path.join(SNAP, "Lazada菲律宾站点202609订单报告.xlsx"),
               "orderNumber", "status", "createTime", None),
    "Tiktok": (os.path.join(SNAP, "Tiktok菲律宾站点202609订单报告.csv"),
               "Order ID", "Order Status", "Created Time", None),
}


def lookup_old_status(plat, oid):
    path, oidc, stc, dtc, trkc = OLD_FILES[plat]
    if path.endswith(".csv"):
        d = pd.read_csv(path, dtype=str, low_memory=False)
    else:
        d = pd.read_excel(path, dtype=str)
    hit = d[d[oidc].map(nid) == oid]
    out = []
    for _, r in hit.iterrows():
        out.append((str(r.get(stc, "")).strip(), d10(r.get(dtc, "")), nsku(r.get("SKU Reference No.") if "SKU Reference No." in d.columns else (r.get("sellerSku") if "sellerSku" in d.columns else r.get("Seller SKU"))),
                    q10(r.get("Quantity")) if "Quantity" in d.columns else 1.0,
                    ntrk(r.get(trkc)) if trkc and trkc in d.columns else ""))
    return out


def read_ids(path):
    if path.lower().endswith(".csv"):
        d = pd.read_csv(path, dtype=str, low_memory=False)
    else:
        d = pd.read_excel(path, dtype=str)
    col = None
    for c in d.columns:
        if str(c).strip().lower() in ("order id", "orderid", "ordernumber",
                                      "order number", "订单号", "平台订单号", "交易号"):
            col = c
            break
    if col is None:
        for c in d.columns:
            lc = str(c).lower()
            if "order" in lc and ("id" in lc or "number" in lc or "号" in str(c)):
                col = c
                break
    if col is None:
        return None, len(d), None
    ids = {nid(v) for v in d[col]} - {""}
    return ids, len(d), col


# ================================================================ P1 文件全貌
RAW = read_new_full()
V = RAW[RAW["订单状态"].isin(["已交运", "已发货"])].copy()

print("=" * 84)
print("P1 新报告文件全貌（菲律宾202609订单报告.xlsx，全集，不滤状态）")
print("=" * 84)
print("shape:", RAW.shape, "（数据行", len(RAW), "）")
print("列:", list(RAW.columns))
print("订单状态为空的行:", int(RAW["订单状态"].isna().sum()))
ds = sorted(x for x in RAW["d"] if x)
print("日期覆盖:", ds[0], "~", ds[-1], "｜日期解析失败行:", int((RAW["d"] == "").sum()))
outsep = RAW[(RAW["d"] < SEP_LO) | (RAW["d"] > SEP_HI)]
print("9 月以外的行:", len(outsep))
if len(outsep):
    print(outsep[["交易号", "plat", "订单状态", "d", "q"]].head(20).to_string())
print("\n逐日件数（行数/件数）:")
g = RAW.groupby("d").agg(rows=("q", "size"), qty=("q", "sum"))
print(g.to_string())
print("\n平台类型 × 订单状态（件数）:")
print(RAW.pivot_table(index="plat", columns="订单状态", values="q",
                      aggfunc="sum", margins=True, fill_value=0).to_string())
print("\n平台类型（行数）:", RAW["plat"].value_counts(dropna=False).to_dict())
print("订单类型:", RAW["otype"].value_counts(dropna=False).to_dict())
man = RAW[RAW["is_man"]]
print("手工单:", len(man), "行 /", man["q"].sum(), "件｜日期:", sorted(set(man["d"])),
      "｜订单ID:", sorted(set(man["订单ID"])), "｜交易号:", sorted(set(man["oid"]))[:3])
print("交易号去重:", RAW["oid"].nunique(), "｜订单ID去重:", RAW["订单ID"].nunique(),
      "｜数据行:", len(RAW), "｜件数合计:", RAW["q"].sum())
known_plat = {"Shopee", "Lazada", "TikTok", "Tiktok shop", "手工订单"}
oth = RAW[~RAW["plat"].isin(known_plat)]
print("非常规平台类型行:", len(oth), oth["plat"].value_counts().to_dict() if len(oth) else "")

# ================================================================ P2 段A + 段B 同段对比
OLD = read_old()
MONR = read_monitor()
recov, oldonly = [], []
grand = {"real": 0.0, "tshift": 0.0, "split": 0.0}

print("\n" + "=" * 84)
print("P2 段A（9/1~9/12）：监控表 vs 旧报告（三层判定，复现 +18）")
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
    for tag, k, q, hits in d1:
        if tag.startswith("真实漏单"):
            recov.append((p, k[0], k[1], q, mA[k][0][1]))
    for tag, k, q, hits in d2:
        if tag.startswith("真实缺失"):
            oldonly.append((p, k[0], k[1], q, oA[k][0][1], oA[k][0][2]))
    print(f"--- {p}: 净差 {net:+.0f}（监 {sum(sum(r[0] for r in v) for v in mA.values()):.0f} vs 旧 {sum(sum(r[0] for r in v) for v in oA.values()):.0f}）"
          f"｜找回 {r1:.0f}｜旧独有 {r2:.0f}｜移位 {t1 + t2:.0f}｜分裂 {s1 + s2:.0f}")

print("\n" + "=" * 84)
print("P2 段B（≥9/13）：新订单报告(非手工,已交运/已发货) vs 旧报告同段（三层判定，复现 +43）")
print("=" * 84)
segB_net = 0.0
for plat, g in sorted(V[(V["d"] >= D2) & (~V["is_man"])].groupby("plat")):
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
    src_rows = [(r["oid"], r["sku"], r["q"], r["d"], r["trk"]) for _, r in g.iterrows()]
    r1, t1, s1, d1 = classify(only_n, "new", OLD[p], src_rows, "真实漏单(新报告有/旧报告无)")
    r2, t2, s2, d2 = classify(only_o, "old", OLD[p], src_rows, "真实缺失(旧报告有/新报告无)")
    grand["real"] += r1; grand["tshift"] += t1; grand["split"] += s1
    grand["real"] -= r2; grand["tshift"] -= t2; grand["split"] -= s2
    for tag, k, q, hits in d1:
        if tag.startswith("真实漏单"):
            recov.append((p, k[0], k[1], q, nB[k][0][1]))
    for tag, k, q, hits in d2:
        if tag.startswith("真实缺失"):
            oldonly.append((p, k[0], k[1], q, oB[k][0][1], oB[k][0][2]))
    print(f"--- {plat}→{p}: 净差 {net:+.0f}（新 {sum(sum(r[0] for r in v) for v in nB.values()):.0f} vs 旧 {sum(sum(r[0] for r in v) for v in oB.values()):.0f}）"
          f"｜找回 {r1:.0f}｜旧独有 {r2:.0f}｜移位 {t1 + t2:.0f}｜分裂 {s1 + s2:.0f}")

print("\n" + "=" * 84)
print("P2b 旧报告独有的订单：全新报告全集再搜（订单号+运单号，不限平台/状态/日期）+ 旧报告原始状态")
print("=" * 84)
goid, gtrk = {}, {}
for _, r in RAW.iterrows():
    goid.setdefault(r["oid"], []).append(r)
    if r["trk"]:
        gtrk.setdefault(r["trk"], []).append(r)
for plat, oid, sku, qty, d, trk in oldonly:
    h1 = goid.get(oid, [])
    h2 = gtrk.get(trk, []) if trk else []
    st = lookup_old_status(plat, oid)
    print(f"  [{plat}] {oid} / {sku} {qty:.0f} 件｜旧报告日期 {d}｜运单 {trk or '—'}")
    print(f"      新报告全集：按订单号命中 {len(h1)} 行、按运单号命中 {len(h2)} 行｜旧报告原始状态: {st}")

print("\n" + "=" * 84)
print("P2 小结")
print("=" * 84)
print(f"  找回真实漏单 {len(recov)} 键 / {sum(x[3] for x in recov):.0f} 件；旧报告独有 {len(oldonly)} 键 / {sum(x[3] for x in oldonly):.0f} 件")
byp = {}
for p, oid, sku, q, d in recov:
    byp.setdefault(p, [0, 0.0])
    byp[p][0] += 1
    byp[p][1] += q
print("  找回按平台:", {k: (v[0], int(v[1])) for k, v in byp.items()})
byp2 = {}
for p, oid, sku, q, d, t in oldonly:
    byp2.setdefault(p, [0, 0.0])
    byp2[p][0] += 1
    byp2[p][1] += q
print("  旧独有按平台:", {k: (v[0], int(v[1])) for k, v in byp2.items()})

# ================================================================ P3 跨境/跨站点双计核查
print("\n" + "=" * 84)
print("P3 跨境/跨站点 双计核查（新报告平台桶 ∩ 其它站点/跨境报告订单号）")
print("=" * 84)
cands = []
for f in sorted(glob.glob(os.path.join(SNAP, "*.xlsx"))) + sorted(glob.glob(os.path.join(SNAP, "*.csv"))) \
        + sorted(glob.glob(os.path.join(SNAP, "跨境", "*.xlsx"))):
    b = os.path.basename(f)
    m = re.match(r"(Shopee|Lazada|Tiktok)(.+?)站点", b)
    if not m:
        continue
    plat, site = m.group(1), m.group(2)
    if site == "菲律宾":
        continue
    cands.append((plat, site, f))
print("候选重叠文件:", [(p, s, os.path.relpath(f, SNAP)) for p, s, f in cands] or "无")
new_buckets = {}
for p in ("Shopee", "Lazada", "Tiktok"):
    new_buckets[p] = set(RAW.loc[RAW["plat"].map(lambda x: PLATMAP.get(x, x)) == p, "oid"]) - {""}
recov_oids = {p: {x[1] for x in recov if x[0] == p} for p in ("Shopee", "Lazada", "Tiktok")}
oldonly_oids = {p: {x[1] for x in oldonly if x[0] == p} for p in ("Shopee", "Lazada", "Tiktok")}
for plat, site, f in cands:
    ids, nrows, col = read_ids(f)
    if ids is None:
        print(f"  [{site}] {os.path.basename(f)}: 找不到订单号列（{nrows} 行），跳过")
        continue
    bucket = new_buckets.get(PLATMAP.get(plat, plat), set())
    inter_b = ids & bucket
    inter_r = ids & recov_oids.get(PLATMAP.get(plat, plat), set())
    inter_o = ids & oldonly_oids.get(PLATMAP.get(plat, plat), set())
    print(f"  [{site}] {plat}: 报告 {nrows} 行/单号 {len(ids)} 个｜∩新报告{plat}桶 {len(inter_b)}"
          f"｜∩找回漏单 {len(inter_r)}｜∩旧独有 {len(inter_o)}")
    if inter_b:
        print("      桶交集例:", sorted(inter_b)[:5])
        for o in sorted(inter_b)[:5]:
            rs = goid.get(o, [])
            for r in rs[:2]:
                print(f"        新报告行: 平台={r['plat']} 状态={r['订单状态']} 日期={r['d']} 类型={r['otype']} qty={r['q']:.0f}")
    if inter_r:
        print("      ⚠️ 找回漏单与本站/跨境报告重叠:", sorted(inter_r))
    if inter_o:
        print("      ⚠️ 旧独有与本站/跨境报告重叠:", sorted(inter_o))

# ================================================================ P4 “13号部分”
print("\n" + "=" * 84)
print("P4 “13号部分”：9/13~9/14 逐平台新旧对比 + 时间戳边界")
print("=" * 84)
ts_sorted = sorted(RAW["ts"])
print("最早 3 条创建时间:", ts_sorted[:3])
print("最晚 3 条创建时间:", ts_sorted[-3:])
for d in ("2026-09-13", "2026-09-14"):
    for p in ("Shopee", "Lazada", "Tiktok"):
        n = V[(V["d"] == d) & (V["plat"].map(lambda x: PLATMAP.get(x, x)) == p)]["q"].sum()
        o = sum(r[0] for r in OLD[p] if r[3] == d)
        print(f"  {d} {p:<7}: 新报告 {n:>6.0f} vs 旧报告 {o:>6.0f}  ({n - o:+.0f})")

# ================================================================ 总账
print("\n" + "=" * 84)
print("总账")
print("=" * 84)
net_total = segA_net + segB_net
print(f"  段A 净差 = {segA_net:+.0f}；段B 净差 = {segB_net:+.0f}；合计 = {net_total:+.0f}")
print(f"  其中：找回真实漏单 {grand['real']:+.0f} 件｜时间移位 {grand['tshift']:+.0f}｜单号格式/命名分裂 {grand['split']:+.0f}")
print(f"  校验：{grand['real']:+.0f} {grand['tshift']:+.0f} {grand['split']:+.0f} = {grand['real'] + grand['tshift'] + grand['split']:+.0f}（应 = {net_total:+.0f}）")
print(f"  ⇒ 新口径 = trunk 3,123 + {net_total:.0f} + 手工145 + RESELLER154 = {3123 + net_total + 299:.0f}")
