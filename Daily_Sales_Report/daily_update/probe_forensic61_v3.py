# -*- coding: utf-8 -*-
"""probe_forensic61_v3.py — 单号长短格式差异核对（只读）

背景：probe_forensic61_v2 按 (交易号, SKU) 精确匹配，发现段B Shopee「旧报告独有 35 键」。
但《核查-菲律宾新订单报告能否一敌多.md》早有结论：9/20~22 三天的少量差异是
「Shopee 单号长短格式差异、运单号 17/17 命中，非丢单」。
⇒ 必须核对这 35 键（以及 50 个新独有键）是不是同一种格式差异，再下"互有漏单"的结论。

做法：
  A. 旧 Shopee 报告 Order ID vs 新报告 交易号 的长度/形态分布；
  B. 35 个旧独有 oid → 在新报告全集里按「前缀/包含」+「同SKU同日期」找对应；
  C. 50 个新独有 oid → 在旧报告里按「前缀」找对应；
  D. 能对上的 → 格式差异（非漏单）；对不上的 → 真漏单/真缺失。
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


def d10(v):
    if v is None:
        return ""
    if hasattr(v, "strftime"):
        return v.strftime("%Y-%m-%d")
    s = str(v).strip()
    m = re.match(r"(\d{4}-\d{2}-\d{2})", s)
    return m.group(1) if m else ""


# ---------- 旧 Shopee 报告（全月，非取消） ----------
od = pd.read_excel(OLD_SHOPEE, dtype=str)
od["订单状态"] = od["Order Status"].astype(str).str.strip()
od.loc[od["订单状态"].str.startswith(RW_PREFIX), "订单状态"] = "Completed"
od["st"] = od["订单状态"].map(SHOPEE_MAP)
od = od[(od["st"].notna()) & (od["st"] != "Cancelled")].copy()
od["oid"] = od["Order ID"].map(nid)
od["sku"] = od["SKU Reference No."].map(nsku)
od["d"] = od["Order Creation Date"].map(d10)
od["q"] = pd.to_numeric(od["Quantity"], errors="coerce").fillna(0)
print("旧Shopee报告列名:", list(od.columns))
print(f"旧Shopee 全月 {len(od)} 行 / {od['q'].sum():.0f} 件")
olen = od["oid"].str.len().value_counts().sort_index()
print("旧报告 订单ID 长度分布:", olen.to_dict())
print("旧报告 含'AAS'的订单ID 行数:", int(od["oid"].str.contains("AAS").sum()))

# ---------- 新报告（全集） ----------
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
shp_new = nd[nd["plat"] == "Shopee"]
print(f"\n新报告 Shopee（已交运/已发货）{len(shp_new)} 行 / {shp_new['q'].sum():.0f} 件")
nlen = shp_new["oid"].str.len().value_counts().sort_index()
print("新报告 交易号 长度分布:", nlen.to_dict())
print("新报告 含'AAS'的交易号 行数:", int(shp_new["oid"].str.contains("AAS").sum()))

# 手工单 oid 形态
man = nd[nd["otype"] == "手工订单"]
print(f"手工单 oid: {man['oid'].unique().tolist()}")

# ---------- B. 段B 旧独有 35 键 → 新报告找对应 ----------
odB = od[(od["d"] >= D2) & (od["d"] <= SEP_HI)]
oldB_keys = {}
for _, r in odB.iterrows():
    if r["oid"] and r["sku"]:
        oldB_keys[(r["oid"], r["sku"])] = oldB_keys.get((r["oid"], r["sku"]), 0) + r["q"]
ndB = shp_new[(shp_new["d"] >= D2) & (shp_new["otype"] != "手工订单")]
newB_keys = {}
for _, r in ndB.iterrows():
    if r["oid"] and r["sku"]:
        newB_keys[(r["oid"], r["sku"])] = newB_keys.get((r["oid"], r["sku"]), 0) + r["q"]
only_o = {k: v for k, v in oldB_keys.items() if k not in newB_keys}
only_n = {k: v for k, v in newB_keys.items() if k not in oldB_keys}
print(f"\n段B Shopee：旧独有 {len(only_o)} 键/{sum(only_o.values()):.0f} 件；新独有 {len(only_n)} 键/{sum(only_n.values()):.0f} 件")

# 新报告全集索引：前缀(前14字符) → [(oid, sku, d, q)]；以及 (sku,d) → [...]
new_pref = {}
new_sd = {}
for _, r in nd.iterrows():
    if r["oid"]:
        new_pref.setdefault(r["oid"][:14], []).append((r["oid"], r["sku"], r["d"], r["q"], r["订单状态"], r["otype"]))
    if r["sku"] and r["d"]:
        new_sd.setdefault((r["sku"], r["d"]), []).append((r["oid"], r["q"], r["订单状态"], r["otype"]))

print("\n===== B. 35 个旧独有键 → 新报告全集核对 =====")
hit_pref, hit_sd, none_at_all = 0, 0, 0
details = []
for (oid, sku), q in sorted(only_o.items()):
    # 1) 前缀命中（长短格式差异）
    cands = new_pref.get(oid[:14], [])
    cands = [c for c in cands if c[1] == sku]
    if cands:
        hit_pref += 1
        details.append(("前缀同SKU", oid, sku, q, cands[:2]))
        continue
    # 2) 同 SKU 同日期（单号完全对不上）
    cands2 = [c for c in new_sd.get((sku, odB[odB["oid"] == oid]["d"].iloc[0] if len(odB[odB["oid"] == oid]) else ""), [])]
    # 上面按 key 直接查可能空，改为遍历该 oid 的日期
    dts = odB[odB["oid"] == oid]["d"].tolist()
    cands2 = []
    for dd in dts:
        cands2.extend(new_sd.get((sku, dd), []))
    if cands2:
        hit_sd += 1
        details.append(("同SKU同日期(单号不同)", oid, sku, q, cands2[:2]))
    else:
        none_at_all += 1
        details.append(("新报告全无", oid, sku, q, []))
print(f"  前缀同SKU命中(长短格式差异) = {hit_pref} 键")
print(f"  同SKU同日期但单号不同     = {hit_sd} 键")
print(f"  新报告全集确实没有         = {none_at_all} 键")
for tag, oid, sku, q, c in details:
    if tag != "前缀同SKU":
        print(f"    [{tag}] {oid} {sku} {q:.0f}件 → 新报告候选: {c}")

# ---------- C. 50 个新独有键 → 旧报告找对应 ----------
old_pref = {}
for _, r in od.iterrows():
    if r["oid"]:
        old_pref.setdefault(r["oid"][:14], []).append((r["oid"], r["sku"], r["d"], r["q"]))
print("\n===== C. 50 个新独有键 → 旧报告全集核对 =====")
hit_pref2, none2 = 0, 0
for (oid, sku), q in sorted(only_n.items(), key=lambda x: -x[1]):
    cands = [c for c in old_pref.get(oid[:14], []) if c[1] == sku]
    if cands:
        hit_pref2 += 1
        print(f"    [前缀同SKU] {oid} {sku} {q:.0f}件 → 旧报告: {cands[:2]}")
    else:
        none2 += 1
print(f"  前缀同SKU命中 = {hit_pref2} 键；旧报告全无 = {none2} 键")

# ---------- D. 若按前缀归一，段B Shopee 净差变成多少 ----------
def norm_oid(o):
    return o[:14] if o else o
odB2 = odB.copy(); odB2["nk"] = odB2["oid"].map(norm_oid)
ndB2 = ndB.copy(); ndB2["nk"] = ndB2["oid"].map(norm_oid)
o2 = {}
for _, r in odB2.iterrows():
    if r["nk"] and r["sku"]:
        o2[(r["nk"], r["sku"])] = o2.get((r["nk"], r["sku"]), 0) + r["q"]
n2 = {}
for _, r in ndB2.iterrows():
    if r["nk"] and r["sku"]:
        n2[(r["nk"], r["sku"])] = n2.get((r["nk"], r["sku"]), 0) + r["q"]
oo = {k: v for k, v in o2.items() if k not in n2}
nn = {k: v for k, v in n2.items() if k not in o2}
print("\n===== D. 按 oid 前14字符归一后，段B Shopee =====")
print(f"  旧 {sum(o2.values()):.0f} 件 vs 新(非手工) {sum(n2.values()):.0f} 件，净差 {sum(n2.values())-sum(o2.values()):+.0f}")
print(f"  旧独有 {len(oo)} 键/{sum(oo.values()):.0f} 件；新独有 {len(nn)} 键/{sum(nn.values()):.0f} 件")
for k, v in sorted(oo.items(), key=lambda x: -x[1])[:12]:
    print(f"    旧独 {k} {v:.0f} 件")
