# -*- coding: utf-8 -*-
"""生成三份国家次表分发件（英文表头，给外国人看）：
- Philippines_202609_sales.xlsx：用「菲律宾重算_202609」副本的新口径（3,329，含手工单 145）重新生成；
- Indonesia / Thailand_202609_sales.xlsx：复制 trunk 已修正表头的现成件（4,026 / 3,238）。
trunk 原件一律不动；分发目录 = 项目根\东南亚次表分发_202609\。
生成逻辑复刻 southeast_asia_sale_file.py §7（列结构/表头/排序/过滤/累计口径完全一致）。
"""
import os
import re
import shutil
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter as gcl

ROOT = r"D:\Desktop\work-main\sale_file_gen"
TRUNK = os.path.join(ROOT, "trunk")
COPY = os.path.join(ROOT, "菲律宾重算_202609", "2026年东南亚销量表Southeast Asia Sales table.xlsx")
TRUNK_DIR = os.path.join(TRUNK, "result", "东南亚销量表", "202609")
OUT_DIR = os.path.join(ROOT, "东南亚次表分发_202609")
MONTH_OFFSET = 9  # 9月
EN_HEADERS = ["Zhongshan Warehouse (Lazada)", "Philippine Warehouse (Lazada)",
              "Philippine Warehouse (Shopee)", "TikTok", "Offline Orders"]
TOT_H = "Sep 2026 Total"
CUM_H = "Jan-Sep 2026 Cumulative"


def num(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return 0.0


# ---------- 1) 商品资料映射（复刻脚本 L222-227 / L602-603） ----------
info = pd.read_excel(os.path.join(TRUNK, "商品资料.xlsx"), sheet_name="商品资料", skiprows=1)
info["SKU"] = info["SKU"].astype(str).str.strip()
for c in ("型号", "类型", "品牌"):
    if c not in info.columns:
        info[c] = None
model_map = info.set_index("SKU")["型号"].to_dict()
type_map = info.set_index("SKU")["类型"].to_dict()
brand_map = info.set_index("SKU")["品牌"].to_dict()

# ---------- 2) 从副本菲律宾 sheet 取数 ----------
wb = load_workbook(COPY, read_only=True, data_only=True)
ws = wb["Philippines 菲律宾销量表"]
WIDTH = 6  # 5 渠道 + 月合计
BASE = 3
rows_raw = list(ws.iter_rows(values_only=True))
wb.close()

sku_rows = [(str(r[0]).strip(), i) for i, r in enumerate(rows_raw) if i >= 2 and r[0] is not None and str(r[0]).strip()]
sku_rows = [(s, i) for s, i in sku_rows if re.match(r"^[GK]\d", s)]
print(f"副本菲律宾 sheet：匹配 ^[GK]\\d 的 SKU 行 {len(sku_rows)} 个")

recs = []
for s, i in sku_rows:
    r = rows_raw[i]
    blk = BASE + (MONTH_OFFSET - 1) * WIDTH          # 9月块首列(1-based) = 51
    vals = [num(r[blk - 1 + k]) for k in range(5)]   # 5 个渠道列
    ytd = 0.0
    for m in range(1, MONTH_OFFSET + 1):             # 1~9 月合计列累加
        v = r[BASE + (m - 1) * WIDTH + 5 - 1]
        ytd += v if isinstance(v, (int, float)) else 0
    recs.append([s, model_map.get(s), type_map.get(s), brand_map.get(s)] + vals + [sum(vals), ytd])

ph = pd.DataFrame(recs, columns=["SKU", "Model", "Type", "Brand"] + EN_HEADERS + [TOT_H, CUM_H])
print(f"菲律宾新口径次表：{len(ph)} 行  当月合计={ph[TOT_H].sum():.0f}  累计合计={ph[CUM_H].sum():.0f}")
print("渠道明细：", {h: ph[h].sum() for h in EN_HEADERS})

# ---------- 3) 写出分发目录 ----------
os.makedirs(OUT_DIR, exist_ok=True)
ph_path = os.path.join(OUT_DIR, "Philippines_202609_sales.xlsx")
ph.to_excel(ph_path, index=False)
print(f"\n已生成 {ph_path}")

for f in ["Indonesia_202609_sales.xlsx", "Thailand_202609_sales.xlsx"]:
    dst = os.path.join(OUT_DIR, f)
    shutil.copy2(os.path.join(TRUNK_DIR, f), dst)
    print(f"已复制 {dst}")

# ---------- 4) 验证 ----------
print("\n" + "=" * 78)
print("验证")
print("=" * 78)

# 4.1 菲律宾新表 vs 副本 9 月块逐 SKU 一致
wb = load_workbook(COPY, read_only=True, data_only=True)
ws = wb["Philippines 菲律宾销量表"]
rows_raw2 = list(ws.iter_rows(values_only=True))
wb.close()
bad = 0
for s, i in sku_rows:
    r = rows_raw2[i]
    blk = BASE + (MONTH_OFFSET - 1) * WIDTH
    src = [num(r[blk - 1 + k]) for k in range(5)]
    row = ph[ph["SKU"] == s].iloc[0]
    new = [num(row[h]) for h in EN_HEADERS]
    if src != new:
        bad += 1
        if bad <= 5:
            print(f"  ★不一致 {s}: 副本={src} 次表={new}")
print(f"4.1 菲律宾次表 vs 副本 9 月块逐 SKU：{len(sku_rows)} 行，不一致 {bad}")

# 4.2 与 trunk 原版次表对比（差异 SKU 清单）
old = pd.read_excel(os.path.join(TRUNK_DIR, "Philippines_202609_sales.xlsx"))
old.columns = [str(c) for c in old.columns]
om = old.set_index(old.columns[0])["Sep 2026 Total"].to_dict()
nm = ph.set_index("SKU")[TOT_H].to_dict()
only_new = {k: v for k, v in nm.items() if k not in om}
only_old = {k: v for k, v in om.items() if k not in nm}
chg = {k: (om[k], nm[k]) for k in om if k in nm and abs(num(om[k]) - num(nm[k])) > 1e-9}
print(f"4.2 新口径 vs trunk 旧口径：行数 {len(om)} → {len(nm)}")
print(f"    新增 SKU：{only_new}")
print(f"    消失 SKU：{only_old or '无'}")
print(f"    值变化 SKU（前 8 个，共 {len(chg)} 个）：")
for k, (a, b) in list(chg.items())[:8]:
    print(f"      {k}: {a:.0f} → {b:.0f}")
print(f"    合计：{sum(num(v) for v in om.values()):.0f} → {nm and sum(nm.values()):.0f}")

# 4.3 三份分发件表头 + 合计
print("\n4.3 分发目录三份表检查：")
for f in ["Philippines_202609_sales.xlsx", "Indonesia_202609_sales.xlsx", "Thailand_202609_sales.xlsx"]:
    p = os.path.join(OUT_DIR, f)
    wbx = load_workbook(p, read_only=True, data_only=True)
    wsx = wbx["Sheet1"]
    rr = list(wsx.iter_rows(values_only=True))
    wbx.close()
    head = [str(x) for x in rr[0] if x is not None]
    hi = next(i for i, h in enumerate(head) if "Total" in h)
    ci = next(i for i, h in enumerate(head) if "Cumulative" in h)
    tot = sum(num(r[hi]) for r in rr[1:] if r[0] is not None)
    cum = sum(num(r[ci]) for r in rr[1:] if r[0] is not None)
    ok_h = ("Sep 2026 Total" in head) and ("Jan-Sep 2026 Cumulative" in head) and not any("Aug" in h for h in head)
    print(f"  [{f}]  行数={len([r for r in rr[1:] if r[0] is not None])}  当月={tot:.0f}  累计={cum:.0f}  表头Sep/Jan-Sep={'✅' if ok_h else '★错'}")

# 4.4 印尼/泰国复制件与 trunk 原件字节一致
print("\n4.4 复制件 vs trunk 原件（应字节一致）：")
for f in ["Indonesia_202609_sales.xlsx", "Thailand_202609_sales.xlsx"]:
    a = open(os.path.join(TRUNK_DIR, f), "rb").read()
    b = open(os.path.join(OUT_DIR, f), "rb").read()
    print(f"  [{f}] {'✅ 一致' if a == b else '★ 不一致'}")

print(f"\n分发目录：{OUT_DIR}")
