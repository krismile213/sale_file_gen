# -*- coding: utf-8 -*-
"""探针：沙特/阿联酋 interflow 库存报告 ASIN 空缺情况 + 商品资料能否补

2026-10-30 用户反馈：interflow 沙特那张表很多 ASIN 是空的，为此加了一份商品资料。
本脚本回答：缺多少、能不能用商品资料 SKU→ASIN 补、补不上的有多少。
"""
import csv
import os

import openpyxl

SNAP = r"C:\Users\dell\Desktop\work-main\sale_file_gen\trunk\data\数据归属周期\202610\快照日期\20261030"
PROD_CANDIDATES = [
    r"C:\Users\dell\Desktop\work-main\资产文件\商品资料.xlsx",
    r"C:\Users\dell\Desktop\work-main\sale_file_gen\商品资料.xlsx",
]


def rows_of(path):
    with open(path, encoding="utf-8-sig", errors="replace") as f:
        return list(csv.DictReader(f))


print("=== 快照里的 interflow / 中东 库存报告 ===")
for fn in sorted(os.listdir(SNAP)):
    if ("沙特" in fn or "阿联酋" in fn or "阿拉伯" in fn) and ("库存" in fn or "inventory" in fn.lower()):
        p = os.path.join(SNAP, fn)
        rs = rows_of(p)
        if not rs:
            print(f"  {fn}: 空")
            continue
        cols = list(rs[0].keys())
        ac = next((c for c in cols if c.strip().lower() == "asin"), None)
        sc = next((c for c in cols if c.strip().lower() == "sku"), None)
        if ac is None:
            print(f"  {fn}: 没有 asin 列，列名={cols[:8]}")
            continue
        blank = [r for r in rs if not (r.get(ac) or "").strip()]
        skus = [r.get(sc, "").strip() for r in rs]
        print(f"  {fn}")
        print(f"    行数 {len(rs)}  ASIN列={ac!r}  SKU列={sc!r}")
        print(f"    ASIN 为空 {len(blank)} 行 ({len(blank)/len(rs)*100:.0f}%)"
              f"，涉及 SKU {len(set(r.get(sc,'').strip() for r in blank))} 个")
        print(f"    空 ASIN 的 SKU 样例：{sorted(set(r.get(sc,'').strip() for r in blank))[:8]}")

print("\n=== 商品资料 ASIN 覆盖 ===")
prod_path = next((p for p in PROD_CANDIDATES if os.path.exists(p)), None)
print("用:", prod_path)
wb = openpyxl.load_workbook(prod_path, read_only=True, data_only=True)
print("sheets:", wb.sheetnames[:10])
ws = wb[wb.sheetnames[0]]
hdr = [("" if v is None else str(v)).strip() for v in next(ws.iter_rows(min_row=1, max_row=1, values_only=True))]
print("表头前 18:", hdr[:18])
i_asin = next((i for i, h in enumerate(hdr) if h.upper() == "ASIN"), None)
i_sku = next((i for i, h in enumerate(hdr) if h.upper() == "SKU"), None)
print("ASIN 列:", i_asin, "SKU 列:", i_sku)

sku2asin = {}
n_blank = 0
for r in ws.iter_rows(min_row=2, max_col=max(i_asin, i_sku) + 1, values_only=True):
    sku = ("" if r[i_sku] is None else str(r[i_sku])).strip()
    if not sku:
        continue
    a = ("" if r[i_asin] is None else str(r[i_asin])).strip()
    if a:
        sku2asin[sku] = a
    else:
        n_blank += 1
wb.close()
print(f"商品资料 SKU 数 {len(sku2asin) + n_blank}，有 ASIN {len(sku2asin)}，ASIN 空 {n_blank}")

# 交叉：沙特空 ASIN 能不能补
for fn in sorted(os.listdir(SNAP)):
    if ("沙特" in fn or "阿拉伯" in fn or "阿联酋" in fn) and "库存" in fn:
        p = os.path.join(SNAP, fn)
        rs = rows_of(p)
        if not rs:
            continue
        cols = list(rs[0].keys())
        ac = next((c for c in cols if c.strip().lower() == "asin"), None)
        sc = next((c for c in cols if c.strip().lower() == "sku"), None)
        if not ac or not sc:
            continue
        blanks = {(r.get(sc) or "").strip() for r in rs if not (r.get(ac) or "").strip()}
        fillable = sorted(s for s in blanks if s in sku2asin)
        unfill = sorted(s for s in blanks if s not in sku2asin)
        print(f"\n{fn}: 空 ASIN 的 SKU {len(blanks)} 个 → 商品资料可补 {len(fillable)}，补不上 {len(unfill)}")
        print(f"  可补样例: {[(s, sku2asin[s]) for s in fillable[:5]]}")
        print(f"  补不上: {unfill[:10]}")
