# -*- coding: utf-8 -*-
"""只读核查：菲律宾202609订单报告.xlsx 中 订单ID=403990300315398144 / 运单号=YFDZ123456789
是否计入 9 月东南亚销量。全链路：新报告 → 旧三报告(+线下) → 中间产物 → 成品表。"""
import os, sys
import pandas as pd

R = r"D:\Desktop\work-main\sale_file_gen"
SNAP = os.path.join(R, "trunk", "data", "数据归属周期", "202609", "快照日期", "20261006", "东南亚")
TMP = os.path.join(R, "trunk", "sale_file_tmp")
TARGET_OID = "403990300315398144"
TARGET_WAYBILL = "YFDZ123456789"

def num(v):
    try:
        return float(v)
    except Exception:
        return 0.0

print("=" * 78)
print("① 新报告《菲律宾202609订单报告.xlsx》中的原始行")
print("=" * 78)
p_new = os.path.join(SNAP, "菲律宾202609订单报告.xlsx")
df = pd.read_excel(p_new)
print("列名：", list(df.columns))
print("总行数：", len(df))
# 找订单ID列与运单号列
oid_col = way_col = date_col = sku_col = qty_col = st_col = None
for c in df.columns:
    cs = str(c).strip()
    if "订单ID" in cs or "订单号" in cs and "运单" not in cs:
        if oid_col is None: oid_col = c
    if "运单" in cs: way_col = c
    if "下单" in cs and "时间" in cs: date_col = c
    if "SKU" in cs.upper() and sku_col is None: sku_col = c
    if "数量" in cs and qty_col is None: qty_col = c
    if "状态" in cs: st_col = c
print(f"\n定位列：订单ID={oid_col!r} 运单号={way_col!r} 日期={date_col!r} SKU={sku_col!r} 数量={qty_col!r} 状态={st_col!r}")

m = pd.Series(False, index=df.index)
if oid_col is not None:
    m |= df[oid_col].astype(str).str.strip() == TARGET_OID
if way_col is not None:
    m |= df[way_col].astype(str).str.strip().str.upper() == TARGET_WAYBILL
sub = df[m]
print(f"\n命中行数：{len(sub)}")
if len(sub):
    with pd.option_context("display.width", 250, "display.max_colwidth", 40):
        print(sub.to_string())

print()
print("=" * 78)
print("② 旧三报告 + 线下报告中是否含该订单/运单号（按字符串全局搜）")
print("=" * 78)
for f in ["Lazada菲律宾站点202609订单报告.xlsx", "Shopee菲律宾站点202609订单报告.xlsx",
          "Tiktok菲律宾站点202609订单报告.csv", "菲律宾KSON仓202609线下订单报告.xlsx"]:
    fp = os.path.join(SNAP, f)
    if not os.path.exists(fp):
        print(f"  [缺] {f}"); continue
    hit_o = hit_w = 0
    try:
        if f.endswith(".csv"):
            d2 = pd.read_csv(fp, dtype=str, low_memory=False)
        else:
            d2 = pd.read_excel(fp, dtype=str)
        for c in d2.columns:
            s = d2[c].astype(str).str.strip()
            hit_o += int((s == TARGET_OID).sum())
            hit_w += int((s.str.upper() == TARGET_WAYBILL).sum())
        print(f"  {f:<42} 行数={len(d2):>6}  订单ID命中={hit_o}  运单号命中={hit_w}")
    except Exception as e:
        print(f"  {f}: 读取失败 {e}")

print()
print("=" * 78)
print("③ 中间产物：菲律宾 9 月是否已含该 SKU 的量")
print("=" * 78)
# 3a. total_southeast_asia_summary.csv
try:
    summ = pd.read_csv(os.path.join(TMP, "total_southeast_asia_summary.csv"), index_col=0)
    print("total_southeast_asia_summary.csv（菲律宾行，9月列）：")
    ph_rows = summ[summ.index.astype(str).str.contains("菲")]
    print(ph_rows.to_string())
except Exception as e:
    print("summary 读取失败：", e)

# 3b. se_a_sku_sale_df.csv 找该 SKU
if len(sub) and sku_col is not None:
    skus = sub[sku_col].astype(str).str.strip().unique().tolist()
    print(f"\n该订单涉及 SKU：{skus}")
    try:
        sk = pd.read_csv(os.path.join(TMP, "se_a_sku_sale_df.csv"), index_col=0)
        sk.columns = [int(c) for c in sk.columns]
        sku_col_idx = sk.columns.max()
        hit = sk[sk[sku_col_idx].astype(str).str.strip().isin(skus)]
        print(f"\nse_a_sku_sale_df.csv 中 {skus} 的行（列=Excel列号，菲律宾块 9月列应为 BC/BD 一带）：")
        print(hit.to_string())
    except Exception as e:
        print("se_a_sku_sale_df 读取失败：", e)

print()
print("=" * 78)
print("④ 成品表《2026年东南亚销量表》菲律宾 sheet：该 SKU 的 9 月值")
print("=" * 78)
wb_path = os.path.join(R, "trunk", "result", "东南亚销量表", "202609", "2026年东南亚销量表Southeast Asia Sales table.xlsx")
try:
    from openpyxl import load_workbook
    from openpyxl.utils import get_column_letter as gcl
    wb = load_workbook(wb_path, read_only=True, data_only=True)
    sn = [s for s in wb.sheetnames if "Philippines" in s][0]
    ws = wb[sn]
    print(f"sheet: {sn}  max_row={ws.max_row} max_col={ws.max_column}")
    if len(sub) and sku_col is not None:
        for r in ws.iter_rows(min_row=1, max_row=3):
            vals = [(gcl(c.column), str(c.value)[:10]) for c in r if c.value is not None]
            print(f"  r{r[0].row}:", vals[:26])
    wb.close()
except Exception as e:
    print("成品表读取失败：", e)
