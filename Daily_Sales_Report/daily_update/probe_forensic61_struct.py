# -*- coding: utf-8 -*-
"""probe_forensic61_struct.py — 61 件取证第一步：摸清各原始数据源的列结构（只读）"""
import os
import pandas as pd
from openpyxl import load_workbook

SNAP = r"D:\Desktop\work-main\sale_file_gen\trunk\data\数据归属周期\202609\快照日期\20261006\东南亚"
MON = r"C:\Users\dell\Desktop\DAILY MONITORING - SEPTEMBER 1-13(1).xlsx"

print("=" * 78)
print("旧三报告 + 新报告列结构")
for f in ["Shopee菲律宾站点202609订单报告.xlsx", "Lazada菲律宾站点202609订单报告.xlsx",
          "Tiktok菲律宾站点202609订单报告.csv", "菲律宾202609订单报告.xlsx"]:
    p = os.path.join(SNAP, f)
    if f.endswith(".csv"):
        d = pd.read_csv(p, nrows=3, low_memory=False)
    else:
        d = pd.read_excel(p, nrows=3, dtype=str)
    print(f"\n--- {f}  ({len(d.columns)} 列)")
    for i, c in enumerate(d.columns):
        print(f"    [{i}] {c!r}  样例={d.iloc[0, i]!r}")

print()
print("=" * 78)
print("监控表 sheet 列表 + 关键页结构")
wb = load_workbook(MON, read_only=True, data_only=True)
print(f"  sheets ({len(wb.sheetnames)}): {wb.sheetnames}")
for sn in ["SHP Data", "LZD Data", "Tiktok Data"]:
    if sn not in wb.sheetnames:
        continue
    ws = wb[sn]
    rows = ws.iter_rows(min_row=1, max_row=2, values_only=True)
    hdr = next(rows)
    smp = next(rows, None)
    print(f"\n--- {sn}  (max_row={ws.max_row}, max_col={ws.max_column})")
    for i, h in enumerate(hdr):
        print(f"    [{i}] {h!r}  样例={smp[i] if smp else None!r}")
wb.close()
