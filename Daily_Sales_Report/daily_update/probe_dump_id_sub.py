# -*- coding: utf-8 -*-
"""probe_dump_id_sub.py — dump 印尼次表结构，查 sheet2 渠道列真实布局"""
import os
import openpyxl

REPO = r"D:\Desktop\work-main\sale_file_gen"
P_ID = os.path.join(REPO, "东南亚次表分发_202609", "Indonesia_202609_sales.xlsx")
P_PH = os.path.join(REPO, "东南亚次表分发_202609", "Philippines_202609_sales.xlsx")
P_TH = os.path.join(REPO, "东南亚次表分发_202609", "Thailand_202609_sales.xlsx")


def dump(path, tag, cn_key, en="Sheet1", maxr=4):
    wb = openpyxl.load_workbook(path, data_only=True)
    print("=" * 78)
    print(f"{tag}: sheets = {wb.sheetnames}")
    s1 = wb[en]
    print(f"  [{en}] dims={s1.dimensions} max_row={s1.max_row} max_col={s1.max_column}")
    for r in range(1, maxr + 1):
        print(f"    row{r}: " + " | ".join(f"{c}={s1.cell(r, c).value!r}" for c in range(1, s1.max_column + 1)))
    s2 = None
    for sn in wb.sheetnames:
        if cn_key in sn:
            s2 = wb[sn]
            break
    print(f"  [{s2.title!r}] dims={s2.dimensions} max_row={s2.max_row} max_col={s2.max_column}")
    for r in range(1, maxr + 1):
        print(f"    row{r}: " + " | ".join(f"{c}={s2.cell(r, c).value!r}" for c in range(1, s2.max_column + 1)))
    wb.close()
    return s2


s2_id = dump(P_ID, "印尼次表", "Indonesia")
s2_ph = dump(P_PH, "菲次表（对照）", "Philippines")
s2_th = dump(P_TH, "泰次表（对照）", "Thailand")

# 逐 SKU 对照：Sheet1(E~H) vs sheet2 各列，找 sheet2 哪几列 == Sheet1 的 9 月渠道值
print("=" * 78)
print("印尼次表逐 SKU 对照（Sheet1 E/F/G/H vs sheet2 全部列）")
wb = openpyxl.load_workbook(P_ID, data_only=True)
s1, s2 = wb["Sheet1"], None
for sn in wb.sheetnames:
    if "Indonesia" in sn:
        s2 = wb[sn]
s1row = {str(s1.cell(r, 1).value).strip(): r for r in range(2, s1.max_row + 1) if s1.cell(r, 1).value}
s2row = {str(s2.cell(r, 1).value).strip(): r for r in range(3, s2.max_row + 1) if s2.cell(r, 1).value}
print(f"  Sheet1 SKU 行 {len(s1row)}；sheet2 SKU 行 {len(s2row)}")
common = [s for s in s1row if s in s2row]
print(f"  共同 SKU {len(common)}；仅 Sheet1: {sorted(set(s1row) - set(s2row))[:8]}；仅 sheet2: {sorted(set(s2row) - set(s1row))[:8]}")
# 对前 5 个共同 SKU 打印两侧值
for s in common[:5]:
    r1, r2 = s1row[s], s2row[s]
    v1 = [s1.cell(r1, c).value for c in range(5, 9)]
    v2 = [s2.cell(r2, c).value for c in range(1, s2.max_column + 1)]
    print(f"    {s}: Sheet1 E~H = {v1}")
    print(f"    {s}: sheet2 全列 = {v2}")
# 列级匹配：sheet2 哪列加总 == Sheet1 各渠道加总
import collections
s1sum = {c: sum(s1.cell(r, c).value or 0 for r in s1row.values() if isinstance(s1.cell(r, c).value, (int, float)))
         for c in range(5, 9)}
print(f"  Sheet1 渠道加总 E~H = {s1sum}")
for c in range(1, s2.max_column + 1):
    t = sum(s2.cell(r, c).value or 0 for r in s2row.values() if isinstance(s2.cell(r, c).value, (int, float)))
    print(f"    sheet2 列{c} 加总 = {t}")
wb.close()
