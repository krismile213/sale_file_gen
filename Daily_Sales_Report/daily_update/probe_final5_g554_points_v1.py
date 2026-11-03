# -*- coding: utf-8 -*-
"""probe_final5_g554_points_v1.py — dump 3,485 口径需要 +2 的全部传播点现值

传播链（G554 Shopee 9月 +2）：
  东南亚销量表: Philippines!G554行 BA+2/BD+2｜row2 BA/BD｜SEA Total 菲9月｜销量情况 AA10
  月销售总表  : 【2026年】G554行 EM+2/HS+2/KA+2/KJ+2｜行3 SUM(公式)｜销量情况 G10/AA10
  菲律宾次表  : G554 行对应渠道 +2（先 dump 结构）
  印尼/泰国次表: 不变
"""
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

SEA = r"D:\Desktop\work-main\sale_file_gen\菲律宾重算_202609\2026年东南亚销量表Southeast Asia Sales table.xlsx"
MON = r"D:\Desktop\work-main\sale_file_gen\菲律宾重算_202609\月销售总表.xlsx"
PH_SUB = r"D:\Desktop\work-main\sale_file_gen\东南亚次表分发_202609\Philippines_202609_sales.xlsx"

print("=" * 88)
print("① 东南亚销量表")
print("=" * 88)
wb = load_workbook(SEA, data_only=True)
ws = wb["Philippines 菲律宾销量表"]
g554_row = None
for r in range(3, ws.max_row + 1):
    if str(ws.cell(r, 1).value or "").strip().upper() == "G554":
        g554_row = r
        break
print(f"  Philippines!G554 行号 = {g554_row}")
print(f"  现值: AY(跨境LZD)={ws.cell(g554_row, 51).value} AZ(本土LZD)={ws.cell(g554_row, 52).value} "
      f"BA(SHP)={ws.cell(g554_row, 53).value} BB(TK)={ws.cell(g554_row, 54).value} "
      f"BC(线下)={ws.cell(g554_row, 55).value} BD(9月合计)={ws.cell(g554_row, 56).value}")
print(f"  row2 合计: BA={ws.cell(2, 53).value} BD={ws.cell(2, 56).value}")

ws2 = wb["SEA Total 东南亚总销量表"]
hdr2 = [(get_column_letter(i + 1), v) for i, v in
        enumerate(next(ws2.iter_rows(min_row=1, max_row=1, values_only=True))) if v]
print(f"  SEA Total row1: {hdr2}")
g554_row2 = None
for r in range(3, ws2.max_row + 1):
    if str(ws2.cell(r, 1).value or "").strip().upper() == "G554":
        g554_row2 = r
        break
print(f"  SEA Total G554 行号 = {g554_row2}")
if g554_row2:
    vals = [(get_column_letter(i + 1), v) for i, v in
            enumerate(next(ws2.iter_rows(min_row=g554_row2, max_row=g554_row2, values_only=True))) if v is not None]
    print(f"  SEA Total G554 现值: {vals}")
print(f"  SEA Total row2 合计行: {[(get_column_letter(i+1), v) for i, v in enumerate(next(ws2.iter_rows(min_row=2, max_row=2, values_only=True))) if v is not None]}")

ws3 = wb["2023-2026年销量情况 Sales"]
hdr3 = [(get_column_letter(i + 1), v) for i, v in
        enumerate(next(ws3.iter_rows(min_row=1, max_row=1, values_only=True))) if v]
print(f"  销量情况 row1: {hdr3}")
r10 = [(get_column_letter(i + 1), v) for i, v in
       enumerate(next(ws3.iter_rows(min_row=10, max_row=10, values_only=True))) if v is not None]
print(f"  销量情况 row10(9月): {r10}")
r14 = [(get_column_letter(i + 1), v) for i, v in
       enumerate(next(ws3.iter_rows(min_row=14, max_row=14, values_only=True))) if v is not None]
print(f"  销量情况 row14(汇总): {r14}")
wb.close()

print("\n" + "=" * 88)
print("② 月销售总表")
print("=" * 88)
wb = load_workbook(MON, data_only=True)
print("  sheets:", wb.sheetnames)
ws = wb["【2026年】"]
# 列锚点：EM(135)菲｜HS(219)东南亚｜KA(279)总计｜KJ(296)当年总销量｜KK(297)SKU
g554_row = None
for r in range(4, ws.max_row + 1):
    if str(ws.cell(r, 297).value or "").strip().upper() == "G554":
        g554_row = r
        break
print(f"  【2026年】G554 行号 = {g554_row}")
if g554_row:
    print(f"  现值: EM(135)={ws.cell(g554_row, 135).value} HS(219)={ws.cell(g554_row, 219).value} "
          f"KA(279)={ws.cell(g554_row, 279).value} KJ(296)={ws.cell(g554_row, 296).value}")
print(f"  行3 合计: EM3={ws.cell(3, 135).value} HS3={ws.cell(3, 219).value} KA3={ws.cell(3, 279).value}")
# 公式版（非 data_only）看行3
wb2 = load_workbook(MON, data_only=False)
wsf = wb2["【2026年】"]
print(f"  行3 公式: EM3={wsf.cell(3, 135).value} HS3={wsf.cell(3, 219).value} KA3={wsf.cell(3, 279).value}")
wb2.close()

if "2021-2026年销量情况" in wb.sheetnames:
    ws3 = wb["2021-2026年销量情况"]
    r10 = [(get_column_letter(i + 1), v) for i, v in
           enumerate(next(ws3.iter_rows(min_row=10, max_row=10, values_only=True))) if v is not None]
    print(f"  月销总表销量情况 row10(9月): {r10}")
else:
    print("  ⚠️ 无 2021-2026年销量情况 sheet")
wb.close()

print("\n" + "=" * 88)
print("③ 菲律宾次表")
print("=" * 88)
wb = load_workbook(PH_SUB, data_only=True)
print("  sheets:", wb.sheetnames)
ws = wb[wb.sheetnames[0]]
print(f"  dims: {ws.max_row} x {ws.max_column}")
hdr = [(get_column_letter(i + 1), v) for i, v in
       enumerate(next(ws.iter_rows(min_row=1, max_row=1, values_only=True))) if v is not None]
print(f"  row1: {hdr}")
for rn in (2, 3):
    vals = [(get_column_letter(i + 1), v) for i, v in
            enumerate(next(ws.iter_rows(min_row=rn, max_row=rn, values_only=True))) if v is not None]
    print(f"  row{rn}: {vals}")
g554_row = None
for r in range(2, ws.max_row + 1):
    if str(ws.cell(r, 1).value or "").strip().upper() == "G554":
        g554_row = r
        break
print(f"  G554 行号 = {g554_row}")
if g554_row:
    vals = [(get_column_letter(i + 1), v) for i, v in
            enumerate(next(ws.iter_rows(min_row=g554_row, max_row=g554_row, values_only=True))) if v is not None]
    print(f"  G554 现值: {vals}")
wb.close()
