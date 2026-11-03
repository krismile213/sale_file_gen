# -*- coding: utf-8 -*-
"""只读实测 v6：行2小计行公式形态 + G165/G165PRO 行名 + SEA Total 行2"""
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter as gcl

M = r"D:\Desktop\work-main\sale_file_gen\trunk\result\数据归属周期\202609\月销售总表.xlsx"
R = r"D:\Desktop\work-main\sale_file_gen\trunk\result\东南亚销量表\202609\2026年东南亚销量表Southeast Asia Sales table.xlsx"

wb2 = load_workbook(R, data_only=False)
ph = wb2["Philippines 菲律宾销量表"]; sea = wb2["SEA Total 东南亚总销量表"]
print("== 菲律宾 sheet 行2 小计行（51~56）公式形态 ==")
for col in range(51, 57):
    c = ph.cell(row=2, column=col)
    print(f"   {gcl(col)} = {c.value!r}")
print("\n== SEA Total 行2 小计行（79~87）公式形态 ==")
for col in range(79, 88):
    c = sea.cell(row=2, column=col)
    print(f"   {gcl(col)} = {c.value!r}")
# G165 / G165PRO 行
for sn, ws in [("菲律宾", ph), ("SEA Total", sea)]:
    rows = {}
    for r in range(3, ws.max_row + 1):
        v = ws.cell(row=r, column=1).value
        if v: rows[str(v).strip().upper()] = r
    print(f"\n{sn}: G165在行{rows.get('G165')}  G165PRO在行{rows.get('G165PRO')}")
    if "G165" in rows:
        rr = rows["G165"]
        print(f"   G165行9月块: ", [(gcl(c), ph.cell(row=rr, column=c).value) for c in range(51, 57)] if sn == "菲律宾" else "")
wb2.close()

wb = load_workbook(M, data_only=False)
ms = wb["2026年"]
print("\n== 月销总表行2 ==")
for col in [3, 135, 143, 219, 227, 279, 287]:
    c = ms.cell(row=2, column=col)
    print(f"   {gcl(col)} = {c.value!r}")
rows = {}
for r in range(4, ms.max_row + 1):
    v = ms.cell(row=r, column=1).value
    if v: rows[str(v).strip().upper()] = r
print("\n月销总表: G165在行", rows.get("G165"), " G165PRO在行", rows.get("G165PRO"))
if "G165PRO" in rows:
    rr = rows["G165PRO"]
    print("   G165PRO 行: EM(143)=", ms.cell(row=rr, column=143).value, " HS(227)=", ms.cell(row=rr, column=227).value, " KA(287)=", ms.cell(row=rr, column=287).value)
wb.close()
