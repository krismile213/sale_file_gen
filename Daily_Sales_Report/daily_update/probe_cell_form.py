# -*- coding: utf-8 -*-
"""只读实测 v5：数据行单元格公式形态（决定写表探针的写法）"""
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter as gcl

M = r"D:\Desktop\work-main\sale_file_gen\trunk\result\数据归属周期\202609\月销售总表.xlsx"
R = r"D:\Desktop\work-main\sale_file_gen\trunk\result\东南亚销量表\202609\2026年东南亚销量表Southeast Asia Sales table.xlsx"

wb = load_workbook(M, data_only=False)   # 非 read_only，看公式
ms = wb["2026年"]
print("月销总表 max_row =", ms.max_row)
# 定位 G445 / G376 / G569 行
targets = {"G445", "G376", "G569", "G175"}
found = {}
for r in range(4, ms.max_row + 1):
    v = ms.cell(row=r, column=1).value
    if v and str(v).strip().upper() in targets:
        found[str(v).strip().upper()] = r
        if len(found) == len(targets): break
for k, r in found.items():
    print(f"\n{k} (行{r})：")
    for col, name in [(143, "EM 菲9月"), (227, "HS 东南亚9月"), (287, "KA 总计9月"), (142, "EL 菲8月"), (226, "HR 东南亚8月")]:
        c = ms.cell(row=r, column=col)
        print(f"   {name}({gcl(col)}) = {c.value!r}  type={c.data_type}")
# 行3合计行
print("\n行3合计行：")
for col in [143, 227, 287, 3]:
    c = ms.cell(row=3, column=col)
    print(f"   {gcl(col)} = {c.value!r}")
# 最后一行（补行的样式参考）
last = ms.max_row
print(f"\n末行({last}) A={ms.cell(row=last, column=1).value!r} EM={ms.cell(row=last, column=143).value!r}")
wb.close()

wb2 = load_workbook(R, data_only=False)
ph = wb2["Philippines 菲律宾销量表"]
print("\n\n菲律宾 sheet max_row =", ph.max_row, " max_col =", ph.max_column)
found2 = {}
for r in range(3, ph.max_row + 1):
    v = ph.cell(row=r, column=1).value
    if v and str(v).strip().upper() in targets:
        found2[str(v).strip().upper()] = r
for k, r in found2.items():
    print(f"\n{k} (行{r})：")
    for col, name in [(51, "AY 跨境LZD"), (52, "AZ 本土LZD"), (53, "BA 本土SHP"), (54, "BB Tiktok"), (55, "BC 线下"), (56, "BD 9月合计")]:
        c = ph.cell(row=r, column=col)
        print(f"   {name}({gcl(col)}) = {c.value!r}")
print(f"\n行2（小计行）：AY={ph.cell(row=2,column=51).value!r} AZ={ph.cell(row=2,column=52).value!r} BA={ph.cell(row=2,column=53).value!r} BB={ph.cell(row=2,column=54).value!r} BC={ph.cell(row=2,column=55).value!r} BD={ph.cell(row=2,column=56).value!r}")
print(f"末行({ph.max_row}) A={ph.cell(row=ph.max_row,column=1).value!r} BD={ph.cell(row=ph.max_row,column=56).value!r}")
wb2.close()
