# -*- coding: utf-8 -*-
"""只读实测③：2026年 sheet 行1分组+行2月标签全量扫描，精确定位各组9月列"""
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter as gcl

M = r"D:\Desktop\work-main\sale_file_gen\菲律宾重算_202609\月销售总表.xlsx"
wb = load_workbook(M, data_only=False)
ws = wb["2026年"]
# 行1 分组标签（取非空且非数字的）
r1 = [(c.column, str(c.value)) for c in ws[1] if c.value is not None and not isinstance(c.value, (int, float))]
print("=== 行1 分组标签 ===")
for col, v in r1:
    print(f"  {gcl(col)}({col}) = {v[:30]}")
# 行2 月标签
r2 = [(c.column, str(c.value)) for c in ws[2] if c.value is not None]
print(f"\n=== 行2 月标签（{len(r2)} 个）===")
for col, v in r2:
    print(f"  {gcl(col)}({col}) = {v[:20]}")
wb.close()
