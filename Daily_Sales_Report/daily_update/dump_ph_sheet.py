# -*- coding: utf-8 -*-
"""dump_ph_sheet.py — dump 菲律宾 sheet 行1~3 + 9月块列定位"""
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

p = r'D:\Desktop\work-main\sale_file_gen\菲律宾重算_202609\2026年东南亚销量表Southeast Asia Sales table.xlsx'
wb = load_workbook(p, data_only=True)
ws = wb['Philippines 菲律宾销量表']
print('dims:', ws.max_row, 'x', ws.max_column)
for rn in (1, 2, 3, 4):
    vals = [(get_column_letter(i + 1), v) for i, v in
            enumerate(next(ws.iter_rows(min_row=rn, max_row=rn, values_only=True))) if v is not None]
    print(f'row{rn}:', vals)
