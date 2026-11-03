# -*- coding: utf-8 -*-
"""逐格 diff：trunk 成品 vs 菲律宾重算_202609 副本（对齐前确认传播面）。
只读，不写任何文件。
"""
import os
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

REPO = r'D:\Desktop\work-main\sale_file_gen'
TRUNK_MON = os.path.join(REPO, 'trunk', 'result', '数据归属周期', '202609', '月销售总表.xlsx')
TRUNK_SEA = os.path.join(REPO, 'trunk', 'result', '东南亚销量表', '202609', '2026年东南亚销量表Southeast Asia Sales table.xlsx')
TRUNK_PH = os.path.join(REPO, 'trunk', 'result', '东南亚销量表', '202609', 'Philippines_202609_sales.xlsx')
COPY_MON = os.path.join(REPO, '菲律宾重算_202609', '月销售总表.xlsx')
COPY_SEA = os.path.join(REPO, '菲律宾重算_202609', '2026年东南亚销量表Southeast Asia Sales table.xlsx')
COPY_PH = os.path.join(REPO, '菲律宾重算_202609', 'Philippines_202609_sales.xlsx')


def n0(v):
    if v is None:
        return 0
    if isinstance(v, str):
        s = v.strip()
        if s in ('00', ''):
            return 0
        try:
            return float(s)
        except ValueError:
            return 0
    return v if isinstance(v, (int, float)) else 0


def diff_book(path_a, path_b, label, max_show=40):
    print(f'\n===== {label} =====')
    print(f'  A(trunk): {path_a}')
    print(f'  B(副本):  {path_b}')
    wa = load_workbook(path_a, data_only=True)
    wb = load_workbook(path_b, data_only=True)
    print(f'  sheets A: {wa.sheetnames}')
    print(f'  sheets B: {wb.sheetnames}')
    only_a = [s for s in wa.sheetnames if s not in wb.sheetnames]
    only_b = [s for s in wb.sheetnames if s not in wa.sheetnames]
    if only_a:
        print(f'  ⚠️ 仅 A 有: {only_a}')
    if only_b:
        print(f'  ⚠️ 仅 B 有: {only_b}')
    total = 0
    for sn in wa.sheetnames:
        if sn not in wb.sheetnames:
            continue
        sa, sb = wa[sn], wb[sn]
        diffs = []
        maxr = max(sa.max_row, sb.max_row)
        maxc = max(sa.max_column, sb.max_column)
        for r in range(1, maxr + 1):
            for c in range(1, maxc + 1):
                va, vb = sa.cell(r, c).value, sb.cell(r, c).value
                fa = isinstance(va, str) and va.startswith('=')
                fb = isinstance(vb, str) and vb.startswith('=')
                if fa or fb:
                    continue  # 公式格单独看
                if n0(va) != n0(vb):
                    diffs.append((f'{get_column_letter(c)}{r}', n0(va), n0(vb)))
        if diffs:
            total += len(diffs)
            print(f'  [{sn}] {len(diffs)} 处静态值差异（前 {min(max_show, len(diffs))} 处）:')
            for ref, va, vb in diffs[:max_show]:
                print(f'      {ref}: trunk={va:g}  副本={vb:g}  Δ={n0(vb)-n0(va):+g}')
            if len(diffs) > max_show:
                print(f'      … 其余 {len(diffs)-max_show} 处省略')
    print(f'  静态值差异合计: {total}')
    wa.close(); wb.close()
    return total


t1 = diff_book(TRUNK_MON, COPY_MON, '月销售总表')
t2 = diff_book(TRUNK_SEA, COPY_SEA, '东南亚销量表')
t3 = diff_book(TRUNK_PH, COPY_PH, '菲律宾次表')
print('\n===== 汇总 =====')
print(f'月销总表差异 {t1}｜东南亚表差异 {t2}｜菲次表差异 {t3}')
