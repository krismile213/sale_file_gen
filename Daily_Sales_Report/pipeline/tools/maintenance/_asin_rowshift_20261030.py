# -*- coding: utf-8 -*-
"""ASIN 列「行号错位」检测：分析表每一行的 ASIN 公式是否引用本行 A 列。

判据：公式形如 =XLOOKUP(A{row},源!A:A,源!C:C)
     若捕获的 {row} != 当前行号 → 错位（会显示成别的 SKU 的 ASIN，或空）。
只读。
"""
import os, re
import openpyxl
from openpyxl.utils import column_index_from_string as CI

ROOT = r'C:\Users\dell\Desktop\work-main\sale_file_gen\Daily_Sales_Report'
BOOK = os.path.join(ROOT, '1.每日销售数据表')

JOBS = [
    ("北美加拿大库存销售表.xlsx", "美国-每日库存分析表", "CR"),
    ("北美加拿大库存销售表.xlsx", "加拿大-每日库存分析表", "CR"),
    ("北美加拿大库存销售表.xlsx", "墨西哥-每日库存分析表 ", "CR"),
    ("澳大利亚库存销售表.xlsx", "澳大利亚-每日库存分析表", "CH"),
    ("Interflowd澳大利亚库存销售表.xlsx", "澳大利亚-每日库存分析表", "CH"),
    ("Interflowd沙特库存销售表.xlsx", "沙特-每日库存分析表", "CH"),
    ("沙特库存销售表.xlsx", "沙特阿拉伯-每日库存分析表", "CH"),
    ("阿联酋库存销售表.xlsx", "阿联酋-每日库存分析表", "CH"),
    ("Interflowd阿联酋库存销售表.xlsx", "阿联酋-每日库存分析表", "CH"),
    ("印度库存销售表.xlsx", "印度-每日库存分析表", "CH"),
    ("日本库存销售表.xlsx", "日本-每日库存分析表", "CH"),
    ("欧盟英国库存销售表.xlsx", "英国-每日库存分析表", "CR"),
    ("欧盟英国库存销售表.xlsx", "欧洲-每日库存分析表", "DB"),
]

RE_REF = re.compile(r'XLOOKUP\(\s*\$?A\$?(\d+)', re.I)

print('%-34s %-22s %4s %5s %5s %6s' % ('簿', '分析表', '行', '有公式', '错位', '无公式'))
print('-' * 92)

for f, s, cl in JOBS:
    p = os.path.join(BOOK, f)
    if not os.path.exists(p):
        print('!! 缺文件', f); continue
    wb = openpyxl.load_workbook(p, read_only=True, data_only=False)
    sh = next((x for x in wb.sheetnames if x.strip() == s.strip()), None)
    if sh is None:
        print('!! 无 sheet', f, s); wb.close(); continue
    ws = wb[sh]
    ci = CI(cl)
    rows = nf = bad = nof = 0
    samples = []
    for ri, row in enumerate(ws.iter_rows(min_row=7, max_col=ci, values_only=True), start=7):
        sku = row[0] if row else None
        if not sku or str(sku).strip() in ('', 'SKU', '合计'):
            continue
        rows += 1
        v = row[ci - 1] if len(row) >= ci else None
        if isinstance(v, str) and v.startswith('='):
            nf += 1
            m = RE_REF.search(v)
            if m:
                ref = int(m.group(1))
                if ref != ri:
                    bad += 1
                    if len(samples) < 5:
                        samples.append('r%d(%s)->A%d' % (ri, str(sku).strip(), ref))
        elif v is None or str(v).strip() == '':
            nof += 1
    wb.close()
    flag = '' if bad == 0 and nof == 0 else '  <<<'
    print('%-34s %-22s %4d %5d %5d %6d%s' % (f[:34], s.strip()[:22], rows, nf, bad, nof, flag))
    if samples:
        print('      错位样例:', ' | '.join(samples))
