# -*- coding: utf-8 -*-
"""全局扫描（快速版）：13 本簿里所有含 ASIN 表头的列。
用 iter_rows 顺序遍历，read_only 下比 cell() 快一个量级。
只读。
"""
import csv, os, glob, openpyxl
from openpyxl.utils import get_column_letter

SNAP = r'C:\Users\dell\Desktop\work-main\sale_file_gen\trunk\data\数据归属周期\202610\快照日期\20261030'
DSR = r'C:\Users\dell\Desktop\work-main\sale_file_gen\Daily_Sales_Report'

def enc_order(name):
    if '日本' in name or '澳大利亚' in name:
        return ['shift_jis','utf-8-sig','gb18030','latin-1']
    return ['utf-8-sig','gb18030','shift_jis','latin-1']

def read_csv_any(p, name=''):
    for enc in enc_order(name):
        try:
            with open(p, encoding=enc, newline='') as fh:
                return list(csv.reader(fh))
        except (UnicodeDecodeError, LookupError):
            continue
    return []

stock = {}
for p in sorted(glob.glob(os.path.join(SNAP, '*管理库存报告*.csv'))):
    name = os.path.basename(p)
    rows = read_csv_any(p, name)
    if not rows: continue
    hdr = [str(c).strip().lower() for c in rows[0]]
    try:
        i_sku, i_asin = hdr.index('sku'), hdr.index('asin')
    except ValueError:
        if len(hdr) >= 24: i_sku, i_asin = 0, 2
        else: continue
    for r in rows[1:]:
        if len(r) <= max(i_sku,i_asin): continue
        s = str(r[i_sku]).strip(); a = str(r[i_asin]).strip()
        if s and a: stock.setdefault(s, set()).add(a)

print('库存端可查 SKU:', len(stock))
print()

for f in sorted(glob.glob(os.path.join(DSR, '1.每日销售数据表', '*.xlsx'))):
    bn = os.path.basename(f)
    if bn.startswith('~$'):
        continue
    try:
        wb = openpyxl.load_workbook(f, read_only=True, data_only=False)
    except Exception as e:
        print('!! 打不开', bn, e); continue
    for ws in wb.worksheets:
        try:
            it = ws.iter_rows(values_only=True)
            head = []
            asin_c = sku_c = None
            nf = nv = nblank = 0
            blanks = []
            for ri, row in enumerate(it, 1):
                if ri <= 8:
                    head.append(row)
                    up = [(str(v).strip().upper() if v is not None else '') for v in row]
                    if asin_c is None and 'ASIN' in up:
                        asin_c = up.index('ASIN')
                        for j, u in enumerate(up):
                            if u in ('SKU', 'MSKU'):
                                sku_c = j; break
                        continue
                    else:
                        continue
                if asin_c is None:
                    continue
                av = row[asin_c] if len(row) > asin_c else None
                sv = row[sku_c] if (sku_c is not None and len(row) > sku_c) else (row[0] if row else None)
                if av is None or (isinstance(av, str) and not av.strip()):
                    if isinstance(sv, str) and sv.strip():
                        nblank += 1
                        if len(blanks) < 300: blanks.append(sv.strip())
                elif isinstance(av, str) and av.startswith('='):
                    nf += 1
                else:
                    nv += 1
            if asin_c is None:
                continue
            rec = [s for s in blanks if s in stock]
            print('%-36s | %-18s | %s | 公式%4d 值%4d 空%4d | 库存可补%3d'
                  % (bn[:36], ws.title[:18], get_column_letter(asin_c+1), nf, nv, nblank, len(rec)))
            if nblank and len(blanks) <= 10:
                print('        空ASIN的SKU:', blanks)
        except Exception as e:
            print('!! sheet', bn, ws.title, e)
    wb.close()
