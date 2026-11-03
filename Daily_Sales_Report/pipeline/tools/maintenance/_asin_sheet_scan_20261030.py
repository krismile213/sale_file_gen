# -*- coding: utf-8 -*-
"""逐簿扫描 ASIN 列：多少行有 SKU 但 ASIN 空 / #N/A，库存报告能否补上。
只读，不改文件。
"""
import csv, os, glob, openpyxl

SNAP = r'C:\Users\dell\Desktop\work-main\sale_file_gen\trunk\data\数据归属周期\202610\快照日期\20261030'
DSR = r'C:\Users\dell\Desktop\work-main\sale_file_gen\Daily_Sales_Report'

JP24 = ["sku","fnsku","asin","product-name"]

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

# 库存：sku -> {site_file: asin}
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
    m = {}
    for r in rows[1:]:
        if len(r) <= max(i_sku,i_asin): continue
        s = str(r[i_sku]).strip(); a = str(r[i_asin]).strip()
        if s and a: m[s] = a
    stock[name] = m

# 分析表定位
files = sorted(glob.glob(os.path.join(DSR, '1.每日销售数据表', '*.xlsx')))
print('分析表簿数:', len(files))
for f in files:
    if os.path.basename(f).startswith('~$'): continue
    try:
        wb = openpyxl.load_workbook(f, read_only=True, data_only=False)
    except Exception as e:
        print('!!', os.path.basename(f), e); continue
    for ws in wb.worksheets:
        if '分析表' not in ws.title: continue
        # 找 ASIN 列：行 5/6 附近
        asin_col = None
        for r in (5,6,4,7):
            for c in range(1, 60):
                v = ws.cell(r,c).value
                if isinstance(v,str) and v.strip().upper() == 'ASIN':
                    asin_col = c; asin_row = r; break
            if asin_col: break
        if not asin_col:
            continue
        # SKU 列
        sku_col = None
        for r in (5,6,4,7):
            for c in range(1,26):
                v = ws.cell(r,c).value
                if isinstance(v,str) and v.strip().upper() in ('SKU','A'):
                    sku_col = c; break
            if sku_col: break
        sku_col = sku_col or 1
        n_sku = 0; n_formula = 0; n_static = 0
        bad_sku = []
        for r in range(asin_row+1, ws.max_row+1):
            sku = ws.cell(r, sku_col).value
            if not isinstance(sku,str) or not sku.strip(): continue
            n_sku += 1
            av = ws.cell(r, asin_col).value
            if isinstance(av,str) and av.startswith('='): n_formula += 1
            elif av not in (None,''): n_static += 1
            else: bad_sku.append(sku.strip())
        # 这些空 ASIN 的 SKU 在库存端有没有
        recover = [s for s in bad_sku if any(s in m for m in stock.values())]
        print('%-42s | %-22s | SKU行%4d | 公式%4d | 静态值%3d | 空%3d | 库存可补%3d'
              % (os.path.basename(f), ws.title, n_sku, n_formula, n_static, len(bad_sku), len(recover)))
        if bad_sku and len(bad_sku) <= 12:
            print('      空ASIN的SKU:', bad_sku)
    wb.close()
