# -*- coding: utf-8 -*-
"""ASIN 来源对比：商品资料表 vs 各站点管理库存报告
只读分析，不改任何文件。
"""
import csv, os, glob, sys, io
import openpyxl

SNAP = r'C:\Users\dell\Desktop\work-main\sale_file_gen\trunk\data\数据归属周期\202610\快照日期\20261030'
PROD = r'C:\Users\dell\Desktop\work-main\sale_file_gen\trunk\商品资料.xlsx'

# ---------- 1. 商品资料 ----------
wb = openpyxl.load_workbook(PROD, read_only=True, data_only=True)
ws = wb['商品资料']
prod = {}          # sku -> asin raw
prod_status = {}   # sku -> 状态
for row in ws.iter_rows(min_row=3, max_col=13, values_only=True):
    sku = (row[0] or '').strip() if isinstance(row[0], str) else row[0]
    if not sku:
        continue
    sku = str(sku).strip()
    prod[sku] = row[1]
    prod_status[sku] = row[12]
wb.close()

def is_bad(v):
    if v is None:
        return True
    s = str(v).strip()
    return s == '' or s.upper().startswith('#N/A') or s.upper() in ('NONE', 'NAN', '-')

bad = {k for k, v in prod.items() if is_bad(v)}
print('[商品资料] 总 SKU %d，ASIN 有效 %d，无效(#N/A/空) %d' % (
    len(prod), len(prod) - len(bad), len(bad)))

# ---------- 2. 库存报告 ----------
JP24 = [
    "sku", "fnsku", "asin", "product-name", "condition", "your-price",
    "mfn-listing-exists", "mfn-fulfillable-quantity", "afn-listing-exists",
    "afn-warehouse-quantity", "afn-fulfillable-quantity", "afn-unsellable-quantity",
    "afn-reserved-quantity", "afn-total-quantity", "per-unit-volume",
    "afn-inbound-working-quantity", "afn-inbound-shipped-quantity",
    "afn-inbound-receiving-quantity", "afn-researching-quantity",
    "afn-reserved-future-supply", "afn-future-supply-buyable",
    "afn-fc-transfer-quantity", "afn-onhand-buyable-quantity", "store",
]

def enc_order(name):
    # 日文界面导出的报告：gb18030 会"成功"解成乱码，必须 shift_jis 优先
    if '日本' in name or '澳大利亚' in name:
        return ['shift_jis', 'utf-8-sig', 'gb18030', 'latin-1']
    return ['utf-8-sig', 'gb18030', 'shift_jis', 'latin-1']

def read_csv_any(p, name=''):
    for enc in enc_order(name):
        try:
            with open(p, encoding=enc, newline='') as fh:
                return list(csv.reader(fh))
        except (UnicodeDecodeError, LookupError):
            continue
    raise RuntimeError('cannot decode ' + p)

sites = {}   # 文件名 -> {sku: asin}
for p in sorted(glob.glob(os.path.join(SNAP, '*管理库存报告*.csv'))):
    name = os.path.basename(p)
    rows = read_csv_any(p, name)
    if not rows:
        continue
    hdr = [str(c).strip().lower() for c in rows[0]]
    try:
        i_sku = hdr.index('sku'); i_asin = hdr.index('asin')
    except ValueError:
        # 日文界面 24 列：列头乱码，按位置映射
        if len(hdr) >= 24:
            hdr = JP24
            i_sku, i_asin = 0, 2
        else:
            print('  !! 无 asin 列:', name, hdr[:5])
            continue
    m = {}
    empty = 0
    for r in rows[1:]:
        if len(r) <= max(i_sku, i_asin):
            continue
        sku = str(r[i_sku]).strip()
        asin = str(r[i_asin]).strip()
        if not sku:
            continue
        if not asin or asin.upper().startswith('#N/A'):
            empty += 1
            m.setdefault(sku, None)
        else:
            m[sku] = asin
    sites[name] = m
    print('[库存] %-46s SKU %3d  其中ASIN空 %d' % (name, len(m), empty))

# ---------- 3. 交叉分析 ----------
print('\n===== A. 商品资料 #N/A 的 SKU 能否从库存补 =====')
stock_all = {}   # sku -> set(asin)
for name, m in sites.items():
    for sku, asin in m.items():
        if asin:
            stock_all.setdefault(sku, set()).add(asin)

recover = {}
for sku in bad:
    if sku in stock_all:
        recover[sku] = stock_all[sku]

print('无效 SKU 共 %d 个；库存报告里有 ASIN 的 %d 个，仍补不上 %d 个' % (
    len(bad), len(recover), len(bad) - len(recover)))

demo = sorted(recover.items())[:15]
for sku, asins in demo:
    print('   %-9s -> %s' % (sku, sorted(asins)))

unrec = sorted(bad - set(recover))
print('   补不上的前 30 个:', unrec[:30])

# ---------- 4. 冲突：商品资料 ASIN vs 库存 ASIN ----------
print('\n===== B. 商品资料 ASIN 与库存 ASIN 是否一致 =====')
ok = 0; conflict = {}; notin = 0
for sku, asin in prod.items():
    if sku in bad:
        continue
    s = str(asin).strip()
    if sku not in stock_all:
        notin += 1
        continue
    aset = stock_all[sku]
    if s in aset:
        ok += 1
        if len(aset) > 1:
            conflict.setdefault(sku, (s, sorted(aset)))
    else:
        conflict[sku] = (s, sorted(aset))
print('一致 %d，不一致 %d，商品资料有但库存报告无此 SKU %d' % (ok, len(conflict), notin))
for sku, (p_asin, aset) in sorted(conflict.items())[:25]:
    print('   %-9s 资料=%-12s 库存=%s' % (sku, p_asin, aset))
print('   冲突合计:', len(conflict))

# ---------- 5. 一 SKU 多 ASIN（跨站点正常）----------
multi = {k: v for k, v in stock_all.items() if len(v) > 1}
print('\n===== C. 同一 SKU 在库存报告里出现多个 ASIN =====')
print('SKU 总数 %d，其中多 ASIN 的 %d 个' % (len(stock_all), len(multi)))
for k, v in sorted(multi.items())[:10]:
    print('   %-9s %s' % (k, sorted(v)))
