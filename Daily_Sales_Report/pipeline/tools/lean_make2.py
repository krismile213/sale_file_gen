# -*- coding: utf-8 -*-
"""日销表精简版生成器 v2（lean_make.py）

策略
----
* **值来源**：`.workbuddy/recalc/<同名>.xlsx` —— LibreOffice 重算过的副本，
  公式格带正确缓存值 ⇒ 无需在 Python 里重写 SUMIFS/SUMPRODUCT。
* **删**：透视类 sheet（用户已确认）；明细表里用不到的列（用户已确认）。
* **留**：库存记录 / 分析表 / 订单数据 / 库存明细 / 产品类型 / Sheet1 /
  10 张国家镜像表 —— 全保留，只把公式换成值。
* **结果**：零公式、零透视缓存、零空壳 → 体积大降，打开即见数。

用法
----
    python lean_make.py                  # 全量
    python lean_make.py --only 日本        # 指定
    python lean_make.py --no-trim-cols    # 只值化、不删列
"""
# --- pipeline 路径引导（2026-10-14 目录收敛自动注入）---
import os as _os, sys as _sys
_bp = _os.path.dirname(_os.path.abspath(__file__))
for _i in range(6):
    if _os.path.isdir(_os.path.join(_bp, "1.每日销售数据表")):
        break
    _bp = _os.path.dirname(_bp)
DSR = _bp
_bcore = _os.path.join(DSR, "pipeline", "core")
if _bcore not in _sys.path:
    _sys.path.insert(0, _bcore)
# --- 引导结束 ---
import os, sys, io, glob, argparse, warnings
warnings.filterwarnings('ignore')
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', line_buffering=True)

import openpyxl
from openpyxl.utils import get_column_letter as gl

HERE = os.path.dirname(os.path.abspath(__file__))
# [normalize] DSR 由上方引导块赋值
CUR = os.path.join(DSR, '1.每日销售数据表')
RECALC = os.path.join(HERE, 'recalc')
OUT = os.path.join(DSR, '日销表_精简版')

DROP_SHEETS = {
    '数据透视表', '数据透视图', '数据透视图阿联酋', '沙特透视图',
    '新品销售情况数据透视图', '墨西哥透视图', '罗马尼亚透视图',
}

SALES_KEEP = [
    'amazon-order-id', 'merchant-order-id', 'purchase-date', 'order-status',
    'sales-channel', 'sku', 'asin', 'quantity', 'item-price',
    'ship-country', 'order-item-id',
    'purchase date', '有效性', '地区', '型号',
]

STOCK_KEEP = [
    'sku', 'fnsku', 'asin', 'product-name', 'condition', 'your-price',
    'mfn-listing-exists', 'mfn-fulfillable-quantity', 'afn-listing-exists',
    'afn-warehouse-quantity', 'afn-fulfillable-quantity', 'afn-unsellable-quantity',
    'afn-reserved-quantity', 'afn-total-quantity', 'per-unit-volume',
    'afn-inbound-working-quantity', 'afn-inbound-shipped-quantity',
    'afn-inbound-receiving-quantity', 'afn-researching-quantity',
    'afn-reserved-future-supply', 'afn-future-supply-buyable',
    'afn-fulfillable-quantity-local', 'afn-fulfillable-quantity-remote',
    'afn-fc-transfer-quantity', 'afn-onhand-buyable-quantity', 'store',
    'FNSKU', 'ASIN', 'MSKU', 'Disposition', 'Ending Warehouse Balance',
    # 结余式的**中文**表头（日文界面导出的结余表用中文列名，漏了会把表删成 2 列）
    '卖家 SKU', '库存属性', '期末仓库余额',
    '在途库存', '现货（可售商品 + 运营中心转运商品）',
]

# 整表保留（只值化）
KEEP_ALL = ('库存记录', '每日库存分析表', '订单数据', '库存明细', '产品类型',
            '多国库存', 'Sheet1', '有问题SKU号',
            '意大利', '法国', '西班牙', '德国', '土耳其',
            '荷兰', '瑞典', '波兰', '比利时', '爱尔兰')


def norm(v):
    return str(v).strip() if v is not None else ''


def is_formula(v):
    if v is None:
        return False
    if hasattr(v, 'text'):
        return True
    return isinstance(v, str) and v.startswith('=')


def pick_cols(ws, wanted):
    """前 3 行里找最匹配 wanted 的表头行，返回命中的列号（升序）。"""
    best_cols, best_hit = [], -1
    for hr in (2, 1, 3):
        hdr = {}
        for c in range(1, (ws.max_column or 0) + 1):
            h = norm(ws.cell(row=hr, column=c).value)
            if h and h not in hdr:
                hdr[h] = c
        cols = sorted({hdr[w] for w in wanted if w in hdr})
        if len(cols) > best_hit:
            best_cols, best_hit = cols, len(cols)
    return best_cols


def real_extent(ws_f, ws_v):
    """真实有值范围（声明常虚高，如 3000x300）。"""
    maxr = maxc = 0
    for r, row in enumerate(ws_f.iter_rows(), 1):
        for c in row:
            v = c.value
            vv = ws_v.cell(row=r, column=c.column).value if is_formula(v) else v
            if vv is not None and (not isinstance(vv, str) or vv.strip() != ''):
                if r > maxr:
                    maxr = r
                if c.column > maxc:
                    maxc = c.column
    return maxr, maxc


def build(src_path, recalc_path, out_path, trim_cols=True, verbose=False):
    wbf = openpyxl.load_workbook(src_path, data_only=False)
    wbv = openpyxl.load_workbook(recalc_path, data_only=True)
    out = openpyxl.Workbook()
    out.remove(out.active)
    report = []

    for sn in wbf.sheetnames:
        if sn in DROP_SHEETS:
            report.append((sn, 'DROP', 0, 0))
            continue
        wsf = wbf[sn]
        wsv = wbv[sn] if sn in wbv.sheetnames else wsf

        cols = None
        if trim_cols:
            if '每日销售数据' in sn:
                cols = pick_cols(wsf, SALES_KEEP)
            elif not any(k in sn for k in KEEP_ALL):
                cols = pick_cols(wsf, STOCK_KEEP)
                # 安全闸：命中列数太少（表头没认出来，例如结余式中/英表头混杂），
                # 宁可整表保留也不要删成残表 —— 删列不可逆。
                _, _maxc = real_extent(wsf, wsv)
                if len(cols) < max(2, _maxc * 0.4):
                    print(f'      [warn] {sn!r} 仅命中 {len(cols)}/{_maxc} 列，改整表保留')
                    cols = None
        if not cols:
            cols = list(range(1, (wsf.max_column or 0) + 1))

        maxr, maxc = real_extent(wsf, wsv)
        maxr = max(maxr, 1)
        cols = [c for c in cols if c <= max(maxc, 1)] or [1]

        ws = out.create_sheet(title=sn)
        for r in range(1, maxr + 1):
            row = []
            for c in cols:
                fv = wsf.cell(row=r, column=c).value
                row.append(wsv.cell(row=r, column=c).value if is_formula(fv) else fv)
            ws.append(row)

        try:
            for i, c in enumerate(cols, 1):
                L = gl(c)
                if L in wsf.column_dimensions and wsf.column_dimensions[L].width:
                    ws.column_dimensions[gl(i)].width = wsf.column_dimensions[L].width
        except Exception:
            pass
        try:
            if wsf.freeze_panes:
                ws.freeze_panes = wsf.freeze_panes
        except Exception:
            pass
        report.append((sn, 'KEEP', maxr, len(cols)))

    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    out.save(out_path)
    wbf.close(); wbv.close()
    if verbose:
        for sn, act, nr, nc in report:
            print(f'      {act:4s} {sn!r:32s} rows={nr:6d} cols={nc}')
    return report


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--only', default='')
    ap.add_argument('--no-trim-cols', action='store_true')
    args = ap.parse_args()

    os.makedirs(OUT, exist_ok=True)
    files = sorted(os.path.basename(p) for p in glob.glob(os.path.join(CUR, '*.xlsx')))
    if args.only:
        keys = args.only.split(',')
        files = [f for f in files if any(k in f for k in keys)]

    print(f'==> 生成精简版 → {OUT}')
    tot_s = tot_d = 0
    for f in files:
        src, rc, dst = (os.path.join(CUR, f), os.path.join(RECALC, f), os.path.join(OUT, f))
        if not os.path.exists(rc) or os.path.getsize(rc) == 0:
            print(f'  [跳过] {f}（无重算副本）')
            continue
        s0 = os.path.getsize(src)
        try:
            build(src, rc, dst, trim_cols=not args.no_trim_cols, verbose=False)
        except Exception as e:
            print(f'  [失败] {f}: {e!r}')
            continue
        s1 = os.path.getsize(dst)
        tot_s += s0; tot_d += s1
        print(f'  {f:44s} {s0/1024/1024:6.2f}MB → {s1/1024/1024:6.2f}MB ({(1-s1/s0)*100:4.1f}%↓)')
    if tot_s:
        print(f'  {"合计":44s} {tot_s/1024/1024:6.2f}MB → {tot_d/1024/1024:6.2f}MB '
              f'({(1-tot_d/tot_s)*100:4.1f}%↓)')


if __name__ == '__main__':
    main()
