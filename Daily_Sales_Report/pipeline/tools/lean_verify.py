# -*- coding: utf-8 -*-
"""验收：精简版 vs 重算标准答案 —— **按表头名**逐格对撞

精简版删掉了部分列，列位会左移，所以不能按位置比。规则：
  * 对"列子集"sheet（每日销售数据 / 库存类）：用表头名把精简版的每一列映射回
    recalc 版的同名列，再逐行比值。
  * 对"整表保留"sheet（库存记录 / 分析表 / 订单数据 / 库存明细…）：直接按位置比。
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
import os, sys, io, argparse, warnings
warnings.filterwarnings('ignore')
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', line_buffering=True)
import openpyxl
from openpyxl.utils import get_column_letter as gl

HERE = os.path.dirname(os.path.abspath(__file__))
# [normalize] DSR 由上方引导块赋值
RECALC = os.path.join(HERE, 'recalc')
LEAN = os.path.join(DSR, '日销表_精简版')

FULL_SHEETS = ('库存记录', '每日库存分析表', '订单数据', '库存明细', '产品类型',
               '多国库存', 'Sheet1', '有问题SKU号')


def norm(v):
    if v is None:
        return None
    if isinstance(v, float) and v == int(v):
        return int(v)
    if isinstance(v, str):
        s = v.strip()
        return s if s != '' else None
    return v


def read_rows(ws):
    """顺序读全表为 list[list]，避免 read_only 模式随机访问触发重复解析。"""
    return [list(r) for r in ws.iter_rows(values_only=True)]


def find_header_row(rows, limit=4):
    """选"非空表头名最多"的前几行作为表头行。"""
    best, best_n = 0, -1
    for i in range(min(limit, len(rows))):
        n = sum(1 for v in rows[i] if norm(v) is not None)
        if n > best_n:
            best, best_n = i, n
    return best  # 0-based


def header_map(rows, hr):
    m = {}
    row = rows[hr] if hr < len(rows) else []
    for c, v in enumerate(row, start=1):
        h = norm(v)
        if h is not None and h not in m:
            m[h] = c
    return m


def compare(fn, limit=25):
    lp, rp = os.path.join(LEAN, fn), os.path.join(RECALC, fn)
    wbl = openpyxl.load_workbook(lp, read_only=True, data_only=True)
    wbr = openpyxl.load_workbook(rp, read_only=True, data_only=True)
    total = diff = 0
    samples = []
    per_sheet = []

    for sn in wbl.sheetnames:
        if sn not in wbr.sheetnames:
            samples.append((sn, '-', 'SHEET 缺失', ''))
            continue
        wsl, wsr = wbl[sn], wbr[sn]
        rl, rr = read_rows(wsl), read_rows(wsr)

        if any(k in sn for k in FULL_SHEETS):
            # 按位置比
            maxr = min(len(rl), len(rr))
            st = sd = 0
            for r in range(maxr):
                arow, brow = rl[r], rr[r]
                maxc = min(len(arow), len(brow))
                for c in range(maxc):
                    a = norm(arow[c]); b = norm(brow[c])
                    if a is None and b is None:
                        continue
                    st += 1
                    if a != b:
                        sd += 1
                        if len(samples) < limit:
                            samples.append((sn, f'{gl(c+1)}{r+1}', a, b))
            total += st; diff += sd
            per_sheet.append((sn, st, sd, 'position', len(rl), len(rr)))
        else:
            # 按表头名比
            hr_l = find_header_row(rl)
            hr_r = find_header_row(rr)
            hl = header_map(rl, hr_l)
            hrr = header_map(rr, hr_r)
            st = sd = 0
            for name, cl in hl.items():
                cr = hrr.get(name)
                if cr is None:
                    samples.append((sn, f'{gl(cl)}{hr_l+1}', f'列 {name!r} 在 recalc 中找不到', ''))
                    continue
                for r in range(hr_l + 1, max(len(rl), len(rr))):
                    a = norm(rl[r][cl-1]) if r < len(rl) and cl-1 < len(rl[r]) else None
                    b = norm(rr[r][cr-1]) if r < len(rr) and cr-1 < len(rr[r]) else None
                    if a is None and b is None:
                        continue
                    st += 1
                    if a != b:
                        sd += 1
                        if len(samples) < limit:
                            samples.append((sn, f'{name}@{r+1}', a, b))
            total += st; diff += sd
            per_sheet.append((sn, st, sd, 'by-header', len(rl), len(rr)))

    wbl.close(); wbr.close()
    print(f'  {fn}   ({os.path.getsize(lp)/1024:.0f}KB)')
    for sn, st, sd, mode, nrl, nrr in per_sheet:
        flag = 'OK ' if sd == 0 else 'DIFF'
        extra = f'  行数 {nrl}/{nrr}' if nrl != nrr else ''
        print(f'      {flag} {sn!r:30s} {st:7d} 格 差异 {sd:5d}  [{mode}]' + extra)
    for sn, coord, a, b in samples:
        print(f'        · {sn!r} {coord}: 精简={str(a)[:40]!r} recalc={str(b)[:40]!r}')
    return total, diff


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--only', default='')
    args = ap.parse_args()
    import glob
    files = sorted(os.path.basename(p) for p in
                   glob.glob(os.path.join(LEAN, '*.xlsx')))
    if args.only:
        keys = args.only.split(',')
        files = [f for f in files if any(k in f for k in keys)]
    gt = gd = 0
    for f in files:
        t, d = compare(f)
        gt += t; gd += d
        print()
    print(f'== 合计 {gt} 格，差异 {gd} ==')


if __name__ == '__main__':
    main()
