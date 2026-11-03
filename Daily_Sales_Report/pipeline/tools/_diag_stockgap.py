# -*- coding: utf-8 -*-
"""诊断：库存源里有、分析表里没有的 SKU（按三态 + 是否有销量分组）。

对每张分析表输出：
  · 源 SKU 数 / 表内 SKU 数
  · 缺口按三态分组：在售 / 在途 / 未上架
  · 每个缺口 SKU 标注：本月销量、首现日期、是否被月份闸挡住
"""
import os
import sys
from collections import defaultdict

HERE = os.path.dirname(os.path.abspath(__file__))
PIPE = os.path.dirname(HERE)
DSR = os.path.dirname(PIPE)
sys.path.insert(0, PIPE)
sys.path.insert(0, os.path.join(PIPE, "core"))

import config as C                                    # noqa: E402
import analysis as An                                 # noqa: E402
import runner as R                                    # noqa: E402
from openpyxl import load_workbook                    # noqa: E402

SNAP = "20261023"
MONTH_START = SNAP[:6] + "01"
first_seen = R._first_seen_map() or {}

ONLY = sys.argv[1:]
total_gap = defaultdict(int)

for b in C.WORKBOOKS:
    fn = b.get("file", "")
    if ONLY and not any(x in fn for x in ONLY):
        continue
    p = os.path.join(C.WORK_DIR, fn)
    if not os.path.exists(p):
        continue
    stripped = p + ".dg.xlsx"
    n = 0
    try:
        n = R.strip_pivot_records(p, stripped)
        wb = load_workbook(stripped if n else p, data_only=False)
    except Exception as e:
        print(f"!! {fn}: {e}")
        continue

    # 本月销量（按站点）
    _cfg = next((x for x in C.WORKBOOKS if x.get("file") == fn), None)
    _specs = (_cfg or {}).get("sales")
    _specs = _specs if isinstance(_specs, list) else ([_specs] if _specs else [])
    _dsn = next((s.get("sheet") for s in _specs
                 if s.get("sheet") in wb.sheetnames and s.get("expect_channel")), None)
    _ym = f"{SNAP[:4]}-{SNAP[4:6]}"

    print(f"\n{'='*78}\n{fn}\n{'='*78}")
    for sn in wb.sheetnames:
        if "分析表" not in sn:
            continue
        ws = wb[sn]
        srcs = An.source_sheets_of(ws, wb)
        if not srcs:
            print(f"  {sn}: 无库存源")
            continue
        stock = An.stock_snapshot_merged(wb, srcs)
        have = An.sku_set(ws)
        sold = {}
        if _dsn:
            ch, reg = R._channel_of_sheet(sn)
            if ch or reg:
                sold = An.month_sales(wb[_dsn], _ym, channel=ch, region=reg)

        gap = {k: v for k, v in stock.items() if k not in have}
        by_state = defaultdict(list)
        for k, v in gap.items():
            by_state[An.tri_state(v)].append(k)

        print(f"\n  {sn}  源 {len(stock)} ｜ 表内 {len(have)} ｜ 缺口 {len(gap)}")
        for st in ("在售", "在途", "未上架"):
            ks = sorted(by_state.get(st, []))
            if not ks:
                continue
            print(f"    [{st}] {len(ks)} 个")
            for k in ks[:14]:
                rec = stock[k]
                fs = first_seen.get(k, "")
                gate = ""
                if st != "未上架":
                    if not fs:
                        gate = "⚠未登记"
                    elif fs < MONTH_START:
                        gate = f"被月份闸挡(首现{fs})"
                q = sold.get(k, 0)
                print(f"      {k:<12} avail={rec['avail']:>6.0f} inb={rec['inb']:>6.0f} "
                      f"listing={str(rec['listing']):<5} 月销={q:>4.0f} {gate}")
                total_gap[st] += 1
            if len(ks) > 14:
                print(f"      … 还有 {len(ks)-14} 个")
    wb.close()
    if n and os.path.exists(stripped):
        try:
            os.remove(stripped)
        except OSError:
            pass

print(f"\n{'='*78}\n合计缺口：{dict(total_gap)}")
