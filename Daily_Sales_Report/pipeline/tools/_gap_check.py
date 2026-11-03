"""放宽 SKU 正则后，检查各簿「源里有证据、但分析表没有」的 SKU。

证据 = 可售>0 或 在途>0（与 runner 补行闸同口径；销量这条这里不查，由 runner 跑批时兜）。
只读不写。
"""
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(HERE))
sys.path.insert(0, os.path.join(os.path.dirname(HERE), "core"))

import config as C                      # noqa: E402
import analysis as An                   # noqa: E402
import runner as R                      # noqa: E402
from openpyxl import load_workbook      # noqa: E402


def main():
    total = 0
    for b in C.WORKBOOKS:
        if b.get("skip"):
            continue
        fname = b["file"]
        p = os.path.join(C.WORK_DIR, fname)
        if not os.path.exists(p):
            continue
        tmp = None
        if os.path.getsize(p) > 8 * 1024 * 1024:
            tmp = os.path.join(C.DSR, ".workbuddy", "_gap_strip.xlsx")
            R.strip_pivot_records(p, tmp)
            load_p = tmp
        else:
            load_p = p
        wb = load_workbook(load_p, data_only=False)
        for sn in wb.sheetnames:
            if "分析表" not in sn:
                continue
            ws = wb[sn]
            srcs = An.source_sheets_of(ws, wb)
            if not srcs:
                continue
            stock = An.stock_snapshot_merged(wb, srcs)
            have = An.sku_set(ws)
            gaps = []
            for sku, rec in stock.items():
                if sku in have:
                    continue
                av = rec.get("avail") or 0
                ib = rec.get("inb") or 0
                if av > 0 or ib > 0:
                    gaps.append((sku, av, ib, An.tri_state(rec)))
            if gaps:
                total += len(gaps)
                print(f"  {fname} / {sn}: {len(gaps)} 个")
                for g in sorted(gaps, key=lambda x: -(x[1] + x[2]))[:10]:
                    print(f"      {g[0]:<12} 可售={g[1]:>5.0f} 在途={g[2]:>5.0f} → {g[3]}")
        wb.close()
        if tmp and os.path.exists(tmp):
            os.remove(tmp)
    print(f"\n合计缺口 {total} 个")


if __name__ == "__main__":
    main()
