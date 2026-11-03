# -*- coding: utf-8 -*-
"""模拟零证据补行闸：在欧盟簿上算出「会新补哪些行」，不落盘。

用法：
  python _sim_zero_evidence.py                 # 只看欧盟簿
  python _sim_zero_evidence.py --book 英国     # 按名字包含匹配
"""
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
CORE = os.path.join(HERE, "..", "core")
sys.path.insert(0, os.path.abspath(CORE))

import config as C
import analysis as An
import runner as R
from openpyxl import load_workbook

BOOKS = os.path.join(C.PROJECT, "Daily_Sales_Report", "1.每日销售数据表")


def main():
    want = ""
    if "--book" in sys.argv:
        want = sys.argv[sys.argv.index("--book") + 1]

    files = sorted(f for f in os.listdir(BOOKS) if f.endswith(".xlsx"))
    if want:
        files = [f for f in files if want in f]

    for fname in files:
        p = os.path.join(BOOKS, fname)
        print("=" * 72)
        print(fname)
        try:
            wb = load_workbook(p, data_only=True)
        except Exception as e:
            print("  打开失败:", e)
            continue

        stock_map = R._stock_map_per_sheet(wb)
        sold = R._sales_cache_per_sheet(wb, fname, C.SNAPSHOT_DATE)

        for sn in wb.sheetnames:
            if "分析表" not in sn:
                continue
            ws = wb[sn]
            srcs = An.source_sheets_of(ws, wb)
            if not srcs:
                continue
            stock = stock_map.get(sn.strip()) or {}
            solds = (sold or {}).get(sn.strip()) or {}
            have = An.sku_set(ws)

            cands = [(s, n, a) for s, n, a in An.listing_skus_merged(wb, srcs)]
            seen = {s for s, _, _ in cands}
            for s, rec in stock.items():
                if s not in seen:
                    cands.append((s, rec.get("name", ""), None))

            new_rows, blocked = [], 0
            for sku, nm, alive in cands:
                if sku in have:
                    continue
                rec = stock.get(sku) or {}
                av = rec.get("avail") or 0
                ib = rec.get("inb") or 0
                q = solds.get(sku, 0)
                st = An.tri_state_of_sku(stock, sku)
                if av > 0 or ib > 0 or q > 0:
                    continue                       # 老闸本来就放行
                ok = (C.ZERO_EVIDENCE_SKU_RE
                      and re.match(C.ZERO_EVIDENCE_SKU_RE, sku)
                      and st != An.STATE_NOT_LISTED
                      and bool(rec))
                if ok:
                    new_rows.append((sku, nm, st, rec.get("listing")))
                else:
                    blocked += 1

            if new_rows:
                print("  [%s] 新增 %d 行，另有 %d 个被闸挡住"
                      % (sn.strip(), len(new_rows), blocked))
                for sku, nm, st, ls in new_rows:
                    print("      + %-8s %-28s 状态=%s listing=%s" % (sku, (nm or "")[:26], st, ls))
            elif blocked:
                print("  [%s] 新增 0 行（%d 个被闸挡住）" % (sn.strip(), blocked))
        wb.close()


if __name__ == "__main__":
    main()
