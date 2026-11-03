# -*- coding: utf-8 -*-
"""新品分布矩阵：G59x~G60x 等新品在 16 张分析表 / 库存源 / 货件里的分布。"""
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
PIPE = os.path.dirname(HERE)
sys.path.insert(0, PIPE)
sys.path.insert(0, os.path.join(PIPE, "core"))

import config as C                                    # noqa: E402
import analysis as An                                 # noqa: E402
import runner as R                                    # noqa: E402
import shipments as Sh                                # noqa: E402
from openpyxl import load_workbook                    # noqa: E402

LO, HI = 585, 605
WANT = [f"G{i}" for i in range(LO, HI + 1)]

ships = Sh.load_shipments(C.SHIP_DIR)
# 按 (book, sheet) 分桶，不能只取全局并集 —— 否则「货」会标到不相干的簿上
ship_by = {}
for s in ships:
    bk = s.get("book") or ""
    ship_by.setdefault(bk, {}).setdefault(s.get("sheet") or "", set()).update(s["items"].keys())

rows = {}
for b in C.WORKBOOKS:
    fn = b.get("file", "")
    p = os.path.join(C.WORK_DIR, fn)
    if not os.path.exists(p):
        continue
    stripped = p + ".mx.xlsx"
    n = 0
    try:
        n = R.strip_pivot_records(p, stripped)
        wb = load_workbook(stripped if n else p, data_only=False)
    except Exception as e:
        print(f"!! {fn}: {e}")
        continue
    for sn in wb.sheetnames:
        if "分析表" not in sn:
            continue
        ws = wb[sn]
        have = An.sku_set(ws)
        srcs = An.source_sheets_of(ws, ws.parent)
        stock = An.stock_snapshot_merged(wb, srcs) if srcs else {}
        key = f"{fn.replace('库存销售表.xlsx','')}·{sn.split('-')[0]}"
        for sku in WANT:
            in_t = "表" if sku in have else ""
            in_s = "源" if sku in stock else ""
            _sh = ship_by.get(fn, {})
            _hs = set(_sh.get(sn.strip(), set()))
            if not _hs:
                for _k, _v in _sh.items():
                    if _k and (_k in sn or sn in _k):
                        _hs |= _v
            in_h = "货" if sku in _hs else ""
            if in_t or in_s or in_h:
                rows.setdefault(sku, {})[key] = (in_t + in_s + in_h) or "-"
    wb.close()
    if n and os.path.exists(stripped):
        try:
            os.remove(stripped)
        except OSError:
            pass

cols = sorted({k for v in rows.values() for k in v})
print(f"覆盖 {len(cols)} 个站点表：")
for i, c in enumerate(cols):
    print(f"  {i:>2} {c}")
print()
hdr = "SKU   " + " ".join(f"{i:>3}" for i in range(len(cols)))
print(hdr)
print("-" * len(hdr))
for sku in WANT:
    v = rows.get(sku)
    if not v:
        continue
    line = f"{sku:<6}" + " ".join(f"{v.get(c,'.'):>3}" for c in cols)
    print(line)
print("\n图例：表=分析表已有 ｜ 源=库存源有 ｜ 货=货件有 ｜ . = 都没有")
