"""G599~G602 全链路追踪（2026-10-24）。

查四层：
  ① 快照原始报告（库存 / 订单 / 货件）
  ② 各簿分析表的库存源 sheet（stock_snapshot 识别结果）
  ③ 各簿分析表 A 列是否已有该 SKU、D 列状态
  ④ 型号首库存登记.xlsx（首现 / 首可售 / 首在途 / 当前 / 状态 / 起售日期）
"""
import csv
import io
import os
import sys
from collections import defaultdict

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(HERE))
sys.path.insert(0, os.path.join(os.path.dirname(HERE), "core"))

import config as C                      # noqa: E402
import analysis as An                   # noqa: E402
import runner as R                      # noqa: E402
from openpyxl import load_workbook      # noqa: E402

WANT = ["G599", "G600", "G601", "G602"]
SNAP = os.path.join(C.SNAP_BASE, C.CURRENT_PERIOD, "快照日期", "20261023")


def dec(p):
    raw = open(p, "rb").read()
    for enc in ("utf-8-sig", "gbk", "utf-16"):
        try:
            return raw.decode(enc)
        except Exception:
            continue
    return raw.decode("utf-8", "replace")


def part1_raw():
    print("=" * 70)
    print("① 快照原始报告（20261023）")
    print("=" * 70)
    hits = defaultdict(list)
    for fn in sorted(os.listdir(SNAP)):
        p = os.path.join(SNAP, fn)
        if not os.path.isfile(p) or not fn.lower().endswith((".csv", ".txt", ".tsv")):
            continue
        try:
            txt = dec(p)
        except Exception:
            continue
        for w in WANT:
            if w in txt:
                # 统计出现次数
                hits[w].append((fn, txt.count(w)))
    for w in WANT:
        if hits[w]:
            print(f"  {w}:")
            for fn, n in hits[w]:
                print(f"      {fn}  ×{n}")
        else:
            print(f"  {w}: ❌ 任何原始报告里都没有")

    # 货件
    print("\n  ── 货件 ──")
    import shipments as Sh
    ships = Sh.load_shipments(C.SHIP_DIR)
    found = False
    for s in ships:
        hit = [k for k in s["items"] if k in WANT]
        if hit:
            found = True
            print(f"    {s['id']}  wh={s.get('wh')}  book={s.get('book')}  "
                  f"sheet={s.get('sheet')}  件数={s.get('total')}  SKU={hit}")
    if not found:
        print("    85 票货件里没有这些 SKU")


def part2_sources(filt=()):
    print("\n" + "=" * 70)
    print("② 各簿库存源识别结果（stock_snapshot_merged）")
    print("=" * 70)
    for b in C.WORKBOOKS:
        if b.get("skip"):
            continue
        fname = b["file"]
        if filt and not any(k in fname for k in filt):
            continue
        p = os.path.join(C.WORK_DIR, fname)
        if not os.path.exists(p):
            continue
        tmp = None
        if os.path.getsize(p) > 8 * 1024 * 1024:
            tmp = os.path.join(C.DSR, ".workbuddy", "_g_strip.xlsx")
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
            hit = {k: stock[k] for k in WANT if k in stock}
            if hit:
                print(f"  {fname} / {sn}  （源 {'+'.join(srcs)}）")
                for k in WANT:
                    if k in hit:
                        r = hit[k]
                        print(f"      {k}  可售={r['avail']:.0f}  在途={r['inb']:.0f}  "
                              f"listing={r['listing']}  → {An.tri_state(r)}")
        wb.close()
        if tmp and os.path.exists(tmp):
            os.remove(tmp)


def part3_tables(filt=()):
    print("\n" + "=" * 70)
    print("③ 各簿分析表 A 列（已在表内的位置与状态）")
    print("=" * 70)
    any_hit = False
    for b in C.WORKBOOKS:
        if b.get("skip"):
            continue
        fname = b["file"]
        if filt and not any(k in fname for k in filt):
            continue
        p = os.path.join(C.WORK_DIR, fname)
        if not os.path.exists(p):
            continue
        tmp = None
        if os.path.getsize(p) > 8 * 1024 * 1024:
            tmp = os.path.join(C.DSR, ".workbuddy", "_g2_strip.xlsx")
            R.strip_pivot_records(p, tmp)
            load_p = tmp
        else:
            load_p = p
        wb = load_workbook(load_p, data_only=True)
        for sn in wb.sheetnames:
            if "分析表" not in sn:
                continue
            ws = wb[sn]
            last = An._last_data_row(ws)
            rows = {}
            for r in range(An.DATA_FIRST_ROW, last + 1):
                v = ws.cell(r, 1).value
                if v:
                    rows[str(v).strip()] = r
            for k in WANT:
                if k in rows:
                    any_hit = True
                    r = rows[k]
                    print(f"  {k:<6} {fname[:22]:<24}{sn[:14]:<16} 行{r}  "
                          f"D={ws.cell(r, 4).value}  F={ws.cell(r, 6).value}  "
                          f"G={ws.cell(r, 7).value}")
        wb.close()
        if tmp and os.path.exists(tmp):
            os.remove(tmp)
    if not any_hit:
        print("  ❌ 13 本分析表里一个都没有")


def part4_register():
    print("\n" + "=" * 70)
    print("④ 型号首库存登记.xlsx")
    print("=" * 70)
    p = os.path.join(C.DSR, "断货分析", "型号首库存登记.xlsx")
    wb = load_workbook(p, data_only=True, read_only=True)
    ws = wb["首库存登记"]
    hdr = next(ws.iter_rows(min_row=1, max_row=1, values_only=True))
    print("  列:", list(hdr))
    found = {}
    for r in ws.iter_rows(min_row=2, values_only=True):
        if r and r[0] and str(r[0]).strip() in WANT:
            found[str(r[0]).strip()] = r
    for k in WANT:
        if k in found:
            print(f"  {k}: {found[k]}")
        else:
            print(f"  {k}: ❌ 未登记")
    print()
    ws2 = wb["站点明细"]
    h2 = next(ws2.iter_rows(min_row=1, max_row=1, values_only=True))
    print("  站点明细列:", list(h2))
    for r in ws2.iter_rows(min_row=2, values_only=True):
        if r and r[0] and str(r[0]).strip() in WANT:
            print(f"    {r}")
    wb.close()


if __name__ == "__main__":
    _filt = tuple(sys.argv[1:])
    part1_raw()
    part2_sources(_filt)
    part3_tables(_filt)
    part4_register()
