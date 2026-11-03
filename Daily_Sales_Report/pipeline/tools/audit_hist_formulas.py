# -*- coding: utf-8 -*-
"""统计 2025-2026 归档日销簿的分析表关键列是「公式」还是「静态值」。

用途：换月/改口径前摸清历史簿形态（尤其「上月日均」列有没有公式），
      避免把「程序刷的静态值」误当成「表内一直有的活公式」，或反过来。

用法：
    python audit_hist_formulas.py [--years 2025,2026] [--jobs 8] [--out 结果.csv]
输出：逐「簿×分析表×字段」明细 CSV + 终端汇总（公式/静态/缺失计数）。
"""
import argparse
import csv
import os
import re
import sys
from multiprocessing import Pool

import openpyxl

DSR = None  # 由 worker 初始化（spawn 安全）


def _find_dsr():
    global DSR
    if DSR is not None:
        return DSR
    p = os.path.abspath(__file__)
    while True:
        if os.path.isdir(os.path.join(p, "1.每日销售数据表")):
            DSR = p
            return DSR
        np = os.path.dirname(p)
        if np == p:
            DSR = os.getcwd()
            return DSR
        p = np


# 字段 → 行5 标签匹配规则（contains）
FIELDS = [
    ("prev_avg", lambda s: s == "上月日均"),
    ("speed", lambda s: s == "实时速度"),
    ("days_sale", lambda s: "可售天数" in s),
    ("month_total", lambda s: s.endswith("月总销量")),
]


def _is_formula(v):
    if v is None:
        return False
    if isinstance(v, str) and v.startswith("="):
        return True
    if hasattr(v, "text"):  # ArrayFormula
        return True
    return False


def _col_letter(ci):
    from openpyxl.utils import get_column_letter
    return get_column_letter(ci)


def scan_file(path):
    """返回 [(ym, book, sheet, field, col, kind, sample)]；kind ∈ FORMULA/STATIC/EMPTY/NOLABEL"""
    root = _find_dsr()
    rel = os.path.relpath(path, root)
    m = re.match(r"1\.每日销售数据表[\\/](\d{4})年每日数据[\\/](\d{4})年(\d{1,2})月[\\/](.+)", rel)
    if not m:
        ym, book = "?", os.path.basename(path)
    else:
        ym, book = f"{m.group(2)}-{int(m.group(3)):02d}", m.group(4)
    out = []
    try:
        wb = openpyxl.load_workbook(path, read_only=True)
    except Exception as e:
        out.append((ym, book, "-", "-", "-", f"OPENFAIL({type(e).__name__})", ""))
        return out
    try:
        for sn in wb.sheetnames:
            if "分析" not in sn:
                continue
            ws = wb[sn]
            # 行5 标签 → 列号（取第一个匹配的列）
            hits = {}
            for row in ws.iter_rows(min_row=5, max_row=5):
                for c in row:
                    if c.value is None:
                        continue
                    s = str(c.value).strip()
                    for fname, pred in FIELDS:
                        if fname not in hits and pred(s):
                            hits[fname] = c.column
                break  # 行5 只有一行
            # 数据区前 3 行（r7-r9）判形态
            need = sorted(set(hits.values()))
            vals = {}
            if need:
                lo, hi = min(need), max(need)
                for ri, row in enumerate(ws.iter_rows(min_row=7, max_row=9,
                                                      min_col=lo, max_col=hi,
                                                      values_only=False), 7):
                    for c in row:
                        if c.column in vals:
                            continue
                        vals[c.column] = c.value
            for fname, _ in FIELDS:
                col = hits.get(fname)
                if col is None:
                    out.append((ym, book, sn, fname, "-", "NOLABEL", ""))
                    continue
                v = vals.get(col)
                if _is_formula(v):
                    kind = "FORMULA"
                elif v is None:
                    kind = "EMPTY"
                else:
                    kind = "STATIC"
                sample = (v.text if hasattr(v, "text") else str(v))[:80]
                out.append((ym, book, sn, fname, _col_letter(col), kind, sample))
    finally:
        wb.close()
    return out


def main(argv=None):
    ap = argparse.ArgumentParser()
    ap.add_argument("--years", default="2025,2026")
    ap.add_argument("--jobs", type=int, default=8)
    ap.add_argument("--out", default="")
    args = ap.parse_args(argv)

    root = _find_dsr()
    files = []
    for y in args.years.split(","):
        base = os.path.join(root, "1.每日销售数据表", f"{y}年每日数据")
        for dirpath, _dirs, names in os.walk(base):
            for n in names:
                if n.endswith(".xlsx") and not n.startswith("~$"):
                    files.append(os.path.join(dirpath, n))
    files.sort()
    print(f"扫描 {len(files)} 本簿（{', '.join(args.years.split(','))} 年归档）…", flush=True)

    rows = []
    if args.jobs > 1:
        with Pool(args.jobs) as pool:
            for i, res in enumerate(pool.imap_unordered(scan_file, files), 1):
                rows.extend(res)
                if i % 20 == 0:
                    print(f"  {i}/{len(files)}", flush=True)
    else:
        for i, f in enumerate(files, 1):
            rows.extend(scan_file(f))

    # 汇总
    from collections import Counter
    stat = Counter((r[3], r[5]) for r in rows)
    books = len({(r[0], r[1]) for r in rows})
    print(f"\n===== 汇总（{books} 本簿 / {len(rows)} 条记录）=====")
    for fname, _ in FIELDS:
        sub = {k[1]: v for k, v in stat.items() if k[0] == fname}
        total = sum(sub.values())
        print(f"{fname:12s} 总{total:5d}  " +
              "  ".join(f"{k}={v}" for k, v in sorted(sub.items())))
    # 公式样本
    fs = [r for r in rows if r[5] == "FORMULA"]
    if fs:
        print("\n--- 公式样本（前 10 条）---")
        for r in fs[:10]:
            print(" | ".join(str(x) for x in r))

    if args.out:
        with open(args.out, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            w.writerow(["ym", "book", "sheet", "field", "col", "kind", "sample"])
            w.writerows(rows)
        print(f"\n明细已写 {args.out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
