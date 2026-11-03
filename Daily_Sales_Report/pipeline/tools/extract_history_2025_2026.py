# -*- coding: utf-8 -*-
"""
提取 2025年1月 ~ 2026年8月 人工做表历史存档的重要信息（只读，无副作用）。

对每本工作簿记录：
  1) 全部 sheet 名单与尺寸
  2) 分析表：表头行位、关键列位（预留/可售/在途/调拨在途合计/在途差额/实时速度/
     上月日均/在售可售天数/在售在途可售天数/月总销量类）、合计行值、SKU 数据行数、
     日窗覆盖范围（表头行里的日期列 min~max）
  3) 库存快照类 sheet 的 A1（快照锚点）
  4) 订单/销售类 sheet 行数

输出：JSONL（每本一行）-> history_extract_2025_2026.jsonl（与本脚本同目录）
进度与异常打印到 stdout。
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
import json
import os
import re
import sys
from datetime import datetime, date

sys.stdout.reconfigure(encoding="utf-8")
import openpyxl
from openpyxl.utils import get_column_letter as GL

HERE = os.path.dirname(os.path.abspath(__file__))
# [normalize] DSR 由上方引导块赋值
BASE = os.path.join(DSR, "1.每日销售数据表")
YEARS = ["2025年每日数据", "2026年每日数据"]
MONTHS = [f"{y}年{m}月" for y, m in [(2025, m) for m in range(1, 13)]
          + [(2026, m) for m in range(1, 9)]]  # 2025-01 .. 2026-08
OUT = os.path.join(HERE, "history_extract_2025_2026.jsonl")

KEY_HEADERS = ["预留库存", "可售库存", "在途", "调拨在途合计", "在途差额",
               "实时速度", "上月日均", "在售可售天数", "在售在途可售天数"]


def is_date(v):
    return isinstance(v, (datetime, date))


def num(v):
    return float(v) if isinstance(v, (int, float)) else None


def extract_book(path):
    rec = {"file": os.path.basename(path), "size_kb": os.path.getsize(path) // 1024,
           "mtime": datetime.fromtimestamp(os.path.getmtime(path)).strftime("%Y-%m-%d")}
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    rec["sheets"] = [{"name": sn, "rows": ws.max_row, "cols": ws.max_column}
                     for sn, ws in ((sn, wb[sn]) for sn in wb.sheetnames)]

    # ---- 分析表 ----
    an_name = next((sn for sn in wb.sheetnames if "分析表" in sn), None)
    if an_name:
        ws = wb[an_name]
        rec["analysis"] = {"sheet": an_name, "rows": ws.max_row, "cols": ws.max_column}
        # 表头行：前 12 行里同时含 预留库存+可售库存 的行
        hdr_row, hdr = None, {}
        for r in range(1, 13):
            vals = [(c, str(c.value).strip()) for c in next(ws.iter_rows(min_row=r, max_row=r, max_col=min(ws.max_column, 130))) if c.value is not None]
            names = [t for _, t in vals]
            if "预留库存" in names and "可售库存" in names:
                hdr_row = r
                hdr = {t: c.column for c, t in vals}
                break
        rec["analysis"]["header_row"] = hdr_row
        if hdr_row:
            rec["analysis"]["key_cols"] = {k: GL(hdr[k]) for k in KEY_HEADERS if k in hdr}
            # 月总销量类列（9月总/8月总/月总 等）
            mon_cols = {t: GL(c) for t, c in hdr.items() if re.search(r"(总销量|月总)", t)}
            if mon_cols:
                rec["analysis"]["month_total_cols"] = mon_cols
            # 合计行 = 表头行+1（校验：c1 或 c2~c4 出现 合计/商品分类/数字）
            tot = {}
            cr = hdr_row + 1
            for k in KEY_HEADERS + list(mon_cols):
                if k in hdr:
                    v = ws.cell(row=cr, column=hdr[k]).value
                    tot[k] = num(v)
            rec["analysis"]["total_row"] = cr
            rec["analysis"]["totals"] = {k: v for k, v in tot.items() if v is not None}
            # 数据行数：SKU 列从合计+1 起非空计数（真扫）
            sku_col = hdr.get("SKU")
            n = 0
            last = cr
            if sku_col:
                for row in ws.iter_rows(min_row=cr + 1, max_row=ws.max_row, min_col=sku_col, max_col=sku_col):
                    v = row[0].value
                    if v is not None and str(v).strip():
                        n += 1
                        last = row[0].row
            rec["analysis"]["sku_rows"] = n
            rec["analysis"]["last_data_row"] = last
            # 日窗范围：表头行里 c>hdr[在途差额]（通常在 S 之后）的日期型值
            anchor = hdr.get("在途差额", 0)
            days = []
            for i, t in enumerate(next(ws.iter_rows(min_row=hdr_row, max_row=hdr_row,
                                                    max_col=min(ws.max_column, 130), values_only=True)), 1):
                if i > anchor and is_date(t):
                    days.append(t.date() if isinstance(t, datetime) else t)
            if days:
                rec["analysis"]["window"] = {"first": str(min(days)), "last": str(max(days)), "n_days": len(days)}
    # ---- 库存快照类 A1 锚点 ----
    snaps = {}
    for sn in wb.sheetnames:
        if re.search(r"(库存明细|库存$)", sn) or sn.endswith("库存"):
            try:
                a1 = wb[sn].cell(row=1, column=1).value
            except Exception:
                a1 = None
            if a1 is not None:
                snaps[sn] = str(a1)[:19]
    if snaps:
        rec["stock_anchors"] = snaps
    # ---- 订单/销售类行数 ----
    sales = {}
    for sn in wb.sheetnames:
        if "订单" in sn or "每日销售数据" in sn:
            sales[sn] = wb[sn].max_row
    if sales:
        rec["sales_rows"] = sales
    wb.close()
    return rec


def main():
    n_ok = n_err = 0
    with open(OUT, "w", encoding="utf-8") as fout:
        for month in MONTHS:
            for year_dir in YEARS:
                d = os.path.join(BASE, year_dir, month)
                if not os.path.isdir(d):
                    continue
                for fn in sorted(os.listdir(d)):
                    if not fn.endswith(".xlsx") or fn.startswith("~$"):
                        continue
                    path = os.path.join(d, fn)
                    try:
                        rec = extract_book(path)
                        rec["month"] = month
                        fout.write(json.dumps(rec, ensure_ascii=False) + "\n")
                        fout.flush()
                        n_ok += 1
                        an = rec.get("analysis", {})
                        print(f"[OK] {month} {rec['file']} "
                              f"sheets={len(rec['sheets'])} "
                              f"analysis(hdr r{an.get('header_row')}, sku {an.get('sku_rows')}行, "
                              f"窗口 {an.get('window', {}).get('first')}~{an.get('window', {}).get('last')})")
                    except Exception as e:
                        n_err += 1
                        print(f"[ERR] {month} {fn}: {type(e).__name__} {e}")
    print(f"=== 完成: OK {n_ok} / ERR {n_err} -> {OUT}")


if __name__ == "__main__":
    main()
