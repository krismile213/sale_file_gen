# -*- coding: utf-8 -*-
"""把 10 月日销表 vs 月销表的逐单差异整理成一份可核对的 Excel。

数据源：pipeline/tools/audit_sales_gap.py 产出的 归因_{period}_订单明细.csv
输出：断货分析/_对账/差异订单清单_{period}.xlsx
"""
import csv
import os
import sys

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

HERE = os.path.dirname(os.path.abspath(__file__))
DSR = os.path.abspath(os.path.join(HERE, "..", ".."))
sys.path.insert(0, os.path.join(DSR, "pipeline", "core"))
import config as C  # noqa: E402

_args = [a for a in sys.argv[1:] if not a.startswith("--")
         and (a.isdigit() or a == "period")]
_clean = []
for i, a in enumerate(sys.argv[1:]):
    if a == "--period":
        _clean.append(sys.argv[i + 2])
    elif not a.startswith("--"):
        _clean.append(a)
PERIOD = next((a for a in _clean if a.isdigit()), "202610")
OUT_DIR = os.path.join(DSR, "断货分析", "_对账")
CSV = os.path.join(OUT_DIR, "归因_%s_订单明细.csv" % PERIOD)
OUT = os.path.join(OUT_DIR, "差异订单清单_%s.xlsx" % PERIOD)

HDR_FILL = PatternFill("solid", fgColor="FFD9E1F2")
BOLD = Font(bold=True)
CENTER = Alignment(horizontal="center", vertical="center")


def read_rows():
    with open(CSV, encoding="utf-8-sig") as f:
        return list(csv.DictReader(f))


def style_head(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    for c in ws[1]:
        c.fill = HDR_FILL
        c.font = BOLD
        c.alignment = CENTER
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions


def add_sheet(wb, title, head, rows, widths):
    ws = wb.create_sheet(title)
    ws.append(head)
    for r in rows:
        ws.append(r)
    style_head(ws, widths)
    return ws


def main():
    if not os.path.exists(CSV):
        print("找不到 %s，先跑 audit_sales_gap.py" % CSV)
        return 2
    rows = read_rows()
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # ---------- 1 汇总 ----------
    order = ["③ 取消后未回退", "④ 日销明细漏记", "⑤ 跨月(日销多)",
             "⑤ 跨月(月销多)", "① 当时已取消后复活", "② 报告里已无此单"]
    agg = {}
    for r in rows:
        k = r["归因"]
        a = agg.setdefault(k, [0, 0.0])
        a[0] += 1
        a[1] += abs(float(r["数量"]))
    ws = wb.create_sheet("汇总")
    ws.append(["归因", "单数", "件数", "方向", "是不是真差"])
    tot_n = tot_q = 0
    for k in order:
        if k not in agg:
            continue
        n, q = agg[k]
        if k.startswith("③") or k.startswith("②") or k.startswith("①"):
            d, real = "日销表多算", "否（已校正后归零）"
        elif k.startswith("④"):
            d, real = "日销表少算", "**是**（可补录）"
        elif "日销多" in k:
            d, real = "日销表多算", "否（归月基准不同）"
        else:
            d, real = "日销表少算", "否（归月基准不同）"
        ws.append([k, n, int(q), d, real])
        tot_n += n
        tot_q += q
    ws.append(["合计差异单", tot_n, int(tot_q), "", ""])
    style_head(ws, [22, 8, 8, 16, 24])
    ws["A" + str(ws.max_row)].font = BOLD

    # ---------- 2 漏记 33 单（真正要补的） ----------
    c4 = [r for r in rows if r["归因"].startswith("④")]
    c4.sort(key=lambda r: (r["分组"], r["purchase-date"]))
    add_sheet(wb, "④ 日销表漏记(要补)",
              ["分组", "渠道", "SKU", "订单号", "数量", "下单时间", "最终状态", "日销表里"],
              [[r["分组"], r["渠道"], r["SKU"], r["amazon-order-id"],
                int(abs(float(r["数量"]))), r["purchase-date"], r["最终状态"], "没有这张单"]
               for r in c4],
              [10, 20, 14, 26, 8, 28, 12, 14])

    # ---------- 3 跨月（日销多） ----------
    c5a = [r for r in rows if r["归因"] == "⑤ 跨月(日销多)"]
    c5a.sort(key=lambda r: (r["分组"], r["purchase-date"]))
    add_sheet(wb, "⑤ 跨月-日销多",
              ["分组", "渠道", "SKU", "订单号", "数量", "下单时间", "最终状态", "说明"],
              [[r["分组"], r["渠道"], r["SKU"], r["amazon-order-id"],
                int(abs(float(r["数量"]))), r["purchase-date"], r["最终状态"],
                "UTC 属 10 月，站点本地已 11 月"] for r in c5a],
              [10, 20, 14, 26, 8, 28, 12, 26])

    # ---------- 4 跨月（月销多） ----------
    c5b = [r for r in rows if r["归因"] == "⑤ 跨月(月销多)"]
    c5b.sort(key=lambda r: (r["分组"], r["purchase-date"]))
    add_sheet(wb, "⑤ 跨月-月销多",
              ["分组", "渠道", "SKU", "订单号", "数量", "下单时间", "最终状态", "说明"],
              [[r["分组"], r["渠道"], r["SKU"], r["amazon-order-id"],
                int(abs(float(r["数量"]))), r["purchase-date"], r["最终状态"],
                "站点本地属 10 月，UTC 还在 9 月"] for r in c5b],
              [10, 20, 14, 26, 8, 28, 12, 26])

    # ---------- 5 订单源与口径 ----------
    ws = wb.create_sheet("订单源与口径")
    ws.append(["工作簿", "明细 sheet", "订单报告源（{m}=周期，如 202610）"])
    for w in C.WORKBOOKS:
        s = w.get("sales")
        if isinstance(s, list):
            rep = " ; ".join(str(x.get("report")) for x in s)
            sh = " ; ".join(str(x.get("sheet")) for x in s)
        elif isinstance(s, dict):
            r = s.get("report")
            rep = " | ".join(r) if isinstance(r, list) else str(r)
            sh = str(s.get("sheet"))
        else:
            continue
        rep = rep.replace("C:\\Users\\dell\\Desktop\\work-main\\sale_file_gen\\trunk\\data\\", "trunk/data/")
        ws.append([w.get("file"), sh, rep])
    style_head(ws, [46, 22, 88])
    ws.append([])
    ws.append(["【两边口径差异】", "", ""])
    for line in [
        "日销表：每跑一次批，把报告里的行追加进明细 sheet，只增不删。",
        "  过滤：order-status ≠ Cancelled 且 sales-channel 前 7 位 = 'Amazon.'（北美簿额外放行 Non-Amazon/TM）",
        "  归月：按 purchase-date 的 UTC 日期（明细辅助列 AE = LEFT(C,10)，SUMIFS 按它匹配）",
        "月销表：trunk/month_sale_file.py 用同一批订单报告整月重算。",
        "  过滤：order-status ≠ Cancelled，并剔除 SKU 以 Amazon.Found. / amzn.gr. 开头的行",
        "  归月：purchase-date 先 tz_convert 到站点本地时区，再取 .month",
        "因此两边必然出现的三类差：①状态时点（已用 status_sync 在封存前校正）",
        "  ②归月基准不同（UTC vs 本地，单月几十件，全年抵消）",
        "  ③封存前那一版报告里的单没来得及追加（本次 33 件）",
    ]:
        ws.append([line, "", ""])
    ws["A" + str(ws.max_row - 9)].font = BOLD

    wb.save(OUT)
    print("→", OUT)
    print("  汇总：%d 单 / %d 件" % (tot_n, tot_q))
    for k in order:
        if k in agg:
            print("   %-18s %4d 单 %6d 件" % (k, agg[k][0], int(agg[k][1])))
    return 0


if __name__ == "__main__":
    sys.exit(main())
