# -*- coding: utf-8 -*-
"""全簿验证平移结果：zip/日期块/星期/折叠/条件格式（对比备份）。"""
import os, sys, datetime, zipfile
HERE = os.path.dirname(os.path.abspath(__file__))
DSR = os.path.dirname(os.path.dirname(os.path.dirname(HERE)))
BKDIR = os.path.join(DSR, "1.每日销售数据表")
BAKDIR = os.path.join(DSR, ".workbuddy", "backup_shift_daycols_20261103")
import openpyxl
from openpyxl.utils import get_column_letter

WEEK = ["周一", "周二", "周三", "周四", "周五", "周六", "周日"]
EXP0, EXP1 = datetime.date(2026, 10, 25), datetime.date(2026, 12, 31)
only = None
for a in sys.argv[1:]:
    if a.startswith("--only="):
        only = a.split("=", 1)[1]

names = sorted(f for f in os.listdir(BKDIR) if f.endswith(".xlsx") and not f.startswith("~$"))
fails = []
for name in names:
    if only and only not in name:
        continue
    bk = os.path.join(BKDIR, name)
    bak = os.path.join(BAKDIR, name)
    probs = []
    zz = zipfile.ZipFile(bk)
    if zz.testzip() is not None:
        probs.append("zip损坏")
    zz.close()
    wb = openpyxl.load_workbook(bk, data_only=False)
    wb2 = openpyxl.load_workbook(bak) if os.path.exists(bak) else None
    for ws in wb.worksheets:
        if "每日库存分析表" not in ws.title:
            continue
        days = []
        for row in ws.iter_rows(min_row=5, max_row=5, min_col=20, max_col=135):
            for c in row:
                if hasattr(c.value, "strftime"):
                    days.append((c.column, c.value.date()))
        if not days:
            probs.append("%s: 无日列" % ws.title)
            continue
        if days[0][1] != EXP0 or days[-1][1] != EXP1:
            probs.append("%s: 首末=%s~%s" % (ws.title, days[0][1], days[-1][1]))
        exp = [EXP0 + datetime.timedelta(days=i) for i in range(len(days))]
        if [d for _, d in days] != exp:
            probs.append("%s: 日期断档" % ws.title)
        wbad = [(get_column_letter(c), ws.cell(1, c).value, str(d))
                for c, d in days
                if ws.cell(1, c).value not in (None, "") and ws.cell(1, c).value != WEEK[d.weekday()]]
        if wbad:
            probs.append("%s: 星期错 %s" % (ws.title, wbad[:2]))
        vis = [d for c, d in days if not ws.column_dimensions[get_column_letter(c)].hidden]
        hid = [d for c, d in days if ws.column_dimensions[get_column_letter(c)].hidden]
        if not (vis and vis[0] == EXP0 and vis[-1] == datetime.date(2026, 11, 3) and len(vis) == 10):
            probs.append("%s: 可见=%s~%s n=%d" % (ws.title, vis[0] if vis else '-', vis[-1] if vis else '-', len(vis)))
        if not (hid and hid[0] == datetime.date(2026, 11, 4) and hid[-1] == EXP1):
            probs.append("%s: 隐藏异常 n=%d" % (ws.title, len(hid)))
        if wb2 is not None and ws.title in wb2.sheetnames:
            n_new = len(list(ws.conditional_formatting))
            n_old = len(list(wb2[ws.title].conditional_formatting))
            if n_new != n_old:
                probs.append("%s: CF %d→%d" % (ws.title, n_old, n_new))
    wb.close()
    if wb2:
        wb2.close()
    tag = "PASS" if not probs else "FAIL"
    print("%s %s" % (tag, name))
    for p in probs:
        print("     ", p)
    if probs:
        fails.append(name)
print("=== %d/%d PASS ===" % (len(names) - len(fails), len(names)))
if fails:
    print("失败:", fails)
