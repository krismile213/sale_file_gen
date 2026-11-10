# -*- coding: utf-8 -*-
"""对已平移的簿重跑折叠（sync_day_outline 新边界：快照日次日）。"""
import os, sys, datetime
HERE = os.path.dirname(os.path.abspath(__file__))
DSR = os.path.dirname(os.path.dirname(os.path.dirname(HERE)))
PIPE = os.path.join(DSR, "pipeline")
for p in (PIPE, os.path.join(PIPE, "core")):
    if p not in sys.path:
        sys.path.insert(0, p)
import openpyxl
from core import analysis as An

SNAP = datetime.date(2026, 11, 3)
BK = os.path.join(DSR, "1.每日销售数据表", "印度库存销售表.xlsx")
wb = openpyxl.load_workbook(BK)
for ws in wb.worksheets:
    if "每日库存分析表" in ws.title:
        An.sync_day_outline(ws, SNAP)
tmp = BK + ".folding.xlsx"
wb.save(tmp)
wb.close()
bak = BK + ".prefold.xlsx"
if os.path.exists(bak):
    os.remove(bak)
os.replace(BK, bak)
os.replace(tmp, BK)
print("refold done")
