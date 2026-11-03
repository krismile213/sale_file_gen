# -*- coding: utf-8 -*-
"""最小可行实验：能否正常 load 欧盟簿并访问分析表（验证是否被 SIGTERM）"""
import os, sys, time
sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "core"))
p = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))),
                 "1.每日销售数据表", "欧盟英国库存销售表.xlsx")
print("path", p, os.path.exists(p), flush=True)
t = time.time()
from openpyxl import load_workbook
print("importing done", time.time() - t, flush=True)
wb = load_workbook(p)
print("loaded", time.time() - t, "sheets:", len(wb.sheetnames), flush=True)
ws = wb["英国-每日库存分析表"]
print("max_row", ws.max_row, "max_col", ws.max_column, flush=True)
import shipments as Sh
fc, lc = Sh.detect_sheet_cols(ws)
print("clear range", fc, lc, flush=True)
import analysis as An
print("template", An.pick_template_row(ws, An._last_data_row(ws)), flush=True)
print("done", time.time() - t, flush=True)
