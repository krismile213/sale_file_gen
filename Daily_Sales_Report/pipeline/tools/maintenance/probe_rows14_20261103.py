# -*- coding: utf-8 -*-
"""侦察：日列区行1~4 是否有随日期走的标注（星期等）。"""
import os, sys, datetime
HERE = os.path.dirname(os.path.abspath(__file__))
DSR = os.path.dirname(os.path.dirname(os.path.dirname(HERE)))
import openpyxl
from openpyxl.utils import get_column_letter

P = os.path.join(DSR, "1.每日销售数据表", "欧盟英国库存销售表.xlsx")
wb = openpyxl.load_workbook(P, read_only=True, data_only=False)
ws = wb["欧洲-每日库存分析表"]
rows = {}
for r in (1, 2, 3, 4, 5):
    vals = []
    for row in ws.iter_rows(min_row=r, max_row=r, min_col=40, max_col=115):
        for c in row:
            if c.value not in (None, ""):
                vals.append((c.coordinate, str(c.value)[:16]))
    rows[r] = vals[:8]
    print("行%d 非空(AS..DK段): %d 个" % (r, len(vals)), rows[r])
wb.close()
