# -*- coding: utf-8 -*-
"""清理平移中间残留：*.pre_shift.xlsx / *.prefold.xlsx 挪进备份目录。"""
import os, shutil
DSR = os.path.dirname(os.path.dirname(os.path.abspath(__file__))) if False else None
HERE = os.path.dirname(os.path.abspath(__file__))
DSR = os.path.dirname(os.path.dirname(os.path.dirname(HERE)))
BKDIR = os.path.join(DSR, "1.每日销售数据表")
DST = os.path.join(DSR, ".workbuddy", "backup_shift_daycols_20261103", "中间残留")
os.makedirs(DST, exist_ok=True)
moved = 0
for f in sorted(os.listdir(BKDIR)):
    if (".pre_shift.xlsx" in f or ".prefold.xlsx" in f) and f.endswith(".xlsx"):
        shutil.move(os.path.join(BKDIR, f), os.path.join(DST, f))
        moved += 1
print("moved:", moved)
rest = [f for f in os.listdir(BKDIR) if f.endswith(".xlsx")]
print("剩余 xlsx:", len(rest))
for f in rest:
    print("  ", f)
