# -*- coding: utf-8 -*-
"""回归检查：_month_summary_sheets 在 13 本簿里到底会处理哪些 sheet。"""
import os, sys, glob
HERE = os.path.dirname(os.path.abspath(__file__))
CORE = os.path.abspath(os.path.join(HERE, "..", "core"))
sys.path.insert(0, CORE)
import config as C
import runner as R
import analysis as An
from openpyxl import load_workbook

files = sorted(glob.glob(os.path.join(C.WORK_DIR, "*.xlsx")))
for p in files:
    nm = os.path.basename(p)
    try:
        wb = load_workbook(p, read_only=True, data_only=False)
    except Exception as e:
        print("%-40s !! 打不开 %r" % (nm, e)); continue
    try:
        old = [s for s in wb.sheetnames if "分析" in s]
        new = R._month_summary_sheets(wb, An)
        added = [s for s in new if s not in old]
        print("%-38s 分析表%d → 总%d  新增: %s" % (nm, len(old), len(new), added or "-"))
    except Exception as e:
        print("%-38s !! %r" % (nm, e))
    finally:
        wb.close()
