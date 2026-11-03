# -*- coding: utf-8 -*-
"""一次性补缺：用 202609 周期的 EMAG 订单报告补齐工作簿漏粘的 66 单。
日常增量走 runner（模板 EMAG全站点{m}订单报告.xlsx）。"""
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
import os, sys, io, shutil
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from runner import append_sales, restore_parts
from openpyxl import load_workbook

# [normalize] DSR 由上方引导块赋值
WORK = os.path.join(DSR, "1.每日销售数据表", "罗马尼亚库存销售表.xlsx")
REPORT = r"C:\Users\dell\Desktop\work-main\sale_file_gen\trunk\data\数据归属周期\202609\快照日期\20261006\EMAG全站点202609订单报告.xlsx"
BAK = os.path.join(DSR, ".workbuddy", "backup_emag_backfill_20261009")

# 锁检测
try:
    fd = os.open(WORK, os.O_RDWR); os.close(fd)
except OSError as e:
    raise SystemExit(f"文件被 WPS 占用(errno {e.errno})，请关闭后重跑")

os.makedirs(BAK, exist_ok=True)
shutil.copy2(WORK, os.path.join(BAK, "罗马尼亚库存销售表.xlsx"))
print("备份 →", BAK)

wb = load_workbook(WORK, data_only=False)
ws = wb["订单数据"]
before = ws.max_row
res = append_sales(ws, REPORT, sep=None, key="订单编号")
if res:
    tmp = WORK + ".new.xlsx"
    wb.save(tmp)
    wb.close()
    restored, piv = restore_parts(WORK, tmp)
    os.replace(tmp, WORK)
    print(f"订单数据 r{before} → r{res and ''}{before}", f"追加{res['appended']}行 | 保真回填{len(restored)}部件 ✓")
else:
    wb.close()
    print("无可追加订单")
print("BACKFILL DONE")
