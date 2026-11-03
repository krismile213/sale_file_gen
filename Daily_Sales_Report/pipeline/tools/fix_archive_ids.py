# -*- coding: utf-8 -*-
"""修 12 份 9月归档件销售表的订单号类型：int → 文本（对齐人工口径，消灭 1.7E+14 显示）。
被占用的文件自动跳过并提示。"""
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
from runner import fix_id_text, restore_parts, _header_row
from openpyxl import load_workbook

# [normalize] DSR 由上方引导块赋值
SEP = os.path.join(DSR, "1.每日销售数据表", "2026年每日数据", "2026年9月")
BAK = os.path.join(DSR, ".workbuddy", "backup_archive_ids_20261009")
os.makedirs(BAK, exist_ok=True)

files = sorted(f for f in os.listdir(SEP) if f.endswith(".xlsx") and not f.startswith("~$"))
for f in files:
    p = os.path.join(SEP, f)
    try:
        fd = os.open(p, os.O_RDWR); os.close(fd)
    except OSError as e:
        print(f"[跳过] {f}: 文件被占用(errno {e.errno})——请关闭 WPS 后重跑")
        continue
    wb = load_workbook(p, data_only=False)
    n_total = 0
    for sn in wb.sheetnames:
        if "每日销售" not in sn:
            continue
        try:
            ws = wb[sn]
            n = fix_id_text(ws)
            if n:
                print(f"  {sn}: {n} 个 int→文本")
            n_total += n
        except Exception as e:
            print(f"  [!] {sn}: {e!r}")
    if n_total:
        shutil.copy2(p, os.path.join(BAK, f))
        tmp = p + ".new.xlsx"
        wb.save(tmp)
        wb.close()
        restored, piv = restore_parts(p, tmp)
        os.replace(tmp, p)
        print(f"[{f}] 修{n_total}处 保真回填{len(restored)}部件 ✓")
    else:
        wb.close()
        print(f"[{f}] 无需修复")
print("ARCHIVE ID FIX DONE")
