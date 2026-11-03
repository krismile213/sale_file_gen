# -*- coding: utf-8 -*-
"""给美客多 4 票「无发货日期」货件列补行5 标签的 M/D 前缀（2026-10-16，用户确认）。

背景：美客多 PDF 无发货日期（`ship_md=None`），`label_of` 只写运输方式 ⇒ 行5 为「空运」；
分析表行 2「发货日期」= `=IFERROR(LEFT(LEFT(<col>5,5),FIND("/",LEFT(<col>5,5))-1)&…,"")`
从行5 前 5 字解析 ⇒ 无 M/D 前缀则行2/3/4 全空。

日期依据：货件列位次 = 发货顺序，前 6 票严格 7 天一期
（8/11→8/18→8/25→9/1→9/8→9/15）⇒ 后 4 票 9/22、9/29、10/6、10/13（用户已确认按此改）。

管线：strip_pivot_records → openpyxl 编辑 → restore_parts → testzip → 备份 → 落盘。
用法：python fix_ml_ship_dates_20261016.py [--apply]
"""
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

import os, sys, shutil, zipfile, traceback
from openpyxl.utils import get_column_letter as gl
import runner as R

APPLY = "--apply" in sys.argv
BOOK = "美客多库存销售表.xlsx"
SHEET = "墨西哥-每日库存分析表"
BOOKS_DIR = os.path.join(DSR, "1.每日销售数据表")
BK = os.path.join(DSR, ".workbuddy", "backup_ml_ship_dates_20261016")
TMP = os.path.join(DSR, ".workbuddy", "_tmp_ml_dates")

# 货件号 → 新的行5 标签（M/D + 原运输方式）
MAP = {
    "77739558": "9/22空运",
    "78704820": "9/29空运",
    "78793377": "10/6空运",
    "79402478": "10/13空运",
}

os.makedirs(BK, exist_ok=True)
os.makedirs(TMP, exist_ok=True)
_lines = []


def emit(s=""):
    _lines.append(str(s))
    print(s)


cur = os.path.join(BOOKS_DIR, BOOK)
n = R.strip_pivot_records(cur, os.path.join(TMP, "stripped.xlsx"))
src = os.path.join(TMP, "stripped.xlsx") if n else cur
wb = R.load_workbook(src)
ws = wb[SHEET]

# 行1 找货件号 → 列
found = {}
for c in range(9, ws.max_column + 1):
    v = ws.cell(1, c).value
    if v in (None, ""):
        continue
    k = str(v).strip()
    if k in MAP:
        found[k] = c

emit(f"# {BOOK} / {SHEET}")
emit(f"  目标 {len(MAP)} 票，命中 {len(found)} 票")
touched = 0
for k, newlab in MAP.items():
    if k not in found:
        emit(f"  !! {k} 未在行1 找到，跳过")
        continue
    c = found[k]
    old = ws.cell(5, c).value
    emit(f"  {gl(c)}5 [{k}]  {old!r}  →  {newlab!r}")
    if APPLY:
        ws.cell(5, c).value = newlab
    touched += 1

if not APPLY:
    emit("  [dry-run] 未落盘")
    wb.close()
    raise SystemExit(0)

saved = os.path.join(TMP, "saved.xlsx")
wb.save(saved)
wb.close()
R.restore_parts(cur, saved)
with zipfile.ZipFile(saved) as z:
    bad = z.testzip()
if bad:
    raise SystemExit(f"zip 损坏: {bad}")
dst = os.path.join(BK, BOOK)
if not os.path.exists(dst):
    shutil.copy2(cur, dst)
try:
    os.replace(saved, cur)
except PermissionError:
    shutil.copyfile(saved, cur)
emit(f"  [OK] 已落盘 {touched} 票；备份→{BK}")

with open(os.path.join(DSR, ".workbuddy", "fix_ml_ship_dates_out.txt"), "w",
          encoding="utf-8") as f:
    f.write("\n".join(_lines) + "\n")
