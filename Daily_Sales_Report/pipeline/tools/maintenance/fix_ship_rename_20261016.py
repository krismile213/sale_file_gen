# -*- coding: utf-8 -*-
"""印度货件换号：FBA15MP0QHYN → FBA15MP79NRK（同量 48 SKU/484 件）。

只改两格：分析表 行1 货件号(P1)、行5 标签(P5 → 10/14空运)；
P 列数据(48 SKU/484 件)原样保留 —— 避免重跑时被当成新票另占一列 → 双计。

幂等：P1 已是新号则跳过。
用法：python fix_ship_rename_20261016.py           # 干跑
     python fix_ship_rename_20261016.py --apply   # 落盘
管线：strip_pivot_records → openpyxl 编辑 → restore_parts → testzip → 备份 → 落盘。
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

BOOKS_DIR = os.path.join(DSR, "1.每日销售数据表")
BK = os.path.join(DSR, ".workbuddy", "backup_ship_rename_20261016")
TMP = os.path.join(DSR, ".workbuddy", "_tmp_ship_rename")
OUT = os.path.join(DSR, ".workbuddy", "fix_ship_rename_out.txt")
import runner as R

APPLY = "--apply" in sys.argv
OLD_ID, NEW_ID = "FBA15MP0QHYN", "FBA15MP79NRK"
NEW_LABEL = "10/14空运"
BOOK, SHEET, COL = "印度库存销售表.xlsx", "印度-每日库存分析表", 16   # P

os.makedirs(BK, exist_ok=True)
os.makedirs(TMP, exist_ok=True)
_lines = []
def emit(s=''):
    _lines.append(str(s)); print(s)

try:
    cur = os.path.join(BOOKS_DIR, BOOK)
    emit("=" * 70)
    emit(f"# {BOOK} | {SHEET} | {gl(COL)} 列：{OLD_ID} → {NEW_ID}")
    n = R.strip_pivot_records(cur, os.path.join(TMP, "stripped.xlsx"))
    src = os.path.join(TMP, "stripped.xlsx") if n else cur
    wb = R.load_workbook(src)
    real = next((x for x in wb.sheetnames if x.strip() == SHEET.strip()), None)
    if real is None:
        raise RuntimeError(f"找不到 sheet {SHEET!r}")
    ws = wb[real]

    old1 = ws.cell(1, COL).value
    old5 = ws.cell(5, COL).value
    data = [(r, ws.cell(r, COL).value) for r in range(7, ws.max_row + 1)
            if isinstance(ws.cell(r, COL).value, (int, float)) and ws.cell(r, COL).value]
    emit(f"  现状: 行1={old1!r}  行5={old5!r}  P列有值行={len(data)}  合计={sum(v for _, v in data)}")

    if str(old1).strip() == NEW_ID:
        emit("  已是新号，跳过"); wb.close()
    else:
        if str(old1).strip() != OLD_ID:
            raise RuntimeError(f"行1 不是 {OLD_ID}（实际 {old1!r}），中止")
        ws.cell(1, COL).value = NEW_ID
        ws.cell(5, COL).value = NEW_LABEL
        emit(f"  改为: 行1={NEW_ID!r}  行5={NEW_LABEL!r}  （P列 {len(data)} 行数据不动）")
        if not APPLY:
            emit("  [dry-run] 未落盘"); wb.close()
        else:
            saved = os.path.join(TMP, "saved_%s" % BOOK)
            wb.save(saved); wb.close()
            R.restore_parts(cur, saved)
            with zipfile.ZipFile(saved) as z:
                bad = z.testzip()
            if bad:
                raise RuntimeError(f"zip 损坏: {bad}")
            dst = os.path.join(BK, BOOK)
            if not os.path.exists(dst):
                shutil.copy2(cur, dst)
            try:
                os.replace(saved, cur)
            except PermissionError:
                shutil.copyfile(saved, cur)
            emit(f"  [OK] 落盘（备份→{BK}）")
except Exception as e:
    emit("[FAIL] " + repr(e))
    emit(traceback.format_exc())

with open(OUT, 'w', encoding='utf-8') as f:
    f.write('\n'.join(_lines) + '\n')
