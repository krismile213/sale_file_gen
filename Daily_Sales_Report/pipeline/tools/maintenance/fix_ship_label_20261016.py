# -*- coding: utf-8 -*-
"""定点修正：按「当前解析器」重写某票货件列的行5标签（修历史误解析留下的错标签）。

背景（2026-10-16）：
  中东文件夹未登记进 FOLDER_RULES → load_shipments 用默认 DMY 解析货件日期；
  沙特 RUH8 的 `10/06/2026`（MM/DD）被读成 6/10（应 10/6）。
  代码已修（resolve_target 按仓库给格式 + TSV 按解析格式重算），但该票上一轮已落列，
  apply_sheet 对「已存在列」只补空格、不重写行5 → 旧错标签会一直留着。

本脚本：给定 (簿, 分析表, 货件号)，用 shipments 解析结果重算行5 标签并覆盖写。
管线：strip_pivot_records → openpyxl 编辑 → restore_parts → testzip → 备份 → 落盘。

用法：python fix_ship_label_20261016.py           # 干跑
     python fix_ship_label_20261016.py --apply
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
import shipments as S
import runner as R

APPLY = "--apply" in sys.argv
SHIP_DIR = r"C:/Users/dell/Desktop/work-main/sale_file_gen/trunk/data/数据归属周期/202610/货件"
BOOKS_DIR = os.path.join(DSR, "1.每日销售数据表")
BK = os.path.join(DSR, ".workbuddy", "backup_ship_label_20261016")
TMP = os.path.join(DSR, ".workbuddy", "_tmp_shiplabel")
# 目标：(簿文件名, 分析表名, 货件号)
TARGETS = [
    ("沙特库存销售表.xlsx", "沙特阿拉伯-每日库存分析表", "FBA15MM4V5DC"),
]

os.makedirs(BK, exist_ok=True)
os.makedirs(TMP, exist_ok=True)

ships = {str(s["id"]).strip(): s for s in S.load_shipments(SHIP_DIR)}

for book_fn, sheet_name, sid in TARGETS:
    cur = os.path.join(BOOKS_DIR, book_fn)
    s = ships.get(sid)
    if s is None:
        print(f"[跳过] 货件目录里没有 {sid}")
        continue
    want = S.label_of(s, s["transport"], None, s["arrived"])
    print(f"== {book_fn} / {sheet_name} / {sid} ==")
    print(f"   解析：wh={s['wh']} ship_md={s['ship_md']} 运输={s['transport']} → 期望行5='{want}'")
    n = R.strip_pivot_records(cur, os.path.join(TMP, "stripped.xlsx"))
    wb = R.load_workbook(os.path.join(TMP, "stripped.xlsx") if n else cur)
    if sheet_name not in wb.sheetnames:
        print(f"   [跳过] 无分析表 {sheet_name}")
        wb.close()
        continue
    ws = wb[sheet_name]
    col = None
    for c in range(1, 60):
        v = ws.cell(1, c).value
        if v is not None and str(v).strip() == sid:
            col = c
            break
    if col is None:
        print(f"   [跳过] 表内找不到货件号 {sid}")
        wb.close()
        continue
    old = ws.cell(5, col).value
    L = ws.cell(1, col).column_letter
    print(f"   列={L}  现有行5={old!r}")
    if old == want:
        print("   [无需处理] 已一致")
        wb.close()
        continue
    ws.cell(5, col).value = want
    if not APPLY:
        print(f"   [dry-run] 将写 {L}5 = {want!r}（未落盘）")
        wb.close()
        continue
    saved = os.path.join(TMP, "saved_%s" % book_fn)
    wb.save(saved)
    wb.close()
    R.restore_parts(cur, saved)
    with zipfile.ZipFile(saved) as z:
        bad = z.testzip()
    if bad:
        raise RuntimeError(f"zip 损坏: {bad}")
    dst = os.path.join(BK, book_fn)
    if not os.path.exists(dst):
        shutil.copy2(cur, dst)
    try:
        os.replace(saved, cur)
    except PermissionError:
        shutil.copyfile(saved, cur)
    print(f"   [OK] {L}5 = {want!r} 已落盘；备份→{BK}")
