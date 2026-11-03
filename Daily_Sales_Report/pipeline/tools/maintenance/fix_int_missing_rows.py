# -*- coding: utf-8 -*-
"""补 Interflowd北美加拿大簿 加拿大/美国 两张分析表缺失的 G594-G597 新 SKU 行。
模板=行75（完整公式行）；公式列用 openpyxl Translator 平移；值格照抄后覆写
A/B/BV/CJ/K(货件数量)。幂等：已存在的 SKU 跳过。
管线：strip_pivot_records → openpyxl 编辑 → restore_parts → testzip → 备份 → 落盘。
"""
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
import os, sys, shutil, zipfile, traceback
from openpyxl.formula.translate import Translator

# [normalize] DSR 由上方引导块赋值
CUR = os.path.join(DSR, "1.每日销售数据表", "Interflowd北美加拿大库存销售表.xlsx")
BK = os.path.join(DSR, ".workbuddy", "backup_introws_20261014")
TMP = os.path.join(DSR, ".workbuddy", "_tmp")
OUT = os.path.join(DSR, ".workbuddy", "fix_introws_out.txt")
pass  # [normalize] 路径已由引导块设置
import runner as R
import shipments as Sh

# (分析表, TSV 货件号, 货件列)
PLAN = [
    ("加拿大-每日库存分析表", "FBA19RJ5R9LK", "K"),
    ("美国-每日库存分析表", "FBA19QQ8H4L2", "K"),
]
TMPL_ROW = 75          # 完整公式模板行
MAX_NEW_ROW = 90       # 安全上限

os.makedirs(BK, exist_ok=True)
os.makedirs(TMP, exist_ok=True)
_lines = []
def emit(s=''):
    _lines.append(str(s))

ships = {s['id']: s for s in Sh.load_shipments(
    os.path.join(r'C:/Users/dell/Desktop/work-main/sale_file_gen/trunk/data/数据归属周期/202610/货件'))}

def brand_of(name):
    n = name or ''
    if 'Sony' in n: return 'Sony'
    if 'Poco' in n: return 'Poco'
    if 'Xiaomi' in n or 'Redmi' in n: return 'Redmi'
    if 'Motorola' in n: return 'Motorola'
    return ''

try:
    n = R.strip_pivot_records(CUR, os.path.join(TMP, "introws.stripped.xlsx"))
    wb = R.load_workbook(os.path.join(TMP, "introws.stripped.xlsx") if n else CUR)
    from openpyxl.utils import column_index_from_string

    for sheet, sid, col_letter in PLAN:
        ws = wb[sheet]
        ci = column_index_from_string(col_letter)
        s = ships[sid]
        have = set()
        last = 0
        for r in range(7, ws.max_row + 1):
            v = ws.cell(r, 1).value
            if v is not None and not str(v).startswith('='):
                have.add(str(v).strip())
                last = r
        emit(f"== {sheet}: 现有 {len(have)} SKU, 最后数据行 {last}, 货件 {sid} 共 {len(s['items'])} SKU")
        new_row = last + 1 if last >= TMPL_ROW else TMPL_ROW + 1
        for sku, qty in s['items'].items():
            if sku in have:
                emit(f"   跳过已有行 {sku}")
                continue
            if new_row > MAX_NEW_ROW:
                raise RuntimeError("行号超出安全上限")
            name = s['names'].get(sku, '')
            brand = brand_of(name)
            # 逐列从模板行平移/照抄
            for col in range(1, 120):
                tv = ws.cell(TMPL_ROW, col).value
                if tv is None:
                    continue
                if isinstance(tv, str) and tv.startswith('='):
                    ws.cell(new_row, col).value = Translator(
                        tv, origin=f"{ws.cell(TMPL_ROW, col).coordinate}"
                    ).translate_formula(f"{ws.cell(new_row, col).coordinate}")
                else:
                    ws.cell(new_row, col).value = tv
            ws.cell(new_row, 1).value = sku
            ws.cell(new_row, 2).value = name
            ws.cell(new_row, 22).value = '钢化膜'      # BV 品类
            ws.cell(new_row, 88).value = brand         # CJ 品牌
            ws.cell(new_row, ci).value = qty           # 货件列数量
            emit(f"   +行{new_row}: {sku} {qty:g}件 [{brand}] {name[:40]} → {col_letter}{new_row}")
            new_row += 1
        # 行6 SUBTOTAL/SUM 为开区间(至2545/2467)，新行自动纳入——无需改
    wb.save(os.path.join(TMP, "introws.saved.xlsx"))
    wb.close()
    R.restore_parts(CUR, os.path.join(TMP, "introws.saved.xlsx"))
    with zipfile.ZipFile(os.path.join(TMP, "introws.saved.xlsx")) as z:
        bad = z.testzip()
    if bad:
        raise RuntimeError(f"zip 损坏: {bad}")
    dst = os.path.join(BK, os.path.basename(CUR))
    if not os.path.exists(dst):
        shutil.copy2(CUR, dst)
    try:
        os.replace(os.path.join(TMP, "introws.saved.xlsx"), CUR)
    except PermissionError:
        shutil.copyfile(os.path.join(TMP, "introws.saved.xlsx"), CUR)
    emit("[OK] 落盘完成（备份→" + BK + "）")
except Exception as e:
    emit("[FAIL] " + repr(e))
    emit(traceback.format_exc())

with open(OUT, 'w', encoding='utf-8') as f:
    f.write('\n'.join(_lines) + '\n')
print('done')
