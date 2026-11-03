# -*- coding: utf-8 -*-
"""给 墨西哥 / 印度 两张分析表补 SKU 行 G598（红米17 钢化膜）。
做法=整行克隆模板行（最后数据行）：公式列 Translator 平移、静态值/样式照抄，
然后覆写 A=SKU、B=型号、货件列(9..差额列-1)清空、品牌列按名称修正。
幂等：A 列已存在该 SKU 则跳过。
用法：python fix_g598_rows.py            # 干运行，只打印计划
     python fix_g598_rows.py --apply    # 落盘
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
from copy import copy
from openpyxl.formula.translate import Translator
from openpyxl.worksheet.formula import ArrayFormula
from openpyxl.utils import get_column_letter as gl

# [normalize] DSR 由上方引导块赋值
BOOKS_DIR = os.path.join(DSR, "1.每日销售数据表")
BK = os.path.join(DSR, ".workbuddy", "backup_g598rows_20261014")
TMP = os.path.join(DSR, ".workbuddy", "_tmp_g598")
OUT = os.path.join(DSR, ".workbuddy", "fix_g598_out.txt")
SHIP_DIR = r"C:/Users/dell/Desktop/work-main/sale_file_gen/trunk/data/数据归属周期/202610/货件"
pass  # [normalize] 路径已由引导块设置
import runner as R
import shipments as Sh

APPLY = "--apply" in sys.argv

# (簿, 分析表, 模板行, SKU, 对应 TSV 货件号)
TARGETS = [
    ("北美加拿大库存销售表.xlsx", "墨西哥-每日库存分析表 ", 409, "G598", "FBA19SY7T08X"),
    ("印度库存销售表.xlsx", "印度-每日库存分析表", 438, "G598", "FBA15MP0QHYN"),
]
BRANDS = ("Sony", "Poco", "Redmi", "Xiaomi", "Motorola", "Samsung", "Apple", "iPhone", "Honor", "Oppo", "Vivo", "Realme", "Google", "OnePlus", "Nokia")

os.makedirs(BK, exist_ok=True)
os.makedirs(TMP, exist_ok=True)
_lines = []
def emit(s=''):
    _lines.append(str(s))

def brand_of(name):
    n = name or ''
    for b in ("Sony", "Poco", "Motorola", "Samsung", "Apple", "Honor", "Oppo", "Vivo", "Realme", "Google", "OnePlus", "Nokia", "iPhone"):
        if b.lower() in n.lower():
            return b
    if 'Xiaomi' in n or 'Redmi' in n:
        return 'Redmi'
    return ''

ships = {s['id']: s for s in Sh.load_shipments(SHIP_DIR)}

try:
    for fn, sheet, tmpl, sku, sid in TARGETS:
        cur = os.path.join(BOOKS_DIR, fn)
        emit("=" * 70)
        emit(f"# {fn} | {sheet!r} | 模板行 {tmpl} | 补 {sku} ({sid})")
        n = R.strip_pivot_records(cur, os.path.join(TMP, "stripped.xlsx"))
        src = os.path.join(TMP, "stripped.xlsx") if n else cur
        wb = R.load_workbook(src)
        real = next((x for x in wb.sheetnames if x.strip() == sheet.strip()), None)
        if real is None:
            raise RuntimeError(f"找不到 sheet {sheet!r}")
        ws = wb[real]

        # 幂等检查
        have = set()
        last = 0
        for r in range(7, ws.max_row + 1):
            v = ws.cell(r, 1).value
            if v is not None and not str(v).startswith('='):
                have.add(str(v).strip()); last = r
        if sku in have:
            emit(f"  已存在 {sku}，跳过"); wb.close(); continue

        nr = last + 1
        fc, lc = Sh.detect_sheet_cols(ws)
        name = ships[sid]['names'].get(sku, '')
        brand = brand_of(name)
        # 品牌列探测：模板行 60..110 中值命中品牌集合的列
        brand_col = None
        for c in range(60, min(sw := (ws.max_column + 1), 111)):
            v = ws.cell(tmpl, c).value
            if isinstance(v, str) and v.strip() in BRANDS:
                brand_col = c; break
        emit(f"  最后数据行 {last} → 新行 {nr} | 货件列 {gl(fc)}..{gl(lc)} | 品牌列={gl(brand_col) if brand_col else '未找到'} | name={name[:60]}")

        maxc = ws.max_column
        copied = 0
        for c in range(1, maxc + 1):
            tc = ws.cell(tmpl, c); nc = ws.cell(nr, c)
            v = tc.value
            if fc <= c <= lc:            # 货件列：清空
                nc.value = None
            elif isinstance(v, ArrayFormula):
                txt = Translator(v.text, origin=tc.coordinate).translate_formula(nc.coordinate)
                nc.value = ArrayFormula(ref=nc.coordinate, text=txt)
                copied += 1
            elif isinstance(v, str) and v.startswith('='):
                nc.value = Translator(v, origin=tc.coordinate).translate_formula(nc.coordinate)
                copied += 1
            else:
                nc.value = v
            # 样式照抄
            try:
                nc._style = copy(tc._style)
            except Exception:
                pass
        ws.cell(nr, 1).value = sku
        ws.cell(nr, 2).value = name
        if brand_col:
            ws.cell(nr, brand_col).value = brand
        emit(f"  克隆 {copied} 个公式格；A={sku} B={name[:30]} 品牌{gl(brand_col) if brand_col else '-'}={brand}")

        if not APPLY:
            emit("  [dry-run] 未落盘")
            wb.close(); continue

        saved = os.path.join(TMP, "saved_%s" % fn)
        wb.save(saved); wb.close()
        R.restore_parts(cur, saved)
        with zipfile.ZipFile(saved) as z:
            bad = z.testzip()
        if bad:
            raise RuntimeError(f"zip 损坏: {bad}")
        dst = os.path.join(BK, fn)
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
print('done, apply=%s' % APPLY)
