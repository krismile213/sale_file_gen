# -*- coding: utf-8 -*-
"""补缺失 SKU 行：Interflowd澳「澳大利亚-每日库存分析表」缺 G598（红米17 钢化膜），
导致货件 FBA15GK3NYR1 的 G598(2 件) 落不进表（日志「未命中 1」）。

做法（同 fix_g598_rows.py）：克隆**末数据行** → 公式列 Translator 平移、样式照抄 →
货件列(9..差额列-1)清空 → 覆写 A=SKU、B=型号、品牌列。幂等：A 列已存在该 SKU 则跳过。
管线：strip_pivot_records → openpyxl 编辑 → restore_parts → testzip → 备份 → 落盘。
用法：python fix_missing_sku_rows_20261016.py            # 干跑
     python fix_missing_sku_rows_20261016.py --apply
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
from copy import copy
from openpyxl.formula.translate import Translator
from openpyxl.worksheet.formula import ArrayFormula
from openpyxl.utils import get_column_letter as gl
import runner as R
import shipments as Sh

APPLY = "--apply" in sys.argv
BOOKS_DIR = os.path.join(DSR, "1.每日销售数据表")
BK = os.path.join(DSR, ".workbuddy", "backup_missing_sku_20261016")
TMP = os.path.join(DSR, ".workbuddy", "_tmp_missku")
SHIP_DIR = r"C:/Users/dell/Desktop/work-main/sale_file_gen/trunk/data/数据归属周期/202610/货件"
# (簿, 分析表, SKU, 对应货件号)
TARGETS = [
    ("Interflowd澳大利亚库存销售表.xlsx", "澳大利亚-每日库存分析表", "G598", "FBA15GK3NYR1"),
]
BRANDS = ("Sony", "Poco", "Redmi", "Xiaomi", "Motorola", "Samsung", "Apple",
          "iPhone", "Honor", "Oppo", "Vivo", "Realme", "Google", "OnePlus", "Nokia")

os.makedirs(BK, exist_ok=True)
os.makedirs(TMP, exist_ok=True)


def brand_of(name):
    n = name or ""
    for b in ("Sony", "Poco", "Motorola", "Samsung", "Apple", "Honor", "Oppo",
              "Vivo", "Realme", "Google", "OnePlus", "Nokia", "iPhone"):
        if b.lower() in n.lower():
            return b
    if "Xiaomi" in n or "Redmi" in n:
        return "Redmi"
    return ""


ships = {str(s["id"]).strip(): s for s in Sh.load_shipments(SHIP_DIR)}

for fn, sheet, sku, sid in TARGETS:
    cur = os.path.join(BOOKS_DIR, fn)
    print("=" * 72)
    print(f"# {fn} | {sheet!r} | 补 {sku}（{sid}）")
    n = R.strip_pivot_records(cur, os.path.join(TMP, "stripped.xlsx"))
    src = os.path.join(TMP, "stripped.xlsx") if n else cur
    wb = R.load_workbook(src)
    real = next((x for x in wb.sheetnames if x.strip() == sheet.strip()), None)
    if real is None:
        print(f"  [跳过] 找不到 sheet {sheet!r}")
        wb.close()
        continue
    ws = wb[real]

    have, last = set(), 0
    for r in range(7, ws.max_row + 1):
        v = ws.cell(r, 1).value
        if v is not None and not str(v).startswith("="):
            have.add(str(v).strip())
            last = r
    if sku in have:
        print(f"  已存在 {sku}，跳过")
        wb.close()
        continue
    if sku not in ships[sid]["items"]:
        print(f"  [跳过] 货件 {sid} 里没有 {sku}")
        wb.close()
        continue

    nr = last + 1
    fc, lc = Sh.detect_sheet_cols(ws)
    name = ships[sid]["names"].get(sku, "")
    brand = brand_of(name)
    brand_col = None
    for c in range(60, min(ws.max_column + 1, 111)):
        v = ws.cell(last, c).value
        if isinstance(v, str) and v.strip() in BRANDS:
            brand_col = c
            break
    print(f"  末数据行={last} → 新行={nr} | 货件列 {gl(fc)}..{gl(lc)} | "
          f"品牌列={gl(brand_col) if brand_col else '未找到'} | 末行 A/B={ws.cell(last,1).value!r}/{str(ws.cell(last,2).value)[:30]!r}")
    print(f"  名称={name[:70]}")
    print(f"  待写：A={sku}  B={name[:40]}  品牌={brand}  该票数量={ships[sid]['items'][sku]}")

    maxc = ws.max_column
    copied = 0
    for c in range(1, maxc + 1):
        tc, nc = ws.cell(last, c), ws.cell(nr, c)
        v = tc.value
        if fc <= c <= lc:
            nc.value = None
        elif isinstance(v, ArrayFormula):
            txt = Translator(v.text, origin=tc.coordinate).translate_formula(nc.coordinate)
            nc.value = ArrayFormula(ref=nc.coordinate, text=txt)
            copied += 1
        elif isinstance(v, str) and v.startswith("="):
            nc.value = Translator(v, origin=tc.coordinate).translate_formula(nc.coordinate)
            copied += 1
        else:
            nc.value = v
        try:
            nc._style = copy(tc._style)
        except Exception:
            pass
    ws.cell(nr, 1).value = sku
    ws.cell(nr, 2).value = name
    if brand_col:
        ws.cell(nr, brand_col).value = brand
    print(f"  克隆 {copied} 个公式格 + 样式；A/B/品牌 已写")

    if not APPLY:
        print("  [dry-run] 未落盘")
        wb.close()
        continue
    saved = os.path.join(TMP, "saved_%s" % fn)
    wb.save(saved)
    wb.close()
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
    print(f"  [OK] 已落盘；备份→{BK}")
