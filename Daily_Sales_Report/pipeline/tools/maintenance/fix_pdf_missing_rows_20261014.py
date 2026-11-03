# -*- coding: utf-8 -*-
"""补 SKU 行（2026-10-14 PDF 货件「未命中」）——美客多墨西哥表 12 行 + 罗马尼亚表 12 行。

为什么需要：非欧盟簿的 apply_shipments **不自动追行**，PDF 货件里出现表内没有的 SKU 时该票
数量被静默跳过（日志留「未命中 N」）。本次未命中（含件数）：
  美客多 墨西哥-每日库存分析表 ← 77739558（G592×37 G593×37 G594×2 G595×2 G596×5
                                     K013-016..020 各×2）、78704820（G597×2）、79402478（G598×2）
  罗马尼亚-每日库存分析表      ← eMAG 285062-333046（G564/565/570/576..584 各×2）

做法（沿用 fix_missing_sku_rows_20261014.py）：整行克隆**末数据行** → 公式列 Translator 平移
（含 ArrayFormula 单独分支）→ 样式 _style 照抄 → 货件列(I..差额列-1)清空 → 覆写 A=SKU、B=型号、
品牌列 BZ。幂等：A 列已存在该 SKU 则跳过。
型号取自 PDF 品名（ML「…Para <型号>」/ eMAG「…pentru <型号>」）并规范化大小写；品牌按两簿
既有 BZ 惯例（Xiaomi/Redmi→小米、Apple→苹果、Samsung→三星…）。**B/品牌为推断值，可人工校正。**
管线：strip_pivot_records → openpyxl 编辑 → restore_parts → testzip → 备份 → 落盘。
用法：python fix_pdf_missing_rows_20261014.py          # 干运行
     python fix_pdf_missing_rows_20261014.py --apply  # 落盘
"""
# --- pipeline 路径引导 ---
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

BOOKS_DIR = os.path.join(DSR, "1.每日销售数据表")
BK = os.path.join(DSR, ".workbuddy", "backup_missingrows_pdf_20261014")
TMP = os.path.join(DSR, ".workbuddy", "_tmp_missingrows_pdf")
OUT = os.path.join(DSR, ".workbuddy", "fix_pdf_missingrows_out.txt")
import runner as R
import shipments as Sh

APPLY = "--apply" in sys.argv

BRAND_COL = 78      # BZ = 品牌

# (簿, 分析表, [(SKU, 型号B, 品牌), ...])
TARGETS = [
    ("美客多库存销售表.xlsx", "墨西哥-每日库存分析表", [
        ("G592", "Poco F9 Pro/Redmi K100 Pro 5G", "Poco"),
        ("G593", "Poco F9 Ultra/K100 Pro Max", "Poco"),
        ("G594", "Xiaomi Redmi Note 17", "小米"),
        ("G595", "Redmi Note 17 Pro 5G", "小米"),
        ("G596", "Redmi Note 17 Pro Max", "小米"),
        ("K013-016", "iPhone 17 Pro Max/18 Pro Max 手机壳 Aria", "苹果"),
        ("K013-017", "iPhone 17 Pro Max/18 Pro Max 手机壳 Ripple", "苹果"),
        ("K013-018", "iPhone 17 Pro Max/18 Pro Max 手机壳 Strum", "苹果"),
        ("K013-019", "iPhone 17 Pro Max/18 Pro Max 手机壳 Bloop", "苹果"),
        ("K013-020", "iPhone 17 Pro Max/18 Pro Max 手机壳 Cadence", "苹果"),
        ("G597", "Sony Xperia 10 VIII 5G", "Sony"),
        ("G598", "Xiaomi Redmi 17 4G/5G", "小米"),
    ]),
    ("罗马尼亚库存销售表.xlsx", "罗马尼亚-每日库存分析表", [
        ("G564", "Apple iPhone 18 Pro 5G", "苹果"),
        ("G565", "Apple iPhone 18 Pro Max 5G", "苹果"),
        ("G570", "Samsung Galaxy S26 FE 5G", "三星"),
        ("G576", "Oppo Reno 16 5G", "Oppo"),
        ("G577", "Oppo Reno 16 Pro 5G", "Oppo"),
        ("G578", "Oppo Reno 16c/16 F 5G", "Oppo"),
        ("G579", "Huawei Pura 80 Pro 4G", "华为"),
        ("G580", "Huawei Pura 80 Ultra 4G", "华为"),
        ("G581", "Nothing Phone (4b) 5G", "Nothing"),
        ("G582", "Samsung Galaxy Z Fold8 5G", "三星"),
        ("G583", "Samsung Galaxy Z Fold8 Ultra 5G", "三星"),
        ("G584", "Realme Narzo 100x 5G", "Realme"),
    ]),
]

os.makedirs(BK, exist_ok=True)
os.makedirs(TMP, exist_ok=True)
_lines = []


def emit(s=''):
    _lines.append(str(s))


try:
    for fn, sheet, items in TARGETS:
        cur = os.path.join(BOOKS_DIR, fn)
        emit("=" * 72)
        emit(f"# {fn} | {sheet!r} | 补 {len(items)} 行")
        n = R.strip_pivot_records(cur, os.path.join(TMP, "stripped.xlsx"))
        src = os.path.join(TMP, "stripped.xlsx") if n else cur
        wb = R.load_workbook(src)
        real = next((x for x in wb.sheetnames if x.strip() == sheet.strip()), None)
        if real is None:
            raise RuntimeError(f"找不到 sheet {sheet!r}")
        ws = wb[real]
        fc, lc = Sh.detect_sheet_cols(ws)
        have, last = set(), 0
        for r in range(7, ws.max_row + 1):
            v = ws.cell(r, 1).value
            if v is not None and not str(v).startswith('='):
                have.add(str(v).strip())
                last = r
        tmpl = last
        emit(f"  货件列 {gl(fc)}..{gl(lc)} | 末数据行/模板行 {tmpl}")

        maxc = ws.max_column
        nr = last
        for sku, bval, bname in items:
            if sku in have:
                emit(f"  {sku} 已存在，跳过")
                continue
            nr += 1
            copied = 0
            for c in range(1, maxc + 1):
                tc = ws.cell(tmpl, c)
                nc = ws.cell(nr, c)
                v = tc.value
                if fc <= c <= lc:                       # 货件列：清空
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
                try:
                    nc._style = copy(tc._style)
                except Exception:
                    pass
            ws.cell(nr, 1).value = sku
            ws.cell(nr, 2).value = bval
            ws.cell(nr, BRAND_COL).value = bname
            have.add(sku)
            emit(f"  新行 {nr}: 克隆 {copied} 公式格；A={sku} B={bval} 品牌={bname}")

        if not APPLY:
            emit("  [dry-run] 未落盘")
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
        emit(f"  [OK] 落盘（备份→{BK}）")
except Exception as e:
    emit("[FAIL] " + repr(e))
    emit(traceback.format_exc())

with open(OUT, 'w', encoding='utf-8') as f:
    f.write('\n'.join(_lines) + '\n')
print('done, apply=%s' % APPLY)
