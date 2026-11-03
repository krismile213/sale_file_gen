# -*- coding: utf-8 -*-
"""补 SKU 行（2026-10-14 货件重导后「未命中」的 4 票）——三张非欧盟分析表共 7 行。

为什么需要：非欧盟簿的 apply_shipments **不自动追行**，TSV 里出现表内没有的 SKU 时该票数量
被静默跳过（只在日志留「未命中 N」）。本次未命中：
  Interflowd北美 美国-分析表  ← FBA19RJBNWW8 (G597)、FBA19SYS2BC1 (G598)
  Interflowd澳大利亚-分析表   ← FBA15GJDCN8J (G594/G595/G596/G597)
  日本-分析表                ← FBA15GK1JFHC (G598)

做法（沿用 fix_g598_rows.py 的成熟套路）：整行克隆末数据行 → 公式列 Translator 平移
（含 ArrayFormula 单独分支）→ 样式 _style 照抄 → 货件列(I..差额列-1)清空 →
覆写 A=SKU、B=型号、品牌列。幂等：A 列已存在该 SKU 则跳过。
B 列写法跟随本表惯例（Int北美=长标题；Int澳/日本=短机型）；品牌列按本表惯例（中/英）。
管线：strip_pivot_records → openpyxl 编辑 → restore_parts → testzip → 备份 → 落盘。
用法：python fix_missing_sku_rows_20261014.py          # 干运行
     python fix_missing_sku_rows_20261014.py --apply  # 落盘
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
BK = os.path.join(DSR, ".workbuddy", "backup_missingrows_20261014")
TMP = os.path.join(DSR, ".workbuddy", "_tmp_missingrows")
OUT = os.path.join(DSR, ".workbuddy", "fix_missingrows_out.txt")
import runner as R
import shipments as Sh

APPLY = "--apply" in sys.argv

# (簿, 分析表, 模板行, [(SKU, 型号B, 品牌列号, 品牌值), ...])
TARGETS = [
    ("Interflowd北美加拿大库存销售表.xlsx", "美国-每日库存分析表", 79, [
        ("G597", "Ibywind for Sony Xperia 10 VIII 5G Screen Protector,2 Pack,9H Hardness Tempered Glass and Camera Lens Protector,with Alignment Frame,Anti-Bubble,Fingerprint", 88, "Sony"),
        ("G598", "Ibywind for Xiaomi Redmi 17 4G/5G Screen Protector,2 Pack,9H Hardness Tempered Glass and Camera Lens Protector,with Alignment Frame,Anti-Bubble,Fingerprint", 88, "Redmi"),
    ]),
    ("Interflowd澳大利亚库存销售表.xlsx", "澳大利亚-每日库存分析表", 76, [
        ("G594", "Xiaomi Redmi Note 17 4G/5G", 78, "小米"),
        ("G595", "Xiaomi Redmi Note 17 Pro 5G", 78, "小米"),
        ("G596", "Xiaomi Redmi Note 17 Pro Max 5G", 78, "小米"),
        ("G597", "Sony Xperia 10 VIII 5G", 78, "Sony"),
    ]),
    ("日本库存销售表.xlsx", "日本-每日库存分析表", 447, [
        ("G598", "Xiaomi Redmi 17 4G/5G", 78, "小米"),
    ]),
]

os.makedirs(BK, exist_ok=True)
os.makedirs(TMP, exist_ok=True)
_lines = []


def emit(s=''):
    _lines.append(str(s))


try:
    for fn, sheet, tmpl, items in TARGETS:
        cur = os.path.join(BOOKS_DIR, fn)
        emit("=" * 72)
        emit(f"# {fn} | {sheet!r} | 模板行 {tmpl} | 补 {[x[0] for x in items]}")
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
        emit(f"  货件列 {gl(fc)}..{gl(lc)} | 末数据行 {last}")

        maxc = ws.max_column
        nr = last
        for sku, bval, bcol, bname in items:
            if sku in have:
                emit(f"  {sku} 已存在，跳过")
                continue
            nr += 1
            copied = 0
            for c in range(1, maxc + 1):
                tc = ws.cell(tmpl, c)
                nc = ws.cell(nr, c)
                v = tc.value
                if fc <= c <= lc:                       # 货件列：清空（新行不该带旧数量）
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
            ws.cell(nr, bcol).value = bname
            have.add(sku)
            emit(f"  新行 {nr}: {gl(fc)}..{gl(lc)} 已清空；克隆 {copied} 个公式格；"
                 f"A={sku} B={bval[:46]} 品牌{gl(bcol)}={bname}")

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
