# -*- coding: utf-8 -*-
"""正确性修复（2026-10-13）：
A) E 预留列 = 真源 afn-reserved（26 列库存表 M 列）：把仍是旧源（可售库存 col7 / 空）的 7 张分析表改成
   与已修好的 7 张完全一致的写法 `=IFERROR(SUMIF('<src>'!A:A,A:A,'<src>'!M:M),0)`。
B) Interflowd北美加拿大 的「美国/加拿大预留库存记录」：清掉 69506 个 #REF! 假公式；
   B 列 = 与兄弟表「XX库存记录」同源的 SKU 列表（字面值）；A 列保留 VLOOKUP(型号)；
   日列只在快照日（2026-10-13 → 第 13 天 → O 列）写入**真实 afn-reserved 快照字面值**，其余日列留空
   （遵循用户「哪天导出写哪天 / 清假公式 / 用报告补真值」的硬要求）。
保真：strip_pivot_records → openpyxl 编辑 → restore_parts（透视部件按原字节覆盖）。
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
import os, sys, shutil, zipfile
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

# [normalize] DSR 由上方引导块赋值
CUR = os.path.join(DSR, "1.每日销售数据表")
BK  = os.path.join(DSR, ".workbuddy", "backup_fix_e_ref_20261013")
TMP = os.path.join(DSR, ".workbuddy", "_tmp")
pass  # [normalize] 路径已由引导块设置
import runner as R  # strip_pivot_records / restore_parts

SNAP_DAY = 13          # 2026-10-13
os.makedirs(BK, exist_ok=True)
os.makedirs(TMP, exist_ok=True)

# ---------------- A) E 预留列 ----------------
E_FIX = [
    ("Interflowd澳大利亚库存销售表.xlsx", "澳大利亚-每日库存分析表", "在途库存"),
    ("澳大利亚库存销售表.xlsx",         "澳大利亚-每日库存分析表", "在途库存"),
    ("沙特库存销售表.xlsx",             "沙特阿拉伯-每日库存分析表", "沙特阿拉伯在途库存"),
    ("印度库存销售表.xlsx",             "印度-每日库存分析表", "在途库存"),
    ("阿联酋库存销售表.xlsx",           "阿联酋-每日库存分析表", "阿联酋在途库存"),
    ("Interflowd阿联酋库存销售表.xlsx", "阿联酋-每日库存分析表", "阿联酋在途库存"),
    ("北美加拿大库存销售表.xlsx",       "墨西哥-每日库存分析表 ", "墨西哥可售+在途库存"),
]

def open_edit(path, tmpname):
    """原件 → strip → openpyxl 打开（可编辑）。返回 (wb, tmp_stripped, tmp_saved)"""
    tmp1 = os.path.join(TMP, tmpname + ".stripped.xlsx")
    tmp2 = os.path.join(TMP, tmpname + ".saved.xlsx")
    n = R.strip_pivot_records(path, tmp1)
    wb = load_workbook(tmp1 if n else path)
    return wb, tmp1, tmp2, n

def commit(wb, orig, tmp2):
    wb.save(tmp2)
    wb.close()
    R.restore_parts(orig, tmp2)
    with zipfile.ZipFile(tmp2) as z:
        bad = z.testzip()
    if bad:
        raise RuntimeError(f"zip 损坏: {bad}")
    shutil.copy2(orig, os.path.join(BK, os.path.basename(orig)))
    try:
        os.replace(tmp2, orig)
    except PermissionError:
        shutil.copyfile(tmp2, orig)
    return os.path.getsize(orig)

print("="*100)
print("### A) E 预留列 → afn-reserved（M 列）")
for fn, an, src in E_FIX:
    p = os.path.join(CUR, fn)
    if not os.path.exists(p):
        print(f"  跳过（缺）{fn}"); continue
    wb, t1, t2, npiv = open_edit(p, "A_" + fn)
    ws = wb[an]
    # 源表 M 列合计（真值，供核验）
    sws = wb[src]
    msum = 0.0
    for r in sws.iter_rows(min_row=3, values_only=True):
        if len(r) > 12 and isinstance(r[12], (int, float)):
            msum += r[12]
    f = f"=IFERROR(SUMIF('{src}'!A:A,A:A,'{src}'!M:M),0)"
    nrow = 0
    for r in range(7, ws.max_row + 1):
        a = ws.cell(row=r, column=1).value
        if a not in (None, ""):
            ws.cell(row=r, column=5).value = f
            nrow += 1
    size = commit(wb, p, t2)
    print(f"  ✓ {fn} [{an}] 写 E × {nrow} 行；源 M 列合计={int(msum)}；size={size/1024:.0f}KB；strip_pivot={npiv}")

# ---------------- B) 预留库存记录 #REF! ----------------
print("="*100)
print("### B) 预留库存记录（#REF!）")
REF_FIX = [
    ("Interflowd北美加拿大库存销售表.xlsx", "美国预留库存记录",   "美国库存记录",   "美国可售+在途库存", "美国-每日库存分析表"),
    ("Interflowd北美加拿大库存销售表.xlsx", "加拿大预留库存记录", "加拿大库存记录", "加拿大可售+在途库存", "加拿大-每日库存分析表"),
]
fn = "Interflowd北美加拿大库存销售表.xlsx"
p = os.path.join(CUR, fn)
wb, t1, t2, npiv = open_edit(p, "B_" + fn)

today_col = 3 + (SNAP_DAY - 1)   # C 列 = 当月 1 日 → 第 13 天 = O(15)

for book_fn, ref_sheet, sib_sheet, src, an in REF_FIX:
    rws = wb[ref_sheet]
    sws = wb[sib_sheet]
    # 兄弟表 SKU 列表（字面值）
    skus = []
    for r in sws.iter_rows(min_row=5, max_col=2, values_only=True):
        v = r[1] if len(r) > 1 else None
        if v not in (None, ""):
            skus.append(str(v).strip())
    # 源表 M 列快照（按 SKU 汇总）
    snap = {}
    for r in wb[src].iter_rows(min_row=3, values_only=True):
        if not r or r[0] in (None, ""):
            continue
        sku = str(r[0]).strip()
        if len(r) > 12 and isinstance(r[12], (int, float)):
            snap[sku] = snap.get(sku, 0) + r[12]
    # 清理：A/B/C..BR 全清，再按 SKU 列表回填
    maxr = rws.max_row
    for r in range(5, maxr + 1):
        rws.cell(row=r, column=1).value = None
        rws.cell(row=r, column=2).value = None
        for c in range(3, rws.max_column + 1):
            rws.cell(row=r, column=c).value = None
    filled = 0; nz = 0
    for i, sku in enumerate(skus):
        r = 5 + i
        rws.cell(row=r, column=2).value = sku
        rws.cell(row=r, column=1).value = f"=VLOOKUP(B:B,'{an}'!A:D,4,0)"
        v = snap.get(sku, 0)
        rws.cell(row=r, column=today_col).value = int(v) if v == int(v) else v
        filled += 1
        if v: nz += 1
    print(f"  ✓ [{ref_sheet}] SKU={len(skus)} 行；快照日列={get_column_letter(today_col)}(第{SNAP_DAY}天)；"
          f"非零 {nz} 个；源 {src} 覆盖 {len(snap)} SKU")

size = commit(wb, p, t2)
print(f"  ✓ {fn} 保存 size={size/1024:.0f}KB；strip_pivot={npiv}")
print("\n完成。")
