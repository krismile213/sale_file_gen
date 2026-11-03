# -*- coding: utf-8 -*-
"""修复日列「组合」导致的列宽突变（2026-10-14）。

问题：`analysis.sync_day_outline` 按列设置 outline 时，对「区间列定义」内的列
（如 `<col min=35 max=42 width=4.75/>` 里的 AJ..AQ）会新建默认 width=13 的
ColumnDimension，与残留区间定义重叠 → Excel 取单列定义 → 这些列宽由 4.75 变 13，
表现为「表格宽度不一样了」。

修复：用修好的 `analysis.sync_day_outline`（内部先 `_materialize_col_ranges` 把区间
拆成单列、逐列继承原宽度）重跑一遍列分组。**只动列定义（宽度/分组），不改任何数据。**

用法：
  python fix_col_widths_20261014.py            # 干跑：只报每簿将修正的列
  python fix_col_widths_20261014.py --apply    # 落盘（先自动备份）
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
import os
import openpyxl
from openpyxl.utils import get_column_letter as gl

import config as C
import analysis as An
import runner as R

APPLY = "--apply" in _sys.argv


def eff_width(holder, c):
    """列的**有效**宽度：优先单列定义，否则取覆盖它的区间定义（不新建列定义）。"""
    exact = rng = None
    for cd in holder.values():
        mn, mx = cd.min or 0, cd.max or 0
        if cd.width is None:
            continue
        if mn == mx == c:
            exact = cd.width
        elif mn <= c <= mx and rng is None:
            rng = cd.width
    return exact if exact is not None else rng


snap_dir, snap_day = C.latest_snapshot()
snap_date = f"{snap_day[:4]}-{snap_day[4:6]}-{snap_day[6:]}"
print(f"快照日 {snap_date}  |  模式 {'落盘' if APPLY else '干跑'}")

files = [w["file"] for w in C.WORKBOOKS if not w.get("skip")]
total_fix = 0
for f in files:
    path = os.path.join(C.WORK_DIR, f)
    if not os.path.exists(path):
        print(f"\n==== {f}: 不存在，跳过")
        continue
    print(f"\n==== {f}")
    stripped = path + ".pivotstrip.xlsx"
    n = R.strip_pivot_records(path, stripped)
    load_path = stripped if n else path
    wb = openpyxl.load_workbook(load_path, data_only=False)
    if n and os.path.exists(stripped):
        try:
            os.remove(stripped)
        except OSError:
            pass
    book_fix = 0
    for sn in list(wb.sheetnames):
        if "分析" not in sn:
            continue
        ws = wb[sn]
        holder = ws.column_dimensions
        cols = [c for c, _ in An.day_columns(ws)]
        before = {c: eff_width(holder, c) for c in cols}
        hid, btn = An.sync_day_outline(ws, snap_date)
        after = {c: eff_width(holder, c) for c in cols}
        diffs = [(gl(c), before[c], after[c]) for c in cols if before[c] != after[c]]
        note = f"宽度修正 {len(diffs)} 列 " + str(diffs[:6]) if diffs else "宽度无需修正"
        print(f"    [{sn}] 折叠 {len(hid)} 列（按钮 {gl(btn) if btn else '-'}）｜{note}")
        book_fix += len(diffs)
    total_fix += book_fix
    if APPLY and book_fix:
        try:
            R.backup([f], "colwidth_20261014")
            tmp = path + ".new.xlsx"
            wb.save(tmp)
            wb.close()
            R.restore_parts(path, tmp)
            try:
                os.replace(tmp, path)
            except OSError:
                import shutil
                shutil.copyfile(tmp, path)
                os.remove(tmp)
            print(f"    → 已落盘（修正 {book_fix} 列宽）")
        except Exception as e:
            print(f"    !! 落盘失败（多半文件被 WPS/Excel 占用）：{e!r}")
            try:
                wb.close()
            except Exception:
                pass
    else:
        wb.close()

print(f"\n合计修正 {total_fix} 列宽。")
if not APPLY:
    print("（干跑，未改盘；加 --apply 落盘）")
