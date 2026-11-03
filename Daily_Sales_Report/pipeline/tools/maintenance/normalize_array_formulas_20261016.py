# -*- coding: utf-8 -*-
"""把分析表里「含整列引用的数组公式」降级为普通公式（修复 VLOOKUP 全 0）。

问题（2026-10-16 用户报：阿联酋日销表 可售库存 全 0）：
  阿联酋-每日库存分析表 F/G 两列是**遗留 CSE 数组公式**：
      <f t="array" ref="F7">IFERROR(VLOOKUP(A:A,阿联酋在途库存!A:Y,COLUMN(...$W$3),0),0)</f>
  数组语义下 `A:A` 是「整列数组」，VLOOKUP 返回整个结果数组，单元格只取**第 1 个元素**
  （等于拿 A1 去查）→ #N/A → IFERROR → **0**。
  同一公式在 Interflowd阿联酋簿里是普通公式（有隐式交集，取本行 A7）→ 正常有值。

修复：**只去掉数组属性，公式文本一字不改**（与能正常工作的簿保持完全一致）。
  仅处理「公式文本里含整列引用（如 A:A / A:Y，列号后无数字）」的数组公式——
  这类才是会被数组语义破坏的；SUMPRODUCT 型数组公式（印度 F/G、罗马尼亚 G、美客多 E/G）
  引用都带行号，数组/普通结果相同，**不动**。

管线：strip_pivot_records → openpyxl 编辑 → restore_parts → testzip → 备份 → 落盘。
用法：python normalize_array_formulas_20261016.py          # 干跑
     python normalize_array_formulas_20261016.py --apply
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

import os, re, sys, shutil, zipfile, traceback
from openpyxl.worksheet.formula import ArrayFormula

BOOKS_DIR = os.path.join(DSR, "1.每日销售数据表")
BK = os.path.join(DSR, ".workbuddy", "backup_arrayfix_20261016")
TMP = os.path.join(DSR, ".workbuddy", "_tmp_arrayfix")
OUT = os.path.join(DSR, ".workbuddy", "normalize_array_out.txt")
import runner as R

APPLY = "--apply" in sys.argv
SHEET_SUFFIX = "每日库存分析表"
# 整列引用（列号后不跟数字，排除 $A$3:$A$30999 / I7:R7）
_COL = r"\$?[A-Z]{1,3}:\$?[A-Z]{1,3}(?![0-9A-Z])"
# 真正有害的两种形态：
#  ① 整列引用当「查找值」：VLOOKUP(A:A,…) / MATCH(A:A,…) —— 数组语义下退化成用第 1 行去查
#  ② 整列引用当「条件参数」：SUMIF(range, A:A, …) —— 同上
RE_HARM_LOOKUP = re.compile(r"(?:VLOOKUP|HLOOKUP|XLOOKUP|MATCH|LOOKUP)\s*\(\s*" + _COL, re.I)
RE_HARM_CRIT = re.compile(
    r"(?:SUMIF|SUMIFS|COUNTIF|COUNTIFS)\s*\(\s*[^,()]+,\s*" + _COL, re.I)


def harmful_wholecol(text):
    """返回第一个会被数组语义破坏的整列引用；没有则 None。

    只有「查找值 / 条件参数」位置的整列引用才有害。
    `SUM(库存明细!O:O)`、`SUMIF(range, A7, sum_range)` 里的整列引用是数组入参，
    数组/普通结果相同，**不算有害**。
    """
    for rx in (RE_HARM_LOOKUP, RE_HARM_CRIT):
        m = rx.search(text)
        if m:
            return m.group(0)
    return None

os.makedirs(BK, exist_ok=True)
os.makedirs(TMP, exist_ok=True)
_lines = []


def emit(s=""):
    _lines.append(str(s))
    print(s)


books = sorted(f for f in os.listdir(BOOKS_DIR) if f.lower().endswith(".xlsx"))
emit(f"扫描 {len(books)} 本簿（{'APPLY 落盘' if APPLY else 'dry-run'}）…\n")
stats = {"books": 0, "cells": 0, "skipped_arr": 0, "fail": 0}

for fn in books:
    cur = os.path.join(BOOKS_DIR, fn)
    emit("=" * 78)
    emit(f"# {fn}")
    try:
        n = R.strip_pivot_records(cur, os.path.join(TMP, "stripped.xlsx"))
        src = os.path.join(TMP, "stripped.xlsx") if n else cur
        wb = R.load_workbook(src)
        targets = [s for s in wb.sheetnames if s.strip().endswith(SHEET_SUFFIX)]
        if not targets:
            emit("  （无 *每日库存分析表，跳过）")
            wb.close()
            continue

        touched = False
        for sheet in targets:
            ws = wb[sheet]
            per_col = {}
            for row in ws.iter_rows():
                for c in row:
                    v = c.value
                    if not isinstance(v, ArrayFormula):
                        continue
                    text = str(v.text or "")
                    if harmful_wholecol(text) is None:
                        stats["skipped_arr"] += 1
                        continue
                    # 只去数组属性：写成普通公式，文本原样
                    c.value = "=" + text.lstrip("=")
                    touched = True
                    stats["cells"] += 1
                    per_col.setdefault(c.column_letter, [0, text])[0] += 1
            if per_col:
                emit(f"  [{sheet}]")
                for col, (cnt, sample) in sorted(per_col.items()):
                    emit(f"     {col} 列 {cnt} 格 → 普通公式   例: ={sample[:100]}")
            else:
                emit(f"  [{sheet}] 无需处理")

        if not touched:
            wb.close()
            continue

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
        stats["books"] += 1
        emit(f"  [OK] 已落盘 备份→{BK}")
    except Exception as e:
        stats["fail"] += 1
        emit(f"  [FAIL] {e!r}")
        emit(traceback.format_exc())

emit("")
emit(f"汇总：修正数组公式 {stats['cells']} 格（{stats['books']} 本簿落盘）；"
     f"保留（引用带行号、无害）{stats['skipped_arr']} 格；失败 {stats['fail']} 本")

with open(OUT, "w", encoding="utf-8") as f:
    f.write("\n".join(_lines) + "\n")
