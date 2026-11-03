# -*- coding: utf-8 -*-
"""清洗阿联酋系簿里「渠道不属于本簿」的历史脏行（2026-10-16 用户点头）。

背景：
  · `阿联酋库存销售表.xlsx` / `每日销售数据阿联酋`：1472 行 = ae 844 + **sa 623** + Non-Amazon 5。
    sa 行覆盖 2026-09-30 ~ 2026-10-13（10/14 起已停增 —— `expect_channel` 守卫生效），
    成因是 10/13 之前阿联酋簿被喂了沙特侧数据（下载/配置串档）。
  · `Interflowd阿联酋库存销售表.xlsx` / `每日销售数据阿联酋`：61 ae + **sa 1**。
  · 影响是**惰性**的（分析表日列 SUMIFS 带 `$D$5='Amazon.ae'` 渠道过滤，sa 行永不进日列/粉标），
    本次清洗属「洁癖级」，用户 2026-10-16 明确要求执行。

做法（逐本）：
  1. `strip_pivot_records` → openpyxl 载入（保公式）→ 定位 `每日销售数据*` sheet；
  2. 按**表头动态**找 `sales-channel` 列，删掉「以 `Amazon.` 开头且不在本簿 expect_channel 里」的行
     （只删整行；`Non-Amazon*` / 空 一律保留）；
  3. ⚠️ openpyxl 的 `delete_rows` **不平移公式**：删完后必须把该 sheet 的**逐行公式列**
     （阿联酋簿 `AE=LEFT(C..)` / `AF=IF(AND(E..,MID(G..)))`；Interflowd阿簿 `AI`/`AJ`）
     按新行号**整体重写**，否则会「引用错行」；
  4. 日销公式用的是**整列引用**（`$O:$O`/`$G:$G`/`$AE:$AE`/`$AF:$AF`），删行不影响；
  5. `restore_parts` → `testzip` → 备份 → 落盘。

用法：python clean_ae_sa_rows_20261016.py            # 干跑（只报告将删多少行）
     python clean_ae_sa_rows_20261016.py --apply
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
from openpyxl.utils import get_column_letter as gl
import runner as R

BOOKS_DIR = os.path.join(DSR, "1.每日销售数据表")
BK = os.path.join(DSR, ".workbuddy", "backup_ae_clean_20261016")
TMP = os.path.join(DSR, ".workbuddy", "_tmp_ae_clean")
OUT = os.path.join(DSR, ".workbuddy", "clean_ae_sa_out.txt")

APPLY = "--apply" in sys.argv
SHEET_PREFIX = "每日销售数据"
CH_COL_NAME = "sales-channel"
# 只清这两本（阿联酋系）；不碰其它簿
TARGETS = ["阿联酋库存销售表.xlsx", "Interflowd阿联酋库存销售表.xlsx"]

os.makedirs(BK, exist_ok=True)
os.makedirs(TMP, exist_ok=True)
_lines = []


def emit(s=""):
    _lines.append(str(s))
    print(s)


def allowed(book):
    """本簿允许的渠道集合（来自 config.WORKBOOKS[...].sales.expect_channel）。"""
    import config as C
    for cfg in C.WORKBOOKS:
        if cfg.get("file") == book:
            s = cfg.get("sales")
            if isinstance(s, list):
                s = s[-1]
            return set(s.get("expect_channel") or [])
    return set()


def find_channel_col(ws):
    """数据表头在第 2 行（第 1 行空）。返回 1-based 列号。"""
    for r in (1, 2, 3):
        for c in range(1, (ws.max_column or 1) + 1):
            v = ws.cell(r, c).value
            if isinstance(v, str) and v.strip() == CH_COL_NAME:
                return c, r
    return None, None


def runs(idx):
    """把行号列表压成连续区间 [(start, end), ...]。"""
    out = []
    for r in idx:
        if out and r == out[-1][1] + 1:
            out[-1][1] = r
        else:
            out.append([r, r])
    return [tuple(x) for x in out]


def template_formula_cols(ws, first=3):
    """找数据区的逐行公式列：返回 {col: 模板(以 first 行号写死)}。"""
    tpl = {}
    for c in range(1, (ws.max_column or 1) + 1):
        v = ws.cell(first, c).value
        if isinstance(v, str) and v.startswith("="):
            tpl[c] = v
    return tpl


def rebase(fml, old_row, new_row):
    """把公式里的行号 old_row 换成 new_row（列字母/函数名紧跟的数字才算）。"""
    if old_row == new_row:
        return fml
    return re.sub(r"(?<=[A-Z])%d(?![0-9A-Z])" % old_row, str(new_row), fml)


emit(f"目标 {len(TARGETS)} 本；备份 → {BK}\n")
stats = {"books": 0, "deleted": 0, "kept": 0, "fail": 0}

for fn in TARGETS:
    cur = os.path.join(BOOKS_DIR, fn)
    emit("=" * 78)
    emit(f"# {fn}")
    if not os.path.exists(cur):
        emit("  !! 文件不存在，跳过")
        stats["fail"] += 1
        continue
    try:
        n = R.strip_pivot_records(cur, os.path.join(TMP, "stripped.xlsx"))
        src = os.path.join(TMP, "stripped.xlsx") if n else cur
        wb = R.load_workbook(src)
        ok = set(allowed(fn))
        emit(f"  本簿 expect_channel = {sorted(ok) or '（无）'}")
        touched = False
        for sheet in [s for s in wb.sheetnames if s.startswith(SHEET_PREFIX)]:
            ws = wb[sheet]
            cc, hr = find_channel_col(ws)
            if not cc:
                emit(f"  [{sheet}] ⚠️ 找不到 {CH_COL_NAME} 列，跳过")
                continue
            first = hr + 1
            end = ws.max_row
            tpl = template_formula_cols(ws, first)
            emit(f"  [{sheet}] 表头行={hr} 渠道列={gl(cc)} 数据 {first}..{end}"
                 f"  逐行公式列={[gl(c) for c in sorted(tpl)]}")

            kill = []
            dist = {}
            for r in range(first, end + 1):
                ch = ws.cell(r, cc).value
                chs = ch.strip() if isinstance(ch, str) else ""
                dist[chs or "(空)"] = dist.get(chs or "(空)", 0) + 1
                if chs.startswith("Amazon.") and chs not in ok:
                    kill.append(r)
            emit(f"      渠道分布: {dist}")
            if not kill:
                emit("      ✓ 无脏行")
                continue

            emit(f"      将删除 {len(kill)} 行（示例: {kill[:3]} … {kill[-3:]}）")
            if not APPLY:
                emit("      [dry-run] 未落盘")
                touched = False
                continue

            # 从下往上按连续段删
            for a, b in reversed(runs(kill)):
                ws.delete_rows(a, b - a + 1)
            new_end = end - len(kill)
            # 重写逐行公式列（delete_rows 不平移公式）
            fixed = 0
            for r in range(first, new_end + 1):
                for c, fml in tpl.items():
                    ws.cell(r, c).value = rebase(fml, first, r)
                    fixed += 1
            emit(f"      ✓ 已删 {len(kill)} 行；重写公式 {fixed} 格；新末行={new_end}")
            touched = True
            stats["deleted"] += len(kill)
            stats["kept"] += (end - first + 1 - len(kill))

        if not touched:
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
        emit(f"  [OK] 已落盘  备份→{BK}")
    except Exception as e:
        stats["fail"] += 1
        emit(f"  [FAIL] {e!r}")
        emit(traceback.format_exc())

emit("")
emit(f"汇总：处理 {stats['books']} 本；删除脏行 {stats['deleted']}；保留 {stats['kept']}；失败 {stats['fail']}")
with open(OUT, "w", encoding="utf-8") as f:
    f.write("\n".join(_lines) + "\n")
