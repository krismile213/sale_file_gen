# -*- coding: utf-8 -*-
"""任务㉔：日销列窗口整体平移（2026-11-03 用户方案）
68 日列 9/24~11/30 → 10/25~12/31（+31 天），不删列、不插列、不动结构：
- 行5 日期值统一 +shift 天；SUMIFS 引本列 $5 → 公式文本零改动，数据自动跟日期走
- 条件格式/列宽/跨表引用不受影响
- 平移后重跑 analysis.sync_day_outline 折叠（快照日之后折叠）
侦察模式（默认）：打印每簿日列块范围、引用日列区间的左侧公式（漂移风险清单）
apply：写回日期 + 修正受影响公式 + 重算折叠 + safe_replace 落盘
"""
import os, sys, datetime, shutil

HERE = os.path.dirname(os.path.abspath(__file__))
DSR = os.path.dirname(os.path.dirname(os.path.dirname(HERE)))
PIPE = os.path.join(DSR, "pipeline")
for p in (PIPE, os.path.join(PIPE, "core")):
    if p not in sys.path:
        sys.path.insert(0, p)

import openpyxl
from openpyxl.utils import get_column_letter
from core import analysis as An

BOOKS_DIR = os.path.join(DSR, "1.每日销售数据表")
NEW_START = datetime.date(2026, 10, 25)
SNAP = datetime.date(2026, 11, 3)
BACKUP = os.path.join(DSR, ".workbuddy", "backup_shift_daycols_20261103")

BOOKS = sorted(
    f for f in os.listdir(BOOKS_DIR)
    if f.endswith(".xlsx") and not f.startswith("~$")
)

import re
COLREF = re.compile(r"\$?([A-Z]{1,3})\$?7")
WEEK_CN = ["周一", "周二", "周三", "周四", "周五", "周六", "周日"]


def shift_book(path, apply=False, new_start=None, snap=None, backup_dir=None):
    """把一本簿的分析表日列窗口整体平移到 new_start 起（保持列数/结构不变）。

    - 行5 日期 +shift 天、行1 星期同步重写；SUMIFS 引本列 $5 → 公式零改动
    - 折叠按 snap 次日起（analysis.sync_day_outline）
    - apply=True 时备份到 backup_dir（如给）+ safe_replace 落盘
    - 返回平移的日期格数（0=无需平移）
    """
    new_start = new_start or NEW_START
    snap = snap or SNAP
    wb = openpyxl.load_workbook(path, data_only=False)
    acts = 0
    for ws in wb.worksheets:
        if "每日库存分析表" not in ws.title:
            continue
        days = An.day_columns(ws, row=5, max_scan=250)
        if not days:
            print("    [跳过] %s: 行5 无日期列" % ws.title)
            continue
        cols = [c for c, _ in days]
        lo, hi = min(cols), max(cols)
        d_lo, d_hi = days[0][1], days[-1][1]
        shift = (new_start - d_lo).days
        if d_lo >= new_start:
            print("    [跳过] %s: 起点 %s 已 >= %s" % (ws.title, d_lo, new_start))
            continue
        # 左侧固定列公式漂移扫描：1..lo-1 列中引用日列区间的公式（应全部为日期感知型）
        risky = []
        for c in range(1, lo):
            for r in (4, 6, 7, 8):
                v = ws.cell(r, c).value
                if isinstance(v, str) and v.startswith("="):
                    for m in COLREF.finditer(v):
                        ref_ci = openpyxl.utils.column_index_from_string(m.group(1))
                        if lo <= ref_ci <= hi:
                            risky.append((get_column_letter(c) + str(r), v[:110]))
                            break
                if risky and r == 4:
                    break
        print("    %s: 日列 %s..%s (%d列 %s~%s) shift=+%d" %
              (ws.title, get_column_letter(lo), get_column_letter(hi), len(cols), d_lo, d_hi, shift))
        for coord, f in risky[:4]:
            print("      ⚠️漂移风险(需人工确认日期感知) %s: %s" % (coord, f))
        if len(risky) > 4:
            print("      … 另 %d 处" % (len(risky) - 4))
        if apply:
            for c, d in days:
                cell = ws.cell(5, c)
                cell.value = datetime.datetime.combine(d + datetime.timedelta(days=shift), datetime.time())
                # 行1 星期标注跟着平移（原非空的列才重写，保持原覆盖范围）
                w = ws.cell(1, c)
                if w.value not in (None, ""):
                    w.value = WEEK_CN[(d + datetime.timedelta(days=shift)).weekday()]
            # 同一内存内直接按快照日重折叠
            An.sync_day_outline(ws, snap)
            acts += len(cols)
    if not apply:
        wb.close()
        return 0
    if backup_dir:
        os.makedirs(backup_dir, exist_ok=True)
        bp = os.path.join(backup_dir, os.path.basename(path))
        if not os.path.exists(bp):
            shutil.copy2(path, bp)
            print("   已备份 → %s/" % os.path.basename(backup_dir))
    tmp = path + ".shifting.xlsx"
    wb.save(tmp)
    wb.close()
    # 沙箱兼容：先挪走旧文件再顶上
    bak = path + ".pre_shift.xlsx"
    if os.path.exists(bak):
        os.remove(bak)
    os.replace(path, bak)
    os.replace(tmp, path)
    return acts


# 兼容旧名
def analyze_book(path, apply=False):
    return shift_book(path, apply=apply)


def refold_book(path):
    """平移后按快照日重跑折叠（sync_day_outline 读行5 日期，会按新窗口折叠）。"""
    wb = openpyxl.load_workbook(path)
    for ws in wb.worksheets:
        if "每日库存分析表" not in ws.title:
            continue
        try:
            An.sync_day_outline(ws, SNAP)
        except Exception as e:
            print("    [折叠失败] %s: %r" % (ws.title, e))
    tmp = path + ".folding.xlsx"
    wb.save(tmp)
    wb.close()
    bak = path + ".prefold.xlsx"
    if os.path.exists(bak):
        os.remove(bak)
    os.replace(path, bak)
    os.replace(tmp, path)


if __name__ == "__main__":
    apply = "--apply" in sys.argv
    only = None
    for a in sys.argv[1:]:
        if a.startswith("--only="):
            only = a.split("=", 1)[1]
    for name in BOOKS:
        path = os.path.join(BOOKS_DIR, name)
        if not os.path.exists(path):
            print("[%s] 不存在，跳过" % name)
            continue
        if only and only not in name:
            continue
        print("== %s ==" % name)
        if apply and not os.path.exists(os.path.join(BACKUP, name)):
            os.makedirs(BACKUP, exist_ok=True)
            shutil.copy2(path, os.path.join(BACKUP, name))
            print("   已备份 → backup_shift_daycols_20261103/")
        n = analyze_book(path, apply=apply)
        if apply and n:
            print("   ✓ 平移 %d 个日期 + 星期同步 + 重折叠 完成" % n)
