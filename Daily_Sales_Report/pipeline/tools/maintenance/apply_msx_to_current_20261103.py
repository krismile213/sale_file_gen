# -*- coding: utf-8 -*-
"""分析表上月封存列一次性对齐月销总表（任务㉒ 的一次性执行，2026-11-03）

背景
----
runner.apply_month_summary 的上月源已改为优先月销总表（任务㉒），日常跑批
此后自动保持上报口径。但本月（2026-10）的封存值是在改动前冻结的（resolve
重算口径，与月销表差跨月单 + 33 件漏记修正），需要一次性把当期已封存的
两处落点刷成月销总表值：

  · 当前簿  1.每日销售数据表/{fname}          —— 10 月汇总列（用户上报落点）
  · 封存簿  2026年每日数据/2026年10月/2026年10月-{fname}

范围：只碰「一个站点只有一本簿」的 3 本（欧盟英国=欧洲+英国、日本、印度）。
北美（美/加/墨）、澳洲、中东（沙特/阿联酋）的 Interflowd 孪生簿持有差额
（实测跨簿相加=月销表：澳洲 1656+276=1932 分毫不差、沙特 1280+17=1297=归因
日销侧），单独覆盖主簿会重复计数——保持不动，由跨簿合计对齐。
站点表（意大利/法国…）不动——上报口径只看分析表。

每张「XX-每日库存分析表」：
  10月列 ← _month_summary_from_msx(2026,10,XX)   （freeze_month 覆盖写）
  上月日均列联动刷新（fill_prev_avg，公式=IF(封存列="","",ROUND(封存列/天数,2))）

用法
----
    python pipeline/tools/maintenance/apply_msx_to_current_20261103.py            # 预演
    python pipeline/tools/maintenance/apply_msx_to_current_20261103.py --apply    # 实跑
"""
import os
import sys
import argparse

HERE = os.path.dirname(os.path.abspath(__file__))   # pipeline/tools/maintenance
PIPE = os.path.dirname(os.path.dirname(HERE))       # pipeline
DSR = os.path.dirname(PIPE)                         # Daily_Sales_Report
for p in (PIPE, os.path.join(PIPE, "core")):
    if p not in sys.path:
        sys.path.insert(0, p)

import openpyxl
from core import config as C
from core import runner
from core import analysis as An

# ⚠️ 范围收紧（2026-11-03 实测）：只有「一个站点一本簿」的站点才能直取月销总表。
# 北美/澳洲/中东的 Interflowd 孪生簿持有差额（澳洲 1656+276=1932 与月销表分毫不差、
# 沙特 1280+17=1297=归因日销侧），单独覆盖主簿会与孪生簿重复计数——保持不动。
BOOKS = [
    "欧盟英国库存销售表.xlsx",       # 欧洲 + 英国 两张分析表
    "日本库存销售表.xlsx",
    "印度库存销售表.xlsx",
]
PM_Y, PM_M = 2026, 10
BKDIR = os.path.join(DSR, ".workbuddy", "backup_monthsale_fix_20261103")


def safe_replace(tmp, tgt):
    os.makedirs(BKDIR, exist_ok=True)
    bak = os.path.join(BKDIR, os.path.basename(tgt) + ".msx对齐前备份.xlsx")
    if os.path.exists(bak):
        os.remove(bak)
    os.replace(tgt, bak)
    try:
        os.replace(tmp, tgt)
    except Exception:
        os.replace(bak, tgt)
        raise
    return bak


def _static_sum(ws, col, first_row=7, max_row=3200):
    """读某列现有静态值合计（改前对照）。同 status_sync._col_static_sum。"""
    s = 0.0
    for r in range(first_row, first_row + max_row):
        if ws.cell(r, 1).value is None:
            break
        v = ws.cell(r, col).value
        if isinstance(v, (int, float)):
            s += v
    return s


def patch_book(path, dry):
    """一本簿：所有「分析」sheet 的 10 月列 ← 月销总表值。返回改动描述列表。"""
    name = os.path.basename(path)
    if not os.path.exists(path):
        return [(name, "不存在，跳过", None, None)]
    wb = openpyxl.load_workbook(path)
    changes = []
    touched = False
    for sn in [s for s in wb.sheetnames if "分析" in s]:
        ws = wb[sn]
        try:
            sc = An.scan_analysis(ws)
        except Exception as e:
            changes.append((name, sn, None, "扫描失败 %r" % (e,)))
            continue
        if not sc:
            changes.append((name, sn, None, "无法解析口径"))
            continue
        col = sc["month_cols"].get(PM_M)
        if col is None:
            changes.append((name, sn, None, "行5无%d月标签列" % PM_M))
            continue
        site = sn.split("-")[0].strip()
        site = {"澳大利亚": "澳洲", "沙特阿拉伯": "沙特"}.get(site, site)
        ms = runner._month_summary_from_msx(PM_Y, PM_M, site)
        if not ms:
            changes.append((name, sn, None, "月销总表无「%s」块，跳过" % site))
            continue
        before = _static_sum(ws, col)
        after = sum(ms.values())
        changes.append((name, sn,
                        openpyxl.utils.get_column_letter(col),
                        "%8.0f → %8.0f (Δ%+.0f)" % (before, after, after - before)))
        if not dry:
            An.freeze_month(ws, col, ms)
            import calendar as _cal
            days = _cal.monthrange(PM_Y, PM_M)[1]
            avg = {k: (v / days) for k, v in ms.items()}
            An.fill_prev_avg(ws, avg, month_col=col, days=days)
            touched = True
    if touched and not dry:
        tmp = path + ".msx.new.xlsx"
        wb.save(tmp)
        wb.close()
        try:
            runner.restore_parts(path, tmp)
        except Exception as e:
            print("   [!] restore_parts 跳过：%r" % (e,))
        safe_replace(tmp, path)
        print("   %s 已落盘" % name)
    else:
        wb.close()
    return changes


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--apply", action="store_true")
    a = ap.parse_args()
    dry = not a.apply
    mode = "预演" if dry else "实跑"
    print("=== 分析表 2026-10 封存列对齐月销总表（%s）===" % mode)
    ok = err = 0
    for fname in BOOKS:
        for tag, base in (("当前簿", os.path.join(C.WORK_DIR, fname)),
                          ("封存簿", os.path.join(C.WORK_DIR,
                                                  "%d年每日数据" % PM_Y,
                                                  "%d年%d月" % (PM_Y, PM_M),
                                                  "%d年%d月-%s" % (PM_Y, PM_M, fname)))):
            for bname, sn, colL, info in patch_book(base, dry):
                if info is None:
                    print("  [%s] %s / %s: %s" % (tag, bname, sn or "-", info))
                    err += 1
                elif info.startswith(("不存在", "月销总表无", "行5无", "无法解析", "扫描失败")):
                    print("  [%s] %s / %s: %s" % (tag, bname, sn or "-", info))
                else:
                    print("  [%s] %-34s %-22s %s列 %s" % (tag, bname, sn, colL, info))
                    ok += 1
    print("\n对齐 %d 张分析表（%s）" % (ok, mode))
    return 0


if __name__ == "__main__":
    sys.exit(main())
