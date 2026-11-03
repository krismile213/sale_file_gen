# -*- coding: utf-8 -*-
"""把 202611/快照日期/20261102 的日销报告归位到 202610/快照日期/20261102。

背景（2026-11-02）：
  用户按自然月新建了 202611 目录，但 38 份报告里 37 份是 10 月的（订单报告全叫
  202610）。按用户拍板「数据归属周期=订单报告月」，这批应归 202610。

只移动**报告文件**（.csv/.txt/.xlsx），不动子目录：
  · 202611/20261102/东南亚    —— 42 份东南亚原始订单报告（月度交付件材料）
  · 202611/20261102/result    —— 东南亚/北美销量表成品
  · 202611/20261102/京东国内202609订单报告 —— 京东目录
这些都不属于日销表输入，留在原处。

库龄报告单独处理：202611/库龄/* → 202610/库龄/20261102/（_autodetect_age 递归扫、
按 mtime 升序后者覆盖前者，所以旧日期的库龄不会盖掉新的）。

用法：
  python pipeline/tools/maintenance/move_snap_1030_20261102.py            # 只打印计划
  python pipeline/tools/maintenance/move_snap_1030_20261102.py --apply    # 实跑
"""
import os
import sys
import shutil
import argparse

BASE = r"C:\Users\dell\Desktop\work-main\sale_file_gen\trunk\data\数据归属周期"
SRC_DAY = os.path.join(BASE, "202611", "快照日期", "20261102")
DST_DAY = os.path.join(BASE, "202610", "快照日期", "20261102")
SRC_AGE = os.path.join(BASE, "202611", "库龄")
DST_AGE = os.path.join(BASE, "202610", "库龄", "20261102")

REPORT_EXT = (".csv", ".txt", ".xlsx", ".xls")


def plan():
    moves, skips = [], []
    if not os.path.isdir(SRC_DAY):
        raise SystemExit("源目录不存在: " + SRC_DAY)
    for name in sorted(os.listdir(SRC_DAY)):
        src = os.path.join(SRC_DAY, name)
        if os.path.isdir(src):
            skips.append((name, "子目录，保留原地"))
            continue
        if name.startswith("~$"):
            skips.append((name, "Excel 临时锁文件，跳过"))
            continue
        if not name.lower().endswith(REPORT_EXT):
            skips.append((name, "非报告扩展名，跳过"))
            continue
        dst = os.path.join(DST_DAY, name)
        if os.path.exists(dst):
            skips.append((name, "目标已存在同名文件，跳过"))
            continue
        moves.append((src, dst))
    return moves, skips


def plan_age():
    moves = []
    if not os.path.isdir(SRC_AGE):
        return moves
    for root, _dirs, files in os.walk(SRC_AGE):
        for f in sorted(files):
            src = os.path.join(root, f)
            rel = os.path.relpath(src, SRC_AGE)
            dst = os.path.join(DST_AGE, rel)
            if os.path.exists(dst):
                continue
            moves.append((src, dst))
    return moves


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--apply", action="store_true")
    args = ap.parse_args()

    moves, skips = plan()
    age_moves = plan_age()

    print("=== 快照报告 %d 个 ===" % len(moves))
    for s, d in moves:
        print("   %-60s -> %s" % (os.path.basename(s), os.path.dirname(d)))
    print("=== 库龄报告 %d 个 -> %s ===" % (len(age_moves), DST_AGE))
    for s, _d in age_moves:
        print("   ", os.path.relpath(s, SRC_AGE))
    print("=== 原地保留 %d 项 ===" % len(skips))
    for n, why in skips:
        print("   %-50s %s" % (n, why))

    if not args.apply:
        print("\n[dry] 未移动。加 --apply 执行。")
        return 0

    os.makedirs(DST_DAY, exist_ok=True)
    ok = 0
    for s, d in moves + age_moves:
        os.makedirs(os.path.dirname(d), exist_ok=True)
        try:
            shutil.move(s, d)
            ok += 1
        except Exception as e:
            print("  !! 移动失败 %s: %s" % (os.path.basename(s), e))
    print("\n[完成] 已移动 %d / %d 个文件" % (ok, len(moves) + len(age_moves)))
    return 0 if ok == len(moves) + len(age_moves) else 1


if __name__ == "__main__":
    sys.exit(main())
