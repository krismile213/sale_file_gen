#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
工具-更新商品资料得分列.py
=====================================
用途：按商品资料.xlsx 第一行的评级规则，用「成品月销售总表」的月销量自动生成/刷新
      「N月销量得分」列，替换原来每月人工维护的动作（solo 数据部：一人也能周更）。

评级规则（摘自 商品资料.xlsx A1 注释，2026-10-17 按原样实现）：
  A: 1个月销量500套及以上
  B: 3个月销量500套及以上
  C: 6个月销量250套及以上
  D: 6个月销量200-250套
  E: 6个月销量200套以下
  X类为新品（首销在评分月前6个月窗口内）：
    XA: sku日均≥16.7  XB: 16.7>日均≥5.5  XC: 5.5>日均≥1.38
    XD: 1.38>日均≥1.11 XE: 1.11>日均      X: 暂无销售数据
  停售标准与「手机壳不参与评级」由人工把握，本工具不自动改「状态」列。

口径（2026-10-17 用 202609 成品反推验证：老品 A-E 复刻 346/350 = 98.9%，
4 个不符均为贴线个案，差异 1~11 件，判定数据源与人工维护时点略有出入）：
  - 销量取成品月销售总表「<年>年」sheet 总计组（全球口径，含东南亚/京东/eMAG/TEMU/美客多）；
  - 「1/3/6个月」窗口含评分月当月；评分月之前不足的月份从上一 sheet（去年）补；
  - 分类继承：上期为 A-E 的老品继续按 A-E；上期 X 类新品满 6 个月历史才「毕业」转 A-E；
    上期空值（停售/不销售/未评级）一律不碰；成品里没有的 SKU → X。

用法：
  # 1) 自验：复刻上期得分列，与现有列逐行比对（不写任何文件）
  python 工具-更新商品资料得分列.py --src 202609 --score-month 8 --verify

  # 2) 生成：为 10 月归属月的跑准备「9月销量得分」列（先备份再写回）
  python 工具-更新商品资料得分列.py --src 202609 --score-month 9 --apply

  # 3)  dry-run：只打印将发生的变更，不写文件
  python 工具-更新商品资料得分列.py --src 202609 --score-month 9

依赖：openpyxl（与主脚本同环境）。
"""
import argparse
import os
import re
import shutil
import sys
from copy import copy
from datetime import datetime

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

TRUNK = os.path.dirname(os.path.abspath(__file__))
PROD_PATH = os.path.join(TRUNK, "商品资料.xlsx")
PROD_SHEET = "商品资料"
PROD_HEADER_ROW = 2
PROD_DATA_START = 3


def n0(v):
    """成品月格 → 数值。停售当月 0 写的是文本 \"00\"；公式无缓存为 None。"""
    if v is None:
        return 0
    if isinstance(v, str):
        v = v.strip()
        if v in ("", "00"):
            return 0
        try:
            return float(v)
        except ValueError:
            return 0
    try:
        return float(v)
    except TypeError:
        return 0


def find_total_first_col(ws):
    """行1 找值为「总计」的列 = 总计组 1 月列。"""
    for c in range(1, ws.max_column + 1):
        if str(ws.cell(1, c).value).strip() == "总计":
            return c
    return None


def load_month_sale(src):
    """读成品月销售总表 → (sku -> 24 个月销量序列 [去年1月..今年12月], 当年年份)。
    序列下标 i 对应月份：i//12 = 偏移年，i%12+1 = 月。"""
    path = os.path.join(TRUNK, "result", "数据归属周期", src, "月销售总表.xlsx")
    if not os.path.exists(path):
        sys.exit(f"❌ 找不到成品月销总表：{path}（先跑主脚本出成品，或用 --src 指定已有归属月）")
    year = int(src[:4])
    wb = load_workbook(path, data_only=True)

    def read_sheet(sheet_name):
        if sheet_name not in wb.sheetnames:
            return {}
        ws = wb[sheet_name]
        first = find_total_first_col(ws)
        if first is None:
            sys.exit(f"❌ {path} 的 [{sheet_name}] 行1 找不到「总计」组，模板布局变了？")
        out = {}
        for r in range(4, ws.max_row + 1):
            sku = ws.cell(r, 1).value
            if not sku or not isinstance(sku, str) or not sku.strip():
                continue
            out[sku.strip()] = [n0(ws.cell(r, first + i).value) for i in range(12)]
        return out

    prev = read_sheet(f"{year - 1}年")
    cur = read_sheet(f"{year}年")
    wb.close()
    data = {}
    for sku, m in list(prev.items()) + list(cur.items()):
        seq = prev.get(sku, [0.0] * 12) + cur.get(sku, [0.0] * 12)
        data[sku] = seq
    return data, year


def score_sku(seq, score_idx, prev_score, inherit=True):
    """按规则给单个 SKU 打分。

    seq        : 24 个月销量序列（下标 0 = 去年 1 月）
    score_idx  : 评分月在序列中的下标
    prev_score : 上期得分（None/NaN 表示上期未评级）
    inherit    : True=沿用上期分类（apply 生成新月度用）；False=纯规则（verify 复刻用）
    返回 (得分, 依据说明)
    """
    m1 = seq[score_idx]
    m3 = sum(seq[max(0, score_idx - 2): score_idx + 1])
    m6 = sum(seq[max(0, score_idx - 5): score_idx + 1])
    first_idx = next((i for i, v in enumerate(seq[: score_idx + 1]) if v > 0), None)

    # 上期未评级（停售/不销售/未登记评级）→ 一律不碰
    if prev_score is None or (isinstance(prev_score, float) and prev_score != prev_score):
        return None, "上期未评级，保持空"

    prev = str(prev_score).strip()
    # 新品 = 首销距评分月 ≤4 个月（202609 成品反推：A-E 老品首销最晚 3 月、
    # X 类新品首销全部 ≥4 月；评分月当月首销的算当月新品）
    is_newbie = first_idx is None or (score_idx - first_idx) <= 4

    # ---- 第一步：决定走 X 类还是 A-E ----
    use_x = False
    if inherit and prev.startswith("X"):
        # 新品继承：从无销售 / 当月无销售 → X；当月仍是 XE 档（<33 件）或
        # 距首销仍 ≤4 个月 → 继续 X 档；当月起量且满 5 个月 → 毕业转 A-E
        if first_idx is None or m1 == 0:
            return "X", "无销售数据" if first_idx is None else "当月无销售"
        if m1 < 33.3 or (score_idx - first_idx) <= 4:
            use_x = True
    elif not inherit:
        # 纯规则：无 6 个月完整历史 → X 类当月分档
        if first_idx is None:
            return "X", "无销售数据"
        if is_newbie:
            use_x = True
        elif m1 == 0:
            use_x = False  # 老品当月 0，按 3/6 个月窗口照常评
    elif first_idx is None:
        return "X", "无销售数据"

    # ---- 第二步：X 类按当月日均分档 ----
    if use_x:
        if m1 == 0:
            return "X", "当月无销售"
        daily = m1 / 30.0
        if daily >= 16.7:
            return "XA", f"新品日均{daily:.1f}(当月{m1:.0f})"
        if daily >= 5.5:
            return "XB", f"新品日均{daily:.1f}(当月{m1:.0f})"
        if daily >= 1.38:
            return "XC", f"新品日均{daily:.1f}(当月{m1:.0f})"
        if daily >= 1.11:
            return "XD", f"新品日均{daily:.1f}(当月{m1:.0f})"
        return "XE", f"新品日均{daily:.1f}(当月{m1:.0f})"

    # ---- 第三步：A-E 按 1/3/6 个月窗口 ----
    if m1 >= 500:
        return "A", f"1个月{m1:.0f}"
    if m3 >= 500:
        return "B", f"3个月{m3:.0f}"
    if m6 >= 250:
        return "C", f"6个月{m6:.0f}"
    if m6 >= 200:
        return "D", f"6个月{m6:.0f}"
    return "E", f"6个月{m6:.0f}"


def find_score_cols(ws):
    """返回 [(月份int, 列号, 列名), ...] 按月份升序。"""
    cols = []
    for c in range(1, ws.max_column + 1):
        v = ws.cell(PROD_HEADER_ROW, c).value
        if v and "销量得分" in str(v):
            m = re.search(r"(\d+)月", str(v))
            cols.append((int(m.group(1)) if m else -1, c, str(v)))
    return sorted(cols)


def main():
    ap = argparse.ArgumentParser(description="按成品月销数据自动生成商品资料得分列")
    ap.add_argument("--src", required=True, help="成品月销总表归属月，如 202609")
    ap.add_argument("--score-month", type=int, default=None,
                    help="评分月（默认=src 归属月的月份）")
    ap.add_argument("--verify", action="store_true",
                    help="自验模式：与现有得分列逐行比对，不写文件")
    ap.add_argument("--apply", action="store_true",
                    help="写回商品资料.xlsx（先自动备份）")
    ap.add_argument("--prev-col", default=None,
                    help="上期得分列名（默认自动取月份最大的得分列）")
    args = ap.parse_args()

    score_month = args.score_month or int(args.src[4:6])
    if not 1 <= score_month <= 12:
        sys.exit(f"❌ --score-month 必须是 1~12，收到 {score_month}")

    data, year = load_month_sale(args.src)
    # 评分月序列下标：去年 12 个月 + 当年 (score_month-1)
    score_idx = 12 + score_month - 1
    new_col_name = f"{score_month}月销量得分"
    print(f"数据源：result/数据归属周期/{args.src}/月销售总表.xlsx（{year} 年 sheet）")
    print(f"评分月：{score_month} 月 → 新列「{new_col_name}」；成品 SKU 行数 {len(data)}")

    wb = load_workbook(PROD_PATH)
    ws = wb[PROD_SHEET]

    score_cols = find_score_cols(ws)
    if not score_cols:
        sys.exit("❌ 商品资料找不到任何「N月销量得分」列")
    if args.prev_col:
        prev_col = next((c for _, c, n in score_cols if n == args.prev_col), None)
        if prev_col is None:
            sys.exit(f"❌ --prev-col {args.prev_col} 不存在，现有：{[n for _,_,n in score_cols]}")
    else:
        prev_col = score_cols[-1][1]
    prev_name = ws.cell(PROD_HEADER_ROW, prev_col).value
    print(f"上期列：{prev_name}（列 {get_column_letter(prev_col)}）")

    # 目标列
    exist = [c for _, c, n in score_cols if n == new_col_name]
    if exist:
        target_col = exist[0]
        print(f"⚠️ 「{new_col_name}」列已存在（{get_column_letter(target_col)}），将原地覆盖")
    else:
        target_col = max(c for c in range(1, ws.max_column + 2)
                         if ws.cell(PROD_HEADER_ROW, c).value is not None) + 1
        print(f"新列位置：{get_column_letter(target_col)}")

    # ---------- 逐行评分 ----------
    results = []  # (row, sku, prev_raw, new_score, reason)
    for r in range(PROD_DATA_START, ws.max_row + 1):
        sku = ws.cell(r, 1).value
        if not sku or not isinstance(sku, str) or not sku.strip():
            continue
        sku = sku.strip()
        prev_raw = ws.cell(r, prev_col).value
        seq = data.get(sku)
        if seq is None:
            # 成品无此 SKU：上期也未评级 → 保持空；否则 X（暂无销售数据）
            if prev_raw is None or (isinstance(prev_raw, float) and prev_raw != prev_raw):
                results.append((r, sku, prev_raw, None, "成品无此行且上期未评级，保持空"))
            else:
                results.append((r, sku, prev_raw, "X", "成品无此行"))
            continue
        new_score, reason = score_sku(seq, score_idx, prev_raw, inherit=not args.verify)
        results.append((r, sku, prev_raw, new_score, reason))

    # ---------- 统计 ----------
    from collections import Counter
    dist = Counter(str(s) for _, _, _, s, _ in results)
    total_scored = sum(1 for _, _, _, s, _ in results if s is not None)
    print(f"\n评级分布（{total_scored} 行有分 / 共 {len(results)} 行）：")
    for k in sorted(dist):
        print(f"  {k:>4s}: {dist[k]}")

    # ---------- 与上期比对（三种模式都打印，便于人工复核） ----------
    same = diff = 0
    diffs = []
    for r, sku, prev_raw, new_score, reason in results:
        p = None if prev_raw is None or (isinstance(prev_raw, float) and prev_raw != prev_raw) else str(prev_raw).strip()
        n = new_score if new_score is not None else None
        if p == n:
            same += 1
        else:
            diff += 1
            diffs.append((sku, p, n, reason))
    print(f"\n与上期「{prev_name}」比对：平级 {same}，变化 {diff}")
    if diffs:
        print("变化明细（SKU | 上期→本期 | 依据）：")
        for sku, p, n, reason in sorted(diffs, key=lambda x: (str(x[1]), str(x[2]))):
            print(f"  {sku:<14s} {str(p):>4s} → {str(n):<4s} {reason}")

    # 贴线个案：差一点就升级/降级的，人工复核用
    borderline = []
    for r, sku, prev_raw, new_score, reason in results:
        seq = data.get(sku)
        if seq is None or new_score is None:
            continue
        m1 = seq[score_idx]
        m3 = sum(seq[max(0, score_idx - 2): score_idx + 1])
        m6 = sum(seq[max(0, score_idx - 5): score_idx + 1])
        near = []
        if 470 <= m1 < 500:
            near.append(f"1个月{m1:.0f}贴A线500")
        if 470 <= m3 < 500:
            near.append(f"3个月{m3:.0f}贴B线500")
        if 240 <= m6 < 250:
            near.append(f"6个月{m6:.0f}贴C线250")
        if 190 <= m6 < 200:
            near.append(f"6个月{m6:.0f}贴D线200")
        if near:
            borderline.append((sku, new_score, "；".join(near)))
    if borderline:
        print(f"\n⚠️ 贴线个案 {len(borderline)} 行（差 1~30 件就换级，建议人工抽查）：")
        for sku, s, why in borderline:
            print(f"  {sku:<14s} → {s:<4s} {why}")

    if not args.apply:
        print("\n（未加 --apply，以上为预览，未写任何文件）")
        wb.close()
        return

    # ---------- 写回 ----------
    backup_dir = os.path.join(TRUNK, ".workbuddy", "backup_scorecol")
    os.makedirs(backup_dir, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup = os.path.join(backup_dir, f"商品资料_{stamp}.xlsx")
    shutil.copy2(PROD_PATH, backup)
    print(f"\n已备份 → {backup}")

    # 表头 + 样式复制上期列
    hdr_src = ws.cell(PROD_HEADER_ROW, prev_col)
    hdr = ws.cell(PROD_HEADER_ROW, target_col)
    hdr.value = new_col_name
    hdr._style = copy(hdr_src._style)

    written = 0
    for r, sku, prev_raw, new_score, reason in results:
        cell = ws.cell(r, target_col)
        cell._style = copy(ws.cell(r, prev_col)._style)
        if new_score is None:
            cell.value = None
        else:
            cell.value = new_score
            written += 1
    wb.save(PROD_PATH)
    wb.close()
    print(f"✅ 已写回 {PROD_PATH}：「{new_col_name}」列（{get_column_letter(target_col)}），"
          f"有效评级 {written} 行，未评级保持空")

    # ---------- 自检 ----------
    wb2 = load_workbook(PROD_PATH)
    ws2 = wb2[PROD_SHEET]
    back = [ws2.cell(r, target_col).value for r in range(PROD_DATA_START, ws2.max_row + 1)]
    n_f = sum(1 for v in back if v is not None)
    n_formula = sum(1 for row in ws2.iter_rows() for c in row
                    if isinstance(c.value, str) and c.value.startswith("="))
    sheets_ok = wb2.sheetnames
    merged_ok = [str(m) for m in ws2.merged_cells.ranges]
    wb2.close()
    print(f"自检：读回有效值 {n_f} 行（应={written}）；公式格 {n_formula}（应=595）；"
          f"sheet 名单一致={sheets_ok}；合并区一致={merged_ok == ['A1:M1']}")
    if n_f != written or n_formula != 595 or merged_ok != ["A1:M1"]:
        sys.exit("❌ 自检未过，请从备份还原！")
    print("自检通过。")


if __name__ == "__main__":
    main()
