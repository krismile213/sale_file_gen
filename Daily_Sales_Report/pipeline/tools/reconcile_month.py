# -*- coding: utf-8 -*-
"""
日销表 / 月销表 逐站逐 SKU 对账（常驻工具）

用法：
    python pipeline/tools/reconcile_month.py --period 202610
    python pipeline/tools/reconcile_month.py --period 202610 --out 断货分析/_对账
    python pipeline/tools/reconcile_month.py --period 202610 --sku-limit 40

两个数据源的取数方式（**都不读日列**，日列是活公式且 openpyxl 存后缓存为空）：
  · 日销表：`1.每日销售数据表/*库存销售表.xlsx` → 各「XX-每日库存分析表」
            行5 里找标签恰为「{N}月」的列（月汇总块），行7 起 A 列 = SKU。
            该块是**静态值**（换月时由 resolve_month_sales 反推写入），可直接读。
  · 月销表：`trunk/result/数据归属周期/{period}/月销售总表.xlsx` → sheet「{Y}年」
            行1 = 分组名（欧洲/英国/美国/...），行2 = 1~12月，
            分组起列 + (N-1) 即当月列，逐 SKU 求和（该表无合计行）。

输出：
  · 控制台：逐站对比表（日销表 / 月销表 / 差 / 差异 SKU 数）
  · {out}/对账_{period}.md      人看
  · {out}/对账_{period}_明细.csv 逐 SKU 差异（只在差异 != 0 时写行）

⚠️ 已知口径差（不是 bug，别当 bug 追）：
  1. 日销表 = 逐日 append 只增不删；月销表 = 订单报告整月重算。
     同一月必然有几十件系统差，形态是几百个 SKU 各差 ±1~±6。
  2. 东南亚（菲律宾/泰国/印尼）是**另一套销售表**，不在 13 本分析表里；
     印尼实测逐 SKU 零差异，泰/菲因「已收货可退货」状态归一化 + SKU 行位不全而偏高。
     要核东南亚请用 东南亚日销/ 下的三国表，本脚本不管。
  3. 英国报告与欧盟报告内容重复，月销表里英国是独立分组，别把两份相加。
"""
import os
import re
import sys
import csv
import glob
import argparse
import datetime

import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
DSR = os.path.dirname(os.path.dirname(HERE))          # Daily_Sales_Report
DAY_DIR = os.path.join(DSR, "1.每日销售数据表")
REPO = os.path.dirname(DSR)                            # sale_file_gen

# 分析表 sheet 名 → 月销表分组名。None = 该表不在月销表里（单独业务）
# ⚠️ 同一站点在不同簿里 sheet 名不一样（"沙特" vs "沙特阿拉伯"）→ 必须映射到同一分组后合并；
#    同名不同业务（"墨西哥" 在北美本是亚马逊、在美客多本是美客多）→ 按文件名区分，见 biz()。
SHEET2GROUP = {
    "欧洲": "欧洲", "英国": "英国",
    "美国": "美国", "加拿大": "加拿大",
    "墨西哥": "墨西哥",
    "日本": "日本", "澳大利亚": "澳洲", "澳洲": "澳洲",
    "阿联酋": "阿联酋", "沙特": "沙特", "沙特阿拉伯": "沙特",
    "印度": "印度",
    "罗马尼亚": "eMAG",
}


def biz(fname):
    """按文件名判定业务线，用于区分同名 sheet"""
    if "美客多" in fname:
        return "美客多"
    if "Interflowd" in fname:
        return "interflow"
    return "本地"


def site_group(site, fname):
    """(sheet 站名, 文件名) → 月销表分组名；None 表示不参与对账"""
    if biz(fname) == "美客多" and site == "墨西哥":
        return "美客多"
    return SHEET2GROUP.get(site)


def norm_sku(v):
    if v is None:
        return None
    s = str(v).strip().upper()
    return s or None


# ---------------- 日销表 ----------------
def _day_files(period):
    """日销侧取数：优先月末封存簿（历史留档，封存前已按订单最终状态校正过），
    没有则回落当前目录（该月还没封存的情况）。
    2026-11-02：原实现只读当前目录 → 换月后再对账，读到的还是当月版里
    封存时刻的旧值，看不到封存前校正的效果。"""
    y, m = period[:4], int(period[4:6])
    ad = os.path.join(DAY_DIR, "%s年每日数据" % y, "%s年%d月" % (y, m))
    af = sorted(glob.glob(os.path.join(ad, "%s年%d月-*.xlsx" % (y, m))))
    if af:
        return af, "月末封存簿"
    return sorted(glob.glob(os.path.join(DAY_DIR, "*库存销售表.xlsx"))), "当前簿"


def read_day_side(period, month):
    """返回 {分组: {sku: qty}} 和 {分组: 来源文件列表}（同分组的多本簿合并）"""
    out, src = {}, {}
    files, _kind = _day_files(period)
    for f in files:
        fn = os.path.basename(f)
        try:
            wb = openpyxl.load_workbook(f, read_only=True, data_only=True)
        except Exception as e:
            print(f"  ⚠️ 打不开 {fn}: {e}")
            continue
        for sn in wb.sheetnames:
            if "每日库存分析表" not in sn:
                continue
            site = sn.split("-")[0].strip()
            g = site_group(site, fn)
            if g is None:
                continue
            ws = wb[sn]
            r5 = next(ws.iter_rows(min_row=5, max_row=5, values_only=True), None)
            if not r5:
                continue
            col = None
            label = f"{month}月"
            for ci, v in enumerate(r5, 1):
                if isinstance(v, str) and v.strip() == label:
                    col = ci
                    break
            if col is None:
                src.setdefault(g, []).append(f"{fn}[{sn.strip()}] 缺「{label}」列")
                continue
            d = out.setdefault(g, {})
            for r in ws.iter_rows(min_row=7, max_row=ws.max_row, max_col=col, values_only=True):
                sku = norm_sku(r[0])
                if sku is None:
                    continue
                v = r[col - 1]
                if isinstance(v, (int, float)) and v:
                    d[sku] = d.get(sku, 0) + v
            src.setdefault(g, []).append(f"{fn}[{sn.strip()}]")
        wb.close()
    return out, src


# ---------------- 月销表 ----------------
def _msx_path(period):
    """月销表文件名随月份变（`月销售总表.xlsx` / `月销售总表（10月）.xlsx` 两种历史命名），
    不能硬编码。优先精确名，否则取该周期目录下最新的一份。"""
    d = os.path.join(REPO, "trunk", "result", "数据归属周期", str(period))
    exact = os.path.join(d, "月销售总表.xlsx")
    if os.path.exists(exact):
        return exact
    cands = sorted(glob.glob(os.path.join(d, "月销售总表*.xlsx")),
                   key=os.path.getmtime, reverse=True)
    if not cands:
        raise SystemExit(f"找不到月销表：{d}/月销售总表*.xlsx")
    return cands[0]


def read_month_side(period, month):
    p = _msx_path(period)
    wb = openpyxl.load_workbook(p, read_only=True, data_only=True)
    sn = f"{str(period)[:4]}年"
    if sn not in wb.sheetnames:
        raise SystemExit(f"月销表无 sheet「{sn}」")
    ws = wb[sn]
    r1 = next(ws.iter_rows(min_row=1, max_row=1, values_only=True))
    starts = {}
    for ci, v in enumerate(r1, 1):
        if isinstance(v, str) and v.strip():
            starts.setdefault(v.strip(), ci)
    out = {}
    for g, c0 in starts.items():
        col = c0 + (month - 1)
        d = {}
        for r in ws.iter_rows(min_row=4, max_row=ws.max_row, max_col=col, values_only=True):
            sku = norm_sku(r[0])
            if sku is None:
                continue
            v = r[col - 1]
            if isinstance(v, (int, float)) and v:
                d[sku] = d.get(sku, 0) + v
        out[g] = d
    wb.close()
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--period", required=True, help="归属周期，如 202610")
    ap.add_argument("--out", default=None, help="输出目录，默认 断货分析/_对账")
    ap.add_argument("--sku-limit", type=int, default=25, help="每站最多列多少条差异（默认25）")
    a = ap.parse_args()

    period = str(a.period)
    year, month = int(period[:4]), int(period[4:])
    outdir = a.out or os.path.join(DSR, "断货分析", "_对账")
    os.makedirs(outdir, exist_ok=True)

    print(f"对账周期 {period}（{year}年{month}月）")
    _df, _dk = _day_files(period)
    print(f"日销侧来源：{_dk}（{len(_df)} 本）")
    day, dsrc = read_day_side(period, month)
    mon = read_month_side(period, month)
    if not day:
        raise SystemExit("日销表没读到任何站（检查 1.每日销售数据表/ 与分析表是否就位）")

    rows = []
    detail = []
    for g in sorted(day):
        m = mon.get(g)
        td = sum(day[g].values())
        if m is None:
            rows.append((g, g, td, None, None, 0, f"月销表无分组「{g}」"))
            continue
        tm = sum(m.values())
        diffs = []
        for k in set(day[g]) | set(m):
            dv, mv = day[g].get(k, 0), m.get(k, 0)
            if dv != mv:
                diffs.append((k, dv, mv, dv - mv))
        diffs.sort(key=lambda x: -abs(x[3]))
        rows.append((g, g, td, tm, td - tm, len(diffs), ""))
        for k, dv, mv, dd in diffs:
            detail.append([period, g, g, k, dv, mv, dd])

    # ---- 控制台 ----
    print()
    hdr = "%-8s %-8s %10s %10s %10s %8s" % ("日销站", "月销组", "日销表", "月销表", "差", "差异SKU")
    print(hdr)
    print("-" * len(hdr))
    sd = sm = 0
    for site, g, td, tm, diff, n, note in rows:
        if tm is None:
            print("%-8s %-8s %10s %10s %10s %8s  %s" % (site, g or "-", td, "-", "-", "-", note))
            continue
        sd += td
        sm += tm
        flag = "" if diff == 0 else ("  ←" if abs(diff) <= 5 else "  ←←")
        print("%-8s %-8s %10d %10d %+10d %8d%s" % (site, g, td, tm, diff, n, flag))
    print("-" * len(hdr))
    if sm:
        print("%-8s %-8s %10d %10d %+10d %8s  (%.2f%%)" % (
            "合计", "", sd, sm, sd - sm, "", (sd - sm) / sm * 100 if sm else 0))

    # ---- 写文件 ----
    md = os.path.join(outdir, f"对账_{period}.md")
    with open(md, "w", encoding="utf-8") as f:
        f.write(f"# 日销表 / 月销表 对账 {period}\n\n")
        f.write(f"生成时间：{datetime.datetime.now():%Y-%m-%d %H:%M}\n\n")
        f.write("## 逐站\n\n| 日销站 | 月销组 | 日销表 | 月销表 | 差 | 差异SKU | 备注 |\n")
        f.write("|---|---|---:|---:|---:|---:|---|\n")
        for site, g, td, tm, diff, n, note in rows:
            if tm is None:
                f.write(f"| {site} | {g or '-'} | {td} | - | - | - | {note} |\n")
            else:
                f.write(f"| {site} | {g} | {td} | {tm} | {diff:+d} | {n} | {note} |\n")
        if sm:
            f.write(f"| **合计** | | **{sd}** | **{sm}** | **{sd-sm:+d}** | | |\n")
        f.write("\n## 差异最大的 SKU（每站前 %d）\n\n" % a.sku_limit)
        bysite = {}
        for _, site, g, k, dv, mv, dd in detail:
            bysite.setdefault(site, []).append((k, dv, mv, dd))
        for site, lst in bysite.items():
            lst.sort(key=lambda x: -abs(x[3]))
            f.write(f"### {site}（{len(lst)} 个差异 SKU）\n\n| SKU | 日销表 | 月销表 | 差 |\n|---|---:|---:|---:|\n")
            for k, dv, mv, dd in lst[:a.sku_limit]:
                f.write(f"| {k} | {dv} | {mv} | {dd:+d} |\n")
            f.write("\n")
        f.write("\n> 口径说明：日销表=逐日 append 只增不删；月销表=订单报告整月重算。"
                "同月几十件系统差属正常，形态是大量 SKU 各差 ±1~±6。\n")
    cp = os.path.join(outdir, f"对账_{period}_明细.csv")
    with open(cp, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["周期", "日销站", "月销组", "SKU", "日销表", "月销表", "差"])
        w.writerows(detail)
    print(f"\n已写：{md}")
    print(f"      {cp}（{len(detail)} 行差异）")


if __name__ == "__main__":
    main()
