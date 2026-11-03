# -*- coding: utf-8 -*-
"""三方订单级闭合核对：月销表 vs 订单报告(站点本地时区) vs 日销归档簿(UTC)。

背景：日销表前面的「月汇总」是跑批时 resolve_month_sales 从归档簿明细重算的
（UTC 辅助列 + 明细状态）。要让 10 月汇总 = 月销表 50788，必须知道每一单
差异在哪、属于什么口径差。本工具把三方逐单对齐，输出可执行的处理清单。

用法：
    python pipeline/tools/verify_month_orders.py --period 202610

输出：
    断货分析/_对账/三方订单核对_{period}.xlsx   （4 个 sheet）
    断货分析/_对账/三方订单核对_{period}.md
"""
import os
import sys
import glob
import argparse
import datetime
from collections import defaultdict
from zoneinfo import ZoneInfo

import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
DSR = os.path.abspath(os.path.join(HERE, "..", ".."))
sys.path.insert(0, os.path.join(DSR, "pipeline", "tools"))
import audit_sales_gap as A          # noqa: E402  复用同一套取数/时区判据

REPO = A.REPO
DAY_DIR = A.DAY_DIR
GROUPS = A.GROUPS


# ---------------------------------------------------------------- 月销表
def read_month_side(period, month):
    """月销表逐分组 {sku: qty}（组首列由行1定位，数据从行4起）。"""
    p = os.path.join(REPO, "trunk", "result", "数据归属周期", str(period), "月销售总表.xlsx")
    if not os.path.exists(p):
        raise SystemExit("找不到月销表：" + p)
    wb = openpyxl.load_workbook(p, read_only=True, data_only=True)
    ws = wb[str(period)[:4] + "年"]
    r1 = next(ws.iter_rows(min_row=1, max_row=1, values_only=True))
    starts = {}
    for i, v in enumerate(r1, 1):
        if isinstance(v, str) and v.strip():
            starts.setdefault(v.strip(), i)
    out = {}
    for g, c0 in starts.items():
        col = c0 + (month - 1)
        d = {}
        for r in ws.iter_rows(min_row=4, max_row=ws.max_row, max_col=col, values_only=True):
            sku = str(r[0]).strip() if r[0] is not None else ""
            if not sku or sku.startswith("总计"):
                continue
            v = r[col - 1]
            if isinstance(v, (int, float)) and v:
                d[sku] = d.get(sku, 0) + v
        out[g] = d
    wb.close()
    return out


# ---------------------------------------------------------------- 报告（月销口径：本地时区归月）
def load_report_local(period):
    """最新快照报告 → {(站点组, sku, key): 单据}，本地时区归月。"""
    snap, rdir = A.latest_snap(str(period))
    rep = {}
    for f in sorted(glob.glob(os.path.join(rdir, "*订单报告.txt"))):
        d = A.read_rep(os.path.join(rdir, f))
        if d:
            rep.update(d)
    out = {}
    for k, v in rep.items():
        if v["status"] == "Cancelled" or not v["utc"]:
            continue
        loc = v["utc"].astimezone(A.tz_of(v["ch"]))
        if loc.strftime("%Y-%m") != "%s-%02d" % (str(period)[:4], int(str(period)[4:])):
            continue
        out[(v["grp"], v["ch"], v["pd"][:19], k[0], k[1])] = v
    return out, snap


# ---------------------------------------------------------------- 归档簿（日销口径：UTC 归月）
def load_archive_utc(period):
    adir = os.path.join(DAY_DIR, "%s年每日数据" % str(period)[:4],
                        "%s年%s月" % (str(period)[:4], int(str(period)[4:])))
    out = {}
    for f in sorted(glob.glob(os.path.join(adir, "%s年%s月-*.xlsx" % (str(period)[:4], int(str(period)[4:]))))):
        d = A.read_arc(f)
        if d:
            out.update(d)
    keep = {}
    for k, v in out.items():
        if v["status"] == "Cancelled" or not v["utc"]:
            continue
        if v["utc"].strftime("%Y-%m") != "%s-%02d" % (str(period)[:4], int(str(period)[4:])):
            continue
        keep[(v["grp"], v["ch"], v["pd"][:19], k[0], k[1])] = v
    return keep


def local_ym(v):
    return v["utc"].astimezone(A.tz_of(v["ch"])).strftime("%Y-%m") if v["utc"] else "-"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--period", required=True)
    a = ap.parse_args()
    period = str(a.period)
    year, month = int(period[:4]), int(period[4:])
    outdir = os.path.join(DSR, "断货分析", "_对账")
    os.makedirs(outdir, exist_ok=True)

    print("1/4 读月销表 ...")
    M = read_month_side(period, month)
    print("2/4 读订单报告（本地时区归月）...")
    B, snap = load_report_local(period)
    print("3/4 读日销归档簿（UTC 归月）...")
    Ab = load_archive_utc(period)

    # ---- 逐站三方合计
    rows = [["分组", "月销表", "报告重算(本地)", "归档簿(UTC)", "月销-报告", "报告-归档"]]
    for g in GROUPS:
        sm = sum(M.get(g, {}).values())
        sb = sum(v["qty"] for k, v in B.items() if k[0] == g)
        sa = sum(v["qty"] for k, v in Ab.items() if k[0] == g)
        rows.append([g, sm, sb, sa, sm - sb, sb - sa])
    sm = sum(sum(M.get(g, {}).values()) for g in GROUPS)
    sb = sum(v["qty"] for v in B.values())
    sa = sum(v["qty"] for v in Ab.values())
    rows.append(["合计", sm, sb, sa, sm - sb, sb - sa])

    # ---- 渠道构成（解释月销表 50788 里非 Amazon. 渠道的量）
    chsum = defaultdict(float)
    for v in B.values():
        chsum[v["ch"]] += v["qty"]

    # ---- 逐单差异（key 五元组：分组,渠道,pd,订单号,sku）
    only_a = {k: v for k, v in Ab.items() if k not in B}     # 日销有、月销无
    only_b = {k: v for k, v in B.items() if k not in Ab}     # 月销有、日销无

    def cls_a(v):
        lm = local_ym(v)[:7]
        if lm == "%04d-%02d" % (year, month + 1) if month < 12 else False:
            return "跨月(本地已次月)"
        if lm != "%04d-%02d" % (year, month):
            return "跨月(本地非本月)"
        return "口径异常(本地也在本月,报告查无)"

    def cls_b(v):
        um = v["utc"].strftime("%Y-%m")
        if um < "%04d-%02d" % (year, month):
            return "跨月(UTC上月)"
        if um > "%04d-%02d" % (year, month):
            return "跨月(UTC次月)"
        return "漏记(UTC同月)"

    ca = defaultdict(lambda: [0, 0.0])
    for v in only_a.values():
        c = cls_a(v)
        ca[c][0] += 1
        ca[c][1] += v["qty"]
    cb = defaultdict(lambda: [0, 0.0])
    for v in only_b.values():
        c = cls_b(v)
        cb[c][0] += 1
        cb[c][1] += v["qty"]

    # ---- 写 xlsx
    xp = os.path.join(outdir, "三方订单核对_%s.xlsx" % period)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "三方逐站"
    ws.append(rows[0])
    for r in rows[1:]:
        ws.append(r)
    ws.append([])
    ws.append(["报告渠道构成（本地10月）"])
    ws.append(["渠道", "件数"])
    for c, q in sorted(chsum.items(), key=lambda x: -x[1]):
        ws.append([c, q])

    ws2 = wb.create_sheet("日销多_月销无")
    ws2.append(["分组", "渠道", "purchase-date(UTC)", "本地时间", "订单号", "SKU", "数量", "状态", "原因"])
    for k, v in sorted(only_a.items(), key=lambda x: (x[0][0], x[1]["pd"])):
        loc = v["utc"].astimezone(A.tz_of(v["ch"])).strftime("%Y-%m-%d %H:%M")
        ws2.append([k[0], k[1], k[2], loc, k[3], k[4], v["qty"], v["status"], cls_a(v)])

    ws3 = wb.create_sheet("月销多_日销无")
    ws3.append(["分组", "渠道", "purchase-date(UTC)", "本地时间", "订单号", "SKU", "数量", "状态", "原因"])
    for k, v in sorted(only_b.items(), key=lambda x: (x[0][0], x[1]["pd"])):
        loc = v["utc"].astimezone(A.tz_of(v["ch"])).strftime("%Y-%m-%d %H:%M")
        ws3.append([k[0], k[1], k[2], loc, k[3], k[4], v["qty"], v["status"], cls_b(v)])

    ws4 = wb.create_sheet("月销-归档 逐SKU差")
    ws4.append(["分组", "SKU", "归档簿(UTC)", "月销表", "差"])
    n4 = 0
    for g in GROUPS:
        ma, mm = {}, M.get(g, {})
        for k, v in Ab.items():
            if k[0] == g:
                ma[k[4]] = ma.get(k[4], 0) + v["qty"]
        for sku in sorted(set(ma) | set(mm), key=lambda s: -abs(ma.get(s, 0) - mm.get(s, 0))):
            d = ma.get(sku, 0) - mm.get(sku, 0)
            if d:
                ws4.append([g, sku, ma.get(sku, 0), mm.get(sku, 0), d])
                n4 += 1
    wb.save(xp)
    print("4/4 写出", xp, " 逐SKU差异行:", n4)

    # ---- 写 md
    L = ["# 三方订单核对 %s（月销表 vs 报告重算 vs 日销归档簿）\n" % period]
    L.append("报告快照：%s   归档簿：%s年%s月\n" % (snap, str(period)[:4], int(str(period)[4:])))
    L.append("| 分组 | 月销表 | 报告重算(本地) | 归档簿(UTC) | 月销-报告 | 报告-归档 |")
    L.append("|---|---:|---:|---:|---:|---:|")
    for r in rows[1:]:
        L.append("| %s | %s | %s | %s | %s | %s |" % tuple(r))
    L.append("")
    L.append("## 日销多、月销无（%d 单 / %.0f 件）\n" % (len(only_a), sum(v["qty"] for v in only_a.values())))
    for c, (n, q) in sorted(ca.items()):
        L.append("- %s：%d 单 / %.0f 件" % (c, n, q))
    L.append("")
    L.append("## 月销多、日销无（%d 单 / %.0f 件）\n" % (len(only_b), sum(v["qty"] for v in only_b.values())))
    for c, (n, q) in sorted(cb.items()):
        L.append("- %s：%d 单 / %.0f 件" % (c, n, q))
    L.append("")
    L.append("> 结论：报告重算(本地) 与 月销表 的差 = 报告导出口径残留；")
    L.append("> 归档簿(UTC) 与 报告重算 的差 = 归月基准(UTC vs 本地) + 封存时点漏记。")
    mp = os.path.join(outdir, "三方订单核对_%s.md" % period)
    with open(mp, "w", encoding="utf-8") as f:
        f.write("\n".join(L))
    print("写出", mp)

    for r in rows[1:]:
        print("%-6s 月销%7.0f 报告%7.0f 归档%7.0f" % (r[0], r[1], r[2], r[3]))


if __name__ == "__main__":
    main()
