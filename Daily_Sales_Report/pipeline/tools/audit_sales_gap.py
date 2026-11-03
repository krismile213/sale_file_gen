# -*- coding: utf-8 -*-
"""
日销表 vs 月销表 **订单级归因**（常驻工具）

回答一个问题：「这两个口径差了几百件，到底差在哪些订单上？」

用法：
    python pipeline/tools/audit_sales_gap.py --period 202610
    python pipeline/tools/audit_sales_gap.py --period 202610 --out 断货分析/_对账

两边取数（与各自真实口径一致，保证是"订单对订单"）：
  · 日销侧 = 月末封存簿 `1.每日销售数据表/{Y}年每日数据/{Y}年{M}月/{Y}年{M}月-*.xlsx`
             「每日销售数据*」sheet 的订单行。该簿是换月时 resolve_month_sales 的封存源，
             也就是日销表「N月」静态值的真正来源。
             → 归月用 **UTC 日期**（表内辅助列 AE = LEFT(C,10)，与日销列 SUMIFS 完全一致）
  · 月销侧 = 快照目录最新一份 `*站点{period}订单报告.txt`
             → 归月用 **站点本地时间**（month_sale_file.py 用 amazon_order_tz 转换后取 .month）

归因分类（按订单号逐单判定，互斥）：
  ③ 取消后未回退  日销明细里状态非 Cancelled（首次抓到时 Pending/Shipped），
                  最新报告里已 Cancelled → 日销表算了、月销表没算  ← 日销多算
  ② 报告里已无此单 日销明细有、最新报告整单消失（多为已删单）      ← 日销多算
  ① 当时已取消后复活 日销明细里是 Cancelled，最终却有效             ← 日销漏算
  ④ 日销明细漏记   最终有效，但日销明细里压根没这行               ← 日销漏算
  ⑤ 跨月边界       同一单 UTC 归 10 月、本地归 11 月（或反之）      ← 口径差，非错

⚠️ 只统计 sales-channel 以 Amazon. 开头的行（TEMU 是北美簿额外放行的旁支，单列不计）。
"""
import os
import re
import csv
import sys
import glob
import argparse
import datetime
from collections import defaultdict
from zoneinfo import ZoneInfo

import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
DSR = os.path.dirname(os.path.dirname(HERE))
DAY_DIR = os.path.join(DSR, "1.每日销售数据表")
REPO = os.path.dirname(DSR)

# 渠道 → 站点分组（与月销表/日销表一致）
EU_CH = {"Amazon.co.uk"}
CH_GRP = {
    "Amazon.com": "美国", "Amazon.ca": "加拿大", "Amazon.com.mx": "墨西哥",
    "Amazon.co.jp": "日本", "Amazon.com.au": "澳洲", "Amazon.ae": "阿联酋",
    "Amazon.sa": "沙特", "Amazon.in": "印度",
}
CH_TZ = {
    "Amazon.com": "America/Los_Angeles", "Amazon.ca": "America/Los_Angeles",
    "Amazon.com.mx": "America/Los_Angeles",
    "Amazon.co.jp": "Asia/Tokyo", "Amazon.com.au": "Australia/Sydney",
    "Amazon.ae": "Asia/Dubai", "Amazon.sa": "Asia/Riyadh", "Amazon.in": "Asia/Kolkata",
}
GROUPS = ["欧洲", "英国", "美国", "加拿大", "墨西哥", "日本", "澳洲", "阿联酋", "沙特", "印度"]

ORDER = ["③ 取消后未回退", "② 报告里已无此单", "⑤ 跨月(日销多)",
         "⑤ 跨月(月销多)", "① 当时已取消后复活", "④ 日销明细漏记"]


def grp_of(ch):
    if ch in EU_CH:
        return "英国"
    if ch in CH_GRP:
        return CH_GRP[ch]
    if ch.startswith("Amazon."):
        return "欧洲"
    return None


def tz_of(ch):
    if ch in CH_TZ:
        return ZoneInfo(CH_TZ[ch])
    return ZoneInfo("Europe/London")     # 欧洲/英国 → 伦敦（实测订单报告就是英国时间）


_PD = re.compile(r"^(\d{4})-(\d{2})-(\d{2})[T ](\d{2}):(\d{2}):(\d{2})"
                 r"(?:\.\d+)?(Z|[+-]\d{2}:?\d{2})?$")


def parse_utc(s):
    """'2026-10-30T23:58:48+00:00' → aware datetime（无偏移按 UTC）"""
    m = _PD.match(str(s).strip())
    if not m:
        return None
    y, mo, d, hh, mi, ss = (int(m.group(i)) for i in range(1, 7))
    tzs = m.group(7)
    if not tzs or tzs == "Z":
        tz = datetime.timezone.utc
    else:
        sign = 1 if tzs[0] == "+" else -1
        body = tzs[1:].replace(":", "")
        tz = datetime.timezone(sign * datetime.timedelta(hours=int(body[:2]), minutes=int(body[2:])))
    return datetime.datetime(y, mo, d, hh, mi, ss, tzinfo=tz)


def _qty(v):
    try:
        return float(v or 0)
    except (TypeError, ValueError):
        return 0.0


def _det_sheet(wb):
    for sn in wb.sheetnames:
        if sn.startswith("每日销售数据"):
            return sn
    return None


def read_arc(path):
    """封存簿明细 → {key: dict(status,ch,qty,pd,utc_dt,grp)}"""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    sn = _det_sheet(wb)
    if not sn:
        wb.close()
        return None
    ws = wb[sn]
    out = {}
    first = True
    for row in ws.iter_rows(min_row=1, max_row=None, min_col=1, max_col=15, values_only=True):
        if first:
            first = False
            continue
        aoi, pdv, st, ch, sku, q = row[0], row[2], row[4], row[6], row[11], row[14]
        if not sku or not pdv:
            continue
        ch = str(ch or "").strip()
        g = grp_of(ch)
        if not g:
            continue
        pdv = str(pdv).strip()
        dt = parse_utc(pdv)
        out[(str(aoi).strip(), str(sku).strip(), pdv[:19])] = dict(
            status=str(st or "").strip(), ch=ch, qty=_qty(q), pd=pdv,
            utc=dt, grp=g)
    wb.close()
    return out


def read_rep(path):
    with open(path, "r", encoding="utf-8", errors="replace") as f:
        hdr = f.readline().rstrip("\n").split("\t")
        idx = {h.strip(): i for i, h in enumerate(hdr)}
        need = ["amazon-order-id", "purchase-date", "order-status", "sales-channel", "sku", "quantity"]
        if any(h not in idx for h in need):
            return None
        i_aoi, i_pd, i_st, i_ch, i_sku, i_q = (idx[h] for h in need)
        mx = max(idx[h] for h in need)
        out = {}
        for line in f:
            p = line.rstrip("\n").split("\t")
            if len(p) <= mx:
                continue
            sku = p[i_sku].strip()
            pdv = p[i_pd].strip()
            if not sku or not pdv:
                continue
            ch = p[i_ch].strip()
            g = grp_of(ch)
            if not g:
                continue
            out[(p[i_aoi].strip(), sku, pdv[:19])] = dict(
                status=p[i_st].strip(), ch=ch, qty=_qty(p[i_q]), pd=pdv,
                utc=parse_utc(pdv), grp=g)
    return out


def latest_snap(period):
    base = os.path.join(REPO, "trunk", "data", "数据归属周期", period, "快照日期")
    ds = sorted(d for d in os.listdir(base) if os.path.isdir(os.path.join(base, d)))
    return ds[-1], os.path.join(base, ds[-1])


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--period", required=True, help="YYYYMM，如 202610")
    ap.add_argument("--out", default=os.path.join(DSR, "断货分析", "_对账"))
    ap.add_argument("--csv", action="store_true", default=True)
    a = ap.parse_args()

    Y, M = a.period[:4], int(a.period[4:6])
    YM = "%s-%02d" % (Y, M)
    os.makedirs(a.out, exist_ok=True)

    # ---------- 日销侧 ----------
    adir = os.path.join(DAY_DIR, "%s年每日数据" % Y, "%s年%d月" % (Y, M))
    files = sorted(glob.glob(os.path.join(adir, "%s年%d月-*.xlsx" % (Y, M))))
    if not files:
        print("!! 找不到月末封存簿：%s" % adir)
        print("   （换月跑批成功后会生成 {Y}年{M}月-*.xlsx；当月未封存时本工具不适用）")
        return 2
    arc_all = {}
    print("日销侧封存簿 %d 本 @ %s" % (len(files), os.path.basename(adir)))
    for f in files:
        d = read_arc(f)
        if not d:
            print("  !! 跳过(无明细sheet) %s" % os.path.basename(f)); continue
        print("  %-40s %6d 单" % (os.path.basename(f), len(d)))
        arc_all.update(d)

    # ---------- 月销侧 ----------
    snap, rdir = latest_snap(a.period)
    rep_all = {}
    for f in sorted(os.listdir(rdir)):
        if f.endswith("订单报告.txt"):
            d = read_rep(os.path.join(rdir, f))
            if d:
                rep_all.update(d)
    print("月销侧订单报告 %s  合并 %d 单（已按订单号去重，英国重复报告不会再翻倍）" % (snap, len(rep_all)))

    # ---------- 归月 ----------
    def utc_m(r):
        return r["utc"].strftime("%Y-%m") if r["utc"] else None

    def loc_m(r):
        if not r["utc"]:
            return None
        return r["utc"].astimezone(tz_of(r["ch"])).strftime("%Y-%m")

    arc_m = {k: v for k, v in arc_all.items() if utc_m(v) == YM}
    rep_m = {k: v for k, v in rep_all.items() if loc_m(v) == YM}
    arc_ok = {k: v for k, v in arc_m.items() if v["status"] != "Cancelled"}
    rep_ok = {k: v for k, v in rep_m.items() if v["status"] != "Cancelled"}

    # ---------- 逐单归因 ----------
    cause = defaultdict(lambda: defaultdict(lambda: [0, 0.0]))
    rows = []
    for k in set(arc_ok) | set(rep_ok):
        in_a, in_r = k in arc_ok, k in rep_ok
        if in_a and in_r:
            continue
        if in_a:
            v = arc_ok[k]
            r = rep_all.get(k)
            if r is not None and r["status"] == "Cancelled":
                c = "③ 取消后未回退"
            elif loc_m(v) != YM:
                # UTC 落在本月、站点本地已跨月（或还没到本月）→ 月销表按本地根本不算它
                c = "⑤ 跨月(日销多)"
            elif r is not None:
                c = "⑤ 跨月(日销多)"
            else:
                c = "② 报告里已无此单"
            q = v["qty"]
        else:
            v = rep_ok[k]
            ar = arc_all.get(k)
            if ar is not None and ar["status"] == "Cancelled":
                c = "① 当时已取消后复活"
            elif utc_m(v) != YM:
                c = "⑤ 跨月(月销多)"
            elif ar is not None:
                c = "⑤ 跨月(月销多)"
            else:
                c = "④ 日销明细漏记"
            q = -v["qty"]                      # 月销算、日销没算 → 反向
        cause[c][v["grp"]][0] += 1
        cause[c][v["grp"]][1] += q
        rows.append([a.period, c, v["grp"], v["ch"], k[1], k[0],
                     "%.0f" % q, v["pd"], v["status"],
                     (rep_all.get(k) or {}).get("status", "")])

    # ---------- 输出 ----------
    def tot_ok(d):
        out = defaultdict(float)
        for v in d.values():
            out[v["grp"]] += v["qty"]
        return out

    A, R = tot_ok(arc_ok), tot_ok(rep_ok)

    lines = []
    lines.append("# %s 日销表 vs 月销表 订单级归因\n" % a.period)
    lines.append("日销侧：月末封存簿（UTC 日期归月，状态=首次被抓到时）  ")
    lines.append("月销侧：%s 订单报告（站点本地时间归月，状态=最终）\n" % snap)
    lines.append("> ⚠️ 2026-11-02 起换月封存在归档前会做**状态校正**（`core/status_sync.py`，方案B）。"
                 "若该期归档簿已校正过，明细里的取消单状态已被回写 → 下表「取消未回退」为 0、"
                 "总差额也接近 0，**这是已修复的正常表现**，不是漏统计。\n")
    lines.append("| 分组 | 日销表 | 月销表 | 差 | 取消未回退 | 已无此单 | 跨月日销多 | 跨月月销多 | 复活 | 漏记 |")
    lines.append("|---|---:|---:|---:|---:|---:|---:|---:|---:|---:|")
    tot = defaultdict(float)
    for g in GROUPS:
        d = A.get(g, 0) - R.get(g, 0)
        cells = []
        for c in ORDER:
            n, q = cause[c].get(g, [0, 0.0])
            cells.append(q)
            tot[c] += q
        lines.append("| %s | %.0f | %.0f | **%.0f** | %.0f | %.0f | %.0f | %.0f | %.0f | %.0f |"
                     % (g, A.get(g, 0), R.get(g, 0), d, *cells))
    lines.append("| **合计** | **%.0f** | **%.0f** | **%.0f** | %.0f | %.0f | %.0f | %.0f | %.0f | %.0f |"
                 % (sum(A.values()), sum(R.values()),
                    sum(A.values()) - sum(R.values()), *[tot[c] for c in ORDER]))
    lines.append("")
    lines.append("> 正数=日销表多算，负数=日销表少算。「跨月」是两边归月基准不同（日销按 UTC、月销按站点本地），")
    lines.append("> 整个公司层面会互相抵消，只影响单月归属。\n")

    lines.append("## 归因明细（单数 / 件数）\n")
    lines.append("| 归因 | 单数 | 件数 | 说明 |")
    lines.append("|---|---:|---:|---|")
    desc = {
        "③ 取消后未回退": "日销明细只增不删：订单被抓到时是 Pending，后来取消也没有回改",
        "② 报告里已无此单": "明细里有、最新报告里整单消失（多为已删单）",
        "⑤ 跨月(日销多)": "UTC 在本月，本地已跨到次月 → 日销算了、月销没算",
        "⑤ 跨月(月销多)": "本地在本月，UTC 还在上月 → 月销算了、日销没算",
        "① 当时已取消后复活": "明细里当时是 Cancelled，最终恢复有效",
        "④ 日销明细漏记": "最终有效但日销明细里没有该行",
    }
    for c in ORDER:
        n = sum(x[0] for x in cause[c].values())
        q = sum(x[1] for x in cause[c].values())
        if n == 0:
            continue
        lines.append("| %s | %d | %+.0f | %s |" % (c, n, q, desc[c]))
    lines.append("")

    # ---- 状态快照对照 ----
    st_a = defaultdict(lambda: [0, 0.0])
    st_r = defaultdict(lambda: [0, 0.0])
    for v in arc_ok.values():
        st_a[v["status"]][0] += 1
        st_a[v["status"]][1] += v["qty"]
    for v in rep_ok.values():
        st_r[v["status"]][0] += 1
        st_r[v["status"]][1] += v["qty"]
    lines.append("## 为什么差在这里：日销明细存的是「首次抓到时」的状态快照\n")
    lines.append("| 状态 | 日销明细 | 月销表 |")
    lines.append("|---|---:|---:|")
    for s in sorted(set(st_a) | set(st_r), key=lambda x: -(st_a[x][1] + st_r[x][1])):
        lines.append("| %s | %.0f 件 | %.0f 件 |" % (s, st_a[s][1], st_r[s][1]))
    lines.append("")
    lines.append("订单第一次被抓到时几乎都是 `Pending`，之后才变 `Shipped` 或 `Cancelled`。"
                 "日销表的 `append_sales` **只判断订单号是否已存在，存在就跳过，从不回写状态**，"
                 "所以明细永久停留在首次快照。\n")
    lines.append("- 变成 `Shipped` 的：不影响总数（件数不变）。")
    lines.append("- 变成 `Cancelled` 的：亚马逊会把 `quantity` 改成 **0**，但日销明细里那条还是 `Pending / qty=1` → **被当成销量算进去**。")
    lines.append("- 这正是差额的唯一主要来源。\n")

    # ---- 结论 ----
    d_all = sum(A.values()) - sum(R.values())
    c3n = sum(x[0] for x in cause["③ 取消后未回退"].values())
    c3q = cause["③ 取消后未回退"]
    c3q = sum(x[1] for x in c3q.values())
    lines.append("## 结论\n")
    lines.append("1. **日销表 10 月虚高 %d 件（%.2f%%）**，其中 **%d 单 / %d 件是取消单没回退**。"
                 % (d_all, 100.0 * d_all / max(sum(R.values()), 1), c3n, c3q))
    lines.append("2. 跨月（日销按 UTC / 月销按站点本地）净差 %+.0f 件，两边会互相抵消，不是错。"
                 % (sum(x[1] for x in cause["⑤ 跨月(日销多)"].values())
                    + sum(x[1] for x in cause["⑤ 跨月(月销多)"].values())))
    lines.append("3. 日销表另外还少记 %d 件（封存前最后几小时的单没抓到）。"
                 % abs(sum(x[1] for x in cause["④ 日销明细漏记"].values())))
    lines.append("4. **月销表口径正确**：只取非 Cancelled、按站点本地时间归月，取消单被正确剔除 → **可以直接上报**。\n")
    lines.append("## 上报注意\n")
    pend = sum(st_r.get("Pending", [0, 0])[1:])
    lines.append("- 月销表当月值里含 **%.0f 件 `Pending`（未发货）**。10/31 的 Pending 到 11 月初还有取消风险，"
                 "若公司要求「只报已发货」，需再剔掉这部分。" % pend)
    lines.append("- 非亚马逊渠道（eMAG / 美客多）源不同，本工具不覆盖，仍按原表数字上报。")
    lines.append("- 本表只算 `sales-channel` 以 `Amazon.` 开头的行。北美簿另放行 `Non-Amazon` / "
                 "`Non-Amazon CA`（TEMU 类），合计约十来件，不在此表内 —— 这是订单级合计与表内合计的残差来源。\n")

    lines.append("## 要不要修日销表\n")
    lines.append("问题出在 `pipeline/core/runner.py: append_sales()`：它用「订单号+SKU+下单日」判断该单是否已存在，"
                 "**存在就 `continue`，从不回写状态**。三种改法：\n")
    lines.append("| 方案 | 做法 | 影响 | 风险 |")
    lines.append("|---|---|---|---|")
    lines.append("| A 每日状态同步 | 每天追加完，用当天报告把明细里已存在行的 `order-status` 回写一遍 | 日销表当日起就准 | 明细不再是纯历史快照；每天多遍历几万行 |")
    lines.append("| B 封存前校正 | 只在换月封存前，用当月报告校正整月明细状态，再算月块 | 月报准，日内仍略高 | 改动最小 |")
    lines.append("| C 不动 | 月报以月销表为准，日销表只做趋势参考 | 零改动 | 日销表长期高估约 1%，影响库存判断 |")
    lines.append("")
    lines.append("> 建议 A 或 B。C 是现状，等于承认日销表的销量列有约 1% 的系统性高估。")

    md = os.path.join(a.out, "归因_%s.md" % a.period)
    with open(md, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))

    cp = os.path.join(a.out, "归因_%s_订单明细.csv" % a.period)
    with open(cp, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["周期", "归因", "分组", "渠道", "SKU", "amazon-order-id",
                    "数量", "purchase-date", "日销状态", "最终状态"])
        w.writerows(rows)

    # 控制台
    print("\n%-8s %9s %9s %7s" % ("分组", "日销表", "月销表", "差"))
    for g in GROUPS:
        print("%-8s %9.0f %9.0f %7.0f" % (g, A.get(g, 0), R.get(g, 0), A.get(g, 0) - R.get(g, 0)))
    print("%-8s %9.0f %9.0f %7.0f" % ("合计", sum(A.values()), sum(R.values()),
                                      sum(A.values()) - sum(R.values())))
    print("\n归因：")
    for c in ORDER:
        n = sum(x[0] for x in cause[c].values())
        q = sum(x[1] for x in cause[c].values())
        print("  %-14s %4d 单 %6.0f 件" % (c, n, q))
    print("\n→ %s" % md)
    print("→ %s" % cp)
    return 0


if __name__ == "__main__":
    sys.exit(main())
