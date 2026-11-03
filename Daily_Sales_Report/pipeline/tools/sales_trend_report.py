# -*- coding: utf-8 -*-
"""销量趋势月报（只读，不写任何工作簿）。

回答「这么多数据，怎么看销量变化」：
  ① 异动清单  —— 环比涨/跌超阈值且基数够的 SKU，按影响件数排序（最实用的一张）
  ② 连跌预警  —— 连续 N 个月下滑且仍有动销
  ③ 新品追踪  —— 首销在最近几个月内的 SKU 爬坡曲线
  ④ 汇总      —— 每簿/每表/每站点的月度合计与环比
  ⑤ 趋势明细  —— 全部 SKU 的 1~当月序列（兜底查证用）

数据源 = 13 本工作簿分析表 + 欧盟簿 10 张站点表的**月度块**（与表内口径永远一致）。
月度块是跑批封存的静态值，不重算、不推导。

⚠️ 读格必须看 t 属性：月度块里有大量 t="s" 的格（<v> 是共享串索引不是数值，
   欧盟簿索引可达 16 万+）。本模块的 rows_of() 已按类型过滤。

用法：
    python sales_trend_report.py                      # 默认最新快照、输出到 销量趋势/
    python sales_trend_report.py --snap 20261028
    python sales_trend_report.py --no-site            # 不含站点表
    python sales_trend_report.py --min-base 10 --mom 0.5
"""
import argparse
import glob
import os
import re
import sys
import zipfile
from statistics import mean

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "..", "core"))
sys.path.insert(0, os.path.join(HERE, "maintenance"))

import config as C
import importlib.util

_spec = importlib.util.spec_from_file_location(
    "rb", os.path.join(HERE, "maintenance", "rebuild_monthly_from_daily_20261027.py"))
RB = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(RB)

CELL = re.compile(r'<c r="([A-Z]+)(\d+)"([^>]*?)(?:/>|>(.*?)</c>)', re.S)
VAL = re.compile(r"<v>(-?[\d.]+)</v>")
NUMTYPES = re.compile(r't="(s|str|d|inlineStr)"')


def smap2(z):
    """sheet 名 → 部件路径。属性顺序无关（openpyxl 重写后会变）。"""
    rels = z.read("xl/_rels/workbook.xml.rels").decode()
    rid2t = {}
    for rel in re.findall(r"<Relationship\b[^>]*/?>", rels):
        i = re.search(r'Id="([^"]+)"', rel)
        t = re.search(r'Target="([^"]+)"', rel)
        if i and t:
            t = t.group(1).lstrip("/")
            rid2t[i.group(1)] = t if t.startswith("xl/") else "xl/" + t
    wbx = z.read("xl/workbook.xml").decode()
    return {n: rid2t[r] for n, r in
            re.findall(r'<sheet [^>]*?name="([^"]+)"[^>]*?r:id="(rId\d+)"', wbx)}


def rows_of(x, sst):
    """单遍产出 (行号, {列: 值})。只收真正的数值；文本格单独收进 '_txt'。"""
    for rm in re.finditer(r'<row r="(\d+)"[^>]*>(.*?)</row>', x, re.S):
        r = int(rm.group(1))
        cells, txt = {}, {}
        for c in CELL.finditer(rm.group(2)):
            col, attrs, body = c.group(1), c.group(3), c.group(4) or ""
            if 't="s"' in attrs:
                v = VAL.search(body)
                if v and int(v.group(1)) < len(sst):
                    txt[col] = sst[int(v.group(1))]
                continue
            if 't="inlineStr"' in attrs:
                txt[col] = "".join(re.findall(r"<t[^>]*>(.*?)</t>", body, re.S))
                continue
            if NUMTYPES.search(attrs):
                continue
            v = VAL.search(body)
            if v:
                cells[col] = float(v.group(1))
        yield r, cells, txt


def parse_book(path, cur_m, with_site=True):
    """返回 (sheet 列表, 记录列表)。记录 = dict(sheet, sku, name, months[1..12])。"""
    z = zipfile.ZipFile(path)
    sm = smap2(z)
    sst = RB.sst_of(z)
    out, sheets = [], []
    for nm, sp in sm.items():
        nm = nm.strip()
        is_site = (not with_site) and (nm in getattr(C, "SITE_SHEET_CHANNELS", {}))
        if "分析表" not in nm and nm not in getattr(C, "SITE_SHEET_CHANNELS", {}):
            continue
        if is_site:
            continue
        x = z.read(sp).decode("utf-8")
        try:
            cols = RB.month_cols(x, sst)
        except Exception:
            cols = None
        if not cols or len(cols) < 12:
            continue
        mcols = [cols[m] for m in range(1, 13)]
        kind = "分析表" if "分析表" in nm else "站点表"
        n_rec = 0
        for r, cells, txt in rows_of(x, sst):
            if r < 7:
                continue
            sku = txt.get("A") or ""
            sku = str(sku).strip()
            if not sku:
                continue
            months = [int(cells.get(c) or 0) for c in mcols]
            if sum(months) == 0:
                continue                      # 全 0 行不进趋势（新建行/未上架）
            out.append({
                "sheet": nm, "kind": kind, "sku": sku,
                "name": (txt.get("B") or "").strip(), "months": months,
            })
            n_rec += 1
        sheets.append({"sheet": nm, "kind": kind, "n": n_rec})
    z.close()
    return sheets, out


def metrics(rec, cur_m, days_cur=30, days_prev=30):
    m = rec["months"]
    cur = m[cur_m - 1]
    prev = m[cur_m - 2] if cur_m >= 2 else 0
    d = cur - prev
    # ⚠️ 本月只累计到快照日（MTD），上月是整月 → 直接比件数会一律显示"跌"。
    # 环比一律按**日均**算；件数差（d）只用于排序与展示。
    avg_cur = cur / days_cur if days_cur else cur
    avg_prev = prev / days_prev if days_prev else prev
    pct = ((avg_cur - avg_prev) / avg_prev) if avg_prev else None
    w3 = [m[i] for i in range(max(0, cur_m - 3), cur_m)]
    avg3 = mean(w3) if w3 else 0
    # 连跌：从当月往前数连续严格下降的月数（上一月为 0 就停）
    drop = 0
    i = cur_m - 1
    while i >= 1 and m[i] < m[i - 1] and m[i - 1] > 0:
        drop += 1
        i -= 1
    first = next((k + 1 for k, v in enumerate(m) if v > 0), None)
    rec.update(cur=cur, prev=prev, d=d, pct=pct, avg3=avg3, drop=drop,
               avg_cur=avg_cur, avg_prev=avg_prev,
               first_month=first, peak=max(m), peak_month=m.index(max(m)) + 1,
               total=sum(m))
    return rec


def build(snap, with_site, min_base, mom):
    import calendar
    cur_m = int(snap[4:6])
    day = int(snap[6:])
    py, pm = (2026, cur_m - 1) if cur_m > 1 else (2025, 12)
    days_prev = calendar.monthrange(py, pm)[1]
    books = []
    for f in sorted(glob.glob(os.path.join(C.WORK_DIR, "*.xlsx"))):
        b = os.path.basename(f)
        if ".strip" in b or ".tmp" in b or b.startswith("~$"):
            continue
        sheets, recs = parse_book(f, cur_m, with_site)
        for r in recs:
            r["book"] = b
            metrics(r, cur_m, day, days_prev)
        books.append({"book": b, "sheets": sheets, "recs": recs})
    allrec = [r for b in books for r in b["recs"]]

    # 异动：基数够 + 变化幅度够
    movers = []
    for r in allrec:
        base = max(r["cur"], r["prev"])
        if base < min_base:
            continue
        if r["prev"] == 0:
            movers.append(dict(r, kind2="新增", pct=None))
            continue
        if r["pct"] is not None and abs(r["pct"]) >= mom:
            movers.append(dict(r, kind2="涨" if r["d"] > 0 else ("归零" if r["cur"] == 0 else "跌")))
    movers.sort(key=lambda r: -abs(r["d"]))

    drops = [r for r in allrec if r["drop"] >= 3 and r["cur"] > 0 and r["avg3"] >= 1]
    drops.sort(key=lambda r: (-r["drop"], -r["avg3"]))

    newbies = [r for r in allrec if r["first_month"] and r["first_month"] >= max(1, cur_m - 2)]
    newbies.sort(key=lambda r: (r["first_month"], -r["total"]))

    return books, allrec, movers, drops, newbies


def sheet_totals(books, cur_m, day=30, days_prev=30):
    rows = []
    for b in books:
        for s in b["sheets"]:
            recs = [r for r in b["recs"] if r["sheet"] == s["sheet"]]
            months = [sum(r["months"][i] for r in recs) for i in range(12)]
            cur = months[cur_m - 1]
            prev = months[cur_m - 2] if cur_m >= 2 else 0
            ac = cur / day if day else cur
            ap = prev / days_prev if days_prev else prev
            rows.append({
                "book": b["book"], "sheet": s["sheet"], "kind": s["kind"], "n": len(recs),
                "months": months, "cur": cur, "prev": prev,
                "d": cur - prev, "avg_cur": ac, "avg_prev": ap,
                "pct": ((ac - ap) / ap) if ap else None,
            })
    return rows


def write_xlsx(path, snap, cur_m, movers, drops, newbies, totals, allrec, min_base, mom):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter

    HDR = Font(bold=True, color="FFFFFF")
    FILL = PatternFill("solid", fgColor="4472C4")
    RED = PatternFill("solid", fgColor="FFC7CE")
    GREEN = PatternFill("solid", fgColor="C6EFCE")
    YEL = PatternFill("solid", fgColor="FFEB9C")
    wb = Workbook()

    def head(ws, cols, widths=None):
        ws.append(cols)
        for c in range(1, len(cols) + 1):
            cell = ws.cell(1, c)
            cell.font = HDR
            cell.fill = FILL
            cell.alignment = Alignment(horizontal="center")
        if widths:
            for i, w in enumerate(widths, 1):
                ws.column_dimensions[get_column_letter(i)].width = w
        ws.freeze_panes = "A2"

    # ① 异动清单
    ws = wb.active
    ws.title = "异动清单"
    cols = ["类型", "簿", "表", "SKU", "品名"] + ["%d月" % m for m in range(1, cur_m + 1)] + \
           ["本月", "上月", "变动", "本月日均", "上月日均", "日均环比%"]
    head(ws, cols, [7, 22, 18, 12, 26] + [7] * cur_m + [8, 8, 8, 9, 9, 10])
    for r in movers:
        ws.append([r["kind2"], r["book"], r["sheet"], r["sku"], r["name"]] +
                  [r["months"][m - 1] for m in range(1, cur_m + 1)] +
                  [r["cur"], r["prev"], r["d"],
                   round(r["avg_cur"], 2), round(r["avg_prev"], 2),
                   "" if r["pct"] is None else round(r["pct"] * 100, 1)])
        i = ws.max_row
        fill = GREEN if r["d"] > 0 else (RED if r["d"] < 0 else YEL)
        ws.cell(i, 1).fill = fill
        ws.cell(i, 4).alignment = Alignment(horizontal="left")

    # ② 连跌预警
    ws = wb.create_sheet("连跌预警")
    head(ws, ["连跌月数", "簿", "表", "SKU", "品名"] + ["%d月" % m for m in range(1, cur_m + 1)] +
         ["近3月均值", "峰值", "峰值月"],
         [9, 22, 18, 12, 26] + [7] * cur_m + [9, 7, 8])
    for r in drops:
        ws.append([r["drop"], r["book"], r["sheet"], r["sku"], r["name"]] +
                  [r["months"][m - 1] for m in range(1, cur_m + 1)] +
                  [round(r["avg3"], 1), r["peak"], "%d月" % r["peak_month"]])
        ws.cell(ws.max_row, 1).fill = RED

    # ③ 新品追踪
    ws = wb.create_sheet("新品追踪")
    head(ws, ["首销月", "簿", "表", "SKU", "品名"] + ["%d月" % m for m in range(1, cur_m + 1)] +
         ["累计", "近3月均值"], [8, 22, 18, 12, 26] + [7] * cur_m + [8, 9])
    for r in newbies:
        ws.append(["%d月" % r["first_month"], r["book"], r["sheet"], r["sku"], r["name"]] +
                  [r["months"][m - 1] for m in range(1, cur_m + 1)] +
                  [r["total"], round(r["avg3"], 1)])
        ws.cell(ws.max_row, 1).fill = YEL

    # ④ 汇总
    ws = wb.create_sheet("汇总")
    head(ws, ["类型", "簿", "表", "SKU数"] + ["%d月" % m for m in range(1, cur_m + 1)] +
         ["本月", "上月", "变动", "本月日均", "上月日均", "日均环比%"],
         [7, 26, 20, 7] + [8] * cur_m + [8, 8, 8, 9, 9, 10])
    for t in sorted(totals, key=lambda t: (t["kind"], t["book"], t["sheet"])):
        ws.append([t["kind"], t["book"], t["sheet"], t["n"]] +
                  [t["months"][m - 1] for m in range(1, cur_m + 1)] +
                  [t["cur"], t["prev"], t["d"],
                   round(t["avg_cur"], 2), round(t["avg_prev"], 2),
                   "" if t["pct"] is None else round(t["pct"] * 100, 1)])
        ws.cell(ws.max_row, 1).fill = YEL if t["kind"] == "站点表" else GREEN

    # ⑤ 趋势明细
    ws = wb.create_sheet("趋势明细")
    head(ws, ["簿", "表", "SKU", "品名"] + ["%d月" % m for m in range(1, cur_m + 1)] +
         ["累计", "本月", "上月", "变动", "环比%", "近3月均值", "连跌", "首销月"],
         [26, 20, 12, 26] + [7] * cur_m + [8, 8, 8, 8, 9, 9, 6, 8])
    for r in sorted(allrec, key=lambda r: (r["book"], r["sheet"], -r["total"])):
        ws.append([r["book"], r["sheet"], r["sku"], r["name"]] +
                  [r["months"][m - 1] for m in range(1, cur_m + 1)] +
                  [r["total"], r["cur"], r["prev"], r["d"],
                   "" if r["pct"] is None else round(r["pct"] * 100, 1),
                   round(r["avg3"], 1), r["drop"], r["first_month"]])

    wb.save(path)
    return path


def write_md(path, snap, cur_m, movers, drops, newbies, totals, min_base, mom):
    up = [r for r in movers if r["d"] > 0][:15]
    dn = [r for r in movers if r["d"] < 0][:15]
    L = []
    L.append("# 销量趋势月报 %s\n" % snap)
    L.append("口径：各表月度块（跑批封存的静态值，与表内一致）；本月 = %d 月（截至 %s 日），上月 = %d 月整月。\n"
             % (cur_m, snap[6:], cur_m - 1))
    L.append("⚠️ 本月是**月内累计（MTD）**，不能与上月整月直接比件数——"
             "所以环比一律按**日均**计算（本月日均 vs 上月日均）。\n")
    L.append("异动筛选：基数 max(本月,上月) ≥ %d 件，且日均环比变化 ≥ %.0f%%。\n" % (min_base, mom * 100))

    L.append("\n## 一、大盘（按表）\n")
    L.append("| 类型 | 簿 | 表 | SKU数 | %d月 | %d月 | 本月日均 | 上月日均 | 日均环比 |" % (cur_m, cur_m - 1))
    L.append("|---|---|---|---|---|---|---|---|---|")
    for t in sorted(totals, key=lambda t: -abs(t["d"]))[:14]:
        L.append("| %s | %s | %s | %d | %d | %d | %.1f | %.1f | %s |" % (
            t["kind"], t["book"].replace("库存销售表.xlsx", ""), t["sheet"].replace("-每日库存分析表", ""),
            t["n"], t["cur"], t["prev"], t["avg_cur"], t["avg_prev"],
            "—" if t["pct"] is None else "%+.0f%%" % (t["pct"] * 100)))

    L.append("\n## 二、涨得最多（Top 15，按变动件数）\n")
    L.append("| SKU | 品名 | 表 | 上月 | 本月 | 变动 | 日均环比 |")
    L.append("|---|---|---|---|---|---|---|")
    for r in up:
        L.append("| %s | %s | %s | %d | %d | %+d | %s |" % (
            r["sku"], r["name"][:24], r["sheet"].replace("-每日库存分析表", ""),
            r["prev"], r["cur"], r["d"],
            "新增" if r["pct"] is None else "%+.0f%%" % (r["pct"] * 100)))

    L.append("\n## 三、跌得最多（Top 15，按变动件数）\n")
    L.append("| SKU | 品名 | 表 | 上月 | 本月 | 变动 | 日均环比 |")
    L.append("|---|---|---|---|---|---|---|")
    for r in dn:
        L.append("| %s | %s | %s | %d | %d | %+d | %s |" % (
            r["sku"], r["name"][:24], r["sheet"].replace("-每日库存分析表", ""),
            r["prev"], r["cur"], r["d"],
            "—" if r["pct"] is None else "%+.0f%%" % (r["pct"] * 100)))

    L.append("\n## 四、连跌 ≥3 个月（仍有动销）\n")
    if drops:
        L.append("| 连跌 | SKU | 品名 | 表 | 近3月均值 | 序列 |")
        L.append("|---|---|---|---|---|---|")
        for r in drops[:20]:
            seq = "→".join(str(r["months"][m - 1]) for m in range(max(1, cur_m - 5), cur_m + 1))
            L.append("| %d | %s | %s | %s | %.1f | %s |" % (
                r["drop"], r["sku"], r["name"][:20], r["sheet"].replace("-每日库存分析表", ""),
                r["avg3"], seq))
    else:
        L.append("无。")

    L.append("\n## 五、新品（最近 3 个月首销）\n")
    if newbies:
        L.append("| 首销 | SKU | 品名 | 表 | 累计 | 序列 |")
        L.append("|---|---|---|---|---|---|")
        for r in newbies[:25]:
            seq = "→".join(str(r["months"][m - 1]) for m in range(max(1, r["first_month"]), cur_m + 1))
            L.append("| %d月 | %s | %s | %s | %d | %s |" % (
                r["first_month"], r["sku"], r["name"][:20],
                r["sheet"].replace("-每日库存分析表", ""), r["total"], seq))
    else:
        L.append("无。")

    L.append("\n---\n")
    L.append("完整数据见同目录 xlsx（异动清单 / 连跌预警 / 新品追踪 / 汇总 / 趋势明细）。")
    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(L))
    return path


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--snap", default="")
    ap.add_argument("--out", default="")
    ap.add_argument("--min-base", type=int, default=10)
    ap.add_argument("--mom", type=float, default=0.5)
    ap.add_argument("--no-site", action="store_true")
    a = ap.parse_args()

    snap = a.snap
    if not snap:
        cands = sorted(glob.glob(os.path.join(C.SNAP_BASE, "*", "快照日期", "2026*")))
        snap = os.path.basename(cands[-1]) if cands else ""
    if not snap:
        raise SystemExit("找不到快照日，请用 --snap 指定")
    cur_m = int(snap[4:6])

    outdir = a.out or os.path.join(C.DSR, "销量趋势")
    os.makedirs(outdir, exist_ok=True)

    books, allrec, movers, drops, newbies = build(snap, not a.no_site, a.min_base, a.mom)
    import calendar
    day = int(snap[6:])
    py, pm = (2026, cur_m - 1) if cur_m > 1 else (2025, 12)
    totals = sheet_totals(books, cur_m, day, calendar.monthrange(py, pm)[1])

    xp = os.path.join(outdir, "销量趋势月报-%s.xlsx" % snap)
    mp = os.path.join(outdir, "销量趋势月报-%s.md" % snap)
    write_xlsx(xp, snap, cur_m, movers, drops, newbies, totals, allrec, a.min_base, a.mom)
    write_md(mp, snap, cur_m, movers, drops, newbies, totals, a.min_base, a.mom)

    print("快照 %s｜本月 %d 月" % (snap, cur_m))
    print("扫描 %d 本 / %d 张表 / %d 个有销量 SKU" % (
        len(books), sum(len(b["sheets"]) for b in books), len(allrec)))
    print("异动 %d（涨 %d / 跌 %d）｜连跌≥3 %d｜新品 %d" % (
        len(movers), sum(1 for r in movers if r["d"] > 0),
        sum(1 for r in movers if r["d"] < 0), len(drops), len(newbies)))
    print("已输出:\n  %s\n  %s" % (xp, mp))


if __name__ == "__main__":
    main()
