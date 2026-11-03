# -*- coding: utf-8 -*-
"""东南亚「在售但无库存无在途」SKU —— 断货前 60 天日销明细。

需求来源（2026-10-22 采购决策）：找出东南亚在售、但库存 0 且在途 0 的 SKU，
把它们**最早进入断货状态之前 60 天**的销售数据按 SKU/型号/当月评级/国家/日销量列出，
按（国家、月份）分 sheet。

口径（与用户逐条确认过）：
  在售   印尼用「印尼状态」= 在售；泰国/菲律宾表里没有状态列 → 用「近 60 天有销量」判定
  断货   在库库存 ≤ 0 且 在途 ≤ 0（在途 = 剩余库存+在途库存 − 在库库存）
  起点   日销表每月只有一张库存快照、没有每日库存，故用「该 SKU 最后一次有销量的日期 + 1」
         作为断货起点（销量停止即断货，这是唯一能从现有数据精确算出的代理指标）
         ⚠ 加了 --earliest 后改用「最早进入 0/0 状态的那个月」内的最后销量日，见 earliest_zero_day()
  明细   [最后销量日 − 59, 最后销量日]，逐日列出各渠道销量
  评级   当月销量 A≥100 / B≥30 / C<30（表内没有现成评级列）

数据源：`Daily_Sales_Report/东南亚日销/` 三本人工维护表
  - 月表（当前快照）：`2026年10月` / `10月`，取 SKU/型号/在库库存/剩余+在途/状态
  - 订单明细：`10月虾皮订单` / `10月抖音订单` / `10月易仓LAZADA订单` / `10月线下订单` …
    （五类表头变体，一律**按列名定位**，不按列号硬编码）

用法：
    python pipeline/tools/sea_stockout_detail.py
    python pipeline/tools/sea_stockout_detail.py --days 60 --out <目录>
    python pipeline/tools/sea_stockout_detail.py --min-days 90   # 泰菲「在售」判据改成近90天有销量
    python pipeline/tools/sea_stockout_detail.py --earliest     # 断货起点改「最早进入 0/0」

输出：`Daily_Sales_Report/断货分析/东南亚/{明细|最早口径}/<YYYYMMDD>/`
    （2026-10-23 起从 `Daily_Sales_Report/东南亚断货分析/` 迁入）
"""
import os
import re
import sys
import argparse
from collections import defaultdict, Counter
from datetime import datetime, date, timedelta

_HERE = os.path.dirname(os.path.abspath(__file__))
_CORE = os.path.join(os.path.dirname(_HERE), "core")
for _p in (_CORE, _HERE):
    if _p not in sys.path:
        sys.path.insert(0, _p)

from openpyxl import Workbook, load_workbook            # noqa: E402
from openpyxl.styles import Font, PatternFill, Alignment  # noqa: E402
from openpyxl.utils import get_column_letter            # noqa: E402

_DSR = os.path.dirname(os.path.dirname(_HERE))          # …\Daily_Sales_Report
SEA_DIR = os.path.join(_DSR, "东南亚日销")
# 断货相关产出统一收在 断货分析/ 下，一级按业务块分（亚马逊 / 东南亚）
SEA_OUT = os.path.join(_DSR, "断货分析", "东南亚")

# ---------------- 三本表配置 ----------------
# month_re : 月表 sheet 名（n=月份数字）
# cols     : 0-based 列位（三本表列序不同，见探查）
# status   : 在售状态列（None = 该表没有，用「近 N 天有销量」代替）
SITES = [
    dict(site="印尼", file="2026年印尼智慧物流销售表1-12月.xlsx",
         month_re=r"^2026年(\d{1,2})月$",
         cols=dict(sku=0, name=1, model=2, avail=11, total=12),
         status_col=9, onsale={"在售"}),
    dict(site="泰国", file="2026年泰国易仓销售表1-12月.xlsx",
         month_re=r"^(\d{1,2})月$",
         cols=dict(sku=0, name=1, model=2, avail=10, total=11),
         status_col=None, onsale=set()),
    dict(site="菲律宾", file="2026年菲律宾仓销售表1-12月.xlsx",
         month_re=r"^(\d{1,2})月$",
         cols=dict(sku=0, name=1, model=2, avail=10, total=11),
         status_col=None, onsale=set()),
]

RATING_A, RATING_B = 100, 30          # 当月销量 ≥100 → A，≥30 → B，其余 C
_CANCEL = re.compile(r"cancel|storno|取消|batal|ยกเลิก", re.I)

# ---------------- 基础 ----------------
def _num(v):
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(",", "")
    if not s or s.startswith("#"):
        return None
    try:
        return float(s)
    except ValueError:
        return None


def _date10(v):
    """各种日期写法 → 'YYYY-MM-DD'；认不出返回 None。"""
    if v is None:
        return None
    if isinstance(v, datetime):
        return v.strftime("%Y-%m-%d")
    s = str(v).strip()
    m = re.match(r"^(\d{4})[/-](\d{1,2})[/-](\d{1,2})", s)
    if m:
        try:
            return f"{m.group(1)}-{int(m.group(2)):02d}-{int(m.group(3)):02d}"
        except ValueError:
            return None
    return None


def _rating(qty):
    if qty >= RATING_A:
        return "A"
    if qty >= RATING_B:
        return "B"
    return "C"


# ---------------- 订单明细（五类表头变体，按列名定位） ----------------
def _find_col(hdr, keys, exclude=()):
    for k in keys:
        for j, h in enumerate(hdr):
            hl = (h or "").lower()
            if k in hl and not any(e in hl for e in exclude):
                return j
    return None


def read_order_sheet(ws):
    """订单明细 sheet → [(date10, sku, qty, channel_kind, cancelled)]。

    表头有两种形态：单行（虾皮/线下/LAZADA）和**双行**（抖音：r1 只有 Order ID 起的中文，
    r2 才是 日期/有效性）。统一做法：先找「第一行 col0 能解析成日期」的行当数据起点，
    再把起点之前所有行的非空值按列合并成表头。
    """
    maxc = min(ws.max_column or 0, 30)
    if not maxc:
        return []
    rows = list(ws.iter_rows(min_row=1, max_row=min(ws.max_row or 0, 20000),
                             max_col=maxc, values_only=True))
    if not rows:
        return []
    # ① 先用前 6 行做「预合并表头」（双行表头的抖音：r1 只有部分列，r2 才有 日期/有效性）
    width = max((len(r) for r in rows[:6]), default=0)
    pre = []
    for j in range(width):
        v = ""
        for r in rows[:6]:
            if j < len(r) and r[j] is not None and str(r[j]).strip():
                v = str(r[j]).strip()
                break
        pre.append(v)

    # ② 定「日期列」：优先表头含「日期/date」，找不到再退回「哪一列能解析成日期」
    #    ⚠ LAZADA 表是 `数量|日期|有效性|orderNumber|createTime|sellerSku|status`，
    #      日期在**第 2 列**。2026-10-23 实测：写死第 1 列会让泰国/菲律宾的 LAZADA
    #      整张表被判成「无数据」而静默丢掉（当时 326 / 340 行全丢）。
    i_date = None
    for j, h in enumerate(pre):
        hl = h.lower()
        if "日期" in h or "date" in hl:
            i_date = j
            break
    if i_date is None:
        for j in range(width):
            if any(j < len(r) and _date10(r[j]) for r in rows[:6]):
                i_date = j
                break
    if i_date is None:
        return []

    # ③ 数据起点 = 日期列第一次能解析成日期的那一行
    start = None
    for i, r in enumerate(rows[:8]):
        if r and len(r) > i_date and _date10(r[i_date]):
            start = i
            break
    if start is None:
        return []

    # ④ 正式表头 = 数据起点之前各行按列合并
    hdr = []
    for j in range(width):
        v = ""
        for r in rows[:start]:
            if j < len(r) and r[j] is not None and str(r[j]).strip():
                v = str(r[j]).strip()
                break
        hdr.append(v)

    i_sku = _find_col(hdr, ["nomor referensi", "seller sku", "sellersku",
                            "sku reference", "商品编码", "sku"])
    i_qty = _find_col(hdr, ["jumlah", "quantity", "จำนวน", "数量"])
    i_st = _find_col(hdr, ["status", "订单状态"])
    i_ok = _find_col(hdr, ["有效性"])
    if i_sku is None or i_qty is None:
        return []

    out = []
    for r in rows[start:]:
        if not r or len(r) <= max(i_sku, i_qty, i_date):
            continue
        d = _date10(r[i_date])
        if not d:
            continue
        sku = str(r[i_sku]).strip() if r[i_sku] is not None else ""
        if not sku:
            continue
        if i_ok is not None and i_ok < len(r) and r[i_ok] is not None:
            if str(r[i_ok]).strip() not in ("1", "1.0", ""):
                continue
        st = str(r[i_st]).strip() if (i_st is not None and i_st < len(r) and r[i_st]) else ""
        q = _num(r[i_qty]) or 0.0
        out.append((d, sku, q, bool(_CANCEL.search(st))))
    return out


def channel_kind(sheet_name):
    n = sheet_name.lower()
    for k in ("lazada", "抖音", "虾皮", "线下"):
        if k in n:
            return "LAZADA" if k == "lazada" else k
    return "其他"


def read_all_orders(wb):
    """整本簿所有「N月…订单」sheet → {(sku, date10): Counter(channel→qty)}，已剔除取消行。"""
    daily = defaultdict(Counter)
    for sn in wb.sheetnames:
        if "订单" not in sn or not re.match(r"^\d{1,2}月", sn):
            continue
        kind = channel_kind(sn)
        for d, sku, q, cancelled in read_order_sheet(wb[sn]):
            if cancelled or not q:
                continue
            daily[(sku, d)][kind] += q
    return daily


# ---------------- 月表（当前快照） ----------------
def latest_month_sheet(wb, pattern):
    """→ (sheet 名, 月份数字)；取编号最大的那张。"""
    best = (None, 0)
    rx = re.compile(pattern)
    for sn in wb.sheetnames:
        m = rx.match(sn.strip())
        if m:
            n = int(m.group(1))
            if n > best[1]:
                best = (sn, n)
    return best


# 惠欢三国月表里的评级词表（A/B/C/D/E + X 前缀新品档）
RATING_VOCAB = {"A", "B", "C", "D", "E", "X", "XA", "XB", "XC", "XD", "XE"}


def _find_rating_col(ws, sample=40, scan_right=6):
    """定位「评级」列 —— 惠欢三国月表里这一列**没有表头**，只能靠值形态找。

    2026-10-29 实测（三份表列数各不相同，但规律一致）：
      印尼 31日=c57 → 评级 c60；泰国 31日=c54 → 评级 c57；菲律宾 31日=c56 → 评级 c59
      即「31 日」列右边第 3 列、表头为空；只有 9 月和 10 月两张表有值，6-8 月整列为空。
    ⚠ 原来的代码不知道这列存在，用 `_rating(当月销量)` 现算 A/B/C —— 算出来的字母和表里
    真实的评级（A~E + X 前缀）不是同一个体系，会误导补货判断。现在优先读真列，读不到才回退。
    """
    rows = list(ws.iter_rows(min_row=1, max_row=sample + 2, values_only=True))
    if not rows:
        return None
    hdr = rows[0]

    def looks_like(j):
        vals = [str(r[j]).strip() for r in rows[2:]
                if j < len(r) and r[j] is not None and str(r[j]).strip()]
        # ⚠️ 未评级的格子会被写成数字 0，混在 A~E / X* 里。不剔掉会让 all() 校验失败，
        #    整个评级列识别不到，静默退回 _rating(当月销量) 现算 —— 那套字母是错的。
        vals = [v for v in vals if v not in ("0", "0.0")]
        return (len(vals) >= 10 and all(v in RATING_VOCAB for v in vals)
                and len(set(vals)) >= 2)

    # ① 主路径：锚「31 日」列，往右找第一个空表头的评级列
    for j, h in enumerate(hdr):
        if str(h).strip() == "31":
            for k in range(j + 1, min(j + 1 + scan_right, len(hdr))):
                if not str(hdr[k] or "").strip() and looks_like(k):
                    return k
    # ② 兜底：全表扫空表头列
    for j in range(len(hdr)):
        if not str(hdr[j] or "").strip() and looks_like(j):
            return j
    return None


def read_current_stock(ws, cfg):
    """月表 → [(sku, 型号, 品名, 在库库存, 在途, 状态, 评级)]"""
    c = cfg["cols"]
    rc = _find_rating_col(ws)
    maxc = max(c["total"], c["sku"], rc if rc is not None else 0) + 1
    out = []
    for row in ws.iter_rows(min_row=3, max_row=min(ws.max_row or 0, 2000),
                            max_col=maxc, values_only=True):
        if not row or len(row) <= c["total"]:
            continue
        sku = str(row[c["sku"]]).strip() if row[c["sku"]] is not None else ""
        if not sku or sku in ("合计", "SKU", "泰国SKU"):
            continue
        avail = _num(row[c["avail"]])
        total = _num(row[c["total"]])
        if avail is None or total is None:
            continue
        inbound = total - avail
        status = ""
        if cfg["status_col"] is not None and len(row) > cfg["status_col"] and row[cfg["status_col"]]:
            status = str(row[cfg["status_col"]]).strip()
        out.append(dict(sku=sku,
                        model=str(row[c["model"]]).strip() if row[c["model"]] else "",
                        name=str(row[c["name"]]).strip() if row[c["name"]] else "",
                        avail=avail, inbound=inbound, status=status,
                        rating=(str(row[rc]).strip()
                                if rc is not None and len(row) > rc and row[rc] is not None
                                else "")))
    return out


# ---------------- 历史月表（判「最早进入 0/0」） ----------------
def read_stock_history(wb, cfg):
    """{(sku, 月份号): (在库, 在途, 评级)} —— 各月月表的库存快照 + 当月评级。

    月表 `L 在库库存` 是 `VLOOKUP(上月表) − 本月销量 + 到货 − 退货` 的实时值，
    所以每张月表都留着「那个月当时的库存」；把它们串起来就能看出某 SKU 哪个月起就 0/0 了。
    2026-10-29 起顺带取各月评级（6-8 月该列为空 → 空串），供明细的「当月销售评级」用。
    """
    rx = re.compile(cfg["month_re"])
    hist = {}
    for sn in wb.sheetnames:
        m = rx.match(sn.strip())
        if not m:
            continue
        mn = int(m.group(1))
        for r in read_current_stock(wb[sn], cfg):
            hist[(r["sku"], mn)] = (r["avail"], r["inbound"], r["rating"])
    return hist


def _months_in_window(snap_dt, window_days):
    """快照日往前 window_days 天覆盖到的月份号（升序，只算同年）。"""
    lo = (snap_dt - timedelta(days=window_days)).date()
    out = []
    for mn in range(1, 13):
        me = (date(snap_dt.year, mn + 1, 1) - timedelta(days=1)) if mn < 12 \
            else date(snap_dt.year, 12, 31)
        if me >= lo:
            out.append(mn)
    return out


def earliest_zero_day(sku, hist, last_by_month, snap_dt, window_days=120):
    """该 SKU **最早**「在库=0 且 在途=0」的月份 → (月份号, 该月内最后销量日)。

    需求原文要「最早进入此状态」，而「最后一次有销量日」只给出**最近**一次断货。
    2026-10-23 实测两处不一致：
      菲律宾 G487 —— 08-08 卖完，10-09 一笔「线下 24 件」→ 按最近算窗口内只剩 1 天销量（失真）
                    按最早算 = 8 月（0/0）→ 窗口 06-10~08-08，34 件 / 22 天，评级 C→B
      菲律宾 G518 —— 8/9 月在途 3 件，不是 0/0；最早 0/0 是 10 月 → 仍是 3 件，不变
    历史月表的「在途」正是区分这两种情况的关键。
    """
    for mn in _months_in_window(snap_dt, window_days):
        av_ib = hist.get((sku, mn))
        if not av_ib:
            continue
        if av_ib[0] > 0 or av_ib[1] > 0:
            continue
        d = last_by_month.get((sku, mn))
        if d:
            return mn, d
    return None, None


# ---------------- 主流程 ----------------
def build(days=60, min_days_sold=60, verbose=True, earliest=False, earliest_window=120):
    """→ (detail_rows, summary_rows, snap_date, notes)"""
    detail, summary, notes = [], [], []
    snap_date = None

    for cfg in SITES:
        path = os.path.join(SEA_DIR, cfg["file"])
        if not os.path.exists(path):
            notes.append(f"✗ 缺文件：{cfg['file']}")
            continue
        wb = load_workbook(path, read_only=True, data_only=True)
        try:
            msn, month = latest_month_sheet(wb, cfg["month_re"])
            if not msn:
                notes.append(f"✗ {cfg['site']}：找不到月表（pattern {cfg['month_re']}）")
                continue
            stock = read_current_stock(wb[msn], cfg)
            if stock and not any(r["rating"] for r in stock):
                notes.append(f"⚠ {cfg['site']}：{msn} 找不到评级列（表头为空、值形态也不像），"
                             f"「当月销售评级」将全部显示「—」")
            daily = read_all_orders(wb)
            if verbose:
                print(f"[{cfg['site']}] 月表 {msn}：{len(stock)} 个 SKU；订单明细 {len(daily)} 个 (SKU,日) 组合")

            # 每个 SKU 的最后销量日 + 每个月内的最后销量日（--earliest 要用）
            last_sold = {}
            last_by_month = {}
            for (sku, d), ch in daily.items():
                if sum(ch.values()) <= 0:
                    continue
                if sku not in last_sold or d > last_sold[sku]:
                    last_sold[sku] = d
                k = (sku, int(d[5:7]))
                if k not in last_by_month or d > last_by_month[k]:
                    last_by_month[k] = d

            # 快照日 = 订单里最大日期（也就是表里的「今日日期」）
            site_snap = max(last_sold.values()) if last_sold else None
            if site_snap and (not snap_date or site_snap > snap_date):
                snap_date = site_snap
            # 泰菲「在售」判据的窗口下界：快照日往前 min_days_sold 天
            lo = ""
            if site_snap:
                lo = (datetime.strptime(site_snap, "%Y-%m-%d").date()
                      - timedelta(days=min_days_sold)).isoformat()
            hist = read_stock_history(wb, cfg)
            snap_dt = datetime.strptime(site_snap, "%Y-%m-%d") if site_snap else None

            cand = []
            for r in stock:
                if r["avail"] > 0 or r["inbound"] > 0:
                    continue                       # 有库存或有在途 → 不算断货
                ls_orig = last_sold.get(r["sku"])
                if not ls_orig:
                    continue                       # 从没卖过 → 不是「在售断货」，是死 SKU
                # 在售判定（用**整体**最后销量日，不能用调整后的窗口终点，否则刚断货的会被误杀）
                if cfg["status_col"] is not None:
                    if r["status"] not in cfg["onsale"]:
                        continue
                elif ls_orig < lo:
                    continue                       # 太久没卖 → 不算在售
                ls, basis = ls_orig, ""
                if earliest and snap_dt:
                    m0, d0 = earliest_zero_day(r["sku"], hist, last_by_month,
                                               snap_dt, earliest_window)
                    if d0 and d0 < ls_orig:
                        ls, basis = d0, f"最早{m0}月进入0/0"
                cand.append((r, ls, basis))

            if verbose:
                print(f"[{cfg['site']}] 断货候选 {len(cand)} 个"
                      f"（在售口径：{'状态列' if cfg['status_col'] is not None else f'近{min_days_sold}天有销量'}"
                      f"{'；起点口径：最早进入0/0' if earliest else '；起点口径：最后销量日'}）")

            for r, ls, basis in cand:
                end = datetime.strptime(ls, "%Y-%m-%d").date()
                start = end - timedelta(days=days - 1)
                by_month = defaultdict(lambda: defaultdict(float))   # 月 → 日 → 合计
                by_month_ch = defaultdict(lambda: defaultdict(Counter))
                for i in range(days):
                    d = (start + timedelta(days=i)).isoformat()
                    ch = daily.get((r["sku"], d))
                    if not ch:
                        continue
                    tot = sum(ch.values())
                    if tot <= 0:
                        continue
                    m = d[:7]
                    by_month[m][d] += tot
                    by_month_ch[m][d].update(ch)

                total_q = sum(sum(v.values()) for v in by_month.values())
                summary.append(dict(site=cfg["site"], sku=r["sku"], model=r["model"],
                                    name=r["name"], avail=r["avail"], inbound=r["inbound"],
                                    last_sold=ls, start=start.isoformat(),
                                    qty=total_q, rating=r["rating"] or "—",
                                    months=len(by_month), basis=basis))
                for m in sorted(by_month):
                    # 当月评级：优先用月表里的真评级（只有 9、10 月有值），没有就留「—」
                    mon_rating = (hist.get((r["sku"], int(m[5:7])), ("", "", ""))[2] or "—")
                    for d in sorted(by_month[m]):
                        ch = by_month_ch[m][d]
                        detail.append(dict(
                            site=cfg["site"], month=m, sku=r["sku"], model=r["model"],
                            rating=mon_rating, date=d,
                            qty=by_month[m][d],
                            **{k: ch.get(k, 0) for k in ("虾皮", "LAZADA", "抖音", "线下", "其他")}))
        finally:
            wb.close()

    return detail, summary, snap_date, notes


# ---------------- 输出 ----------------
_HDR_FILL = PatternFill("solid", fgColor="1F4E79")
_HDR_FONT = Font(color="FFFFFF", bold=True, size=10)
_TITLE_FONT = Font(bold=True, size=12)


def _write_sheet(ws, title, cols, rows, widths=None):
    ws.cell(row=1, column=1, value=title).font = _TITLE_FONT
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(cols))
    for j, (key, label) in enumerate(cols, start=1):
        c = ws.cell(row=2, column=j, value=label)
        c.fill, c.font = _HDR_FILL, _HDR_FONT
        c.alignment = Alignment(horizontal="center", vertical="center")
    for i, r in enumerate(rows, start=3):
        for j, (key, _l) in enumerate(cols, start=1):
            ws.cell(row=i, column=j, value=r.get(key))
    ws.freeze_panes = "A3"
    for j, (key, label) in enumerate(cols, start=1):
        w = (widths or {}).get(key) or max(len(str(label)) * 2, 10)
        ws.column_dimensions[get_column_letter(j)].width = min(w, 40)


def write_xlsx(detail, summary, snap_date, out_dir):
    os.makedirs(out_dir, exist_ok=True)
    day = snap_date or datetime.now().strftime("%Y-%m-%d")
    out = os.path.join(out_dir, f"东南亚缺货60天明细_{day.replace('-', '')}.xlsx")

    wb = Workbook()
    ws = wb.active
    ws.title = "汇总"
    cols = [("site", "国家"), ("sku", "SKU"), ("model", "型号"), ("name", "品名"),
            ("avail", "在库库存"), ("inbound", "在途"), ("last_sold", "最后销量日"),
            ("start", "断货起点"), ("qty", "断货前60天销量"), ("rating", "评级"),
            ("months", "跨月数")]
    if any(r.get("basis") for r in summary):      # 只有用了 --earliest 才加这列
        cols.append(("basis", "起点口径"))
    _write_sheet(ws, f"东南亚 · 在售但无库存无在途 SKU（快照 {day}）", cols,
                 sorted(summary, key=lambda x: (x["site"], -x["qty"], x["sku"])),
                 widths={"model": 28, "name": 34})

    # 按（国家、月份）分 sheet
    order = {c["site"]: i for i, c in enumerate(SITES)}
    groups = defaultdict(list)
    for r in detail:
        groups[(r["site"], r["month"])].append(r)
    for key in sorted(groups, key=lambda k: (order.get(k[0], 9), k[1])):
        site, month = key
        rows = sorted(groups[key], key=lambda x: (x["sku"], x["date"]))
        name = f"{site}-{month}"[:31]
        ws = wb.create_sheet(name)
        _write_sheet(ws, f"{site} · {month} 日销明细（断货前 60 天）",
                     [("sku", "SKU"), ("model", "型号"), ("rating", "当月销售评级"),
                      ("site", "国家"), ("date", "日期"), ("qty", "日销量合计"),
                      ("虾皮", "虾皮"), ("LAZADA", "LAZADA"), ("抖音", "抖音"),
                      ("线下", "线下"), ("其他", "其他")],
                     rows, widths={"model": 28})
    wb.save(out)
    return out, len(groups)


def write_md(detail, summary, snap_date, out_dir, days=60, earliest=False):
    os.makedirs(out_dir, exist_ok=True)
    day = (snap_date or datetime.now().strftime("%Y-%m-%d")).replace("-", "")
    out = os.path.join(out_dir, f"东南亚缺货说明_{day}.md")
    by_site = defaultdict(list)
    for r in summary:
        by_site[r["site"]].append(r)
    if earliest:
        basis_txt = (f"> 断货起点 = 该 SKU **最早**「在库=0 且 在途=0」那个月内的最后有销量日 + 1"
                     f"（用历史月表的在途列区分「真断货」与「补货中」）。明细范围 = 断货前 {days} 天。")
    else:
        basis_txt = (f"> 断货起点 = 该 SKU **最后一次**有销量的日期 + 1（日销表无每日库存快照，"
                     f"用「销量停止」作为断货起点）。明细范围 = 断货前 {days} 天。")
    L = [f"# 东南亚 · 在售但无库存无在途（快照 {snap_date}）", "", basis_txt, "",
         "## 汇总", "", "| 国家 | SKU 数 | 断货前 60 天销量合计 | A/B/C |", "|---|---:|---:|---|"]
    for site in [c["site"] for c in SITES]:
        rs = by_site.get(site) or []
        if not rs:
            continue
        cnt = Counter(r["rating"] for r in rs)
        L.append(f"| {site} | {len(rs)} | {sum(r['qty'] for r in rs):.0f} | "
                 f"{cnt.get('A',0)}/{cnt.get('B',0)}/{cnt.get('C',0)} |")
    L += ["", "## 明细", ""]
    for site in [c["site"] for c in SITES]:
        rs = sorted(by_site.get(site) or [], key=lambda x: -x["qty"])
        if not rs:
            continue
        L += [f"### {site}（{len(rs)} 个）", "",
              "| SKU | 型号 | 最后销量日 | 断货起点 | 前60天销量 | 评级 | 备注 |",
              "|---|---|---|---|---:|---|---|"]
        for r in rs:
            L.append(f"| {r['sku']} | {r['model']} | {r['last_sold']} | {r['start']} | "
                     f"{r['qty']:.0f} | {r['rating']} | {r.get('basis') or ''} |")
        L.append("")
    with open(out, "w", encoding="utf-8") as f:
        f.write("\n".join(L))
    return out


def main(argv):
    ap = argparse.ArgumentParser()
    ap.add_argument("--days", type=int, default=60, help="断货前取多少天（默认60）")
    ap.add_argument("--min-days", type=int, default=60,
                    help="泰菲「在售」判据：近 N 天有销量（默认60）")
    ap.add_argument("--earliest", action="store_true",
                    help="断货起点改用「最早进入 0/0 状态的那个月」内的最后销量日"
                         "（需求原文写的是「最早」，默认仍是最后销量日）")
    ap.add_argument("--earliest-window", type=int, default=120,
                    help="--earliest 时往前查多少天的历史月表（默认120天≈4个月）")
    ap.add_argument("--out", default=None,
                    help="输出目录；不填则按快照日落到 "
                         "断货分析/东南亚/{明细|最早口径}/<YYYYMMDD>/")
    a = ap.parse_args(argv)

    detail, summary, snap_date, notes = build(days=a.days, min_days_sold=a.min_days,
                                              earliest=a.earliest,
                                              earliest_window=a.earliest_window)
    for n in notes:
        print(n)
    if not summary:
        print("没有命中的 SKU。")
        return 1
    if a.out:
        out_dir = a.out
    else:
        daytag = (snap_date or datetime.now().strftime("%Y-%m-%d")).replace("-", "")
        out_dir = os.path.join(SEA_OUT, "最早口径" if a.earliest else "明细", daytag)
    xlsx, nsheet = write_xlsx(detail, summary, snap_date, out_dir)
    md = write_md(detail, summary, snap_date, out_dir, days=a.days, earliest=a.earliest)
    print(f"\n命中 SKU {len(summary)} 个；明细 {len(detail)} 行，分 {nsheet} 个（国家,月份）sheet")
    n_adj = sum(1 for r in summary if r.get("basis"))
    if a.earliest:
        print(f"其中 {n_adj} 个 SKU 的起点按「最早进入0/0」前移（其余与最后销量日一致）")
    print(f"已输出:\n  {xlsx}\n  {md}")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
