# -*- coding: utf-8 -*-
"""亚马逊/美客多/eMAG 缺货情况表（2026-10-22 需求，随日销表周期输出）。

在原有日销表数据基础上，把「当前缺货」单独输出成一个文件，供采购直接用。

三级标记（用户给定口径）：
  一级  库存 ≠ 0 且 可售天数 + 7 < 计划备货天数   → 卖得动、但撑不到下次到货
  二级  库存 = 0 且 在途 > 0                      → 断货，补给在路上
  三级  库存 = 0 且在途 = 0                       → 断货且无补给，最紧急

计划备货天数读 `Daily_Sales_Report/断货分析/备货天数.xlsx`（站点 → 天数，可随时改）。
「+7」= 后勤周期（每周三进一次货）。

输出：`Daily_Sales_Report/断货分析/亚马逊/每日缺货/<YYYYMMDD>/亚马逊美客多EMAG缺货情况表.xlsx`
（2026-10-23 起从 `Daily_Sales_Report/每日缺货/` 迁入，断货相关的产出统一收在 `断货分析/` 下）

数据源全部复用 `stockout_report.build()`（同一套库存/订单口径，避免两套算法打架）：
  亚马逊  afn-fulfillable-quantity / inbound-working+shipped+receiving
  美客多  Resumen c19 可销售 / c16 待录入 + c17 转账中
  eMAG    stock_type=ready_to_sell

用法：
    python pipeline/tools/shortage_report.py                # 取最新快照
    python pipeline/tools/shortage_report.py --snap <目录>
    python pipeline/tools/shortage_report.py --lead 7 --out <目录>
"""
import os
import sys
import math
import argparse
from collections import Counter
from datetime import date, datetime, timedelta

_HERE = os.path.dirname(os.path.abspath(__file__))
_CORE = os.path.join(os.path.dirname(_HERE), "core")
for _p in (_CORE, _HERE):
    if _p not in sys.path:
        sys.path.insert(0, _p)

import config as C                                        # noqa: E402
import stockout_report as S                               # noqa: E402
import product_master as PM                               # noqa: E402  商品资料主数据
from openpyxl import Workbook, load_workbook              # noqa: E402
from openpyxl.styles import Font, PatternFill, Alignment  # noqa: E402
from openpyxl.utils import get_column_letter              # noqa: E402

_DSR = os.path.dirname(os.path.dirname(_HERE))            # …\Daily_Sales_Report
# 断货相关产出统一收在 断货分析/ 下，一级按业务块分（亚马逊 / 东南亚）
ANALYSIS_DIR = os.path.join(_DSR, "断货分析")
PLAN_XLSX = os.path.join(ANALYSIS_DIR, "备货天数.xlsx")

# 备货天数表里的站点名 ↔ 本程序 site 名
_ALIAS = {"沙特阿拉伯": "沙特", "澳大利亚": "澳洲"}

# 等级 → (底色, 文字)
_LEVEL_STYLE = {
    "一级": ("FFF2CC", "一级 · 可售不足"),
    "二级": ("FCE4D6", "二级 · 断货有在途"),
    "三级": ("F8CBAD", "三级 · 断货无补给"),
}
_HDR_FILL = PatternFill("solid", fgColor="1F4E79")
_HDR_FONT = Font(color="FFFFFF", bold=True, size=10)


def plan_days_map():
    """读 备货天数.xlsx → {站点: 天数}。实现唯一在 stockout_report（本处只做转发）。"""
    return S.plan_days_map()


def plan_of(site, acct, m):
    """站点/账号 → 计划备货天数。实现唯一在 stockout_report。"""
    return S.plan_of(site, acct, m)


def grade(avail, inbound, dos, plan, lead=7):
    """→ (等级, 判定依据)；不算缺货返回 (None, '')。"""
    if avail is not None and avail > 0:
        if dos is None or plan is None:
            return None, ""
        if dos + lead < plan:
            return "一级", f"可售 {dos:.1f} 天 + {lead} < 计划 {plan} 天（差 {plan - dos - lead:.1f} 天）"
        return None, ""
    if inbound is not None and inbound > 0:
        return "二级", f"库存 0，在途 {inbound} 件"
    return "三级", "库存 0，在途 0"


def _book_of():
    """(site, acct) → 簿文件名。用 WORKBOOKS 里的 report 模板反查，不靠名字猜。

    为什么要反查：同一张分析表名在不同簿里会重名（「墨西哥-每日库存分析表」既在
    北美加拿大簿=亚马逊主账号，也在美客多簿=美客多），只按站点名匹配会张冠李戴。
    """
    rep2book = {}
    for bk in getattr(C, "WORKBOOKS", []):
        for st in bk.get("stocks", []):
            r = st.get("report")
            for x in ([r] if isinstance(r, str) else (r or [])):
                rep2book.setdefault(x, bk["file"])
    out = {}
    for s in S.SITES:
        reps = s["rep"] if isinstance(s["rep"], list) else [s["rep"]]
        for x in reps:
            if x in rep2book:
                out[(s["site"], s["acct"])] = rep2book[x]
                break
    # 美客多 / eMAG 不在 SITES 里（数据源不是亚马逊库存报告），单独补
    out[("墨西哥", "美客多")] = "美客多库存销售表.xlsx"
    out[("罗马尼亚", "eMAG")] = "罗马尼亚库存销售表.xlsx"
    return out


def _norm_sku(s):
    """`amzn.gr.<SKU>-<hash>[-LN|-VG]` → `<SKU>`，与 analysis.normalize_sku 同源。

    ⚠️ 必须归并：缺货表的 SKU 来自**库存报告**（可能是 `amzn.gr.G421-xxx-LN` 这种别名），
    而月度块的 A 列是**归并后**的 `G421`。不归并就查不到历史，销量列会显示 0。
    """
    try:
        import analysis as A
        return A.normalize_sku(s)
    except Exception:
        return str(s or "").strip()


def _one_month_days(start, end, snap_dt):
    """日销表「当月缺货天数」公式的 Python 版（逐字对齐 SWITCH 语义）。

    公式原文（以罗马尼亚透视图 F8 为例，A = 锚点-1）：
        SWITCH(TRUE,
          C="",                                    0,
          MONTH(C)=MONTH(A) AND D=锚点,            锚点-C+1,
          MONTH(C)=MONTH(A) AND D<>锚点,           D-C+1,
          C-A<0 AND MONTH(D)=MONTH(A),             D-(EOMONTH(A,-1)+1)+1,
          0)

    ⚠️ 三处必须照抄的细节：
      ① 第 1 个分支用的是**锚点**（=快照日），不是 A（=锚点-1），两者差一天；
      ② 第 3 个分支从**当月 1 号**起算（EOMONTH(A,-1)+1），不是从缺货开始日；
      ③ 结束时间早于本月的（上个月就补上了）→ 0，不是负数。
    缺货结束时间为空 = 至今仍断 → 按 A（昨天）收尾。
    """
    if start is None:
        return 0
    a = snap_dt - timedelta(days=1)          # A = 锚点-1
    if start.year == a.year and start.month == a.month:
        if end is None or end == snap_dt:
            return (snap_dt - start).days + 1
        return max(0, (end - start).days + 1)
    if start < a and end is not None and end.year == a.year and end.month == a.month:
        return max(0, (end - a.replace(day=1)).days + 1)
    return 0


def oos_days(book_map, snap_day):
    """{(site, acct, sku): 当月缺货天数}，读各簿「透视」表的缺货起止时间算出来。

    为什么需要：日销表的实时速度（日均）= 当月累计 ÷ (已过天数 − 当月缺货天数)。
    断货期间没货可卖，销量自然少；不把缺货天数从分母里剔掉，日均会被系统性低估，
    可售天数就虚高，一级（可售天数+7 < 计划备货天数）会漏判 —— 而这恰恰是断货表
    最该抓住的场景。

    找不到透视表的簿（澳大利亚/印度等 4 本没有该表）→ 缺货天数按 0，
    公式退化为 当月累计 ÷ 已过天数，与该簿自身的日销表行为一致（它也没这列）。
    """
    from openpyxl import load_workbook
    snap_dt = datetime.strptime(str(snap_day), "%Y%m%d").date()

    by_book = {}
    for (site, acct), bk in book_map.items():
        by_book.setdefault(bk, []).append((site, acct))

    out, found_books, found_tbl = {}, 0, 0
    seen_books = set()
    for bk, keys in by_book.items():
        p = os.path.join(C.WORK_DIR, bk)
        if not os.path.isfile(p):
            continue
        try:
            wb = load_workbook(p, read_only=True, data_only=True)
        except Exception:
            continue
        got = {}
        for ws in wb.worksheets:
            nm = ws.title
            # 「新品销售情况数据透视图」是另一回事（新品判定用），跳过
            if "透视" not in nm or "新品" in nm:
                continue
            try:
                rows = list(ws.iter_rows(min_row=1, max_row=60, max_col=12, values_only=True))
            except Exception:
                continue
            hdr_i = start_j = None
            for i, r in enumerate(rows):
                for j, v in enumerate(r):
                    if isinstance(v, str) and "缺货开始" in v:
                        hdr_i, start_j = i, j
                        break
                if hdr_i is not None:
                    break
            if hdr_i is None:
                continue
            found_tbl += 1
            # 表头：SKU 在开始列左边一列，结束列在右边一列，当月列往后找「当月」
            hdr = rows[hdr_i]
            end_j = start_j + 1
            cur_j = None
            for j, v in enumerate(hdr):
                if isinstance(v, str) and v.strip() == "当月":
                    cur_j = j
                    break
            if cur_j is None:
                cur_j = start_j + 3            # 两本簿实测都是 开始/结束/上月/当月
            sku_j = start_j - 1
            for r in rows[hdr_i + 1:]:
                if not r or sku_j >= len(r):
                    continue
                sku = r[sku_j]
                if not isinstance(sku, str) or not sku.strip():
                    continue
                start = r[start_j] if start_j < len(r) else None
                end = r[end_j] if end_j < len(r) else None
                # openpyxl 读出来是 datetime，公式里是日期 → 统一成 date 再比
                if isinstance(start, datetime):
                    start = start.date()
                if isinstance(end, datetime):
                    end = end.date()
                for v in (start, end):
                    if v is not None and not isinstance(v, date):
                        start = end = None
                        break
                d = _one_month_days(start, end, snap_dt)
                if d:
                    got[sku.strip()] = max(got.get(sku.strip(), 0), d)
        wb.close()
        if found_tbl and bk not in seen_books:
            seen_books.add(bk)
            found_books += 1
        if got:
            for (site, acct) in keys:
                for sku, d in got.items():
                    out[(site, acct, _norm_sku(sku))] = d
    if found_books == 0:
        print("[缺货天数] !! 一本透视表都没读到（实时速度退化为 当月累计÷已过天数）", flush=True)
    else:
        print(f"[缺货天数] 读到 {found_books} 本簿的缺货起止表，"
              f"当月有缺货记录的 SKU {len({k[2] for k in out})} 个", flush=True)
    return out


def hist_months(book_map=None):
    """{(site, acct, sku): [1月..12月销量]}。读不到月度块返回 {}。

    数据源 = 各簿「-每日库存分析表」的**月度块**（1~12 月，跑批封存的静态值）。
    月度块里全 0 的行在解析阶段就被跳过，所以「查不到」= 今年一件没卖过。

    用途有两个（2026-10-29）：
      ① 判断「未起量」：可售为 0 的 SKU 本来就没货可卖，近 30 天销量 0 是**结果**
         不是「没需求」的证据。实测 20261028 的 326 个无销量 SKU 里 listing 一个都没下架。
      ② 给每行补「上月 / 近3月 / 今年累计」销量列，让采购一眼看出是刚断还是从没卖过。

    读不到时返回 {}，调用方据此退化为「不拆分」且销量列留空（宁可不拆，不能误拆）。
    """
    book_map = book_map or _book_of()
    want = {}
    for (site, acct), bk in book_map.items():
        want.setdefault(bk, []).append((site, acct))
    if not want:
        return {}
    try:
        import sales_trend_report as T          # 懒加载：不跑分类时不必付这个开销
    except Exception as e:
        print(f"[历史销量] !! 无法加载月度块解析器 {e!r}（本次不拆分未起量 SKU）", flush=True)
        return {}

    out, ok_books = {}, 0
    for bk, keys in want.items():
        p = os.path.join(C.WORK_DIR, bk)
        if not os.path.isfile(p):
            continue
        try:
            _sheets, recs = T.parse_book(p, 10, with_site=False)
        except Exception:
            continue                        # 单本读失败只影响这一本，其余照常
        ok_books += 1
        for r in recs:
            site = r["sheet"].replace("-每日库存分析表", "").strip()
            sku = _norm_sku(r["sku"])
            for (s, a) in keys:
                if s != site:
                    continue
                k = (s, a, sku)
                if k in out:                # 归并后撞号 → 逐月相加，不覆盖
                    out[k] = [x + y for x, y in zip(out[k], r["months"])]
                else:
                    out[k] = list(r["months"])
    if ok_books == 0:
        print("[历史销量] !! 一本月度块都没读到（本次不拆分未起量 SKU）", flush=True)
        return {}
    return out


def hist_sold(book_map=None):
    """今年有过销量的 {(site, acct, sku)}。基于 hist_months()，口径完全一致。"""
    return {k for k, ms in hist_months(book_map).items() if sum(ms) > 0}


def collect(snap_dir, snap_day, lead=7, classify_unsold=True, with_history=True,
            legacy_dos=False, moq=0):
    """→ (items, meta)

    classify_unsold=True（默认）时，把「三级 且 无在途 且 今年从未有过销量」的 SKU
    打上 unsold=True，由 write() 拆到独立 sheet，不计入缺货数。
    二级（有在途）不拆 —— 货已经在路上，属于标准缺货，不是「没需求」。

    with_history=True（默认）时，额外给每行补「上月 / 近3月 / 今年累计」销量列
    （数据源 = 月度块，与 classify_unsold 共用一次读取）。所有等级都补。

    moq>1 时建议补货量按最小批量向上取整（件/箱），与 stockout_report 的 --moq 同义。
    """
    m = plan_days_map()
    rows, sales_only, missing, warn, snap_dt, _ = S.build(
        snap_dir, snap_day, [], lead, 14, use_plan=False)
    # use_plan=False：本函数自己用 plan_days 判级，
    # 不需要 build() 再按同一套规则判一遍（省一次备货天数读取，也避免两处判定打架）。

    months_map = hist_months() if with_history else {}
    if with_history and not months_map:
        classify_unsold = False             # 读不到历史 → 退化为不拆，避免把真缺货误拆
        degraded = True
    else:
        degraded = False

    cur_m = int(str(snap_day)[4:6]) or 10          # 快照月份，用于取上月/近3月

    # ---- 统一日历口径（2026-11-07 用户要求）----
    # 原实现固定 ÷30 / ÷90，跨月天数不等（31/28 天）就会系统性偏差：
    #   10 月 31 天却 ÷30 → 上月日均虚高 3.3%；2 月 28 天却 ÷30 → 虚低 6.7%。
    # 而日销表 AE 列「上月日均」是用 ROUND(封存列/当月日历天数,2) 算的，
    # 断货分析用 ÷30 就跟日销表对不上。这里统一改成按真实日历天数。
    def _month_days(y, m):
        """某年某月的日历天数"""
        if m == 12:
            nxt = date(y + 1, 1, 1)
        else:
            nxt = date(y, m + 1, 1)
        return (nxt - date(y, m, 1)).days

    def _shift_month(y, m, back):
        """往回推 back 个月 → (年, 月)"""
        for _ in range(back):
            if m == 1:
                y, m = y - 1, 12
            else:
                m -= 1
        return y, m

    _py, _pm = _shift_month(snap_dt.year, snap_dt.month, 1)     # 上月
    _py2, _pm2 = _shift_month(snap_dt.year, snap_dt.month, 2)   # 上上月
    _py3, _pm3 = _shift_month(snap_dt.year, snap_dt.month, 3)   # 上上上月
    DAYS_PREV = _month_days(_py, _pm)                          # 上月天数（31/30/28/29）
    DAYS_M3 = _month_days(_py2, _pm2) + _month_days(_py3, _pm3) + DAYS_PREV  # 近3月总天数
    DAYS_Q30 = 30                                              # 近30天是固定窗，不变

    def _key(r):
        return (r["site"], r["acct"], _norm_sku(r["sku"]))

    # 当月缺货天数（日销表实时速度的分母修正项），读各簿透视表
    oos = oos_days(_book_of(), snap_day) if with_history else {}
    elapsed = snap_dt.day - 1                       # DAY(锚点-1)，与日销表同源

    def _daily_mtd(r):
        """实时速度 = 当月累计 ÷ (已过天数 − 当月缺货天数)。日销表 T / AD 列口径。

        断货那几天没货可卖，销量天然少；不剔掉会把日均低估、可售天数虚高。
        美客多没有逐日销量（q_mtd=None）→ 退化为近30天日均，与它自己的口径一致。
        """
        q = r.get("q_mtd")
        if q is None:
            return (r.get("q30") or 0) / 30.0
        if elapsed < 8:
            # 月初分支（与日销表一致）：本月样本太少，改用 (已过+7) 天滚动窗
            q = r.get("q_roll") or 0
            den = elapsed + 7 - oos.get(_key(r), 0)
        else:
            den = elapsed - oos.get(_key(r), 0)
        return (q / den) if den > 0 else 0.0

    # 「今年卖过」= 月度块有销量 **或** 近30天有销量。
    # 后半个条件必须加：新补的行（如 G573/G590）月度块的 SUMIFS 缓存还没算出来，
    # 只按月度块判会把正在卖的 SKU 误判成「未起量」（实测漏了 1 个：沙特 interflow G590
    # 近7天卖 1、近30天卖 2，月度块却是空）。
    sold = {k for k, ms in months_map.items() if sum(ms) > 0}
    sold |= {_key(r) for r in rows if (r.get("q30") or 0) > 0}

    def _hist_cols(key, q30):
        """→ (上月, 近3月, 今年累计)。月度块缺失（新补行）时用近30天兜底。"""
        q30 = q30 or 0
        ms = months_map.get(key)
        if not ms:
            return 0, q30, q30
        prev = ms[cur_m - 2] if cur_m >= 2 else 0
        # 近3月 / 今年累计 都要 ≥ 近30天：q30 是订单报告的硬数据，必然落在这些窗口内。
        # 取 max 是为了兜住「月度块滞后」的行，同时保证 累计 ≥ 近3月 ≥ 近30天 不会倒挂。
        m3 = max(sum(ms[max(0, cur_m - 3):cur_m]), q30)
        ytd = max(sum(ms[:cur_m]), q30)
        return prev, m3, ytd

    items = []
    # 商品资料（2026-11-07 接入）：类型/状态/销量评级/品牌/供应商/镜头膜。
    # 采购最关心两件事 —— 这是壳还是膜、这个 SKU 是不是已经停售了；
    # 另外按用户给的评级规则，E / XE「不要采购」，遇到就直接压掉建议数量。
    pm = PM.load()
    # 商品资料覆盖统计：查不到属性的 SKU 会在「类型/商品状态/销量评级」三列留空，
    # 采购看到空白会以为是数据没算出来。这里累计下来，最后在 meta 里报警。
    _pm_miss = {}
    for r in rows:
        plan = plan_of(r["site"], r["acct"], m)
        key = _key(r)
        m_prev, m3, m_ytd = _hist_cols(key, r.get("q30")) if with_history else (None, None, None)
        # 日销表两个日均 + 含预留的可售天数（2026-10-29 用户要求「搬过来」）
        # 日均先按 3 位小数定值，可售天数再用**这个定值**去除 —— 与日销表
        # 「ROUND(日均,1) 后再算 (F+E)/T」的做法一致，保证表里看到的数能对上。
        d_mtd = round(_daily_mtd(r), 3)
        d_prev = round(m_prev / DAYS_PREV, 3) if m_prev is not None else None   # 日历天数
        if d_mtd > 0:
            dos_v = ((r.get("avail") or 0) + (r.get("reserved") or 0)) / d_mtd
        else:
            dos_v = None                            # T=0 → 日销表显示 "-"
        # 判定用哪个可售天数：默认日销表口径（含预留、本月卖速）；
        # --legacy-dos 回退到 2026-10-28 前的「可售 ÷ 近30天日均」。
        dos_use = r.get("dos") if legacy_dos else dos_v
        g, why = grade(r.get("avail"), r.get("inbound"), dos_use, plan, lead)
        if not g:
            continue
        dos = r.get("dos")
        # 归并注记（2026-11-03）：库存报告的 amzn.gr.* 别名行已并入本行，可售/在途
        # 是成员行之和 —— 不注明的话，采购对不上原始报告里「裸 SKU 那一行」的数。
        if r.get("aliases"):
            why += f"（并 {len(r['aliases'])-1} 个 amzn.gr.* 别名行：{r['aliases'][1]} 等，库存已求和）"
        unsold = bool(
            classify_unsold and g == "三级" and (r.get("inbound") or 0) <= 0
            and key not in sold
        )
        if unsold:
            # stockout_report 的原措辞是「疑似已下架/弃售」，但实测 listing 全活着，
            # 这里按实际情况重写，免得采购照着去查下架、查了半天发现货架还在。
            why = "可售 0、在途 0，且今年 1~10 月从未有过销量（listing 仍在，疑似未起量）"

        # ---------------- 采购建议（2026-11-03 用户需求：新增「采购建议」sheet） ----------------
        # 目标覆盖天数 = 计划备货天数，与「计划备货天数 / 缺口天数」两列同源。
        # 日均回退链：本月实时速度 → 上月日均 → 近3月÷日历天数 → 近30天÷30。
        # ⚠️ 回退链不能省：二级/三级的 SKU 这个月本来就没货可卖，本月日均经常是 0，
        #    直接按 0 算会得出「建议补货 0 件」—— 把最该补的断货行整行漏掉。
        if d_mtd > 0:
            by_daily, by_src = d_mtd, "本月日均"
        elif d_prev is not None and d_prev > 0:
            by_daily, by_src = d_prev, "上月日均"
        elif m3:
            by_daily, by_src = round(m3 / DAYS_M3, 3), f"近3月÷{DAYS_M3}"
        elif (r.get("q30") or 0) > 0:
            by_daily, by_src = round(r["q30"] / DAYS_Q30, 3), "近30天÷30"
        else:
            by_daily, by_src = 0.0, "无销量"

        tgt_days = plan if plan else 60          # 查不到备货天数的站点按 60 天兜底
        target = int(math.ceil(by_daily * tgt_days)) if by_daily > 0 else 0
        # 预留是已锁定给既有订单的量，不当可用库存扣减（与断货明细的口径一致）
        buy = max(0, target - (r.get("avail") or 0) - (r.get("inbound") or 0))
        # 与 buy 同一分母的可撑天数：buy>0 ⇔ dos_in < 计划备货天数，采购可逐行自校验
        _hold = (r.get("avail") or 0) + (r.get("inbound") or 0)
        dos_in = round(_hold / by_daily, 1) if by_daily > 0 else None
        if moq > 1 and buy > 0:
            buy = int(math.ceil(buy / moq) * moq)

        # 采购动作：纯枚举值（立刻补 / 在途不够 / 下周三排），供采购在 Excel 里直接筛选。
        # 原来是一整串文字（「补货 156 件（在途不足）」），筛选只能按文本匹配，很不顺手。
        # 2026-11-03 按用户要求拆成「采购动作」+「建议数量」两列。
        if buy <= 0:
            action = "—"
        elif g == "三级":
            action = "立即补货"
        elif g == "二级":
            action = "在途不足"
        else:
            action = "排入下周三"

        # ---------------- 商品资料属性（2026-11-07） ----------------
        # 三条硬规则直接压采购建议，避免采购照着表去下不该下的单：
        #   ①状态=停售/不销售 → 不补（钢化膜/水凝膜连续 12 个月 E 级、平板膜一年<250 套的
        #     已经被运营标停售，再补就是压库存）
        #   ②销量评级 E / XE → 不补（用户写在 商品资料!A1：「评级为E或者XE,不要采购」）
        #   ③手机壳不参与评级（用户原文「手机壳产品不参与以上产品评级规则！」）：
        #     product_master.should_buy() 对手机壳直接返回 None，不走 E/XE 压制。
        #     实测 26 个手机壳里有 18 个挂着 E/XE —— 不豁免的话它们会被整批误杀。
        _pa = PM.attrs(r["sku"], pm) or {}
        ptype, pstatus, pscore = _pa.get("type"), _pa.get("status"), _pa.get("score")
        if not _pa:
            _pm_miss.setdefault(str(r["sku"]), []).append(f'{r["site"]}/{r["acct"]}')
        buy_blk = None
        if buy > 0:
            if str(pstatus or "") in ("停售", "不销售"):
                buy_blk = f"停售（{pstatus}），不给采购建议"
            elif PM.should_buy(r["sku"], pm) is False:
                buy_blk = f"评级 {pscore}，按规则不采购"
        if buy_blk:
            buy, action, dos_in = 0, "—", dos_in
            why = f"{why}｜{buy_blk}"

        # 优先级与断货明细同源：严重断货=三级、一般断货=二级、预备断货=一级
        q30v = r.get("q30") or 0
        if g == "三级":
            prio = "P0" if q30v >= 5 else ("P1" if q30v >= 1 else "P3")
        elif g == "二级":
            prio = "P1" if q30v >= 1 else "P3"
        else:
            prio = "P2"

        items.append(dict(
            level=g, site=r["site"], acct=r["acct"], sku=r["sku"],
            name=(r.get("name") or "").strip(),
            avail=r.get("avail"), inbound=r.get("inbound"),
            q7=r.get("q7"), q30=r.get("q30"), daily=r.get("daily"),
            m_prev=m_prev, m3=m3, m_ytd=m_ytd,
            daily_mtd=d_mtd, daily_prev=d_prev, dos_v=dos_v,
            reserved=(r.get("reserved") or 0), oos=oos.get(key, 0),
            dos=dos, plan=plan,
            # 缺口天数必须与「可售天数」显示列、以及判级用的是同一个可售天数，
            # 否则会出现「可售 6 天 + 后勤 7 天 + 缺口 48 天 ≠ 计划 75 天」。
            # 2026-11-03 修：原用 r["dos"]（近30天日均口径），与 dos_v 不同源。
            gap=(round(plan - dos_use - lead, 1)
                 if (plan is not None and dos_use is not None) else None),
            why=why, src=r.get("src") or "亚马逊",
            nosales=bool(r.get("nosales")),
            unsold=unsold,
            # 采购建议
            by_daily=by_daily, by_src=by_src, tgt_days=tgt_days,
            target=target, buy=buy, dos_in=dos_in,
            action=action, prio=prio,
            tier=_tier_of(buy),
            # 商品资料属性（2026-11-07）：壳/膜、是否停售、销量评级、品牌、供应商、
            # 镜头膜标记（镜头膜属于配件，采购/运营分线不同，单独一列便于筛）
            ptype=ptype, pstatus=pstatus, pscore=pscore,
            pauto=bool(PM.is_stop_sale(r["sku"], pm)),
            pbrand=_pa.get("brand"), pvendor=_pa.get("vendor"),
            plens=_pa.get("lens"),
            buy_blk=buy_blk,
            # 库龄（2026-11-03 加，源=库龄总表，随月销表更新）：age = 七档 dict；
            # stale = 超90天呆滞件数（亚马逊/interflow/eMAG = 91天以上各档之和，
            # 美客多 = Resumen「陈旧库存」，该账号没有分档）
            age=r.get("tiers"), stale=(r.get("stale") or 0),
        ))
    _age_p, _age_per = S.age_xlsx_info()
    meta = dict(plan=m, missing=missing, warn=warn, snap_dt=snap_dt,
                total=len(rows), sales_only=sales_only, cur_m=cur_m,
                with_history=with_history, elapsed=elapsed, oos_rows=len(oos),
                classified=classify_unsold, degraded=degraded, moq=moq,
                lead=lead,
                age_xlsx=(os.path.basename(_age_p) if _age_p else None),
                age_per=_age_per,
                # 日历口径（2026-11-07）：把三个分母带出去，供说明文案引用。
                # DAYS_PREV / DAYS_M3 是 build() 内的局部变量，_advice_sheet()
                # 在另一个函数里拿不到 —— 必须走 meta 传参（曾因此 NameError）。
                days_prev=DAYS_PREV, days_m3=DAYS_M3, days_q30=DAYS_Q30,
                # 商品资料覆盖（2026-11-07）：{SKU: [站点/账号, ...]}，供报表脚注与告警
                pm_path=PM.find_path(), pm_total=len(pm), pm_miss=_pm_miss)
    return items, meta


# ---------------- 输出 ----------------
_LEVEL_ORDER = {"一级": 0, "二级": 1, "三级": 2}
_COLS = [("level", "缺货等级"), ("site", "站点"), ("acct", "账号"), ("sku", "SKU"),
         ("name", "品名"),
         # 商品资料（2026-11-07）：采购第一眼要看清这是壳还是膜、这个 SKU 还卖不卖
         ("ptype", "类型"), ("pstatus", "商品状态"), ("pscore", "销量评级"),
         ("avail", "可售库存"), ("reserved", "预留"),
         ("inbound", "在途"),
         ("q7", "近7天"), ("q30", "近30天"),
         ("m_prev", "上月销量"), ("m3", "近3月销量"), ("m_ytd", "今年累计"),
         # 两个日均都是日销表口径：本月日均=实时速度（当月累计÷(已过−缺货天数)），
         # 上月日均=上月销量÷30。可售天数=(可售+预留)÷本月日均，与日销表 V 列一致。
         ("daily_mtd", "本月日均"), ("daily_prev", "上月日均"),
         ("dos_v", "可售天数"), ("plan", "计划备货天数"), ("gap", "缺口天数"),
         ("why", "判定依据"), ("src", "数据源")]


_HIST_COLS = ("m_prev", "m3", "m_ytd")

# 采购建议 sheet 的列（2026-11-03 新增）。比明细 sheet 少「判定依据」、「数据源」，
# 多「计算日均 / 日均来源 / 目标库存量 / 采购动作 / 建议数量」——
# 采购只看这一张就能下单，不用回头对照别的 sheet。
_ADV_COLS = [
    ("prio", "优先级"), ("level", "缺货等级"), ("site", "站点"), ("acct", "账号"),
    ("sku", "SKU"), ("name", "品名"),
    # 商品资料（2026-11-07）：类型（钢化膜/水凝膜/手机壳…）、商品状态（在售/停售）、
    # 销量评级（A~E / XA~XE）。评级 E、XE 与停售 SKU 的建议数量已被压成 0，
    # 原因写在「停购原因」列，采购一眼能看出为什么这行不用下单。
    ("ptype", "类型"), ("pstatus", "商品状态"), ("pscore", "销量评级"),
    ("pbrand", "品牌"), ("pvendor", "供应商"),
    ("avail", "可售库存"), ("reserved", "预留"), ("inbound", "在途"),
    ("q7", "近7天"), ("q30", "近30天"), ("m3", "近3月销量"), ("m_ytd", "今年累计"),
    ("by_daily", "计算日均"), ("by_src", "日均来源"),
    # ⚠️ 这里不放明细表的「缺口天数」：那列按近30天日均算，而建议补货量按计算日均
    # （本月优先，为0回退）算，两者并列会出现「可售6天/缺口48天/建议补货0件」这种
    # 采购看不懂的组合。改用「含在途可撑天数」—— 与建议补货量同一分母，逐行可自校验：
    #   含在途可撑天数 < 计划备货天数  ⇔  建议补货量 > 0
    ("dos_in", "含在途可撑天数"), ("plan", "计划备货天数"),
    ("target", "目标库存量"),
    # 动作在前、数量紧跟其后：筛「采购动作」时数量就在旁边，一眼看到要下多少。
    ("action", "采购动作"), ("buy", "建议数量"),
    # 梯队（2026-11-03 用户要求「按建议数量梯队排」）：
    # 数量从大到小编号，同梯队里的第几单写在「梯队内序号」，采购可按梯队分批下单。
    ("tier", "数量梯队"), ("tier_seq", "梯队内序号"),
    # 库龄七档 + 呆滞（2026-11-03 用户要求把断货明细的库龄统计也加进来）。
    # 源 =《亚马逊EMAG库龄情况总表》（trunk/data/库龄报表/，随月销表更新），
    # 按（账号, 站点, SKU）精确挂；美客多行没有分档（七档整行留空），
    # 呆滞取 Resumen「陈旧库存」。第 7 档在总表里标「366+」，这里沿用断货明细的「365+」。
    ("a0", "0-30"), ("a1", "31-60"), ("a2", "61-90"), ("a3", "91-180"),
    ("a4", "181-270"), ("a5", "271-365"), ("a6", "365+"), ("stale", "呆滞>90天"),
    ("buy_blk", "停购原因"),
]

# 七档库龄的 key（总表缺失时这七列整体不出，呆滞列保留——美客多仍有数）
_AGE_TIER_KEYS = ("a0", "a1", "a2", "a3", "a4", "a5", "a6")

# 数量梯队的边界（件）。切法依据 20261103 真实分布，不是拍脑袋的等分：
#   ≥200 → 31 个 SKU，却占 76.5% 的件数（这几个是真正的大单，要单独走采购流程）
#   51~199 → 19 个；11~50 → 78 个；≤10 → 241 个（碎单，采购一般攒几批一起下）
# 想改档位只改这里，排序、序号、表尾汇总会一起跟着变。
_TIER_BOUNDS = (("T1 大单", 200), ("T2 中单", 51), ("T3 小单", 11))
_TIER_NAMES = [t[0] for t in _TIER_BOUNDS] + ["T4 碎单"]
_TIER_ORDER = {n: i for i, n in enumerate(_TIER_NAMES)}


def _tier_of(buy):
    for name, lo in _TIER_BOUNDS:
        if buy >= lo:
            return name
    return "T4 碎单"

_PRIO_ORDER = {"P0": 0, "P1": 1, "P2": 2, "P3": 3, "—": 9}

# 优先级底色：P0 红、P1 橙、P2 淡黄、P3 灰（与三级标记的色系保持一致）
_PRIO_FILL = {"P0": "F8CBAD", "P1": "FCE4D6", "P2": "FFF2CC", "P3": "EDEDED"}


def _sheet(ws, title, items, cols=None):
    cols = cols or _COLS
    ws.cell(row=1, column=1, value=title).font = Font(bold=True, size=12)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(cols))
    for j, (_k, label) in enumerate(cols, start=1):
        c = ws.cell(row=2, column=j, value=label)
        c.fill, c.font = _HDR_FILL, _HDR_FONT
        c.alignment = Alignment(horizontal="center", vertical="center")
    for i, it in enumerate(items, start=3):
        for j, (k, _l) in enumerate(cols, start=1):
            v = it.get(k)
            if isinstance(v, float):
                if k in ("dos_v", "gap"):
                    v = round(v, 1)                # 可售天数/缺口天数：1 位
                elif k in ("daily_mtd", "daily_prev"):
                    v = round(v, 3)                # 日均：3 位（用户 2026-10-29 要求）
            ws.cell(row=i, column=j, value=v)
        # 整行按等级上色（只染前三列，避免大片色块影响读数）
        fill = _LEVEL_STYLE.get(it["level"], (None, None))[0]
        if fill:
            for j in range(1, 4):
                ws.cell(row=i, column=j).fill = PatternFill("solid", fgColor=fill)
    ws.freeze_panes = "A3"
    ws.auto_filter.ref = f"A2:{get_column_letter(len(cols))}{max(len(items) + 2, 3)}"
    for j, (k, label) in enumerate(cols, start=1):
        w = {"name": 40, "why": 42, "sku": 12, "site": 12, "acct": 10}.get(k, len(label) * 2 + 2)
        ws.column_dimensions[get_column_letter(j)].width = w


def _advice_sheet(wb, adv, meta, live_n):
    """「采购建议」sheet（2026-11-03 用户需求）—— 只列建议补货量 > 0 的行。

    2026-11-03 二次调整：主序改成「数量梯队」（T1 大单 → T4 碎单），
    梯队内按建议数量降序编号，方便采购分批下单。
    2026-11-03 三次调整：尾部追加库龄七档 + 「呆滞>90天」（与断货明细的采购建议
    sheet 对齐），源 =《亚马逊EMAG库龄情况总表》（随月销表更新）；总表缺失时
    七档整组不出、只留呆滞列。
    """
    cols = _ADV_COLS
    if not any(it.get("m3") is not None for it in adv):
        cols = [c for c in _ADV_COLS if c[0] not in _HIST_COLS]
    if not any(it.get("age") for it in adv):
        cols = [c for c in cols if c[0] not in _AGE_TIER_KEYS]
    ws = wb.create_sheet("采购建议")
    moq = meta.get("moq") or 0
    _dm3 = meta.get("days_m3") or 90        # 近3月日历天数（实际 89~92，随月份变）
    _dq30 = meta.get("days_q30") or 30      # 近30天固定窗
    tier_txt = "；".join(
        f"{n} = ≥{lo} 件" if i == 0 else
        f"{n} = {_TIER_BOUNDS[i][1]}~{_TIER_BOUNDS[i - 1][1] - 1} 件"
        for i, (n, lo) in enumerate(_TIER_BOUNDS)
    ) + f"；T4 碎单 = ≤{_TIER_BOUNDS[-1][1] - 1} 件"
    note = (f"建议数量 = ceil(计算日均 × 计划备货天数) − 可售库存 − 在途"
            f"（预留已锁定给既有订单，不计入可售）。"
            f"「含在途可撑天数」＜「计划备货天数」⇔ 建议数量＞0，可逐行自校验。"
            f"日均回退链：本月日均(实时速度) → 上月日均 → 近3月÷{_dm3} → 近30天÷{_dq30}"
            f"（断货 SKU 本月常常没货可卖，不回退会把建议量算成 0）。"
            f"【采购动作】立即补货 = 三级（库存在途皆 0，已无货可卖）；"
            f"在途不足 = 二级（有在途但撑不到计划天数，建议下周三前到）；"
            f"排入下周三 = 一级（还能撑几天，按节奏排单）。"
            f"【数量梯队】{tier_txt}，整表按梯队从大到小排，「梯队内序号」= 该档第几单。"
            f"只列建议量 > 0 的 {len(adv)} 个 / 缺货 {live_n} 个；"
            f"未列入 = 在途已覆盖目标库存或无需补货。"
            + (f"建议量已按 {moq} 件/箱向上取整。" if moq > 1 else "")
            + (f"【库龄】七档来源《{meta.get('age_xlsx')}"
               + (f"》（期数 {meta.get('age_per')}，随月销表更新，与当日快照可能有 1~2 天时点差）"
                  if meta.get("age_per") else "》")
               + f"；「呆滞>90天」= 91 天以上各档之和，美客多行取 Resumen「陈旧库存」"
                 f"（该账号无分档）；呆滞大的 SKU 建议先清库再补货。"
               if any(it.get("age") for it in adv) else
               f"【呆滞>90天】美客多取 Resumen「陈旧库存」，其余账号本次无库龄数据。")
            # 商品资料（2026-11-07）：说明三列从哪来、以及两条压制规则
            + (f"【商品资料】类型/商品状态/销量评级/品牌/供应商取自《{os.path.basename(meta['pm_path'])}》，"
               f"共 {meta.get('pm_total')} 个 SKU。两条规则已作用到「建议数量」上："
               f"①状态=停售/不销售 → 不补；②膜类评级 E/XE → 不补（手机壳不参与评级，照常补）。"
               f"被压掉的行「采购动作」显示「—」、原因写在「停购原因」列。"
               if meta.get("pm_path") else
               "【商品资料】未找到 商品资料.xlsx，类型/商品状态/销量评级三列为空，"
               "且停售/评级 E-XE 的压制规则本次未生效。"))
    ws.cell(row=1, column=1, value=note).font = Font(italic=True, size=9, color="666666")
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(cols))
    for j, (_k, label) in enumerate(cols, start=1):
        c = ws.cell(row=2, column=j, value=label)
        c.fill, c.font = _HDR_FILL, _HDR_FONT
        c.alignment = Alignment(horizontal="center", vertical="center")
        ws.column_dimensions[get_column_letter(j)].width = ({
            "name": 40, "action": 14, "sku": 12, "site": 10,
            "by_src": 11, "buy": 11, "tier": 12, "tier_seq": 11,
        }.get(_k) or (8 if _k in _AGE_TIER_KEYS or _k == "stale"
                      else len(label) * 2 + 2))
    for i, it in enumerate(adv, start=3):
        for j, (k, _l) in enumerate(cols, start=1):
            if k in _AGE_TIER_KEYS:
                # 七档不在 item 上平铺，从「age」dict 按档名取（档名 = 断货明细口径）
                v = (it.get("age") or {}).get(S.AGE_TIER_LABEL[int(k[1:])])
            else:
                v = it.get(k)
            if isinstance(v, float):
                if k == "dos_in":
                    v = round(v, 1)                # 天数：1 位
                elif k == "by_daily":
                    v = round(v, 3)                # 日均：3 位（与明细表一致）
            ws.cell(row=i, column=j, value=v)
        fill = _PRIO_FILL.get(it.get("prio"))
        if fill:
            for j in (1, 2):
                ws.cell(row=i, column=j).fill = PatternFill("solid", fgColor=fill)
    # 合计行：采购最先看的就是总量
    r = 3 + len(adv)
    tot = ws.cell(row=r, column=1, value="合计")
    tot.font = Font(bold=True, size=9)
    ws.cell(row=r, column=6, value=f"{len(adv)} 个 SKU").font = Font(bold=True, size=9)
    # 按 key 定位「建议数量」列，别再用列数倒推（2026-11-03 拆列时踩过一次）
    buy_col = [k for k, _ in cols].index("buy") + 1
    grand = sum(it.get("buy") or 0 for it in adv)
    ws.cell(row=r, column=buy_col, value=grand).font = Font(bold=True, size=9)

    # 表尾：按数量梯队汇总（采购一眼看出哪一档吃掉了多少量）
    # 放在筛选范围之外，筛的时候不会混进汇总行
    rt = r + 2
    ws.cell(row=rt, column=1, value="按数量梯队汇总").font = Font(bold=True, size=10)
    for j, h in enumerate(("数量梯队", "SKU 数", "合计件数", "件数占比"), start=1):
        c = ws.cell(row=rt + 1, column=j, value=h)
        c.fill, c.font = _HDR_FILL, _HDR_FONT
        c.alignment = Alignment(horizontal="center", vertical="center")
    for i, t in enumerate(_TIER_NAMES):
        g = [it for it in adv if it.get("tier") == t]
        q = sum(it.get("buy") or 0 for it in g)
        row = rt + 2 + i
        ws.cell(row=row, column=1, value=t).font = Font(bold=True, size=9)
        ws.cell(row=row, column=2, value=len(g)).font = Font(size=9)
        ws.cell(row=row, column=3, value=q).font = Font(size=9)
        pc = ws.cell(row=row, column=4, value=(q / grand if grand else 0))
        pc.number_format, pc.font = "0.0%", Font(size=9)
    row = rt + 2 + len(_TIER_NAMES)
    ws.cell(row=row, column=1, value="合计").font = Font(bold=True, size=9)
    ws.cell(row=row, column=2, value=len(adv)).font = Font(bold=True, size=9)
    ws.cell(row=row, column=3, value=grand).font = Font(bold=True, size=9)
    ws.cell(row=row, column=4, value=1 if grand else 0).number_format = "0.0%"
    ws.cell(row=row, column=4).font = Font(bold=True, size=9)
    ws.freeze_panes = "A3"
    ws.auto_filter.ref = f"A2:{get_column_letter(len(cols))}{max(len(adv) + 2, 3)}"
    return ws


def _hist_tip(meta, snap_day):
    """销量列的口径说明（上月是哪个月、近3月是哪几个月）。"""
    if not meta.get("with_history", True):
        return ""
    cm = meta.get("cur_m") or int(str(snap_day)[4:6])
    prev = "%d月" % (cm - 1) if cm >= 2 else "去年12月"
    m3 = "%d~%d月" % (max(1, cm - 2), cm)
    dp = meta.get("days_prev") or 30
    return (f"（销量列：上月={prev}　近3月={m3}　今年累计=1~{cm}月，源=日销表月度块；"
            f"本月日均=实时速度=当月累计÷(已过{meta.get('elapsed', '')}天−当月缺货天数)，"
            f"上月日均=上月销量÷{dp}（{prev}日历天数）；可售天数=(可售+预留)÷本月日均）")


def write(items, meta, out_dir, snap_day):
    os.makedirs(out_dir, exist_ok=True)
    out = os.path.join(out_dir, "亚马逊美客多EMAG缺货情况表.xlsx")
    wb = Workbook()

    # 未起量 = 三级且今年从未有过销量，单独成表，不计入缺货数（2026-10-29 用户拍板）
    live = [it for it in items if not it.get("unsold")]
    unsold = [it for it in items if it.get("unsold")]

    # --no-history 时三列销量全为 None，直接不出这三列，避免留一排空列
    cols = _COLS if any(it.get("m_ytd") is not None for it in items) else \
        [c for c in _COLS if c[0] not in _HIST_COLS]

    ws = wb.active
    ws.title = "汇总"
    ws.cell(row=1, column=1, value=f"缺货情况汇总（快照 {snap_day}）").font = Font(bold=True, size=12)
    hdr = ["站点", "账号", "一级", "二级", "三级", "合计", "其中近30天有动销",
           "计划备货天数", "未起量(单列,不计入)"]
    for j, h in enumerate(hdr, start=1):
        c = ws.cell(row=2, column=j, value=h)
        c.fill, c.font = _HDR_FILL, _HDR_FONT
        c.alignment = Alignment(horizontal="center")
    agg = {}
    _KEY = {"一级": "l1", "二级": "l2", "三级": "l3"}
    for it in live:
        k = (it["site"], it["acct"])
        a = agg.setdefault(k, dict(l1=0, l2=0, l3=0, live=0, plan=it["plan"], unsold=0))
        a[_KEY[it["level"]]] += 1
        if (it.get("q30") or 0) > 0:
            a["live"] += 1
    for it in unsold:
        k = (it["site"], it["acct"])
        a = agg.setdefault(k, dict(l1=0, l2=0, l3=0, live=0, plan=it["plan"], unsold=0))
        a["unsold"] += 1
    r = 3
    for (site, acct) in sorted(agg, key=lambda x: (x[1] != "主账号", x[0])):
        a = agg[(site, acct)]
        tot = a["l1"] + a["l2"] + a["l3"]
        vals = [site, acct, a["l1"], a["l2"], a["l3"], tot, a["live"], a["plan"],
                a["unsold"] or None]
        for j, v in enumerate(vals, start=1):
            ws.cell(row=r, column=j, value=v)
        r += 1
    ws.cell(row=r, column=1, value="合计").font = Font(bold=True)
    for j, key in enumerate(["l1", "l2", "l3"], start=3):
        ws.cell(row=r, column=j, value=sum(a[key] for a in agg.values())).font = Font(bold=True)
    ws.cell(row=r, column=6, value=len(live)).font = Font(bold=True)
    ws.cell(row=r, column=7,
            value=sum(1 for it in live if (it.get("q30") or 0) > 0)).font = Font(bold=True)
    ws.cell(row=r, column=9, value=len(unsold) or None).font = Font(bold=True)
    for j in range(1, len(hdr) + 1):
        ws.column_dimensions[get_column_letter(j)].width = 14
    if unsold:
        note = ws.cell(row=r + 2, column=1,
                       value=f"「未起量(单列,不计入)」= 三级断货且今年 1~10 月从未有过销量的 SKU，"
                             f"已单列到「未起量」sheet。⚠️ 它不属于左边「三级」那列（三级只统计有销量的），"
                             f"所以这个数比三级大是正常的。")
        note.font = Font(size=9, italic=True, color="808080")

    tip = _hist_tip(meta, snap_day)
    tip = ("　" + tip) if tip else ""          # 无销量列时不要留个尾巴空格
    for lv in ("一级", "二级", "三级"):
        sub = [it for it in live if it["level"] == lv]
        sub.sort(key=lambda x: (x["site"], x["acct"], x["dos"] if x["dos"] is not None else 9e9))
        _sheet(wb.create_sheet(lv), f"{_LEVEL_STYLE[lv][1]}（{len(sub)} 个）{tip}", sub, cols)

    allr = sorted(live, key=lambda x: (_LEVEL_ORDER[x["level"]], x["site"], x["acct"]))
    _sheet(wb.create_sheet("全部缺货"), f"全部缺货（{len(allr)} 个，按等级排序）{tip}", allr, cols)

    # 采购建议（2026-11-03）：只收 live，未起量今年零销量、建议量必为 0，本就不该补货
    adv = [it for it in live if (it.get("buy") or 0) > 0]
    # 主序 = 数量梯队（用户 2026-11-03 要求「按建议数量梯队排」），
    # 梯队内再按建议数量降序、优先级、近30天销量兜底。
    adv.sort(key=lambda x: (_TIER_ORDER.get(x.get("tier"), 9), -x["buy"],
                            _PRIO_ORDER.get(x.get("prio"), 9),
                            -(x.get("q30") or 0)))
    # 梯队内序号：采购按「T1 第 1 单、第 2 单…」分批下单，序号随排序走不会跳号
    for t in _TIER_NAMES:
        for seq, it in enumerate((x for x in adv if x.get("tier") == t), 1):
            it["tier_seq"] = seq
    _advice_sheet(wb, adv, meta, len(live))

    if unsold:
        us = sorted(unsold, key=lambda x: (x["site"], x["acct"], x["sku"]))
        _sheet(wb.create_sheet("未起量"),
               f"未起量 · 今年从未有过销量（{len(us)} 个，不计入缺货数）{tip}", us, cols)

    wb.save(out)
    return out


def main(argv=None):
    ap = argparse.ArgumentParser()
    ap.add_argument("--snap", default=None)
    ap.add_argument("--day", default=None, help="快照日 YYYYMMDD（默认取目录名）")
    ap.add_argument("--lead", type=int, default=7, help="后勤周期天数（默认7）")
    ap.add_argument("--out", default=None)
    ap.add_argument("--keep-unsold", action="store_true",
                    help="未起量 SKU 不拆分，仍计入三级（回到 2026-10-29 前的口径）")
    ap.add_argument("--no-history", action="store_true",
                    help="不读月度块：不输出「上月/近3月/今年累计」三列，"
                         "也不拆未起量（省约 16 秒，回到 2026-10-28 前的行为）")
    ap.add_argument("--legacy-dos", action="store_true",
                    help="一级判定回退到「可售÷近30天日均」（2026-10-29 前口径）；"
                         "列里的两个日均和可售天数仍按日销表公式显示")
    ap.add_argument("--moq", type=int, default=0,
                    help="最小补货批量（件/箱），0=不取整；与断货明细 --moq 同义")
    a = ap.parse_args(argv)

    snap_dir = a.snap
    day = a.day
    if not snap_dir:
        snap_dir, day = C.latest_snapshot()
    if not snap_dir:
        print("找不到快照目录")
        return 1
    day = day or os.path.basename(str(snap_dir).rstrip("\\/"))

    items, meta = collect(snap_dir, day, lead=a.lead,
                          classify_unsold=not (a.keep_unsold or a.no_history),
                          with_history=not a.no_history,
                          legacy_dos=a.legacy_dos, moq=a.moq)
    out_dir = a.out or os.path.join(ANALYSIS_DIR, "亚马逊", "每日缺货", day)
    out = write(items, meta, out_dir, day)

    live = [it for it in items if not it.get("unsold")]
    unsold = [it for it in items if it.get("unsold")]
    cnt = Counter(it["level"] for it in live)
    live_n = sum(1 for it in live if (it.get("q30") or 0) > 0)
    print(f"快照 {day}：在册 {meta['total']} 个 SKU")
    print(f"缺货 {len(live)} 个 —— 一级 {cnt.get('一级',0)} / 二级 {cnt.get('二级',0)} / "
          f"三级 {cnt.get('三级',0)}（其中近30天有动销 {live_n} 个）")
    if unsold:
        print(f"另有 {len(unsold)} 个「未起量」SKU（三级、无在途、今年从未有过销量）"
              f"单列在「未起量」sheet，不计入缺货数")
    elif meta.get("degraded"):
        print("[!] 未能读到月度块，本次未拆分未起量 SKU（已按旧口径全量计入）")
    adv = [it for it in live if (it.get("buy") or 0) > 0]
    if adv:
        by_prio = Counter(it["prio"] for it in adv)
        print(f"采购建议 {len(adv)} 个 SKU，合计建议补货 "
              f"{sum(it['buy'] for it in adv)} 件 —— "
              f"P0 {by_prio.get('P0', 0)} / P1 {by_prio.get('P1', 0)} / "
              f"P2 {by_prio.get('P2', 0)} / P3 {by_prio.get('P3', 0)}")
    if meta["missing"]:
        print(f"[!] 缺报告 {len(meta['missing'])} 份：" + "；".join(meta["missing"][:4]))
    _merge = [w for w in meta["warn"] if w.startswith("[归并]")]
    _anom = [w for w in meta["warn"] if not w.startswith("[归并]")]
    for w in _anom[:5]:
        print(f"[!] {w}")
    for w in _merge[:5]:
        print(w)        # 归并是正常动作，不带 [!] 异常标记
    # 商品资料覆盖（2026-11-07）：缺资料的 SKU 在「类型/商品状态/销量评级」是空白，
    # 采购容易误以为程序漏算。默认不打断流程，只提示去补录。
    _pmm = meta.get("pm_miss") or {}
    if not meta.get("pm_path"):
        print("[!] 未找到 商品资料.xlsx —— 类型/商品状态/销量评级 三列将为空，"
              "且「评级E/XE不采购」「停售不补」两条规则不会生效")
    elif _pmm:
        print(f"[!] 商品资料未收录 {len(_pmm)} 个 SKU（三列留空、不参与停售/评级压制）："
              + "、".join(list(_pmm)[:8]) + ("…" if len(_pmm) > 8 else ""))
    print(f"已输出: {out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
