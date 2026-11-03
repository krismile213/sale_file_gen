# -*- coding: utf-8 -*-
"""订单报告日期覆盖检查 —— 堵「date range 按站点本地时间设、purchase-date 是 UTC」的洞。

背景（2026-09 实测坐实）：
  Seller Central 导订单报告时，日期选择器按**站点本地时间**过滤，而报告里的
  purchase-date 是 UTC。给北美（9 月 = PDT，UTC-7）选 9/1~9/30，实际只拿到
  UTC 9/1 07:00 ~ 10/1 07:00 —— 月初 7 小时整段不在报告里（9 月实测：
  美 64 + 加 30 + 墨 22 = 116 件）。UTC+ 的站点（欧/日/澳）只会**多包**不会丢，
  多包的行被程序按 purchase-date 前缀归属裁掉，无害。

  两套系统的归属口径不同，对账时要心里有数：
    日销表 / 月度汇总   —— 按 **UTC 日期**归属（purchase-date 前 10 位）
    月销售总表          —— 按**站点本地时间**归属（trunk/month_sale_file.py 的
                          amazon_order_tz，北美 = America/Los_Angeles）
  所以同一份报告，月初/月末 7 小时的单在两套表里归不同月，差异属口径不属丢单。

检查项（只对亚马逊 TSV 订单报告；美客多/eMAG 的 xlsx 报告自动跳过）：
  ① 月初截断 —— min(purchase-date) 晚于归属月 1 日 00:00 UTC ⇒ date range 用了本地时间
  ② 月末截断 —— max(purchase-date) 早于归属月最后一天 23:59 UTC ⇒ 同理（UTC+ 站点症状）
  ③ 没导到当天 —— max(purchase-date 日期) 早于数据日 ⇒ 报告下载时间早或缺当天
  ④ 月内单日空洞 —— 报告覆盖区间内有整天 0 单（仅对日均 ≥5 单的报告检查，避免小站点误报）
  ⑨ 渠道错位（channel_mislayout）—— 报告 sales-channel 分布 vs 本簿 expect_channel：
     整份下错站点（外来渠道 ≥50% 或无一条本店渠道）→ runner 会整份拒收

  ① ② 判定的是「按 UTC 归属（日销表/月度汇总口径）会不会缺」；月销总表按
  站点本地时间归属（amazon_order_tz），本地时间设 range 的导出对它反而不缺。

单向检查（报告 → 有没有截断）：
    python pipeline/core/report_coverage.py --snap <快照目录> --day 20261016
    python pipeline/core/report_coverage.py --report <亚马逊北美站点202610订单报告.txt>
反向审计（已导入工作簿 → 哪天丢了凌晨段，定位历史损伤）：
    python pipeline/core/report_coverage.py --book <工作簿.xlsx>
"""
import calendar
import collections
import datetime
import os
import sys

_ENCS = ("utf-8-sig", "utf-8", "shift_jis", "gb18030", "latin-1")


# ---------------- 读报告 ----------------
def read_amazon_order_report(path):
    """读亚马逊 TSV 订单报告 → dict(n, min_ts, max_date, per_day)；非订单报告返回 None。

    per_day: {purchase-date 前 10 位: 行数}；min_ts/max_date 为原始字符串字典序
    （ISO 带时区串按字典序比 == 按时间比）。
    """
    if not path or not path.lower().endswith(".txt"):
        return None
    try:
        with open(path, "rb") as f:
            raw = f.read()
    except OSError:
        return None
    txt = None
    for enc in _ENCS:
        try:
            txt = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    if txt is None:
        return None
    lines = txt.splitlines()
    hi = None
    for k, ln in enumerate(lines[:5]):
        if "amazon-order-id" in ln.lower():
            hi = k
            break
    if hi is None:
        return None
    head = lines[hi].split("\t")
    try:
        i = next(k for k, h in enumerate(head) if h.strip().lower() == "purchase-date")
    except StopIteration:
        return None
    per_day = collections.Counter()
    ch_cnt = collections.Counter()
    mn, n = None, 0
    ci = None
    try:
        ci = next(k for k, h in enumerate(head) if h.strip().lower() == "sales-channel")
    except StopIteration:
        ci = None
    for ln in lines[hi + 1:]:
        if not ln:
            continue
        p = ln.split("\t")
        if len(p) <= i:
            continue
        v = p[i].strip()
        if len(v) < 10:
            continue
        per_day[v[:10]] += 1
        if ci is not None and len(p) > ci and p[ci].strip():
            ch_cnt[p[ci].strip()] += 1
        n += 1
        if mn is None or v < mn:
            mn = v
    if not n:
        return None
    return dict(n=n, min_ts=mn, max_date=max(per_day), per_day=per_day,
                channels=ch_cnt)


# ---------------- 单份报告判定 ----------------
def _instant(s):
    """ISO 串 → aware datetime（UTC）；失败返回 None。"""
    try:
        t = datetime.datetime.fromisoformat(s.replace("Z", "+00:00"))
    except ValueError:
        return None
    if t.tzinfo is None:
        t = t.replace(tzinfo=datetime.timezone.utc)
    return t


def _lost_hours(ts, bound):
    """ts 相对 UTC 界限 bound（'YYYY-MM-01T00:00:00' 形式）晚了多少小时；没晚返回 None。"""
    a, b = _instant(ts), _instant(bound)
    if a is None or b is None or a <= b:
        return None
    return (a - b).total_seconds() / 3600.0


def _today_ym():
    return datetime.date.today().strftime("%Y%m")


def report_issues(path, order_month, day=None):
    """单份亚马逊订单报告的覆盖问题 → [(级别, 文本)]；非订单报告返回 None。"""
    got = read_amazon_order_report(path)
    if got is None:
        return None
    out = []
    y, m = order_month[:4], order_month[4:6]
    last = calendar.monthrange(int(y), int(m))[1]
    month_start = f"{y}-{m}-01T00:00:00"
    month_end = f"{y}-{m}-{last:02d}T23:59:59"

    # ① 月初截断：本地时间 date range 的典型症状（北美 UTC-7 → 丢 7 小时）
    #    注意 min(purchase-date) 是「最早一单」不是「范围起点」，小站点可能只是
    #    头几小时没单 —— 按近期日均单量分级：≥30 单/天 ✗，否则 △
    #    （报告按「最后更新日期」导出会混入 2024/2025 退款散单，取最近 7 天的量）
    per_all = got["per_day"]
    recent = sorted(per_all)[-7:]
    recent_cnt = sorted(per_all[d] for d in recent)
    med = recent_cnt[len(recent_cnt) // 2] if recent_cnt else 0
    lv_edge = "✗" if med >= 30 else "△"
    weak = "" if med >= 30 else "（近期日均单少，也可能是该时段确实无单，需人工确认）"
    lost = _lost_hours(got["min_ts"], month_start)
    if lost:
        out.append((lv_edge, f"月初被截（UTC 归属缺 {lost:.0f} 小时）：min(purchase-date)={got['min_ts']}"
                             f" > {y}-{m}-01 00:00 UTC —— date range 按站点本地时间设的"
                             f"（北美 UTC-7 = 丢 7h；改按 UTC 设或前后各多导一天）{weak}"))

    # ② 月末截断：本地时间 range 设在上月/当月末尾时，UTC 归属会缺尾段
    #    （只对已结束的月份有意义；进行中的月份报告天然只到今天）
    if month_end:
        complete = (day and day[:6] > order_month) or (not day and _today_ym() > order_month)
        if complete:
            mx = _instant(got["max_date"] + "T23:59:59")
            if mx is not None:
                # max_date 只是日期（无时刻），按当天 23:59 估；末尾时刻未知，容差 1 小时
                lost_tail = (_instant(month_end) - mx).total_seconds() / 3600.0
                if lost_tail > 1:
                    out.append((lv_edge, f"月末被截（UTC 归属缺约 {lost_tail:.0f} 小时）："
                                         f"max(purchase-date 日期)={got['max_date']}，"
                                         f"其最晚时刻应 ≥ {y}-{m}-{last:02d} 23:59 UTC{weak}"))

    # ③ 没导到当天（仅当月进行中才有意义；昨天的都缺才是问题）
    if day and day[:6] == order_month:
        ds = f"{day[:4]}-{day[4:6]}-{day[6:]}"
        if got["max_date"] < ds:
            out.append(("⚠", f"没导到数据日：max(purchase-date 日期)={got['max_date']} < {ds}"
                              f"（当天缺可接受=下载早；昨日都缺=报告没导全）"))

    # ④ 覆盖区间内的单日空洞（整天 0 单 ⇒ 导入漏一天；小站点天然有 0 单日，跳过）
    if per_all:
        if med >= 5:
            try:
                cur = datetime.date.fromisoformat(max(min(per_all), f"{y}-{m}-01"))
                end = datetime.date.fromisoformat(got["max_date"])
            except ValueError:
                cur = end = None
            holes = []
            while cur and end and cur < end:
                if f"{cur}" not in per_all:
                    holes.append(f"{cur}")
                cur += datetime.timedelta(days=1)
            if holes:
                out.append(("△", f"覆盖区间内 {len(holes)} 天 0 单：{', '.join(holes)}"))
    return out


# ---------------- ⑨ 渠道错位（整份报告下错站点） ----------------
def _channel_owners_map(workbooks):
    """{sales-channel: {簿文件名}} —— 与 runner._channel_owners() 同规则（此处自含避免循环导入）。"""
    owners = {}
    for bk in workbooks:
        specs = bk.get("sales") or []
        specs = specs if isinstance(specs, list) else [specs]
        for sc in specs:
            for ch in (sc.get("expect_channel") or ()):
                owners.setdefault(ch, set()).add(bk.get("file", "?"))
    return owners


def _resolve_report(snap_dir, day, tmpl):
    """报告候选模板 → 第一个存在的路径（与 precheck.collect / runner.process 同规则）。"""
    for t in (tmpl if isinstance(tmpl, list) else [tmpl]):
        try:
            c = t.format(d=day, d4=day[4:], m=_order_month(day))
        except (KeyError, IndexError):
            c = t
        p = c if os.path.isabs(c) else os.path.join(snap_dir, c)
        if os.path.exists(p):
            return p
    return None


_CFG = None


def _cfg():
    global _CFG
    if _CFG is None:
        try:
            import config as C
            _CFG = C
        except Exception:
            _CFG = None
    return _CFG


def _order_month(day):
    c = _cfg()
    if c is not None and getattr(c, "ORDER_MONTH", None):
        return c.ORDER_MONTH
    return day[:6] if day else ""


def channel_mislayout(snap_dir, day, workbooks):
    """跑前检测订单报告「整份下错站点」→ [(簿, sheet, 报告名, [(级别, 文本)])]。

    2026-10-19 事故：`亚马逊沙特阿拉伯站点202610订单报告.txt` 里 20018 行几乎全是
    欧洲渠道（下载时选错范围）。runner 侧现在会整份拒收，但那是动过工作簿之后；
    这里在跑前就按「本簿 expect_channel vs 报告 sales-channel 分布」判：
      ✗ 整份错位：外来渠道占比 ≥50%，或报告里一条本簿渠道都没有
      △ 部分外来：占比 10%~50%（逐行会被拒，但报告可能选错范围）
    只查 config 里配了的销售报告（没人用的多余文件不查）。
    """
    if not snap_dir or not os.path.isdir(snap_dir):
        return []
    owners = _channel_owners_map(workbooks)
    out = []
    for bk in workbooks:
        if bk.get("skip"):
            continue
        specs = bk.get("sales") or []
        specs = specs if isinstance(specs, list) else [specs]
        fname = bk.get("file", "?")
        for sc in specs:
            rp = _resolve_report(snap_dir, day, sc.get("report"))
            if not rp or not str(rp).lower().endswith(".txt"):
                continue
            got = read_amazon_order_report(rp)
            if got is None:
                continue
            own = set(sc.get("expect_channel") or ())
            chs = got["channels"]
            total = sum(chs.values())
            if not total:
                continue
            foreign = {c: k for c, k in chs.items()
                       if c not in own and not c.startswith("Non-Amazon")
                       and owners.get(c) and fname not in owners[c]}
            n_foreign = sum(foreign.values())
            n_own = sum(k for c, k in chs.items() if c in own)
            ratio = n_foreign / total
            iss = []
            if sc.get("allow_mixed_channel"):
                # 2026-10-24：interflow 沙特/阿联酋两份订单报告是**同一份下载**
                # （sha256 相同），天然混着 Amazon.ae 62 行 + Amazon.sa 5 行。
                # 两个簿各取自己渠道是设计意图，不是"下错站点" → 只做提示，不判硬伤。
                if n_foreign:
                    br = "、".join(f"{c}×{k}" for c, k in
                                   sorted(foreign.items(), key=lambda x: -x[1])[:6])
                    iss.append(("", f"混渠道报告（已按 expect_channel 逐行切分）："
                                    f"本店 {n_own} 行 / 其它渠道 {n_foreign} 行 —— {br}"))
                    out.append((fname, sc.get("sheet", "?"), os.path.basename(rp), iss))
                continue
            if n_own == 0 and n_foreign > 0:
                iss.append(("✗", f"整份错位：报告 {total} 行里没有一条本簿渠道"
                                  f"（{sc.get('expect_channel')}），全是别家的 —— 下载时选错站点/范围"))
            elif ratio >= 0.5:
                iss.append(("✗", f"整份疑似错位：外来渠道占 {ratio:.0%}"
                                  f"（{n_foreign}/{total}，本店渠道仅 {n_own} 行）"
                                  f" —— runner 会整份拒收、本簿当日销售不更新"))
            elif ratio >= 0.1:
                iss.append(("△", f"部分行属于别家簿：外来渠道占 {ratio:.0%}"
                                  f"（{n_foreign}/{total}）—— 会被逐行拒收，但也核对下是否选错范围"))
            if iss:
                br = "、".join(f"{c}×{k}" for c, k in sorted(foreign.items(), key=lambda x: -x[1])[:6])
                iss.append(("", f"外来渠道: {br}"))
                out.append((fname, sc.get("sheet", "?"), os.path.basename(rp), iss))
    return out


def render_channels(rows):
    """体检第 ⑨ 块的文本。rows = channel_mislayout() 的返回值。"""
    if not rows:
        return "⑨ 订单报告渠道错位：✅ 各簿订单报告的 sales-channel 与所属站点一致"
    n_err = sum(1 for r in rows for lv, _t in r[3] if lv == "✗")
    L = [f"⑨ ⚠ 订单报告渠道错位：{len(rows)} 份报告的站点对不上（硬伤 {n_err} 处）"
         f" —— 下载时选错范围，runner 会整份拒收"]
    for fname, sheet, rep, iss in rows:
        L.append(f"      {fname} / {sheet} ← {rep}")
        for lv, txt in iss:
            L.append(f"        {lv + ' ' if lv else '  '}{txt}")
    L.append("      修法：按正确站点/范围重下该报告（覆盖同名文件）后重跑；"
             "订单号去重，不会重复入库")
    return "\n".join(L)


# ---------------- 快照级 ----------------
def snapshot_coverage(snap_dir, day, order_month):
    """扫快照目录里所有「*订单报告.txt」→ [(报告名, [(级别, 文本)])]。"""
    out = []
    if not snap_dir or not os.path.isdir(snap_dir):
        return out
    for f in sorted(os.listdir(snap_dir)):
        p = os.path.join(snap_dir, f)
        if not os.path.isfile(p) or f.startswith("~$"):
            continue
        if "订单报告" not in f or not f.lower().endswith(".txt"):
            continue
        try:
            iss = report_issues(p, order_month, day)
        except Exception as e:
            iss = [("!", f"解析失败 {type(e).__name__}: {e}")]
        if iss:
            out.append((f, iss))
    return out


def render_snapshot(rows):
    """体检第 ⑧ 块的文本。rows = snapshot_coverage() 的返回值。"""
    if not rows:
        return "⑧ 订单报告日期覆盖：✅ 抽查的订单报告均无月初/月末截断、无单日空洞"
    n_err = sum(1 for _f, iss in rows for lv, _t in iss if lv == "✗")
    L = [f"⑧ 订单报告日期覆盖：{len(rows)} 份报告的日期范围可疑（本地时间 date range 后遗症，"
         f"其中硬伤 {n_err} 处）"]
    for f, iss in rows:
        L.append(f"      {f}")
        for lv, txt in iss:
            L.append(f"        {lv} {txt}")
    L.append("      说明：只影响**拿该报告重建月度汇总**的场景（月初/月末几小时不在报告里）；"
             "日销簿是 append-only，当月首日导入过的行不受影响（10 月实测：报告 min=06:12，"
             "簿里 00:00-06:11 的 74 单早在 10/1 首跑就已入库）。")
    L.append("      修法：以后月度报告 date range 按 UTC 设（或前后各多选一天）。"
             "※ 月销总表按站点本地时间归属（amazon_order_tz），本地时间 range 对它不构成缺数")
    return "\n".join(L)


# ---------------- 反向审计：工作簿已导入数据 ----------------
def book_audit(path, sheet_kw="销售数据"):
    """扫工作簿销售 sheet：按 UTC 日的三个时段统计订单行数，找「凌晨段整段丢失」的日子。

    本地时间 date range 的后遗症在已导入数据里的症状：某 UTC 日白天有单、
    凌晨 00:00-07:00（北美）/ 傍晚 17:00-24:00（UTC+ 站点）却 0 单。
    返回 [(日期, 全天, [00,07), [07,17), [17,24))]，已按 UTC 日排序。
    """
    import openpyxl
    out = []
    if not os.path.exists(path):
        return out
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        for sn in wb.sheetnames:
            if sheet_kw not in sn:
                continue
            ws = wb[sn]
            days = collections.defaultdict(lambda: [0, 0, 0, 0])
            pcol = None
            for ri, row in enumerate(ws.iter_rows(min_row=1, max_row=6,
                                                  max_col=40, values_only=True), 1):
                if any(str(v).strip().lower() == "purchase-date" for v in row if v):
                    pcol = next(k for k, v in enumerate(row)
                                if v and str(v).strip().lower() == "purchase-date")
                    data_from = ri + 1
                    break
            if pcol is None:
                continue
            for row in ws.iter_rows(min_row=data_from, max_col=pcol + 1, values_only=True):
                v = row[pcol] if pcol < len(row) else None
                if not (isinstance(v, str) and len(v) >= 13 and v[4] == "-"):
                    continue
                d, hh = v[:10], v[11:13]
                rec = days[d]
                rec[0] += 1
                if hh < "07":
                    rec[1] += 1
                elif hh < "17":
                    rec[2] += 1
                else:
                    rec[3] += 1
            out.extend((d, r[0], r[1], r[2], r[3]) for d, r in sorted(days.items()))
    finally:
        wb.close()
    return out


def render_book(rows, min_total=20):
    """book_audit() 结果的文本渲染；min_total=当天至少这么多单才判「凌晨段丢失」。

    两种正常情况不当丢失报：
      今天            —— 当天数据还在收（报告只导到早上），不判
      窗口下界日(min) —— WINDOW_LO 过滤把凌晨段裁掉，非导数丢失
    """
    if not rows:
        return "未找到销售 sheet（无 purchase-date 列）。"
    first, last = rows[0][0], rows[-1][0]
    today = datetime.date.today().isoformat()
    L = [f"{'UTC 日':<12}{'全天':>6}{'凌晨00-07':>11}{'午间07-17':>11}{'傍晚17-24':>11}  标记"]
    bad = 0
    for d, tot, e, m, l in rows:
        flag = ""
        if d == last == today:
            flag = "（当天数据仍在收集中）"
        elif d == first and tot >= min_total and e == 0:
            flag = "○ 窗口下界日（WINDOW_LO）凌晨段被裁，非导数丢失"
        elif tot >= min_total:
            if e == 0:
                flag = "✗ 凌晨段整段丢失（本地时间 date range 后遗症）"
            elif l == 0:
                flag = ("△ 傍晚段 0 单（当天报告只导到早上，或 UTC+ 站点尾段；"
                        "一般由次日报告补回，月末归档快照会永久缺）")
        if flag:
            if flag[0] in "✗△○":
                bad += 1
        L.append(f"{d:<12}{tot:>6}{e:>11}{m:>11}{l:>11}  {flag}")
    L.append("")
    L.append(f"结论：{len(rows)} 天中 {bad} 天有时段性异常（○/△ 多为口径或采集节奏，✗ 才是真丢失）")
    return "\n".join(L)


# ---------------- CLI ----------------
def main(argv):
    args = [a for a in argv if a != "--quiet"]
    if "--book" in args:
        p = args[args.index("--book") + 1]
        print(render_book(book_audit(p)))
        return 0
    if "--report" in args:
        p = args[args.index("--report") + 1]
        month = None
        if "--month" in args:
            month = args[args.index("--month") + 1]
        if month is None:                      # 从文件名里挖 YYYYMM
            import re
            mm = re.search(r"(?<!\d)(20\d{4})(?!\d)", os.path.basename(p))
            month = mm.group(1) if mm else datetime.date.today().strftime("%Y%m")
        day = args[args.index("--day") + 1] if "--day" in args else None
        iss = report_issues(p, month, day)
        print(f"报告: {p}\n归属月: {month}")
        if iss is None:
            print("（非亚马逊 TSV 订单报告，跳过）")
        elif not iss:
            print("✅ 日期覆盖正常")
        else:
            for lv, t in iss:
                print(f"{lv} {t}")
        return 0
    # --snap <目录> --day YYYYMMDD [--month YYYYMM]
    snap = args[args.index("--snap") + 1] if "--snap" in args else None
    day = args[args.index("--day") + 1] if "--day" in args else None
    month = args[args.index("--month") + 1] if "--month" in args else None
    if month is None and day:
        month = day[:6]
    if snap is None and day:
        sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__))))
        import config as C
        snap = os.path.join(C.SNAP_BASE, C.CURRENT_PERIOD, "快照日期", day)
    rows = snapshot_coverage(snap, day, month or datetime.date.today().strftime("%Y%m"))
    print(f"快照: {snap}\n归属月: {month}")
    print(render_snapshot(rows))
    wbs = getattr(_cfg(), "WORKBOOKS", None) if _cfg() else None
    if wbs:
        print()
        print(render_channels(channel_mislayout(snap, day, wbs)))
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
