# -*- coding: utf-8 -*-
"""东南亚「在售 · 无库存 · 无在途」SKU 及断货前 60 天日销明细 —— V3（商品资料口径）

与 V2（20261021_修正版）的唯一差别：**「在售」改取 `资产文件/商品资料.xlsx`
的「状态」列（在售/停售/不销售），不再用日销表印尼状态列 / 泰菲近60天销量代理。**

为什么改（2026-10-29 用户确认）：
  商品资料是全公司主数据，状态只有三个值，最准。
  日销表「印尼状态」的 4 个值其实是两套东西混在一起 ——
    售完截止 → 与商品资料「停售」完全对应（29/29）
    在售 / 机会 / 需求 → 与商品资料「在售」对应（但「需求」里混了 7 个停售）
  把它当在售判据，会把 39 个商品资料明确在售的 SKU 漏掉。

用法：
  python _sea_rebuild_v3_20261029.py            # 只打印统计
  python _sea_rebuild_v3_20261029.py --write     # 落盘 xlsx
"""
import os
import sys
import calendar
import datetime as dt
import collections
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.abspath(os.path.join(HERE, "..", "..", ".."))
SEA = os.path.join(ROOT, "Daily_Sales_Report", "东南亚日销")
PROD = os.path.join(os.path.dirname(ROOT), "资产文件", "商品资料.xlsx")
OUT_DIR = os.path.join(ROOT, "Daily_Sales_Report", "断货分析", "东南亚", "明细", "20261021")
OUT = os.path.join(OUT_DIR, "东南亚缺货60天明细_20261021_V3_商品资料口径.xlsx")

SNAP = dt.date(2026, 10, 21)          # 库存快照日
CUR_MONTH = 10
WIN = 60

SITES = {
    "印尼": dict(f="2026年印尼智慧物流销售表1-12月.xlsx",
                 sheets={m: f"2026年{m}月" for m in range(1, 13)}),
    "泰国": dict(f="2026年泰国易仓销售表1-12月.xlsx",
                 sheets={m: f"{m}月" for m in range(1, 13)}),
    "菲律宾": dict(f="2026年菲律宾仓销售表1-12月.xlsx",
                  sheets={m: f"{m}月" for m in range(1, 13)}),
}

RATING_VOCAB = {"A", "B", "C", "D", "E", "X", "XA", "XB", "XC", "XD", "XE"}


# ---------------------------------------------------------------- 读商品资料
def load_product():
    wb = openpyxl.load_workbook(PROD, read_only=True, data_only=True)
    ws = wb["商品资料"]
    out = {}
    for r in ws.iter_rows(min_row=3, values_only=True):
        s = str(r[0]).strip() if r[0] else ""
        if not s:
            continue
        out[s] = dict(
            state=str(r[12]).strip() if r[12] else "",
            score=str(r[5]).strip() if r[5] else "",
            typ=str(r[6]).strip() if r[6] else "",
            model=str(r[11]).strip() if r[11] else "",
        )
    wb.close()
    return out


def _num(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return 0.0


def _hdr_map(hdr):
    """按表头定位关键列。

    ⚠️ 印尼 5 月表有**两行表头**（r1 是入仓单号、r2 才是 SKU/在库库存/1..31），
    所以表头行不能写死第 1 行 —— 由调用方先用 _is_hdr() 找到真表头。
    日列也不认「第一个值为 1 的列」，而是找**连续 31 个整数 1..31** 的起点，
    否则 col24「退货入库/盘点/其他」这类文字列或货件列里的数字会误命中。
    """
    pos = {}
    for j, h in enumerate(hdr):
        t = str(h).strip() if h is not None else ""
        if t in ("在库库存", "剩余库存+在途库存"):
            pos[t] = j
        elif (t.endswith("总销售数量") or t.endswith("总销量")) and "mtotal" not in pos:
            pos["mtotal"] = j

    n = len(hdr)

    def _as_int(x):
        try:
            return int(str(x).strip())
        except (TypeError, ValueError):
            return None

    for j in range(n - 30):
        if _as_int(hdr[j]) == 1 and all(_as_int(hdr[j + k]) == k + 1 for k in range(31)):
            pos["d1"] = j
            break
    return pos


def _is_hdr(row):
    """真表头行 = 有 SKU 列名（印尼/菲律宾写「SKU」，泰国写「泰国SKU」）+ 有库存列。"""
    vals = [str(x).strip() for x in row if x is not None]
    has_sku = any(v == "SKU" or v.endswith("SKU") for v in vals)
    has_stock = any(v in ("在库库存", "剩余库存+在途库存") for v in vals)
    return has_sku and has_stock


def _find_rating_col(hdr, rows, d1):
    """评级列无表头，位于「31」列右边第 3 格（印尼 c60/泰 c57/菲 c59）。

    右侧依次是 品类 / 品牌 / 评级 / 「2026年：」/ 12月销量…，三列都无表头，
    只能靠值形态分辨：品类是「钢化膜/水凝膜」，品牌是英文名，评级是 A~E + X*。
    ⚠️ 评级列里会混进数字 0（未评级的格子写成 0），不能直接用词表做 all() 校验，
       否则整列判定失败 —— 之前就是这样漏掉的。
    """
    if d1 is None:
        return None
    j31 = d1 + 30
    for k in range(j31 + 1, min(j31 + 7, len(hdr))):
        if str(hdr[k] or "").strip():
            continue
        vals = [str(r[k]).strip() for r in rows
                if k < len(r) and r[k] is not None and str(r[k]).strip()]
        if len(vals) < 10:
            continue
        if all(v in RATING_VOCAB or v in ("0", "0.0") for v in vals) and len(set(vals)) >= 2:
            return k
    return None


def read_site(site, cfg):
    """返回 (rows_by_sku, meta)。rows_by_sku[sku][月] = dict(stock, avail, days[31], mtotal)"""
    path = os.path.join(SEA, cfg["f"])
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    data = collections.defaultdict(dict)
    names = {}
    ratings = {}
    for m, sn in cfg["sheets"].items():
        if sn not in wb.sheetnames:
            continue
        ws = wb[sn]
        it = ws.iter_rows(min_row=1, values_only=True)
        hdr = None
        for r in it:
            if _is_hdr(r):
                hdr = r
                break
        if hdr is None:
            continue
        pos = _hdr_map(hdr)
        if "在库库存" not in pos or "d1" not in pos:
            continue
        body = list(it)
        rc = _find_rating_col(hdr, body, pos["d1"])
        for r in body:
            sku = str(r[0]).strip() if r and r[0] else ""
            if not sku or sku == "合计":
                continue
            days = []
            for k in range(31):
                j = pos["d1"] + k
                days.append(_num(r[j]) if j < len(r) else 0.0)
            data[sku][m] = dict(
                stock=_num(r[pos["在库库存"]]),
                avail=_num(r[pos["剩余库存+在途库存"]]),
                days=days,
                mtotal=_num(r[pos["mtotal"]]) if "mtotal" in pos and pos["mtotal"] < len(r) else sum(days),
            )
            if m == CUR_MONTH:
                names[sku] = (str(r[1]).strip() if len(r) > 1 and r[1] else "",
                              str(r[2]).strip() if len(r) > 2 and r[2] else "")
                if rc is not None and rc < len(r):
                    ratings[sku] = str(r[rc]).strip() if r[rc] else ""
    wb.close()
    return data, names, ratings


def day_of(sku_month, month, day):
    days = sku_month.get(month, {}).get("days") or []
    return days[day - 1] if 0 < day <= len(days) else 0.0


def main(write=False):
    prod = load_product()
    result = []
    notes_by_site = {}

    for site, cfg in SITES.items():
        data, names, ratings = read_site(site, cfg)
        cur = {s: v[CUR_MONTH] for s, v in data.items() if CUR_MONTH in v}

        # 候选：商品资料在售 + 当前在库<=0 + 剩余+在途<=0
        cand = []
        for sku, v in cur.items():
            p = prod.get(sku)
            if not p or p["state"] != "在售":
                continue
            if v["stock"] <= 0 and v["avail"] <= 0:
                cand.append(sku)

        for sku in cand:
            # 最后有货月（在库库存>0 的最大月份）
            stocked = [m for m, v in data[sku].items() if v["stock"] > 0]
            last_stock_m = max(stocked) if stocked else None
            # 最后销量日
            last_sale = None
            for m in sorted(data[sku]):
                for d, q in enumerate(data[sku][m]["days"], 1):
                    if q > 0:
                        last_sale = dt.date(2026, m, d)
            # 断货起点 = max(最后有货月末+1, 最后销量日+1)
            anchor = None
            if last_stock_m:
                anchor = dt.date(2026, last_stock_m,
                                 calendar.monthrange(2026, last_stock_m)[1]) + dt.timedelta(days=1)
            if last_sale:
                a2 = last_sale + dt.timedelta(days=1)
                anchor = max(anchor, a2) if anchor else a2

            if anchor is None:
                w_start = w_end = None
                q60 = 0.0
            else:
                w_end = anchor - dt.timedelta(days=1)
                w_start = w_end - dt.timedelta(days=WIN - 1)
                q60 = 0.0
                cur_d = w_start
                while cur_d <= w_end:
                    q60 += day_of(data[sku], cur_d.month, cur_d.day)
                    cur_d += dt.timedelta(days=1)

            # 近 3 月（8/1 ~ 10/21）
            q3m = 0.0
            cur_d = dt.date(2026, 8, 1)
            while cur_d <= SNAP:
                q3m += day_of(data[sku], cur_d.month, cur_d.day)
                cur_d += dt.timedelta(days=1)

            nm, md = names.get(sku, ("", ""))
            result.append(dict(
                site=site, sku=sku,
                model=prod[sku]["model"] or md, name=nm or md,
                p_state=prod[sku]["state"], p_score=prod[sku]["score"], typ=prod[sku]["typ"],
                rating=ratings.get(sku, "") or "—",
                stock=cur[sku]["stock"], avail=cur[sku]["avail"],
                last_stock_m=last_stock_m, last_sale=last_sale,
                anchor=anchor, w_start=w_start, w_end=w_end,
                q60=q60, q3m=q3m, days=data[sku],
            ))

        notes_by_site[site] = f"{site}：主表 {len(cur)} 个 SKU，商品资料在售且零库存零在途 {len(cand)} 个"

    # ---------------------------------------------------------------- 统计
    print("=" * 78)
    for s, n in notes_by_site.items():
        print(" ", n)
    tot = collections.Counter(r["site"] for r in result)
    print("  合计：", dict(tot), "→", len(result), "个 SKU")
    print("  断货前 60 天销量合计：", int(sum(r["q60"] for r in result)), "件")
    print("  近 3 月销量合计：", int(sum(r["q3m"] for r in result)), "件")
    print()
    for site in SITES:
        rs = [r for r in result if r["site"] == site]
        q = sum(r["q60"] for r in rs)
        with_sale = sum(1 for r in rs if r["q60"] > 0)
        no_anchor = [r["sku"] for r in rs if r["anchor"] is None]
        print(f"  {site}: {len(rs)} 个 / 60天 {int(q)} 件 / 60天有销量 {with_sale} 个"
              + (f" / 无断货起点 {len(no_anchor)} 个 {no_anchor[:6]}" if no_anchor else ""))
    print("=" * 78)

    if not write:
        print("（未落盘，加 --write 生成 xlsx）")
        return result

    write_xlsx(result)
    print("已生成：", OUT)
    return result


# ---------------------------------------------------------------- 落盘
H_FILL = PatternFill("solid", fgColor="1F4E79")
H_FONT = Font(bold=True, color="FFFFFF", size=10)
T_FONT = Font(bold=True, size=13)
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def _style_header(ws, row, ncol):
    for c in range(1, ncol + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = H_FILL
        cell.font = H_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BOX


def write_xlsx(result):
    wb = openpyxl.Workbook()

    # 「从未有货/从未销量」的 SKU 没有「进入断货」的时点，60 天窗口无意义 → 单列，不计入主表
    main_rows = [r for r in result if r["anchor"] is not None]
    idle_rows = [r for r in result if r["anchor"] is None]

    # ---------- 汇总
    ws = wb.active
    ws.title = "汇总"
    ws["A1"] = "东南亚 · 在售但无库存无在途 SKU（快照 2026-10-21）· V3 商品资料口径"
    ws["A1"].font = T_FONT
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=18)
    ws["A2"] = ("「在售」取商品资料「状态」列；本表只列**有断货起点**的 SKU。"
                "今年从未备过货、从未卖出过的另见「从未备货」页。")
    ws["A2"].font = Font(size=9, color="808080")
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=18)

    hdr = ["国家", "SKU", "型号", "品名", "商品资料状态", "商品资料评级", "当月评级(10月)",
           "在库库存", "剩余+在途", "最后有货月", "最后销量日", "最早进入断货",
           "60天窗口起", "60天窗口止", "断货前60天销量", "近3月销量", "品类", "备注"]
    ws.append([])
    ws.append(hdr)
    _style_header(ws, 4, len(hdr))

    for r in sorted(main_rows, key=lambda x: (["印尼", "泰国", "菲律宾"].index(x["site"]),
                                              -(x["q60"] or 0), x["sku"])):
        note = ""
        if r["w_start"].year < 2026:
            # 日销表只有 2026 年 1-12 月，跨到去年的部分取不到 → 明示是「部分窗口」
            got = (r["w_end"] - dt.date(2026, 1, 1)).days + 1
            note = (f"⚠ 窗口跨 2025-12，2025 年数据不在日销表内，"
                    f"60 天销量只含 2026 年内的 {got} 天")
        elif r["q60"] == 0:
            note = "60 天窗口内零销量（断货已久）"
        elif (r["p_score"] or "").startswith("X"):
            note = f"新品（{r['p_score']}），断货即损失起量窗口"
        ws.append([
            r["site"], r["sku"], r["model"], r["name"], r["p_state"], r["p_score"] or "—",
            r["rating"], int(r["stock"]), int(r["avail"]),
            r["last_stock_m"] or "—",
            r["last_sale"].isoformat() if r["last_sale"] else "—",
            r["anchor"].isoformat(),
            r["w_start"].isoformat(), r["w_end"].isoformat(),
            int(r["q60"]), int(r["q3m"]),
            r["typ"], note,
        ])

    n = ws.max_row
    ws.append(["合计", "", "", "", "", "", "", "", "", "", "", "", "", "",
               int(sum(r["q60"] for r in main_rows)), int(sum(r["q3m"] for r in main_rows)), "", ""])
    for c in range(1, len(hdr) + 1):
        ws.cell(row=n + 1, column=c).font = Font(bold=True)
        ws.cell(row=n + 1, column=c).fill = PatternFill("solid", fgColor="DDEBF7")

    widths = [7, 9, 24, 34, 11, 11, 12, 9, 10, 10, 12, 13, 12, 12, 13, 10, 9, 30]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "C5"

    # ---------- 从未备货
    if idle_rows:
        sh = wb.create_sheet("从未备货")
        sh["A1"] = "在售 · 零库存零在途 · 但 2026 年从未备过货、从未卖出"
        sh["A1"].font = T_FONT
        sh["A2"] = ("这些 SKU 不存在「进入断货」的时点（库存与销量全年为 0），无法计算断货前 60 天，"
                    "故不计入汇总。多为上架后一直未排产的新品，建议单独走新品备货评审。")
        sh["A2"].font = Font(size=9, color="808080")
        sh.append([])
        head = ["国家", "SKU", "型号", "品名", "商品资料评级", "品类", "说明"]
        sh.append(head)
        _style_header(sh, 4, len(head))
        for r in sorted(idle_rows, key=lambda x: (["印尼", "泰国", "菲律宾"].index(x["site"]), x["sku"])):
            sh.append([r["site"], r["sku"], r["model"], r["name"], r["p_score"] or "—",
                       r["typ"], "2026 年库存与销量均为 0"])
        for i, w in enumerate([7, 9, 26, 40, 12, 10, 26], 1):
            sh.column_dimensions[get_column_letter(i)].width = w
        sh.freeze_panes = "C5"

    # ---------- 明细：按 (国家, 月份)
    groups = collections.defaultdict(list)
    for r in main_rows:
        d = r["w_start"]
        while d <= r["w_end"]:
            groups[(r["site"], d.month)].append(r)
            d = dt.date(d.year + (d.month == 12), 1 if d.month == 12 else d.month + 1, 1)

    order = ["印尼", "泰国", "菲律宾"]
    for site in order:
        for m in range(1, CUR_MONTH + 1):
            rs = groups.get((site, m))
            if not rs:
                continue
            sh = wb.create_sheet(f"{site}-2026-{m:02d}")
            sh["A1"] = f"{site} · 2026年{m}月 · 断货前 60 天日销明细"
            sh["A1"].font = T_FONT
            head = ["国家", "SKU", "型号", "品名", "评级(当月)", "断货起点", "60天窗口"]
            head += [f"{d}日" for d in range(1, 32)]
            head += ["本月合计", "窗口60天合计", "备注"]
            sh.append([])
            sh.append(head)
            _style_header(sh, 3, len(head))
            for r in sorted(rs, key=lambda x: -(x["q60"] or 0)):
                days = [int(day_of(r["days"], m, d)) for d in range(1, 32)]
                mt = int(r["days"].get(m, {}).get("mtotal", sum(days)) or sum(days))
                sh.append([r["site"], r["sku"], r["model"], r["name"],
                           r["rating"],
                           r["anchor"].isoformat(),
                           f'{r["w_start"].isoformat()}~{r["w_end"].isoformat()}']
                          + days + [mt, int(r["q60"]), ""])
            sh.column_dimensions["A"].width = 7
            sh.column_dimensions["B"].width = 9
            sh.column_dimensions["C"].width = 22
            sh.column_dimensions["D"].width = 30
            for c in range(5, 8):
                sh.column_dimensions[get_column_letter(c)].width = 12
            for c in range(8, 39):
                sh.column_dimensions[get_column_letter(c)].width = 4.5
            sh.freeze_panes = "E4"

    # ---------- 口径说明
    sh = wb.create_sheet("口径说明")
    lines = [
        ("东南亚「在售 · 无库存 · 无在途」SKU 及断货前 60 天日销明细 · V3", T_FONT),
        ("", None),
        ("零、先排除一个怀疑：SKU 识别没问题", Font(bold=True, size=11)),
        ("三国 10 月主表 SKU 与商品资料的匹配率：印尼 389/389 = 100%，泰国 393/393 = 100%，", None),
        ("菲律宾 409/410 = 99.8%（仅 G001 一个查无此 SKU）。所以上一版与备货表的出入，", None),
        ("不是 SKU 对不上，而是「在售」这个判据本身取错了源。", None),
        ("", None),
        ("一、与上一版（V2 / 修正版）的唯一差别", Font(bold=True, size=11)),
        ("「在售」改取《资产文件/商品资料.xlsx》「商品资料」页 M 列「状态」= 在售。", None),
        ("V2 用的是日销表口径：印尼看「印尼状态」列 = 在售；泰国、菲律宾没有状态列，用「近 60 天有销量」代理。", None),
        ("", None),
        ("二、为什么改（实测依据，快照 2026-10-21）", Font(bold=True, size=11)),
        ("日销表「印尼状态」列的 4 个值，与商品资料「状态」交叉后是两套东西：", None),
        ("  · 售完截止 29 个 → 商品资料 29 个全部「停售」，两边一致", None),
        ("  · 在售   6 个 → 商品资料 6 个全部「在售」，两边一致", None),
        ("  · 机会   7 个 → 商品资料 7 个全部「在售」（其中 6 个 10 月还有销量）", None),
        ("  · 需求  39 个 → 商品资料 32 个「在售」+ 7 个「停售」（在售的 32 个 10 月销量均为 0）", None),
        ("结论：把「印尼状态 = 在售」当唯一判据，会漏掉 39 个商品资料明确标在售的 SKU；", None),
        ("      用「近 60 天有销量」代理，又会把已停售但近期清库存的 SKU 误收。", None),
        ("", None),
        ("三、评级体系（商品资料 A1 原文）", Font(bold=True, size=11)),
        ("A：1 个月销量 500 套及以上；B：3 个月 500 套及以上；C：6 个月 250 套及以上；", None),
        ("D：6 个月 200-250 套；E：6 个月 200 套以下。", None),
        ("X 类为新品：XA 日均≥16.7｜XB 16.7>日均≥5.5｜XC 5.5>日均≥1.38｜XD 1.38>日均≥1.11｜", None),
        ("XE 日均<1.11｜X 暂无销售数据。评级为 E 或 XE 的不要采购。", None),
        ("即 XB / XC 是「新品的 B 档 / C 档」，与 B / C 同档可比，不是另一套体系。", None),
        ("", None),
        ("四、入选口径", Font(bold=True, size=11)),
        ("1. 在售：商品资料「状态」= 在售（停售、不销售均排除）", None),
        ("2. 无库存且无在途：10/21 快照「在库库存」≤0 且「剩余库存+在途库存」≤0", None),
        ("3. 断货起点 = max（最后有货月月末+1，最后销量日+1）；60 天窗口 = [起点-60, 起点-1]", None),
        ("4. 日销数量取各月主表 1..31 日列（不是订单明细表）——批量入账只记在月表日列", None),
        ("5. 今年库存与销量均为 0 的（无「进入断货」时点）单列到「从未备货」页，不计入汇总", None),
        ("", None),
        ("五、两版数量对照", Font(bold=True, size=11)),
        ("          V2（日销表口径）              V3（商品资料口径）", None),
        ("  印尼    4 个 / 54 件                  13 个 / 71 件", None),
        ("  泰国    3 个 / 23 件                   1 个 /  7 件", None),
        ("  菲律宾 20 个 / 587 件                 19 个 / 586 件", None),
        ("  合计   27 个 / 664 件                 33 个 / 664 件", None),
        ("说明：V2 与 V3 共有的 24 个 SKU，60 天销量逐条一致（0 处差异）。", None),
        ("      泰国 3→1 是剔掉 K007、G412（商品资料停售），菲律宾 20→19 是剔掉 G136B（停售）；", None),
        ("      印尼 4→13 是补回 9 个「机会 / 需求」但商品资料在售的 SKU。", None),
        ("      合计件数两版都是 664 属巧合：剔掉的 3 个停售共 17 件，补回的 9 个共 17 件。", None),
        ("", None),
        ("六、已知限制", Font(bold=True, size=11)),
        ("· 商品资料「状态」是全公司唯一值，不区分站点。个别 SKU 可能在 A 站在售、B 站停售，", None),
        ("  此时本表按全公司在售处理。V2 的日销表状态列反而是分站点的，但只有印尼有。", None),
        ("· 商品资料评级列是「9月销量得分」（滞后一期）；本表「当月评级(10月)」取自日销表 10 月表，", None),
        ("  两列并列，可自行比对，不做取舍。", None),
        ("· 「不销售」3 个 SKU 一律排除。", None),
        ("· 日销表只有 2026 年 1-12 月。个别 SKU 断货早于 1 月，60 天窗口会跨到 2025-12，", None),
        ("  那一段取不到，本表按「2026 年内部分」统计并在备注里标注（本期仅印尼 G257 一个）。", None),
        ("· 明细 sheet 只覆盖 2026 年 1-10 月，窗口跨年时 12 月部分不在表内，同上。", None),
    ]
    for i, (t, f) in enumerate(lines, 1):
        c = sh.cell(row=i, column=1, value=t)
        if f:
            c.font = f
    sh.column_dimensions["A"].width = 120

    wb.save(OUT)


if __name__ == "__main__":
    main(write="--write" in sys.argv)
