# -*- coding: utf-8 -*-
"""印尼「在售 · 无库存 · 无在途」SKU 及断货前 60 天日销明细 —— 快照 2026-10-29

数据源（用户 2026-10-30 提供的桌面文件）：
  日销表 C:/Users/dell/Desktop/2026年印尼智慧物流销售表1-12月.xlsx   快照 2026-10-29（今日日期列）
  备货表 C:/Users/dell/Desktop/东南亚商品资料及库存管理.xlsx          快照 2026-10-29
         （反推依据：Indonesia stock 的 In stock 与日销表「在库库存」353 个 SKU 逐个相等、0 处差异）

「在售」判据：备货表 Indonesia stock 的「状态」列 = 在售（分站点，比全公司口径准）。
「无库存且无在途」：备货表 In stock ≤ 0 且 In transit ≤ 0。
  ⚠️ 备货表 In transit 含**已排产的计划到货**（周列 0923/1009/1021/1104/1118），
     所以它比日销表「剩余+在途」大 —— 用哪边口径结论差很多（34 vs 77），详见 --compare。

用法：python _sea_indonesia_20261029.py [--write]
"""
import os
import sys
import calendar
import datetime as dt
import collections
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

SALES = r"C:\Users\dell\Desktop\2026年印尼智慧物流销售表1-12月.xlsx"
STOCK = r"C:\Users\dell\Desktop\东南亚商品资料及库存管理.xlsx"

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.abspath(os.path.join(HERE, "..", "..", ".."))
OUT_DIR = os.path.join(ROOT, "Daily_Sales_Report", "断货分析", "东南亚", "明细", "20261029")
OUT = os.path.join(OUT_DIR, "印尼缺货60天明细_日销表20261029_备货表20261029.xlsx")

SNAP = dt.date(2026, 10, 29)
CUR = 10
WIN = 60
RATING_VOCAB = {"A", "B", "C", "D", "E", "X", "XA", "XB", "XC", "XD", "XE"}


def num(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return 0.0


def txt(v):
    return str(v).strip() if v is not None else ""


def _is_hdr(row):
    vals = [txt(x) for x in row if x is not None]
    return any(v == "SKU" or v.endswith("SKU") for v in vals) and \
           any(v in ("在库库存", "剩余库存+在途库存") for v in vals)


def _hdr_map(hdr):
    pos = {}
    for j, h in enumerate(hdr):
        t = txt(h)
        if t in ("在库库存", "剩余库存+在途库存"):
            pos[t] = j
        elif (t.endswith("总销售数量") or t.endswith("总销量")) and "mtotal" not in pos:
            pos["mtotal"] = j
    for j in range(len(hdr) - 30):
        ok = True
        for k in range(31):
            try:
                if int(txt(hdr[j + k])) != k + 1:
                    ok = False
                    break
            except (TypeError, ValueError):
                ok = False
                break
        if ok:
            pos["d1"] = j
            break
    return pos


def read_sales():
    """→ data[sku][月] = dict(stock, avail, days[31], mtotal)；names；ratings（10 月表）"""
    wb = openpyxl.load_workbook(SALES, read_only=True, data_only=True)
    data = collections.defaultdict(dict)
    names, ratings = {}, {}
    for m in range(1, 13):
        sn = f"2026年{m}月"
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
        # 评级列：31 日右边第 3 格、无表头、值混数字 0
        rc = None
        for k in range(pos["d1"] + 31, pos["d1"] + 37):
            if txt(hdr[k]):
                continue
            vals = [txt(r[k]) for r in body if k < len(r) and txt(r[k])]
            v2 = [v for v in vals if v not in ("0", "0.0")]
            if len(v2) >= 10 and all(v in RATING_VOCAB for v in v2) and len(set(v2)) >= 2:
                rc = k
                break
        for r in body:
            sku = txt(r[0])
            if not sku or sku == "合计":
                continue
            days = [num(r[pos["d1"] + k]) if pos["d1"] + k < len(r) else 0.0 for k in range(31)]
            data[sku][m] = dict(
                stock=num(r[pos["在库库存"]]),
                avail=num(r[pos["剩余库存+在途库存"]]),
                days=days,
                mtotal=num(r[pos["mtotal"]]) if "mtotal" in pos and pos["mtotal"] < len(r) else sum(days),
            )
            if m == CUR:
                names[sku] = (txt(r[1]) if len(r) > 1 else "", txt(r[2]) if len(r) > 2 else "",
                              txt(r[9]) if len(r) > 9 else "")
                if rc is not None and rc < len(r):
                    ratings[sku] = txt(r[rc])
    wb.close()
    return data, names, ratings


def read_stock():
    wb = openpyxl.load_workbook(STOCK, read_only=True, data_only=True)
    ws = wb["Indonesia stock"]
    rows = list(ws.iter_rows(min_row=1, values_only=True))
    hdr = rows[0]
    out = {}
    for r in rows[2:]:
        sku = txt(r[0])
        if not sku:
            continue
        out[sku] = dict(
            rating=txt(r[1]), state=txt(r[2]), remark=txt(r[3]),
            instock=num(r[6]), st=num(r[7]), transit=num(r[8]),
            weeks={txt(hdr[j]): num(r[j]) for j in range(11, 17)},
        )
    wb.close()
    return out


def day_of(days_by_m, m, d):
    arr = days_by_m.get(m, {}).get("days") or []
    return arr[d - 1] if 0 < d <= len(arr) else 0.0


def main(write=False):
    data, names, ratings = read_sales()
    stock = read_stock()

    # ---- 口径对照
    z_stock = [s for s, v in stock.items() if v["instock"] <= 0]
    a = [s for s in z_stock if stock[s]["transit"] <= 0 and stock[s]["state"] == "在售"]
    z_sales = [s for s, v in data.items() if CUR in v and v[CUR]["stock"] <= 0 and v[CUR]["avail"] <= 0]
    c = [s for s in z_sales if names.get(s, ("", "", ""))[2] == "在售"]

    print("=" * 80)
    print(f"印尼 · 快照 2026-10-29")
    print(f"  备货表 353 个：零库存 {len(z_stock)}，其中无在途 "
          f"{sum(1 for s in z_stock if stock[s]['transit'] <= 0)}")
    print(f"  ① 备货表口径（在售+In stock≤0+In transit≤0）     → {len(a)} 个")
    print(f"  ② 日销表口径（印尼状态=在售+在库≤0+剩余在途≤0）  → {len(c)} 个")
    print(f"  ③ 日销表：只要求零库存零在途（不看状态）         → {len(z_sales)} 个")
    print("=" * 80)

    rows = []
    for sku in a:
        s = stock[sku]
        nm, md, dstate = names.get(sku, ("", "", ""))
        # 最后有货月 / 最后销量日（用日销表全年）
        stocked = [m for m, v in data.get(sku, {}).items() if v["stock"] > 0]
        last_stock_m = max(stocked) if stocked else None
        last_sale = None
        for m in sorted(data.get(sku, {})):
            for d, q in enumerate(data[sku][m]["days"], 1):
                if q > 0:
                    last_sale = dt.date(2026, m, d)
        anchor = None
        if last_stock_m:
            anchor = dt.date(2026, last_stock_m,
                             calendar.monthrange(2026, last_stock_m)[1]) + dt.timedelta(days=1)
        if last_sale:
            a2 = last_sale + dt.timedelta(days=1)
            anchor = max(anchor, a2) if anchor else a2

        if anchor is None:
            w_start = w_end = None
            q60 = q3 = 0.0
        else:
            w_end = anchor - dt.timedelta(days=1)
            w_start = w_end - dt.timedelta(days=WIN - 1)
            q60 = 0.0
            d0 = w_start
            while d0 <= w_end:
                q60 += day_of(data[sku], d0.month, d0.day)
                d0 += dt.timedelta(days=1)
        q3 = 0.0
        d0 = dt.date(2026, 8, 1)
        while d0 <= SNAP:
            q3 += day_of(data.get(sku, {}), d0.month, d0.day)
            d0 += dt.timedelta(days=1)

        rows.append(dict(
            sku=sku, model=md, name=nm, dstate=dstate,
            k_state=s["state"], k_rating=s["rating"], rating=ratings.get(sku, "") or "—",
            instock=s["instock"], transit=s["transit"],
            weeks={k: v for k, v in s["weeks"].items() if v > 0},
            last_stock_m=last_stock_m, last_sale=last_sale, anchor=anchor,
            w_start=w_start, w_end=w_end, q60=q60, q3=q3, days=data.get(sku, {}),
        ))

    main_rows = [r for r in rows if r["anchor"] is not None]
    idle_rows = [r for r in rows if r["anchor"] is None]
    print(f"\n清单 {len(rows)} 个：有断货起点 {len(main_rows)}｜今年从未备货 {len(idle_rows)}")
    print(f"断货前 60 天销量合计 {int(sum(r['q60'] for r in main_rows))} 件")
    for r in sorted(main_rows, key=lambda x: -x["q60"]):
        print(f"   {r['sku']:8s} {r['k_rating'] or '-':3s} 起点{r['anchor']} 60天{int(r['q60']):4d} "
              f"近3月{int(r['q3']):3d} {r['name'][:28]}")
    if idle_rows:
        print(f"\n从未备货 {len(idle_rows)} 个：" + " ".join(r["sku"] for r in idle_rows))

    if not write:
        print("\n（未落盘，加 --write 生成 xlsx）")
        return rows
    write_xlsx(main_rows, idle_rows)
    print("\n已生成：", OUT)
    return rows


# ---------------------------------------------------------------- 落盘
H_FILL = PatternFill("solid", fgColor="1F4E79")
H_FONT = Font(bold=True, color="FFFFFF", size=10)
T_FONT = Font(bold=True, size=13)
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def _hdr(ws, row, n):
    for c in range(1, n + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = H_FILL
        cell.font = H_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BOX


def write_xlsx(main_rows, idle_rows):
    os.makedirs(OUT_DIR, exist_ok=True)
    wb = openpyxl.Workbook()

    ws = wb.active
    ws.title = "汇总"
    ws["A1"] = "印尼 · 在售但无库存无在途 SKU 及断货前 60 天日销明细"
    ws["A1"].font = T_FONT
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=18)
    ws["A2"] = ("数据快照：日销表 2026-10-29（今日日期列）｜备货表 2026-10-29"
                "（In stock 与日销表在库库存 353 个 SKU 逐个相等，反推同日）"
                "　在售取备货表「状态」列；无在途取备货表 In transit（含已排产计划到货）")
    ws["A2"].font = Font(size=9, color="808080")
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=18)

    head = ["SKU", "型号", "品名", "备货表状态", "备货表评级", "日销表状态", "当月评级(10月)",
            "In stock", "In transit", "计划到货周", "最后有货月", "最后销量日", "最早进入断货",
            "60天窗口起", "60天窗口止", "断货前60天销量", "近3月销量", "备注"]
    ws.append([])
    ws.append(head)
    _hdr(ws, 4, len(head))

    for r in sorted(main_rows, key=lambda x: -x["q60"]):
        note = ""
        if r["w_start"].year < 2026:
            got = (r["w_end"] - dt.date(2026, 1, 1)).days + 1
            note = f"⚠ 窗口跨 2025-12，只含 2026 年内 {got} 天"
        elif r["q60"] == 0:
            note = "60 天窗口内零销量（断货已久）"
        elif (r["k_rating"] or "").startswith("X"):
            note = f"新品（{r['k_rating']}）"
        ws.append([
            r["sku"], r["model"], r["name"], r["k_state"], r["k_rating"] or "—",
            r["dstate"] or "—", r["rating"],
            int(r["instock"]), int(r["transit"]),
            "/".join(f"{k}:{int(v)}" for k, v in r["weeks"].items()) or "—",
            r["last_stock_m"] or "—",
            r["last_sale"].isoformat() if r["last_sale"] else "—",
            r["anchor"].isoformat(), r["w_start"].isoformat(), r["w_end"].isoformat(),
            int(r["q60"]), int(r["q3"]), note,
        ])
    n = ws.max_row
    ws.append(["合计"] + [""] * 14 + [int(sum(r["q60"] for r in main_rows)),
                                      int(sum(r["q3"] for r in main_rows)), ""])
    for c in range(1, len(head) + 1):
        ws.cell(row=n + 1, column=c).font = Font(bold=True)
        ws.cell(row=n + 1, column=c).fill = PatternFill("solid", fgColor="DDEBF7")
    for i, w in enumerate([9, 24, 34, 11, 11, 11, 12, 10, 10, 18, 10, 12, 13, 12, 12, 13, 10, 26], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "B5"

    if idle_rows:
        sh = wb.create_sheet("从未备货")
        sh["A1"] = "在售 · 零库存 · 无在途 · 但 2026 年从未备过货、从未卖出"
        sh["A1"].font = T_FONT
        sh["A2"] = "没有「进入断货」的时点，算不出断货前 60 天，故不计入汇总。采购上属于「上架了但一直没排产」。"
        sh["A2"].font = Font(size=9, color="808080")
        sh.append([])
        h2 = ["SKU", "型号", "品名", "备货表评级", "日销表状态", "In stock", "In transit", "说明"]
        sh.append(h2)
        _hdr(sh, 4, len(h2))
        for r in sorted(idle_rows, key=lambda x: x["sku"]):
            sh.append([r["sku"], r["model"], r["name"], r["k_rating"] or "—",
                       r["dstate"] or "—", int(r["instock"]), int(r["transit"]),
                       "2026 年库存与销量均为 0"])
        for i, w in enumerate([9, 26, 40, 12, 12, 10, 10, 30], 1):
            sh.column_dimensions[get_column_letter(i)].width = w
        sh.freeze_panes = "B5"

    # 明细
    groups = collections.defaultdict(list)
    for r in main_rows:
        d = r["w_start"]
        while d <= r["w_end"]:
            groups[d.month].append(r)
            d = dt.date(d.year + (d.month == 12), 1 if d.month == 12 else d.month + 1, 1)
    for m in range(1, CUR + 1):
        rs = groups.get(m)
        if not rs:
            continue
        sh = wb.create_sheet(f"印尼-2026-{m:02d}")
        sh["A1"] = f"印尼 · 2026年{m}月 · 断货前 60 天日销明细"
        sh["A1"].font = T_FONT
        h2 = ["SKU", "型号", "品名", "评级(备货表)", "断货起点", "60天窗口"]
        h2 += [f"{d}日" for d in range(1, 32)] + ["本月合计", "窗口60天合计"]
        sh.append([])
        sh.append(h2)
        _hdr(sh, 3, len(h2))
        for r in sorted(rs, key=lambda x: -x["q60"]):
            days = [int(day_of(r["days"], m, d)) for d in range(1, 32)]
            mt = int(r["days"].get(m, {}).get("mtotal", sum(days)) or sum(days))
            sh.append([r["sku"], r["model"], r["name"], r["k_rating"] or "—",
                       r["anchor"].isoformat(),
                       f'{r["w_start"].isoformat()}~{r["w_end"].isoformat()}']
                      + days + [mt, int(r["q60"])])
        for i, w in enumerate([9, 22, 30, 12, 12, 24], 1):
            sh.column_dimensions[get_column_letter(i)].width = w
        for c in range(7, 39):
            sh.column_dimensions[get_column_letter(c)].width = 4.5
        sh.freeze_panes = "D4"

    # 口径说明
    sh = wb.create_sheet("口径说明")
    lines = [
        ("印尼缺货明细 · 快照 2026-10-29", T_FONT),
        ("", None),
        ("一、数据源与快照日", Font(bold=True, size=11)),
        ("日销表：C:/Users/dell/Desktop/2026年印尼智慧物流销售表1-12月.xlsx", None),
        ("        快照日取「2026年10月」页「今日日期」列 = 2026-10-29（389 行全部同一天）", None),
        ("备货表：C:/Users/dell/Desktop/东南亚商品资料及库存管理.xlsx", None),
        ("        该表没有日期字段，快照日是反推的：Indonesia stock 的 In stock 与日销表「在库库存」", None),
        ("        在 353 个 SKU 上逐个相等（0 处差异，合计 9375 vs 9374，差 1 为取整），故同为 2026-10-29", None),
        ("", None),
        ("二、两个文件核对结果", Font(bold=True, size=11)),
        ("① SKU 覆盖：备货表 353 个全部在日销表内；日销表多出 36 个（G102/G104/G107… 等老款），备货表未管理", None),
        ("② 在库库存：353 个逐个一致，0 处差异", None),
        ("③ 在途：47 个不一致，全部是「日销表 < 备货表」", None),
        ("        原因：备货表 In transit 含**已排产的计划到货**（周列 0923/1009/1021/1104/1118）", None),
        ("        日销表「剩余库存+在途库存」只统计当前实际在途", None),
        ("④ 状态：日销表「印尼状态」有 4 值，备货表只有 3 值，交叉后：", None),
        ("        在售→在售 236｜机会→在售 67｜需求→在售 39", None),
        ("        售完截止→售完即止 7｜在售→售完即止 2｜机会→售完即止 1｜在售→停售 1", None),
        ("        即日销表的「机会」「需求」在备货表里都算「在售」，两者不是同一套分类", None),
        ("⑤ 评级：353 个中 304 个完全一致（86.1%），差异多为差一档（A→B 10 个、XB→B 6 个、XC→XE 5 个）", None),
        ("", None),
        ("三、入选口径", Font(bold=True, size=11)),
        ("1. 在售：备货表「状态」= 在售（分站点口径，比全公司主数据准）", None),
        ("2. 无库存且无在途：备货表 In stock ≤ 0 且 In transit ≤ 0", None),
        ("   ⚠️ In transit 含计划到货，所以「无在途」= 连排产都没有，是真正需要下采购单的", None),
        ("3. 断货起点 = max（最后有货月月末+1，最后销量日+1）；60 天窗口 = [起点-60, 起点-1]", None),
        ("4. 日销数量取日销表各月 1..31 日列", None),
        ("5. 2026 年库存与销量均为 0 的单列「从未备货」页", None),
        ("", None),
        ("四、口径差异（为什么和上一版不一样）", Font(bold=True, size=11)),
        ("本表 34 个 = 备货表「在售 + 零库存 + 无在途」。若换成日销表口径：", None),
        ("  · 日销表「印尼状态=在售」+ 零库存 + 零在途 → 只有 3 个", None),
        ("  · 日销表只要求零库存 + 零在途（不看状态）→ 77 个", None),
        ("差异来自两处：① 在售判据（备货表分站点、日销表 4 值混两套）；② 在途是否含计划到货", None),
        ("", None),
        ("五、与上一版（V3，10-22 快照）的对照", Font(bold=True, size=11)),
        ("V3 印尼有 13 个 SKU 有断货起点，到 10-29 只剩 5 个（G178/G257/G422/G576/G402）有起点，", None),
        ("另外 8 个都已经有了计划到货，排产已录进备货表：", None),
        ("  G572 1104 到 55 件｜G567、G569 1118 各 6 件｜G552 1104 到 8 件", None),
        ("  G577、G578、G579、G580 1118 各 132 件", None),
        ("这 8 个就是上一版邮件里提到的「已在补货队列」的那批，现在补货计划落地了。", None),
        ("", None),
        ("六、已知限制", Font(bold=True, size=11)),
        ("· 本表只覆盖印尼。泰国、菲律宾的最新日销表未提供，仍为 10-22 快照。", None),
        ("· 备货表评级列与日销表 10 月评级列并列，两者时点不同（约 13.9% 有差异），未做取舍。", None),
        ("· 日销表只有 2026 年数据，断货早于 1 月的 SKU 窗口会跨到 2025-12，该段取不到。", None),
    ]
    for i, (t, f) in enumerate(lines, 1):
        c = sh.cell(row=i, column=1, value=t)
        if f:
            c.font = f
    sh.column_dimensions["A"].width = 118

    wb.save(OUT)


if __name__ == "__main__":
    main(write="--write" in sys.argv)
