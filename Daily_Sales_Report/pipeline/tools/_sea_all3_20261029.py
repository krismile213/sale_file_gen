# -*- coding: utf-8 -*-
"""东南亚三国「在售 · 无库存 · 无在途」SKU 及断货前 60 天日销明细 —— 快照 2026-10-29

数据源（5 个文件，均为 2026-10-30 用户提供）：
  日销表（三国）  桌面 2026年{印尼智慧物流|泰国易仓|菲律宾仓}销售表1-12月.xlsx
                  快照日 = 「今日日期/今日时间」列，三国 10 月页均为 2026-10-29
  备货表          桌面 东南亚商品资料及库存管理.xlsx
                  无日期字段，反推同为 10-29：三国 stock 页 In stock 与日销表「在库库存」
                  合计逐个对上（印尼 9375/9374 差 1 取整、泰国 8915=8915、菲律宾 7322=7322）
  商品资料        资产文件/商品资料.xlsx —— 只做对照列（品类 / 全平台状态 / 全平台评级）

「在售」判据：备货表对应国家 stock 页的「状态」列 = 在售（分站点，比全公司主数据准）
「无库存且无在途」：In stock ≤ 0 且 In transit ≤ 0（In transit 含已排产的计划到货）

⚠️ 备货表三国 stock 页**列位置不一样**：
   印尼/泰国  状态 col2、In stock col6、In transit col8
   菲律宾     状态 col3（多一列「系统SKU」）、In stock col8、In transit col10
   所以一律按表头名定位，不能写死列号。

用法：python _sea_all3_20261029.py [--write]
"""
import os
import sys
import calendar
import datetime as dt
import collections
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

DESK = r"C:\Users\dell\Desktop"
HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.abspath(os.path.join(HERE, "..", "..", ".."))
PROD = os.path.join(os.path.dirname(ROOT), "资产文件", "商品资料.xlsx")

OUT_DIR = os.path.join(ROOT, "Daily_Sales_Report", "断货分析", "东南亚", "明细", "20261029")
OUT = os.path.join(OUT_DIR, "东南亚缺货60天明细_日销表20261029_备货表20261029.xlsx")

SNAP = dt.date(2026, 10, 29)
CUR = 10
WIN = 60
RATING_VOCAB = {"A", "B", "C", "D", "E", "X", "XA", "XB", "XC", "XD", "XE"}

SITES = [
    dict(name="印尼", sales="2026年印尼智慧物流销售表1-12月.xlsx",
         sheet=lambda m: f"2026年{m}月", stock="Indonesia stock"),
    dict(name="泰国", sales="2026年泰国易仓销售表1-12月.xlsx",
         sheet=lambda m: f"{m}月", stock="Thailand stock"),
    dict(name="菲律宾", sales="2026年菲律宾仓销售表1-12月.xlsx",
         sheet=lambda m: f"{m}月", stock="Philippines stock"),
]
ORDER = ["印尼", "泰国", "菲律宾"]


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


def read_sales(path, sheet_of):
    data = collections.defaultdict(dict)
    names, ratings = {}, {}
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    for m in range(1, 13):
        sn = sheet_of(m)
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


def read_stock(sheet_name):
    """按表头名定位：状态 / In stock / In transit / 周列。"""
    wb = openpyxl.load_workbook(os.path.join(DESK, "东南亚商品资料及库存管理.xlsx"),
                                read_only=True, data_only=True)
    ws = wb[sheet_name]
    rows = list(ws.iter_rows(min_row=1, values_only=True))
    hdr = rows[0]
    c_state = c_instock = c_transit = None
    weeks = []
    for j, h in enumerate(hdr):
        t = txt(h)
        if t == "状态":
            c_state = j
        elif t == "In stock":
            c_instock = j
        elif t == "In transit":
            c_transit = j
        elif t.isdigit() and len(t) == 4:
            weeks.append((t, j))
    out = {}
    for r in rows[2:]:
        sku = txt(r[0])
        if not sku:
            continue
        out[sku] = dict(
            rating=txt(r[1]),
            state=txt(r[c_state]) if c_state is not None else "",
            remark=txt(r[3]),
            instock=num(r[c_instock]),
            transit=num(r[c_transit]),
            weeks={w: num(r[j]) for w, j in weeks if num(r[j]) > 0},
        )
    wb.close()
    return out


def read_product():
    wb = openpyxl.load_workbook(PROD, read_only=True, data_only=True)
    ws = wb["商品资料"]
    out = {}
    for r in ws.iter_rows(min_row=3, values_only=True):
        s = txt(r[0])
        if not s:
            continue
        out[s] = dict(p_state=txt(r[12]), p_score=txt(r[5]), typ=txt(r[6]),
                      model=txt(r[11]))
    wb.close()
    return out


def day_of(days_by_m, m, d):
    arr = days_by_m.get(m, {}).get("days") or []
    return arr[d - 1] if 0 < d <= len(arr) else 0.0


def main(write=False):
    prod = read_product()
    all_rows = []
    idle_all = []
    stats = {}

    for cfg in SITES:
        site = cfg["name"]
        data, names, ratings = read_sales(os.path.join(DESK, cfg["sales"]), cfg["sheet"])
        stock = read_stock(cfg["stock"])

        cur = {s: v[CUR] for s, v in data.items() if CUR in v}
        cand = [s for s in stock
                if stock[s]["state"] == "在售"
                and stock[s]["instock"] <= 0 and stock[s]["transit"] <= 0]

        not_in_stock = [s for s in cur if s not in stock]
        zero_sales = [s for s, v in cur.items() if v["stock"] <= 0 and v["avail"] <= 0]

        stats[site] = dict(
            sales_sku=len(cur), stock_sku=len(stock), cand=len(cand),
            not_in_stock=len(not_in_stock), zero_sales=len(zero_sales),
            in_stock_state=collections.Counter(stock[s]["state"] for s in stock),
        )

        for sku in cand:
            k = stock[sku]
            nm, md, dstate = names.get(sku, ("", "", ""))
            p = prod.get(sku, {})
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
                q60 = 0.0
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

            row = dict(
                site=site, sku=sku, model=p.get("model") or md, name=nm or md,
                k_state=k["state"], k_rating=k["rating"], rating=ratings.get(sku, "") or "—",
                dstate=dstate, p_state=p.get("p_state", ""), p_score=p.get("p_score", ""),
                typ=p.get("typ", ""),
                instock=k["instock"], transit=k["transit"], weeks=k["weeks"],
                last_stock_m=last_stock_m, last_sale=last_sale, anchor=anchor,
                w_start=w_start, w_end=w_end, q60=q60, q3=q3, days=data.get(sku, {}),
            )
            (all_rows if anchor else idle_all).append(row)

    # ---------------- 统计
    print("=" * 86)
    print("东南亚三国 · 快照 2026-10-29 · 备货表口径")
    for s in ORDER:
        st = stats[s]
        print(f"  {s}：日销表 {st['sales_sku']}｜备货表 {st['stock_sku']}"
              f"（{dict(st['in_stock_state'])}）｜不在备货表 {st['not_in_stock']}"
              f"｜零库存零在途 {st['zero_sales']} → **在售且零库存零在途 {st['cand']}**")
    print(f"\n  合计 {len(all_rows) + len(idle_all)} 个：有断货起点 {len(all_rows)}"
          f"｜今年从未备货 {len(idle_all)}")
    print(f"  断货前 60 天销量合计 {int(sum(r['q60'] for r in all_rows))} 件")
    print("=" * 86)
    for s in ORDER:
        rs = [r for r in all_rows if r["site"] == s]
        idle = [r for r in idle_all if r["site"] == s]
        print(f"\n【{s}】断货 {len(rs)} 个 / {int(sum(r['q60'] for r in rs))} 件"
              f"｜从未备货 {len(idle)} 个")
        for r in sorted(rs, key=lambda x: -x["q60"]):
            print(f"   {r['sku']:8s} {r['k_rating'] or '-':3s} 起点{r['anchor']} "
                  f"60天{int(r['q60']):4d} 近3月{int(r['q3']):3d}  {r['name'][:26]}")
        if idle:
            print("   从未备货：" + " ".join(f"{r['sku']}({r['k_rating'] or '-'})"
                                            for r in sorted(idle, key=lambda x: x["sku"])))

    if not write:
        print("\n（未落盘，加 --write 生成 xlsx）")
        return all_rows, idle_all, stats
    write_xlsx(all_rows, idle_all, stats)
    print("\n已生成：", OUT)
    return all_rows, idle_all, stats


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


def write_xlsx(main_rows, idle_rows, stats):
    os.makedirs(OUT_DIR, exist_ok=True)
    wb = openpyxl.Workbook()

    ws = wb.active
    ws.title = "汇总"
    ws["A1"] = "东南亚 · 在售但无库存无在途 SKU 及断货前 60 天日销明细"
    ws["A1"].font = T_FONT
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=19)
    ws["A2"] = ("快照 2026-10-29｜日销表三国「今日日期」列均为 10-29；备货表三国 stock 页 In stock "
                "与日销表「在库库存」合计对上（印尼 9375/9374、泰国 8915、菲律宾 7322）"
                "　在售取备货表分站点「状态」列；无在途取备货表 In transit（含已排产计划到货）")
    ws["A2"].font = Font(size=9, color="808080")
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=19)

    head = ["国家", "SKU", "型号", "品名", "品类", "备货表状态", "备货表评级", "日销表状态",
            "当月评级(10月)", "全平台状态", "全平台评级", "In stock", "In transit", "计划到货周",
            "最后有货月", "最后销量日", "最早进入断货", "60天窗口", "断货前60天销量", "近3月销量", "备注"]
    ws.append([])
    ws.append(head)
    _hdr(ws, 4, len(head))

    for r in sorted(main_rows, key=lambda x: (ORDER.index(x["site"]), -x["q60"])):
        note = ""
        if r["w_start"].year < 2026:
            got = (r["w_end"] - dt.date(2026, 1, 1)).days + 1
            note = f"⚠ 窗口跨 2025-12，只含 2026 年内 {got} 天"
        elif r["q60"] == 0:
            note = "60 天窗口内零销量（断货已久）"
        elif (r["k_rating"] or "").startswith("X"):
            note = f"新品（{r['k_rating']}）"
        ws.append([
            r["site"], r["sku"], r["model"], r["name"], r["typ"] or "—",
            r["k_state"], r["k_rating"] or "—", r["dstate"] or "—", r["rating"],
            r["p_state"] or "—", r["p_score"] or "—",
            int(r["instock"]), int(r["transit"]),
            "/".join(f"{k}:{int(v)}" for k, v in r["weeks"].items()) or "—",
            r["last_stock_m"] or "—",
            r["last_sale"].isoformat() if r["last_sale"] else "—",
            r["anchor"].isoformat(),
            f'{r["w_start"].isoformat()}~{r["w_end"].isoformat()}',
            int(r["q60"]), int(r["q3"]), note,
        ])
    n = ws.max_row
    ws.append(["合计"] + [""] * 17 + [int(sum(r["q60"] for r in main_rows)),
                                      int(sum(r["q3"] for r in main_rows)), ""])
    for c in range(1, len(head) + 1):
        ws.cell(row=n + 1, column=c).font = Font(bold=True)
        ws.cell(row=n + 1, column=c).fill = PatternFill("solid", fgColor="DDEBF7")
    for i, w in enumerate([7, 9, 24, 32, 10, 11, 11, 11, 12, 11, 11, 10, 10, 18,
                           10, 12, 13, 24, 13, 10, 24], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "C5"

    if idle_rows:
        sh = wb.create_sheet("从未备货")
        sh["A1"] = "在售 · 零库存 · 无在途 · 但 2026 年从未备过货、从未卖出"
        sh["A1"].font = T_FONT
        sh["A2"] = ("没有「进入断货」的时点，算不出断货前 60 天，故不计入汇总。"
                    "采购上属于「上架了但一直没排产」，建议单独过一遍。")
        sh["A2"].font = Font(size=9, color="808080")
        sh.append([])
        h2 = ["国家", "SKU", "型号", "品名", "品类", "备货表评级", "日销表状态",
              "全平台评级", "In stock", "In transit", "说明"]
        sh.append(h2)
        _hdr(sh, 4, len(h2))
        for r in sorted(idle_rows, key=lambda x: (ORDER.index(x["site"]), x["sku"])):
            sh.append([r["site"], r["sku"], r["model"], r["name"], r["typ"] or "—",
                       r["k_rating"] or "—", r["dstate"] or "—", r["p_score"] or "—",
                       int(r["instock"]), int(r["transit"]), "2026 年库存与销量均为 0"])
        for i, w in enumerate([7, 9, 26, 38, 10, 12, 12, 12, 10, 10, 28], 1):
            sh.column_dimensions[get_column_letter(i)].width = w
        sh.freeze_panes = "C5"

    # 明细
    groups = collections.defaultdict(list)
    for r in main_rows:
        d = r["w_start"]
        while d <= r["w_end"]:
            groups[(r["site"], d.month)].append(r)
            d = dt.date(d.year + (d.month == 12), 1 if d.month == 12 else d.month + 1, 1)
    for site in ORDER:
        for m in range(1, CUR + 1):
            rs = groups.get((site, m))
            if not rs:
                continue
            sh = wb.create_sheet(f"{site}-2026-{m:02d}")
            sh["A1"] = f"{site} · 2026年{m}月 · 断货前 60 天日销明细"
            sh["A1"].font = T_FONT
            h2 = ["国家", "SKU", "型号", "品名", "评级(备货表)", "断货起点", "60天窗口"]
            h2 += [f"{d}日" for d in range(1, 32)] + ["本月合计", "窗口60天合计"]
            sh.append([])
            sh.append(h2)
            _hdr(sh, 3, len(h2))
            for r in sorted(rs, key=lambda x: -x["q60"]):
                days = [int(day_of(r["days"], m, d)) for d in range(1, 32)]
                mt = int(r["days"].get(m, {}).get("mtotal", sum(days)) or sum(days))
                sh.append([r["site"], r["sku"], r["model"], r["name"], r["k_rating"] or "—",
                           r["anchor"].isoformat(),
                           f'{r["w_start"].isoformat()}~{r["w_end"].isoformat()}']
                          + days + [mt, int(r["q60"])])
            for i, w in enumerate([7, 9, 22, 30, 12, 12, 24], 1):
                sh.column_dimensions[get_column_letter(i)].width = w
            for c in range(8, 40):
                sh.column_dimensions[get_column_letter(c)].width = 4.5
            sh.freeze_panes = "E4"

    # 口径说明
    sh = wb.create_sheet("口径说明")
    L = [
        ("东南亚三国缺货明细 · 快照 2026-10-29", T_FONT),
        ("", None),
        ("一、数据源与快照日（5 个文件）", Font(bold=True, size=11)),
        ("① 日销表 三国（桌面）", None),
        ("   2026年印尼智慧物流销售表1-12月.xlsx｜2026年泰国易仓销售表1-12月.xlsx｜2026年菲律宾仓销售表1-12月.xlsx", None),
        ("   快照日取各表 10 月页「今日日期 / 今日时间」列：印尼 389 行、泰国 393 行、菲律宾 410 行", None),
        ("   **全部为 2026-10-29**", None),
        ("② 备货表（桌面）东南亚商品资料及库存管理.xlsx", None),
        ("   无日期字段，反推为 2026-10-29：三国 stock 页 In stock 合计与日销表「在库库存」合计对上", None),
        ("   印尼 9375 vs 9374（差 1 为取整）｜泰国 8915 = 8915｜菲律宾 7322 = 7322", None),
        ("③ 商品资料 资产文件/商品资料.xlsx —— 只做对照列（品类 / 全平台状态 / 全平台评级），不参与筛选", None),
        ("", None),
        ("二、三国口径差异（备货表 stock 页列位置不同）", Font(bold=True, size=11)),
        ("   印尼/泰国：状态 col2、In stock col6、In transit col8", None),
        ("   菲律宾　：状态 col3（多一列「系统SKU」）、In stock col8、In transit col10", None),
        ("   → 代码一律按表头名定位，不写死列号。", None),
        ("", None),
        ("三、入选口径", Font(bold=True, size=11)),
        ("1. 在售：备货表对应国家 stock 页「状态」= 在售（分站点口径）", None),
        ("2. 无库存且无在途：In stock ≤ 0 且 In transit ≤ 0", None),
        ("   ⚠️ In transit 含已排产的计划到货（周列 0923/1007/1019/1021/1104/1118），", None),
        ("     所以「无在途」= 连排产都没有，是真正需要下采购单的。", None),
        ("3. 断货起点 = max（最后有货月月末+1，最后销量日+1）；60 天窗口 = [起点-60, 起点-1]", None),
        ("4. 日销数量取各月主表 1..31 日列（不是订单明细表）——批量入账只记在月表日列", None),
        ("5. 2026 年库存与销量均为 0 的单列「从未备货」页，不计入汇总", None),
        ("", None),
        ("四、为什么用备货表当在售判据", Font(bold=True, size=11)),
        ("日销表「印尼状态」有 4 个值（在售/机会/需求/售完截止），与备货表交叉后是两套分类：", None),
        ("   机会 67 个、需求 39 个 → 备货表里全部是「在售」", None),
        ("   售完截止 7 个 → 备货表「售完即止」", None),
        ("即「机会 / 需求」并不是「不在售」，只是备货策略不同；只认「在售」两字会漏掉 106 个。", None),
        ("泰国、菲律宾日销表根本没有状态列，只能用备货表。", None),
        ("", None),
        ("五、三国结果", Font(bold=True, size=11)),
    ]
    for s in ORDER:
        st = stats[s]
        L.append((f"   {s}：日销表 {st['sales_sku']}｜备货表 {st['stock_sku']}"
                  f"｜不在备货表 {st['not_in_stock']}｜在售且零库存零在途 {st['cand']}", None))
    L += [
        ("   ⚠️「不在备货表」的 SKU 不纳入：备货表不管理 = 霞姐不负责采购，多为老款清尾。", None),
        ("", None),
        ("六、已知限制", Font(bold=True, size=11)),
        ("· 备货表评级列与日销表 10 月评级列并列，两者时点不同（约 14% 有差异），未做取舍。", None),
        ("· 商品资料「状态/评级」是全公司/全平台口径、不分站点，仅作对照。", None),
        ("· 日销表只有 2026 年数据，断货早于 1 月的 SKU 窗口会跨到 2025-12，该段取不到。", None),
    ]
    for i, (t, f) in enumerate(L, 1):
        c = sh.cell(row=i, column=1, value=t)
        if f:
            c.font = f
    sh.column_dimensions["A"].width = 122

    wb.save(OUT)


if __name__ == "__main__":
    main(write="--write" in sys.argv)
