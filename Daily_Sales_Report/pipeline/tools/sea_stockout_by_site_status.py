# -*- coding: utf-8 -*-
"""
东南亚断货判定 · 分站点商品资料口径（2026-10-30）

为什么有这个脚本
----------------
旧版 sea_stockout_detail.py 的「在售」判据是印尼日销表 J 列（在售/机会/需求/售完截止），
泰菲则用「近 60 天有销量」代理 —— 这套判据跟霞姐维护的商品资料（每个国家一张 stock 表）
对不上，同一批 SKU 两边结论会打架。本版改为：

    在售  = 《东南亚商品资料及库存管理.xlsx》对应国家 stock 页的「状态」列 == 在售
    库存  = 同页 In stock（现货）/ In transit（含已排产计划到货）
    销量  = 三国日销表 1..10 月主表「当月总销售数量」列

三档分类（都是「在售 + 无现货」的子集，按有没有销量/在途拆开）
--------------------------------------------------------------
  A 真断货     无现货 + 无在途 + 今年有销量   → 客户想买买不到，且在补货管道里也没有
  B 缺货在补   无现货 + 有在途 + 今年有销量   → 已排产/已下单，等货到
  C 从未备货   无现货 + 无在途 + 今年零销量   → 在售但今年一件没进，不属于「断货」事件
  D 无货有在途零销量  无现货 + 有在途 + 零销量 → 备了新货还没卖过，观察用

断货起点 = max(最后有货月月末 + 1, 最后销量日 + 1)
60 天窗口 = [起点 - 60, 起点 - 1]

用法
----
  python pipeline/tools/sea_stockout_by_site_status.py             # 全跑 + 出 xlsx
  python pipeline/tools/sea_stockout_by_site_status.py --dry-run   # 只看控制台
"""
from __future__ import annotations

import argparse
import datetime as dt
import os
import sys
from collections import defaultdict

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))


def _find_desktop(start):
    p = os.path.abspath(start)
    while True:
        if os.path.basename(p).lower() == "desktop":
            return p
        parent = os.path.dirname(p)
        if parent == p:
            raise RuntimeError("找不到 Desktop 目录")
        p = parent


DESKTOP = _find_desktop(ROOT)
PRODUCT_FILE = os.path.join(DESKTOP, "东南亚商品资料及库存管理.xlsx")
SALES_FILES = {
    "印尼": os.path.join(DESKTOP, "2026年印尼智慧物流销售表1-12月.xlsx"),
    "泰国": os.path.join(DESKTOP, "2026年泰国易仓销售表1-12月.xlsx"),
    "菲律宾": os.path.join(DESKTOP, "2026年菲律宾仓销售表1-12月.xlsx"),
}
MAIN_SHEET = {
    "印尼": lambda m: f"2026年{m}月",
    "泰国": lambda m: f"{m}月",
    "菲律宾": lambda m: f"{m}月",
}
# 主表列位（0 基）：在库 / 剩余+在途 / 当月总销量
COL = {
    "印尼": {"stock": 11, "stock_transit": 12, "sales": 13},
    "泰国": {"stock": 10, "stock_transit": 11, "sales": 12},
    "菲律宾": {"stock": 10, "stock_transit": 11, "sales": 12},
}
# 商品资料分国家 stock 页列位（0 基）
SITE_SHEET = {
    "印尼": ("Indonesia stock", {"sku": 0, "rating": 1, "status": 2, "remark": 3,
                              "stock": 6, "transit": 8}),
    "泰国": ("Thailand stock", {"sku": 0, "rating": 1, "status": 2, "remark": 3,
                             "stock": 6, "transit": 8}),
    "菲律宾": ("Philippines stock", {"sku": 0, "rating": 1, "status": 3, "remark": 4,
                                 "stock": 8, "transit": 10}),
}
SITES = ("印尼", "泰国", "菲律宾")
TODAY = dt.date(2026, 10, 30)
V3_FILE = os.path.join(ROOT, "断货分析", "东南亚", "明细", "20261021",
                       "东南亚缺货60天明细_20261021_V3_商品资料口径.xlsx")


def _num(v):
    if v is None:
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(",", "")
    if s in ("", "-", ".", "None"):
        return 0.0
    try:
        return float(s)
    except Exception:
        return 0.0


def _txt(v):
    return "" if v is None else str(v).strip()


def month_end(m):
    return dt.date(2026, m + 1, 1) - dt.timedelta(days=1) if m < 12 else dt.date(2026, 12, 31)


# ---------------- 1) 商品资料 ----------------
def load_product():
    wb = openpyxl.load_workbook(PRODUCT_FILE, read_only=True, data_only=True)
    out = {}
    try:
        for site, (sheet, c) in SITE_SHEET.items():
            ws = wb[sheet]
            d = {}
            for r in ws.iter_rows(min_row=3, max_col=ws.max_column, values_only=True):
                sku = _txt(r[c["sku"]]) if len(r) > c["sku"] else ""
                if not sku or sku == ".":
                    continue
                get = lambda k: (r[c[k]] if len(r) > c[k] else None)
                d[sku] = {
                    "状态": _txt(get("status")),
                    "评级": _txt(get("rating")),
                    "备注": _txt(get("remark")),
                    "在库": _num(get("stock")),
                    "在途": _num(get("transit")),
                }
            out[site] = d
    finally:
        wb.close()
    return out


# ---------------- 2) 日销表 ----------------
def _find_day0(hdr):
    for i in range(len(hdr) - 3):
        if all(hdr[i + k] is not None and str(hdr[i + k]).strip() == str(k + 1)
               for k in range(3)):
            return i
    return None


def load_sales(site, months=range(1, 11)):
    wb = openpyxl.load_workbook(SALES_FILES[site], read_only=True, data_only=True)
    res = {}
    try:
        for m in months:
            sheet = MAIN_SHEET[site](m)
            if sheet not in wb.sheetnames:
                continue
            ws = wb[sheet]
            c = COL[site]
            day0 = None
            for ri, r in enumerate(ws.iter_rows(min_row=1, max_row=ws.max_row,
                                                max_col=ws.max_column, values_only=True), 1):
                if ri == 1:
                    day0 = _find_day0(r)
                    continue
                if ri == 2:
                    continue
                sku = _txt(r[0]) if r else ""
                if not sku or sku == ".":
                    continue
                d = res.setdefault(sku, {"月销量": {}, "月库存": {}, "月库存含在途": {},
                                         "日销": defaultdict(float)})
                d["月销量"][m] = _num(r[c["sales"]]) if len(r) > c["sales"] else 0.0
                d["月库存"][m] = _num(r[c["stock"]]) if len(r) > c["stock"] else 0.0
                d["月库存含在途"][m] = _num(r[c["stock_transit"]]) if len(r) > c["stock_transit"] else 0.0
                if day0 is not None:
                    for k in range(31):
                        j = day0 + k
                        if j >= len(r):
                            break
                        q = _num(r[j])
                        if q:
                            d["日销"][dt.date(2026, m, k + 1)] += q
    finally:
        wb.close()
    return res


# ---------------- 3) 分类 ----------------
CLASSES = {
    "A": "A 真断货 · 无现货无在途但有销量",
    "B": "B 缺货在补 · 无现货有在途且有销量",
    "C": "C 从未备货 · 在售但今年零进货零销量",
    "D": "D 无货有在途零销量 · 新备货未开卖",
}


def analyze():
    prod = load_product()
    sales = {s: load_sales(s) for s in SITES}
    buckets = {k: [] for k in CLASSES}
    for site in SITES:
        for sku, info in prod[site].items():
            if info["状态"] != "在售":
                continue
            if info["在库"] > 0:
                continue
            smap = sales[site].get(sku, {"月销量": {}, "月库存": {}, "月库存含在途": {}, "日销": {}})
            tot = sum(smap["月销量"].values())
            days = sorted(d for d, q in smap["日销"].items() if q > 0)
            last_sale = days[-1] if days else None
            with_stock = [m for m, v in smap["月库存含在途"].items() if v > 0]
            last_stock_m = max(with_stock) if with_stock else None
            has_transit = info["在途"] > 0
            if tot > 0:
                key = "B" if has_transit else "A"
            else:
                key = "D" if has_transit else "C"
            cand = []
            if last_stock_m:
                cand.append(month_end(last_stock_m) + dt.timedelta(days=1))
            if last_sale:
                cand.append(last_sale + dt.timedelta(days=1))
            start = max(cand) if cand else None
            row = {
                "国家": site, "SKU": sku, "分类": key,
                "状态": info["状态"], "评级": info["评级"], "备注": info["备注"],
                "在库": info["在库"], "在途": info["在途"],
                "今年销量": tot, "最后有货月": last_stock_m, "最后销量日": last_sale,
                "断货起点": start,
                "已空天数": (TODAY - start).days + 1 if start else None,
                "窗口起": start - dt.timedelta(days=60) if start else None,
                "窗口止": start - dt.timedelta(days=1) if start else None,
                "窗口销量": (sum(q for d, q in smap["日销"].items()
                             if start - dt.timedelta(days=60) <= d <= start - dt.timedelta(days=1))
                         if start else 0),
            }
            buckets[key].append(row)
    order = {s: i for i, s in enumerate(SITES)}
    for k in buckets:
        buckets[k].sort(key=lambda r: (order[r["国家"]],
                                       -(r["已空天数"] or 0)))
    return buckets, prod, sales


def load_v3():
    """V3（20261021）那份的 SKU → 国家。"""
    if not os.path.exists(V3_FILE):
        return {}
    wb = openpyxl.load_workbook(V3_FILE, read_only=True, data_only=True)
    d = {}
    try:
        ws = wb["汇总"]
        for r in ws.iter_rows(min_row=5, max_col=14, values_only=True):
            site, sku = _txt(r[0]), _txt(r[1])
            if site in SITES and sku:
                d[(site, sku)] = {"状态": _txt(r[4]), "评级": _txt(r[5])}
    finally:
        wb.close()
    return d


# ---------------- 4) 输出 ----------------
HDR_FILL = PatternFill("solid", fgColor="DDEBF7")
A_FILL = PatternFill("solid", fgColor="FFC7CE")
TITLE_FONT = Font(bold=True, size=13)

COLS = ["国家", "SKU", "分类", "评级", "备注", "在库", "在途", "今年销量",
        "最后有货月", "最后销量日", "断货起点", "已空天数", "窗口起", "窗口止", "窗口销量"]


def _write_sheet(ws, title, rows, note=None):
    ws["A1"] = title
    ws["A1"].font = TITLE_FONT
    r0 = 3
    if note:
        ws["A2"] = note
        ws["A2"].alignment = Alignment(wrap_text=False)
    for j, name in enumerate(COLS, 1):
        c = ws.cell(row=r0, column=j, value=name)
        c.font = Font(bold=True)
        c.fill = HDR_FILL
        c.alignment = Alignment(horizontal="center")
    for i, row in enumerate(rows, r0 + 1):
        for j, name in enumerate(COLS, 1):
            v = row.get(name)
            if isinstance(v, dt.date):
                v = v.strftime("%Y-%m-%d")
            ws.cell(row=i, column=j, value=v)
        if row["分类"] == "A":
            for j in range(1, len(COLS) + 1):
                ws.cell(row=i, column=j).fill = A_FILL
    widths = [8, 10, 6, 6, 14, 7, 7, 10, 11, 12, 12, 9, 12, 12, 9]
    for j, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(j)].width = w
    ws.freeze_panes = ws.cell(row=r0 + 1, column=1)


def write_xlsx(buckets, outdir):
    os.makedirs(outdir, exist_ok=True)
    path = os.path.join(outdir, f"东南亚断货明细_{TODAY:%Y%m%d}_分站点商品资料口径.xlsx")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "汇总"

    # 汇总页
    ws["A1"] = "东南亚断货判定 · 分站点商品资料口径"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = f"快照 2026-10-29（商品资料 stock 页 + 三国日销表 1-10 月）｜判定日 {TODAY:%Y-%m-%d}"
    ws["A3"] = "在售 = 商品资料对应国家 stock 页「状态」列；库存 = 同页 In stock / In transit（含已排产）"
    head = ["国家", "在售 SKU", "A 真断货", "B 缺货在补", "C 从未备货", "D 新备货未开卖", "无现货合计"]
    for j, h in enumerate(head, 1):
        c = ws.cell(row=5, column=j, value=h)
        c.font = Font(bold=True)
        c.fill = HDR_FILL
        c.alignment = Alignment(horizontal="center")
    for i, site in enumerate(SITES, 6):
        a = len([r for r in buckets["A"] if r["国家"] == site])
        b = len([r for r in buckets["B"] if r["国家"] == site])
        c_ = len([r for r in buckets["C"] if r["国家"] == site])
        d_ = len([r for r in buckets["D"] if r["国家"] == site])
        ws.cell(row=i, column=1, value=site)
        ws.cell(row=i, column=2, value=sum(1 for v in _PROD_CACHE[site].values() if v["状态"] == "在售"))
        for j, v in enumerate([a, b, c_, d_, a + b + c_ + d_], 3):
            ws.cell(row=i, column=j, value=v)
    tot_row = 6 + len(SITES)
    ws.cell(row=tot_row, column=1, value="合计").font = Font(bold=True)
    for j in range(2, 8):
        s = sum((ws.cell(row=i, column=j).value or 0) for i in range(6, tot_row))
        ws.cell(row=tot_row, column=j, value=s).font = Font(bold=True)
    ws.column_dimensions["A"].width = 10
    for col in "BCDEFG":
        ws.column_dimensions[col].width = 15
    ws["A" + str(tot_row + 2)] = "A 档为红底：客户想买买不到、且补货管道里也没有货。"
    ws["A" + str(tot_row + 3)] = "断货起点 = max(最后有货月月末+1, 最后销量日+1)；60 天窗口 = [起点-60, 起点-1]。"

    for k, tab in (("A", "A-真断货"), ("B", "B-缺货在补"),
                   ("C", "C-从未备货"), ("D", "D-新备货未开卖")):
        wsx = wb.create_sheet(tab)
        _write_sheet(wsx, f"{CLASSES[k]}（{len(buckets[k])} 个）", buckets[k])

    # 口径说明
    wsp = wb.create_sheet("口径说明")
    lines = [
        "一、数据源",
        "  商品资料 = 桌面《东南亚商品资料及库存管理.xlsx》（霞姐维护）",
        "    · Indonesia stock   A=SKU  B=评级  C=状态  D=Remarks  G=In stock  I=In transit",
        "    · Thailand stock    同上",
        "    · Philippines stock A=SKU  B=评级  D=状态  E=Remarks  I=In stock  K=In transit",
        "  日销表 = 桌面三国销售表 1-12 月，取 1..10 月主表（不含渠道页/订单页）",
        "",
        "二、判定口径",
        "  在售 = 状态列 == '在售'（'售完即止'、'停售' 均不计入）",
        "  无现货 = In stock <= 0",
        "  有在途 = In transit > 0（**含已排产、尚未发运的计划到货**，对应 stock 页的 0923/1009/1021/1104/1118/1202 到货周列）",
        "  有销量 = 1-10 月「当月总销售数量」合计 > 0",
        "",
        "三、为什么要拆成四档",
        "  A 真断货：无现货 + 无在途 + 有销量 → 真正断掉，需要立刻决策补不补",
        "  B 缺货在补：无现货 + 有在途 + 有销量 → 货已在计划里，属于「等货」不是「断货」",
        "  C 从未备货：在售但今年一件没进 → 是选品/上架问题，不是断货事件",
        "  D 新备货未开卖：备了货还没卖 → 观察",
        "",
        "四、与旧版（V3 / sea_stockout_detail.py）的差别",
        "  旧版在售判据用印尼日销表 J 列（在售/机会/需求/售完截止），泰菲用「近60天有销量」代理；",
        "  旧版库存用日销表「剩余+在途」，只含已发运在途，不含计划到货。",
        "  结果：一批「已排产未发运」的 SKU 被旧版误判成断货。",
    ]
    for i, ln in enumerate(lines, 1):
        wsp.cell(row=i, column=1, value=ln)
    wsp.column_dimensions["A"].width = 120

    wb.save(path)
    return path


_PROD_CACHE = {}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=os.path.join(ROOT, "断货分析", "东南亚", "明细",
                                                  TODAY.strftime("%Y%m%d")))
    ap.add_argument("--dry-run", action="store_true")
    args = ap.parse_args()

    buckets, prod, _ = analyze()
    _PROD_CACHE.update(prod)

    print("=" * 78)
    print("东南亚断货判定 · 分站点商品资料口径")
    print("=" * 78)
    for site in SITES:
        a = [r for r in buckets["A"] if r["国家"] == site]
        b = [r for r in buckets["B"] if r["国家"] == site]
        c = [r for r in buckets["C"] if r["国家"] == site]
        d = [r for r in buckets["D"] if r["国家"] == site]
        on = sum(1 for v in prod[site].values() if v["状态"] == "在售")
        print(f"\n【{site}】商品资料 {len(prod[site])} 个，在售 {on} 个")
        print(f"  A 真断货 {len(a)}｜B 缺货在补 {len(b)}｜C 从未备货 {len(c)}｜D 新备货未开卖 {len(d)}")
        for r in a:
            print(f"    [A] {r['SKU']:8s} {r['评级']:3s} 今年销量{r['今年销量']:>5.0f} "
                  f"断货起点 {r['断货起点']} 已空 {r['已空天数']:>3d} 天 60天窗口销量 {r['窗口销量']:.0f}")

    print("\n" + "=" * 78)
    print("合计：A 真断货 %d｜B 缺货在补 %d｜C 从未备货 %d｜D 新备货未开卖 %d"
          % (len(buckets["A"]), len(buckets["B"]), len(buckets["C"]), len(buckets["D"])))
    if not args.dry_run:
        p = write_xlsx(buckets, args.out)
        print("已输出:", p)
    return buckets


if __name__ == "__main__":
    if sys.platform.startswith("win"):
        sys.stdout.reconfigure(encoding="utf-8")
    main()
