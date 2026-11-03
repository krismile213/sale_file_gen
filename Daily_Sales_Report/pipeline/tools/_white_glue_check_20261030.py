# -*- coding: utf-8 -*-
"""墨西哥两站点「白边液产品」核查 —— 2026-10-30

背景：物流合规原因，墨西哥（亚马逊、美客多）无法发白边液，已持续 2 年以上。
     但美客多机制特殊，可能出现"僵尸 listing"导致客户下单 → 触发备货 → 白边液被采购和发货。

口径：
  白边液产品 = 商品资料「钢化膜组装表」中「白边液」列 > 0 的钢化膜 SKU（52 个）
  近三个月   = 2026-08-01 ~ 2026-10-30
     · 亚马逊：用「墨西哥-每日库存分析表」的 8/9/10 月列（col78/77/76）
                + 「每日销售数据」订单（9/24~10/30）+「9月站点汇总」交叉验证
     · 美客多：用「墨西哥-每日库存分析表」的 8/9/10 月列（col68/67/66）
                + 「订单数据」逐单明细（8/31~10/27）
  Listing 情况：
     · 亚马逊：是否在「墨西哥-每日库存分析表」里（不在 = 无 listing 管理）
     · 美客多：订单里的 MLM 发布编号是否在「库存明细」的发布编号列里
"""
import os
import re
import collections
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from urllib.parse import quote

DSR = r"C:\Users\dell\Desktop\work-main\sale_file_gen\Daily_Sales_Report"
PROD = r"C:\Users\dell\Desktop\work-main\资产文件\商品资料.xlsx"
WB_AMZ = os.path.join(DSR, "1.每日销售数据表", "北美加拿大库存销售表.xlsx")
WB_ML = os.path.join(DSR, "1.每日销售数据表", "美客多库存销售表.xlsx")
SH_AMZ = "墨西哥-每日库存分析表 "   # ⚠ 末尾有空格
SH_ML = "墨西哥-每日库存分析表"

OUT_DIR = os.path.join(DSR, "白边液核查", "20261030")
OUT = os.path.join(OUT_DIR, "墨西哥白边液产品核查_20261030.xlsx")


def txt(v):
    return str(v).strip() if v is not None else ""


def num(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return 0.0


def norm(s):
    return re.sub(r"[^A-Za-z0-9]", "", s).upper()


ES_MONTH = {"enero": 1, "febrero": 2, "marzo": 3, "abril": 4, "mayo": 5, "junio": 6,
            "julio": 7, "agosto": 8, "septiembre": 9, "setiembre": 9, "octubre": 10,
            "noviembre": 11, "diciembre": 12}


def parse_es_date(s):
    """'31 de agosto de 2026 23:37 hs.' → (2026, 8, 31)"""
    m = re.search(r"(\d{1,2})\s+de\s+([a-zA-Záéíóú]+)\s+de\s+(\d{4})", s or "")
    if not m:
        return None
    mo = ES_MONTH.get(m.group(2).lower())
    if not mo:
        return None
    return (int(m.group(3)), mo, int(m.group(1)))


# ---------------------------------------------------------------- 1. 白边液 SKU
def load_white():
    wb = openpyxl.load_workbook(PROD, read_only=True, data_only=True)
    ws = wb["钢化膜组装表"]
    rows = list(ws.iter_rows(min_row=1, values_only=True))
    white = {}
    for r in rows[2:]:
        sku = txt(r[0])
        if not sku:
            continue
        q = txt(r[11]) if len(r) > 11 else ""
        if q in ("1", "2"):
            white[sku] = dict(qty=int(float(q)), asm_name=txt(r[1]))

    # 组装表 SKU（可能带 ZH 后缀）→ 商品资料主 SKU
    ws = wb["商品资料"]
    o2a, info = {}, {}
    for r in ws.iter_rows(min_row=3, values_only=True):
        a = txt(r[0])
        if not a:
            continue
        info[a] = dict(state=txt(r[12]), score=txt(r[5]), typ=txt(r[6]),
                       name=txt(r[7]), model=txt(r[11]), asin=txt(r[1]))
        o = txt(r[14])
        if o:
            o2a[o] = a
    wb.close()

    out = {}
    for sku, v in white.items():
        a = o2a.get(sku, sku)
        if a not in info:          # 极端情况：组装表 SKU 在商品资料里查不到
            a = sku
        out[a] = dict(qty=v["qty"], asm_sku=sku, asm_name=v["asm_name"],
                      **info.get(a, dict(state="", score="", typ="", name="", model="", asin="")))
    return out


# ---------------------------------------------------------------- 2. 亚马逊墨西哥
def load_amz(wn):
    wb = openpyxl.load_workbook(WB_AMZ, read_only=True, data_only=True)
    ws = wb[SH_AMZ]
    rows = list(ws.iter_rows(min_row=7, values_only=True))
    table = {}
    for r in rows:
        sku = txt(r[0])
        if not sku or sku == "合计":
            continue
        if norm(sku) in wn:
            table[sku] = dict(state=txt(r[3]), stock=num(r[5]), transit=num(r[6]),
                              m8=num(r[78]), m9=num(r[77]), m10=num(r[76]),
                              asin=txt(r[95]))
    total_sku = len([r for r in rows if txt(r[0]) and txt(r[0]) != "合计"])

    # 订单（9/24~10/30）按渠道统计白边液
    ws = wb["每日销售数据"]
    orders = []
    for r in ws.iter_rows(min_row=3, values_only=True):
        if not r[0]:
            continue
        if norm(txt(r[11])) in wn:
            orders.append(dict(date=txt(r[2])[:10], channel=txt(r[6]), sku=txt(r[11]),
                               qty=num(r[14]), status=txt(r[4]), oid=txt(r[0])))
    wb.close()
    return table, total_sku, orders


# ---------------------------------------------------------------- 3. 美客多墨西哥
def load_ml(wn):
    wb = openpyxl.load_workbook(WB_ML, read_only=True, data_only=True)
    ws = wb[SH_ML]
    rows = list(ws.iter_rows(min_row=7, values_only=True))
    table = {}
    for r in rows:
        sku = txt(r[0])
        if not sku or sku == "合计":
            continue
        if norm(sku) in wn:
            table[sku] = dict(state=txt(r[3]), stock=num(r[5]), transit=num(r[6]),
                              m8=num(r[68]), m9=num(r[67]), m10=num(r[66]),
                              offline=txt(r[98]))
    total_sku = len([r for r in rows if txt(r[0]) and txt(r[0]) != "合计"])

    # 订单明细
    ws = wb["订单数据"]
    orders = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        if not r[0]:
            continue
        if norm(txt(r[18])) in wn:
            d = parse_es_date(txt(r[1]))
            orders.append(dict(oid=txt(r[0]), raw_date=txt(r[1]), ymd=d,
                               estado=txt(r[2]), qty=num(r[6]), total=num(r[15]),
                               sku=txt(r[18]), mlm=txt(r[19]), title=txt(r[20])))

    # 库存明细：发布编号（col3，无 MLM 前缀）→ 状态
    ws = wb["库存明细"]
    pub = {}
    for r in ws.iter_rows(min_row=5, values_only=True):
        p = txt(r[3])
        if p:
            pub[p] = dict(sku=txt(r[2]), estado=txt(r[9]))
    wb.close()
    return table, total_sku, orders, pub


def main(write=False):
    white = load_white()
    wn = {norm(s) for s in white}

    amz_tab, amz_total, amz_orders = load_amz(wn)
    ml_tab, ml_total, ml_orders, ml_pub = load_ml(wn)

    print("=" * 90)
    print(f"白边液产品（钢化膜组装表 白边液>0）：{len(white)} 个 SKU")
    print(f"亚马逊墨西哥：日销表 {amz_total} 个 SKU，白边液命中 {len(amz_tab)}")
    print(f"美客多墨西哥：日销表 {ml_total} 个 SKU，白边液命中 {len(ml_tab)}")
    print("=" * 90)

    # 亚马逊订单按渠道
    ch = collections.Counter()
    for o in amz_orders:
        ch[(o["channel"], o["status"])] += o["qty"]
    print("\n【亚马逊】9/24~10/30 订单里白边液：")
    for k, v in sorted(ch.items()):
        print(f"   {k[0]:14s} {k[1]:10s} {int(v)} 件")
    mx = [o for o in amz_orders if "mx" in o["channel"].lower()]
    print(f"   → Amazon.com.mx 命中 {len(mx)} 条")

    # 美客多订单
    print(f"\n【美客多】订单明细 {len(ml_orders)} 条 / {int(sum(o['qty'] for o in ml_orders))} 件")
    by_m = collections.Counter()
    for o in ml_orders:
        if o["ymd"]:
            by_m[o["ymd"][1]] += o["qty"]
    print(f"   按月：{ {f'{k}月': int(v) for k, v in sorted(by_m.items())} }")

    # listing 核对
    rows = []
    for sku, info in sorted(white.items()):
        a = amz_tab.get(sku)
        m = ml_tab.get(sku)
        mo = [o for o in ml_orders if norm(o["sku"]) == norm(sku)]
        mlm = sorted({o["mlm"] for o in mo})
        in_pub = [bool(ml_pub.get(x.replace("MLM", ""))) for x in mlm]
        rows.append(dict(
            sku=sku, model=info["model"], name=info["name"], typ=info["typ"],
            asin=info["asin"], p_state=info["state"], p_score=info["score"], qty=info["qty"],
            amz_in=bool(a), amz_state=a["state"] if a else "—",
            amz_m8=a["m8"] if a else 0, amz_m9=a["m9"] if a else 0, amz_m10=a["m10"] if a else 0,
            ml_in=bool(m), ml_state=m["state"] if m else "—",
            ml_m8=m["m8"] if m else 0, ml_m9=m["m9"] if m else 0, ml_m10=m["m10"] if m else 0,
            ml_orders=len(mo), ml_qty=sum(o["qty"] for o in mo),
            ml_amt=sum(o["total"] for o in mo),
            mlm=", ".join(mlm), mlm_in_pub=all(in_pub) if mlm else None,
        ))

    hit = [r for r in rows if r["ml_orders"] > 0]
    print(f"\n【结论】美客多有订单的白边液 SKU：{len(hit)} 个，共 {int(sum(r['ml_qty'] for r in hit))} 件")
    for r in sorted(hit, key=lambda x: -x["ml_qty"]):
        print(f"   {r['sku']:8s} {r['ml_orders']} 单 {int(r['ml_qty'])} 件 "
              f"金额{int(r['ml_amt'])} 商品资料={r['p_state']}/{r['p_score'] or '-'} "
              f"在库存明细={'是' if r['mlm_in_pub'] else '否'}")
    print(f"【结论】亚马逊墨西哥白边液：销量 0（日销表无 SKU、9 月汇总无、9/24 起订单无）")

    if not write:
        print("\n（未落盘，加 --write 生成 xlsx）")
        return rows, ml_orders, amz_orders

    write_xlsx(rows, ml_orders, amz_orders, amz_total, ml_total)
    print("\n已生成：", OUT)
    return rows, ml_orders, amz_orders


# ---------------------------------------------------------------- 落盘
H_FILL = PatternFill("solid", fgColor="1F4E79")
H_FONT = Font(bold=True, color="FFFFFF", size=10)
T_FONT = Font(bold=True, size=13)
RED_FILL = PatternFill("solid", fgColor="FCE4E4")
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def _hdr(ws, row, n):
    for c in range(1, n + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = H_FILL
        cell.font = H_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BOX


COV = [
    ("亚马逊", "美国",   WB_AMZ, "美国-每日库存分析表", 0, 7),
    ("亚马逊", "加拿大", WB_AMZ, "加拿大-每日库存分析表", 0, 7),
    ("亚马逊", "墨西哥", WB_AMZ, SH_AMZ, 0, 7),
    ("亚马逊", "墨西哥", WB_AMZ, "墨西哥可售+在途库存", None, None),
    ("亚马逊", "墨西哥", WB_AMZ, "墨西哥库存记录", None, None),
    ("美客多", "墨西哥", WB_ML, SH_ML, 0, 7),
    ("美客多", "墨西哥", WB_ML, "库存明细", 2, 5),
]


def coverage(rows):
    """白边液 SKU 在各站点管理表里的覆盖情况。

    美国/加拿大表里大量存在、墨西哥两站点 0 —— 这是"墨西哥根本不管理白边液"的硬证据。
    """
    W = {r["sku"] for r in rows}
    out = []
    cache = {}
    for plat, site, path, sheet, col, first in COV:
        wb = cache.get(path)
        if wb is None:
            wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
            cache[path] = wb
        if sheet not in wb.sheetnames:
            out.append((plat, site, sheet, "—", 0, "表不存在"))
            continue
        ws = wb[sheet]
        n, hit = 0, set()
        for i, row in enumerate(ws.iter_rows(values_only=True), 1):
            if not row:
                continue
            if col is not None and first and i >= first and len(row) > col:
                v = row[col]
                if v is not None and str(v).strip():
                    n += 1
            for v in row:
                if v is None:
                    continue
                s = str(v).strip()
                if s in W:
                    hit.add(s)
                    break
        if col is None:
            n = None
        out.append((plat, site, sheet, n if n is not None else "—", len(hit), ""))
    for wb in cache.values():
        wb.close()
    return out


def write_xlsx(rows, ml_orders, amz_orders, amz_total, ml_total):
    os.makedirs(OUT_DIR, exist_ok=True)
    wb = openpyxl.Workbook()

    # ---------- 汇总
    ws = wb.active
    ws.title = "汇总"
    ws["A1"] = "墨西哥两站点 · 白边液产品核查（近三个月：2026-08-01 ~ 2026-10-30）"
    ws["A1"].font = T_FONT
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=16)
    ws["A2"] = ("白边液产品 = 商品资料「钢化膜组装表」中「白边液」列 > 0 的钢化膜 SKU（共 52 个）"
                "｜数据快照：日销表 2026-10-30；美客多订单 2026-08-31 ~ 2026-10-27；亚马逊订单 2026-09-24 ~ 2026-10-30")
    ws["A2"].font = Font(size=9, color="808080")
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=16)

    head = ["SKU", "型号", "品名", "白边液用量", "商品资料状态", "商品资料评级",
            "亚马逊\n在日销表", "亚马逊\n8月", "亚马逊\n9月", "亚马逊\n10月",
            "美客多\n在日销表", "美客多\n8月", "美客多\n9月", "美客多\n10月",
            "美客多\n近三月订单", "美客多\n近三月件数", "美客多 MLM", "MLM\n在库存明细"]
    ws.append([])
    ws.append(head)
    _hdr(ws, 4, len(head))

    for r in sorted(rows, key=lambda x: (-x["ml_qty"], x["sku"])):
        ws.append([
            r["sku"], r["model"], r["name"], r["qty"], r["p_state"], r["p_score"] or "—",
            "是" if r["amz_in"] else "否", int(r["amz_m8"]), int(r["amz_m9"]), int(r["amz_m10"]),
            "是" if r["ml_in"] else "否", int(r["ml_m8"]), int(r["ml_m9"]), int(r["ml_m10"]),
            r["ml_orders"], int(r["ml_qty"]), r["mlm"] or "—",
            ("—" if r["mlm_in_pub"] is None else ("是" if r["mlm_in_pub"] else "否")),
        ])
        if r["ml_orders"] > 0:
            for c in range(1, len(head) + 1):
                ws.cell(row=ws.max_row, column=c).fill = RED_FILL

    n = ws.max_row
    ws.append(["合计"] + [""] * 13 + [sum(r["ml_orders"] for r in rows),
                                      int(sum(r["ml_qty"] for r in rows)), "", ""])
    for c in range(1, len(head) + 1):
        ws.cell(row=n + 1, column=c).font = Font(bold=True)
        ws.cell(row=n + 1, column=c).fill = PatternFill("solid", fgColor="DDEBF7")
    for i, w in enumerate([9, 24, 40, 11, 12, 12, 10, 7, 7, 7, 10, 7, 7, 7, 12, 12, 18, 11], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "B5"

    # ---------- Listing 排查清单（用户逐个点开排查）
    sh = wb.create_sheet("Listing 排查清单", 1)
    sh["A1"] = "墨西哥白边液 · Listing 排查清单（52 个 SKU，按优先级排序）"
    sh["A1"].font = T_FONT
    sh.merge_cells("A1:M1")
    notes = [
        ("亚马逊：链接用 amazon.com.mx/dp/ASIN。已核实同一 SKU 在美/加/墨三站 ASIN 完全相同（306/306），"
         "可直接用商品资料的 ASIN 打开墨西哥站。判读：404/无法购买 = 墨西哥未上架（合规，无需处理）；"
         "能打开且可加购 = 有 listing，需处理。", "C00000"),
        ("美客多：链接用 articulo.mercadolibre.com.mx/MLM-编号；若 404 改试 mercadolibre.com.mx/p/MLM编号。"
         "没有 MLM 的 SKU 说明近三月无单，需登录卖家后台按 SKU 搜发布：mercadolibre.com.mx/publicaciones/listado", "C00000"),
        ("优先级：P1 = 近三月有真实订单（已被证实存在且能接单，先处理）｜P2 = 商品资料在售但无订单（可能还有未出单的 listing）"
         "｜P3 = 已停售且无 ASIN（大概率已下架，确认即可）。最后一列可下拉填写结论。", "808080"),
    ]
    r0 = 2
    for t, col in notes:
        c = sh.cell(row=r0, column=1, value=t)
        c.font = Font(size=9, color=col)
        sh.merge_cells(start_row=r0, start_column=1, end_row=r0, end_column=13)
        r0 += 1

    h5 = ["优先级", "SKU", "机型", "商品资料\n状态", "评级", "ASIN",
          "亚马逊墨西哥 listing", "美客多 MLM", "美客多 listing",
          "近三月\n美客多单", "在美客多\n库存明细", "排查要点", "排查结论（手填）"]
    sh.append([])
    sh.append(h5)
    _hdr(sh, sh.max_row, len(h5))
    hrow = sh.max_row

    LINK_FONT = Font(color="0563C1", underline="single", size=10)

    def _prio(r):
        if r["ml_orders"] > 0:
            return "P1"
        if r["p_state"] == "在售":
            return "P2"
        return "P3"

    ordered = sorted(rows, key=lambda x: (_prio(x), -x["ml_qty"], x["sku"]))
    for r in ordered:
        p = _prio(r)
        asin = r["asin"] if re.fullmatch(r"B0[A-Z0-9]{8}", r["asin"] or "") else ""
        mlm = (r["mlm"] or "").split(",")[0].strip()
        mlm_id = mlm.replace("MLM", "")

        if asin:
            amz_url = f"https://www.amazon.com.mx/dp/{asin}"
            amz_txt = f"amazon.com.mx/dp/{asin}"
        else:
            kw = quote(r["model"] or r["sku"])
            amz_url = f"https://www.amazon.com.mx/s?k={kw}"
            amz_txt = "（无 ASIN）按机型搜索"

        if mlm_id:
            ml_url = f"https://articulo.mercadolibre.com.mx/MLM-{mlm_id}"
            ml_txt = f"MLM-{mlm_id}"
        else:
            ml_url, ml_txt = "", "—"

        if p == "P1":
            tip = "已证实平台能接单但不在库存明细 → 打开确认是否仍可下单，能下单立即下架"
        elif p == "P2":
            tip = "在售但近三月无单 → 确认墨西哥是否有 listing；有则下架，无则记「未上架」"
        else:
            tip = "已停售 → 确认是否已彻底下架，确认后标记"

        sh.append([p, r["sku"], r["model"], r["p_state"], r["p_score"] or "—",
                   asin or "—", amz_txt, mlm or "—", ml_txt,
                   r["ml_orders"] if r["ml_orders"] else "—",
                   ("—" if r["mlm_in_pub"] is None else ("是" if r["mlm_in_pub"] else "否")),
                   tip, ""])
        row = sh.max_row
        sh.cell(row=row, column=7).hyperlink = amz_url
        sh.cell(row=row, column=7).font = LINK_FONT if asin else Font(size=9, color="808080", underline="single")
        if ml_url:
            sh.cell(row=row, column=9).hyperlink = ml_url
            sh.cell(row=row, column=9).font = LINK_FONT
        if p == "P1":
            for c in range(1, len(h5) + 1):
                sh.cell(row=row, column=c).fill = RED_FILL
        sh.cell(row=row, column=13).fill = PatternFill("solid", fgColor="FFF7E6")

    dv = DataValidation(
        type="list",
        formula1='"无 listing,有 listing 已下架,有 listing 待下架,已停售无需处理,待确认"',
        allow_blank=True, showDropDown=False)
    sh.add_data_validation(dv)
    dv.add(f"M{hrow + 1}:M{sh.max_row}")

    for i, w in enumerate([8, 9, 26, 10, 7, 13, 26, 16, 18, 10, 12, 46, 18], 1):
        sh.column_dimensions[get_column_letter(i)].width = w
    sh.freeze_panes = f"C{hrow + 1}"

    # ---------- 美客多订单明细
    sh = wb.create_sheet("美客多订单明细")
    sh["A1"] = "美客多墨西哥 · 白边液产品订单明细（2026-08-31 ~ 2026-10-27）"
    sh["A1"].font = T_FONT
    sh["A2"] = "共 11 条。红底 = 买家取消（金额 0）；其余为有效订单，需按此触发备货。"
    sh["A2"].font = Font(size=9, color="808080")
    sh.append([])
    h2 = ["订单号", "下单日期", "年-月", "SKU", "数量", "金额(MXN)", "订单状态",
          "MLM 发布编号", "是否在库存明细", "商品标题"]
    sh.append(h2)
    _hdr(sh, 4, len(h2))
    for o in sorted(ml_orders, key=lambda x: (x["ymd"] or (0, 0, 0))):
        sh.append([
            o["oid"], o["raw_date"].replace(" hs.", ""), 
            f'{o["ymd"][0]}-{o["ymd"][1]:02d}' if o["ymd"] else "—",
            o["sku"], int(o["qty"]), round(o["total"], 2), o["estado"],
            o["mlm"], "否", o["title"],
        ])
        if "cancel" in o["estado"].lower() or "取消" in o["estado"]:
            for c in range(1, len(h2) + 1):
                sh.cell(row=sh.max_row, column=c).fill = PatternFill("solid", fgColor="F2F2F2")
    for i, w in enumerate([20, 26, 10, 9, 7, 11, 30, 16, 12, 46], 1):
        sh.column_dimensions[get_column_letter(i)].width = w
    sh.freeze_panes = "A5"

    # ---------- 亚马逊对照
    sh = wb.create_sheet("亚马逊订单对照")
    sh["A1"] = "亚马逊 · 白边液产品订单（2026-09-24 ~ 2026-10-30，按站点渠道）"
    sh["A1"].font = T_FONT
    sh["A2"] = ("用于证明：白边液在加拿大、美国正常出单，但 Amazon.com.mx（墨西哥）一条都没有。"
                "→ 墨西哥亚马逊的合规执行到位，问题只在美客多。")
    sh["A2"].font = Font(size=9, color="808080")
    sh.append([])
    agg = collections.Counter()
    for o in amz_orders:
        agg[(o["channel"], o["status"])] += o["qty"]
    h3 = ["渠道", "订单状态", "件数"]
    sh.append(h3)
    _hdr(sh, 4, len(h3))
    for (c, s), q in sorted(agg.items(), key=lambda x: -x[1]):
        sh.append([c, s, int(q)])
    mx = sum(q for (c, s), q in agg.items() if "mx" in c.lower())
    sh.append(["Amazon.com.mx（墨西哥）", "合计", int(mx)])
    for c in range(1, 4):
        sh.cell(row=sh.max_row, column=c).font = Font(bold=True)
        sh.cell(row=sh.max_row, column=c).fill = PatternFill("solid", fgColor="DDEBF7")
    for i, w in enumerate([26, 18, 10], 1):
        sh.column_dimensions[get_column_letter(i)].width = w

    # ---------- 站点覆盖对照
    sh = wb.create_sheet("站点覆盖对照")
    sh["A1"] = "白边液 SKU 在各站点管理表里的覆盖情况（共 52 个白边液 SKU）"
    sh["A1"].font = T_FONT
    sh["A2"] = ("美国 37 个、加拿大 34 个正常在管；墨西哥亚马逊 / 美客多全部为 0。"
                "→ 白边液在墨西哥两站点根本没有纳入库存与 listing 管理。")
    sh["A2"].font = Font(size=9, color="808080")
    sh.append([])
    h4 = ["平台", "站点", "表名", "表内 SKU 数", "白边液命中", "说明"]
    sh.append(h4)
    _hdr(sh, 4, len(h4))
    for plat, site, sheet, n, hit, note in coverage(rows):
        sh.append([plat, site, sheet, n, hit, note])
        if site == "墨西哥" and hit == 0:
            for c in range(1, len(h4) + 1):
                sh.cell(row=sh.max_row, column=c).fill = RED_FILL
    for i, w in enumerate([10, 10, 26, 12, 12, 30], 1):
        sh.column_dimensions[get_column_letter(i)].width = w

    # ---------- 口径说明
    sh = wb.create_sheet("口径说明")
    L = [
        ("墨西哥白边液产品核查 · 2026-10-30", T_FONT),
        ("", None),
        ("一、白边液产品怎么筛", Font(bold=True, size=11)),
        ("商品资料（资产文件/商品资料.xlsx）→「钢化膜组装表」→「白边液」列（第 12 列 L）。", None),
        ("该列取值只有 -/－（无）和 1、2（用量），>0 的共 52 个 SKU。", None),
        ("组装表 SKU 带 ZH 后缀（如 B172ZH），通过商品资料的第二个 SKU 列（O 列）映射回主 SKU。", None),
        ("", None),
        ("二、数据源与时间范围", Font(bold=True, size=11)),
        ("· 亚马逊墨西哥：「北美加拿大库存销售表.xlsx」→「墨西哥-每日库存分析表 」（表名末尾有空格）", None),
        (f"  日销表内 SKU {amz_total} 个；8/9/10 月销量取 col78/77/76；listing 状态取 col3", None),
        ("· 美客多墨西哥：「美客多库存销售表.xlsx」→「墨西哥-每日库存分析表」", None),
        (f"  日销表内 SKU {ml_total} 个；8/9/10 月销量取 col68/67/66；listing 状态取 col3", None),
        ("· 订单交叉：「每日销售数据」9/24~10/30（亚马逊）、「订单数据」8/31~10/27（美客多）", None),
        ("· 9 月交叉验证：「2026年9月-日销站点汇总（09月01日-10月01日）.xlsx」墨西哥 / 墨西哥(2)", None),
        ("", None),
        ("三、核心发现", Font(bold=True, size=11)),
        ("1. 52 个白边液 SKU 在亚马逊墨西哥、美客多墨西哥的日销表里**一个都没有**", None),
        ("   （亚马逊 415 个 SKU 无交集、美客多 360 个 SKU 无交集），即都不在库存/备货跟踪范围内。", None),
        ("2. 亚马逊墨西哥：近三个月白边液销量 = 0", None),
        ("   · 日销表无这些 SKU；9 月站点汇总（403 SKU）无；9/24~10/30 的 Amazon.com.mx 订单无一条", None),
        ("   · 对照：同一时段白边液在 Amazon.com（美国）和 Amazon.ca（加拿大）正常出单", None),
        ("3. 美客多墨西哥：近三个月有 11 条订单 / 11 件（8/31~10/21）", None),
        ("   · 4 条被买家取消（金额 0），7 条为有效订单，合计金额约 924 MXN", None),
        ("   · 这些 SKU 都不在美客多日销表（360 个）里，也不在「库存明细」（353 个发布编号）里", None),
        ("   · 但订单里的 10 个 MLM 发布编号在平台上都是可下单状态", None),
        ("   → 即「listing 存在、能接单，但完全脱离运营的库存/listing 管理表」，", None),
        ("     运营看不到、不会主动下架；客户下单后按订单触发备货 → 白边液被采购和发货。", None),
        ("4. 商品资料里 G171 已标「停售」，10-17 仍出了一单。", None),
        ("", None),
        ("四、已知限制", Font(bold=True, size=11)),
        ("· 亚马逊订单明细只保留最近 37 天（9/24 起），8 月与 9 月上旬无法从订单层验证；", None),
        ("  但 9 月站点汇总与日销表双重证明 9 月为 0，8 月因 SKU 不在表内同样为 0。", None),
        ("· 美客多订单明细覆盖 8/31~10/27，8 月只有 8/31 一天。", None),
        ("· 「是否在库存明细」按发布编号（去掉 MLM 前缀）匹配「库存明细」页 col3（Publicaciones）。", None),
        ("· 采购/发货单据不在本次数据范围内，本表只能证明「订单已产生」，", None),
        ("  是否已实际采购、发货需另查采购单与出库记录。", None),
        ("", None),
        ("五、Listing 排查清单怎么用（第 2 个页签）", Font(bold=True, size=11)),
        ("· 亚马逊链接用 amazon.com.mx/dp/ASIN。已核实：306 个跨站 SKU 在美/加/墨三站的 ASIN 完全相同，", None),
        ("  所以商品资料里的 ASIN 可以直接打开墨西哥站，不需要另找。", None),
        ("· 52 个白边液里 40 个有合法 ASIN；12 个商品资料是 #N/A（全部是已停售的老 SKU），", None),
        ("  这些行改用「按机型搜索」的兜底链接，需在卖家后台按 SKU 反查 ASIN。", None),
        ("· 美客多只有 10 个 SKU 有 MLM（从订单里带出来的）；其余 42 个近三月无单，", None),
        ("  需登录卖家后台按 SKU 搜发布：mercadolibre.com.mx/publicaciones/listado", None),
        ("· 判读规则：亚马逊链接 404 或无法加购 = 墨西哥未上架（合规，无需处理）；", None),
        ("  能打开且可加购 = 有 listing，需要下架。美客多能打开且可下单 = 僵尸 listing 仍在。", None),
        ("· 优先级：P1 有真实订单（10 个，先查）｜P2 在售但无单（23 个）｜P3 已停售（19 个）。", None),
        ("  最后一列「排查结论」带下拉，可直接选。", None),
    ]
    for i, (t, f) in enumerate(L, 1):
        c = sh.cell(row=i, column=1, value=t)
        if f:
            c.font = f
    sh.column_dimensions["A"].width = 118

    wb.save(OUT)


if __name__ == "__main__":
    import sys
    main(write="--write" in sys.argv)
