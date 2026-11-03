# -*- coding: utf-8 -*-
"""印尼断货口径对账（2026-10-30）

三套口径一次算清：
  ① 商品资料「在售 + In stock 0 + In transit 0」      → 34
  ② 其中 2026 年真的卖出过的                          → 5（= 程序报的"真断货"）
  ③ 日销表「剩余+在途 = 0」（不论状态）                → 77

产出 xlsx：断货分析/东南亚/明细/20261030/印尼断货口径对账_20261030.xlsx
"""
import os
from collections import Counter, defaultdict

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

DESKTOP = r"C:\Users\dell\Desktop"
PROD = os.path.join(DESKTOP, "东南亚商品资料及库存管理.xlsx")
DIARY = os.path.join(DESKTOP, "2026年印尼智慧物流销售表1-12月.xlsx")
OUT_DIR = os.path.join(
    r"C:\Users\dell\Desktop\work-main\sale_file_gen\Daily_Sales_Report",
    "断货分析", "东南亚", "明细", "20261030",
)
OUT = os.path.join(OUT_DIR, "印尼断货口径对账_20261030.xlsx")


def num(v):
    if v is None:
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).replace(",", "").strip())
    except Exception:
        return 0.0


def txt(v):
    return "" if v is None else str(v).strip()


def nk(s):
    return "".join(txt(s).split()).lower()


# ---------------- 商品资料 ----------------
wb = openpyxl.load_workbook(PROD, read_only=True, data_only=True)
ws = wb["Indonesia stock"]
hdr = [txt(x) for x in next(ws.iter_rows(min_row=1, max_row=1, values_only=True))]
col = {nk(h): i for i, h in enumerate(hdr) if txt(h)}
i_rate, i_state, i_rem = col["评级"], col["状态"], col["remarks"]
i_in, i_tr = col["instock"], col["intransit"]
week_cols = [(hdr[i], i) for i in range(len(hdr)) if txt(hdr[i]).isdigit() and len(txt(hdr[i])) == 4]

prod = {}
for r in ws.iter_rows(min_row=3, max_col=len(hdr), values_only=True):
    sku = txt(r[0])
    if not sku or sku == ".":
        continue
    weeks = {h: num(r[i]) for h, i in week_cols}
    prod[sku] = {
        "sku": sku, "rate": txt(r[i_rate]), "state": txt(r[i_state]), "rem": txt(r[i_rem]),
        "ins": num(r[i_in]), "tr": num(r[i_tr]), "both": num(r[i_both]), "weeks": weeks,
    }
wb.close()

# ---------------- 销量 1-10 月 ----------------
wb = openpyxl.load_workbook(DIARY, read_only=True, data_only=True)
sheet_of = (lambda m: f"2026年{m}月") if "2026年1月" in wb.sheetnames else (lambda m: f"{m}月")
sales = defaultdict(float)
sale_months = defaultdict(list)
for m in range(1, 11):
    sh = sheet_of(m)
    if sh not in wb.sheetnames:
        continue
    w = wb[sh]
    h = [txt(x) for x in next(w.iter_rows(min_row=1, max_row=1, values_only=True))]
    idx = None
    for cand in (f"{m}月总销售数量", "总销售数量"):
        if cand in h:
            idx = h.index(cand)
            break
    if idx is None:
        for i, v in enumerate(h):
            if "总销售数量" in v:
                idx = i
                break
    if idx is None:
        continue
    for r in w.iter_rows(min_row=3, max_col=idx + 1, values_only=True):
        sku = txt(r[0])
        if not sku or sku == ".":
            continue
        v = num(r[idx])
        if v:
            sales[sku] += v
            sale_months[sku].append((m, v))

# 日销表 10 月页
w10 = wb[sheet_of(10)]
h10 = [txt(x) for x in next(w10.iter_rows(min_row=1, max_row=1, values_only=True))]
c10 = {nk(x): i for i, x in enumerate(h10) if txt(x)}
j_state, j_in = c10["印尼状态"], c10["在库库存"]
j_both = c10.get("剩余库存+在途库存") or c10.get("剩余库存+ 在途库存")
j_avg, j_oct = c10["近3个月日均"], c10["10月总销售数量"]
j_prep = c10["是否备货"]
diary = {}
for r in w10.iter_rows(min_row=3, max_col=len(h10), values_only=True):
    sku = txt(r[0])
    if not sku or sku == ".":
        continue
    diary[sku] = {
        "sku": sku, "state": txt(r[j_state]), "ins": num(r[j_in]),
        "both": num(r[j_both]), "avg3": num(r[j_avg]), "oct": num(r[j_oct]),
        "prep": txt(r[j_prep]),
    }
wb.close()

# ---------------- ① 34 个 ----------------
g34 = [x for x in prod.values() if x["state"] == "在售" and x["ins"] <= 0 and x["tr"] <= 0]
g34.sort(key=lambda x: x["sku"])
for x in g34:
    x["s2026"] = sales.get(x["sku"], 0.0)
    x["months"] = ",".join(f"{m}月{v:g}" for m, v in sale_months.get(x["sku"], []))
    x["cls"] = "A 真断货" if x["s2026"] > 0 else "C 从未备货"

# ---------------- ③ 日销表 77 个 ----------------
g77 = [x for x in diary.values() if x["ins"] <= 0 and x["both"] <= 0]
g77.sort(key=lambda x: (x["state"], x["sku"]))
for x in g77:
    x["s2026"] = sales.get(x["sku"], 0.0)
    p = prod.get(x["sku"])
    x["prod_state"] = p["state"] if p else "(不在商品资料)"
    x["prod_tr"] = p["tr"] if p else -1
    x["wr"] = ""
    if p:
        nz = {h: v for h, v in p["weeks"].items() if v}
        if nz:
            x["wr"] = "/".join(f"{h}={v:g}" for h, v in sorted(nz.items()))

print(f"商品资料 SKU {len(prod)} | ①在售+无货+无在途 = {len(g34)}")
print(f"  ②有销量 = {sum(1 for x in g34 if x['cls'].startswith('A'))}  "
      f"零销量 = {sum(1 for x in g34 if x['cls'].startswith('C'))}")
print(f"③日销表 剩余+在途=0 = {len(g77)}  状态分布 {dict(Counter(x['state'] for x in g77).most_common())}")
print(f"  其中有销量 = {sum(1 for x in g77 if x['s2026'] > 0)}")

# ---------------- 写 xlsx ----------------
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
H_FILL = PatternFill("solid", fgColor="305496")
H_FONT = Font(bold=True, color="FFFFFF", size=10)
F_A = PatternFill("solid", fgColor="FFC7CE")   # 真断货 浅红
F_C = PatternFill("solid", fgColor="FFF2CC")   # 从未备货 浅黄

wb = openpyxl.Workbook()


def style_sheet(ws, headers, widths, rows, fills=None, freeze="A4"):
    ws.append(headers)
    for j in range(1, len(headers) + 1):
        c = ws.cell(row=1, column=j)
        c.fill, c.font = H_FILL, H_FONT
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(j)].width = widths[j - 1]
    ws.row_dimensions[1].height = 32
    for i, row in enumerate(rows, start=2):
        ws.append(row)
        for j in range(1, len(headers) + 1):
            c = ws.cell(row=i, column=j)
            c.border = BORDER
            c.alignment = Alignment(horizontal="center" if j > 2 else "left", vertical="center")
        if fills and fills[i - 2]:
            for j in range(1, len(headers) + 1):
                ws.cell(row=i, column=j).fill = fills[i - 2]
    ws.freeze_panes = freeze
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{len(rows) + 1}"


# 汇总页
ws0 = wb.active
ws0.title = "口径对账"
ws0["A1"] = "印尼断货：三套口径对账（2026-10-30）"
ws0["A1"].font = Font(bold=True, size=14)
ws0.merge_cells("A1:E1")
rows0 = [
    ["①", "商品资料：状态=在售 且 In stock=0 且 In transit=0", 34, "霞姐手动筛的口径（截图里的 34）", "含义：无现货且无任何补给"],
    ["②", "① 之中，2026 年 1-10 月有实际销量", 5, "程序报的「真断货」", "G178 G257 G402 G422 G576"],
    ["③", "① 之中，2026 年零销量", 29, "「从未备货」不是断货", "Remarks 虽写 out of stock，但整年没卖过一件"],
    ["④", "日销表：剩余库存+在途=0（不论状态）", 77, "含需求39/售完截止29/机会6/在售3", "日销表「在途」只算已发运，不含计划到货"],
    ["⑤", "④ 之中，10 月有销量", 10, "用户截图1 的 9 个 + G136B(售完截止)", "其中 7 个商品资料有在途 → 属「缺货在补」"],
]
style_sheet(ws0, ["#", "口径定义", "命中", "说明", "备注"], [5, 46, 8, 38, 44], rows0, freeze="A2")
for i in range(2, 7):
    ws0.row_dimensions[i].height = 34
    ws0.cell(row=i, column=3).font = Font(bold=True, size=12)

# 34 个逐条
rows34 = []
fills34 = []
for x in g34:
    rows34.append([x["sku"], x["rate"], x["state"], x["rem"], x["ins"], x["tr"],
                   x["s2026"], x["months"] or "—", x["cls"]])
    fills34.append(F_A if x["cls"].startswith("A") else F_C)
style_sheet(wb.create_sheet("①34个逐条"),
            ["SKU", "评级", "商品资料状态", "Remarks", "In stock", "In transit",
             "2026销量", "有销量月份", "归类"],
            [10, 7, 13, 15, 9, 10, 10, 30, 14], rows34, fills34)

# 77 个日销口径
rows77 = []
fills77 = []
for x in g77:
    rows77.append([x["sku"], x["state"], x["ins"], x["both"], x["avg3"], x["oct"],
                   x["s2026"], x["prod_state"], x["prod_tr"], x["wr"], x["prep"]])
    fills77.append(F_A if x["s2026"] > 0 else None)
style_sheet(wb.create_sheet("④日销表无货77个"),
            ["SKU", "印尼状态(日销表)", "在库", "剩余+在途", "近3月日均", "10月销量",
             "2026销量", "商品资料状态", "商品资料在途", "计划到货周", "是否备货"],
            [10, 15, 8, 11, 11, 10, 10, 14, 13, 18, 11], rows77, fills77)

os.makedirs(OUT_DIR, exist_ok=True)
wb.save(OUT)
print("已写出:", OUT)

# ---------------- 关键 9 个 ----------------
print("\n=== 日销表 10 月「无货 + 有销量」11 个（含 1 个售完截止）===")
print(f"{'SKU':8s}{'日销状态':10s}{'商品资料状态':12s}{'计划到货周':22s}{'归类'}")
for x in sorted([y for y in g77 if y["oct"] > 0], key=lambda y: -y["oct"]):
    p = prod.get(x["sku"])
    cls = "A 真断货" if (p and p["state"] == "在售" and p["ins"] <= 0 and p["tr"] <= 0) else \
          ("B 缺货在补" if (p and p["tr"] > 0) else "其他(非在售/已剔除)")
    print(f"{x['sku']:8s}{x['state']:10s}{x['prod_state']:12s}{x['wr'] or '—':22s}{cls}")
