# -*- coding: utf-8 -*-
"""印尼「手动统计断货」与「程序判定」对账（2026-10-30 一次性探针）

回答两件事：
 1) 霞姐在 Indonesia stock 里筛出的 34 个，到底是什么条件筛出来的、和程序判定的 5 个差在哪
 2) 日销表 10 月页里那 9 个「剩余+在途=0 且在备货」的 SKU，为什么没被算成断货
"""
import os
import sys
from collections import Counter, defaultdict

import openpyxl

DESKTOP = r"C:\Users\dell\Desktop"
PROD = os.path.join(DESKTOP, "东南亚商品资料及库存管理.xlsx")
DIARY = os.path.join(DESKTOP, "2026年印尼智慧物流销售表1-12月.xlsx")


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


# ---------- 1. 商品资料 Indonesia stock ----------
wb = openpyxl.load_workbook(PROD, read_only=True, data_only=True)
ws = wb["Indonesia stock"]
hdr = [txt(x) for x in next(ws.iter_rows(min_row=1, max_row=1, values_only=True))]
print("=== Indonesia stock 表头 ===")
for i, h in enumerate(hdr[:20]):
    print(f"  idx{i:2d} {openpyxl.utils.get_column_letter(i+1):3s} {h!r}")

def _n(s):
    return "".join(txt(s).split()).lower()


col = {_n(h): i for i, h in enumerate(hdr) if txt(h)}


def ci(*names):
    for n in names:
        k = _n(n)
        if k in col:
            return col[k]
    raise KeyError(names)


C_SKU = 0
C_RATE = ci("评级")
C_STATE = ci("状态")
C_REM = ci("Remarks")
C_IN = ci("In stock")
C_INTR = ci("In transit")
C_BOTH = ci("In stock+ in transit")

rows = []
for r in ws.iter_rows(min_row=3, max_col=len(hdr), values_only=True):
    sku = txt(r[C_SKU])
    if not sku or sku == ".":
        continue
    rows.append(
        {
            "sku": sku,
            "rate": txt(r[C_RATE]),
            "state": txt(r[C_STATE]),
            "rem": txt(r[C_REM]),
            "ins": num(r[C_IN]),
            "tr": num(r[C_INTR]),
            "both": num(r[C_BOTH]) if C_BOTH < len(r) else None,
        }
    )
wb.close()

print(f"\n=== Indonesia stock 共 {len(rows)} 个 SKU ===")
print("状态分布:", dict(Counter(x['state'] for x in rows).most_common()))
print("评级分布:", dict(Counter(x['rate'] for x in rows).most_common()))

print("\n=== 各筛选组合命中的 SKU 数 ===")
combos = {
    "A 在售 & 无现货 & 无在途": lambda x: x["state"] == "在售" and x["ins"] <= 0 and x["tr"] <= 0,
    "B 在售 & 无现货（不论在途）": lambda x: x["state"] == "在售" and x["ins"] <= 0,
    "C 停售 & 无现货 & 无在途": lambda x: x["state"] == "停售" and x["ins"] <= 0 and x["tr"] <= 0,
    "D 停售（全部）": lambda x: x["state"] == "停售",
    "E 所有状态 & 无现货 & 无在途": lambda x: x["ins"] <= 0 and x["tr"] <= 0,
}
for k, fn in combos.items():
    hit = [x for x in rows if fn(x)]
    print(f"  {k:28s} = {len(hit):3d} 个")

# ---------- 2. 2026 年 1-10 月销量 ----------
wb = openpyxl.load_workbook(DIARY, read_only=True, data_only=True)
if "2026年1月" in wb.sheetnames:
    sheet_of = lambda m: f"2026年{m}月"
else:
    sheet_of = lambda m: f"{m}月"

sales = defaultdict(float)
sales_by_month = defaultdict(dict)
for m in range(1, 11):
    sh = sheet_of(m)
    if sh not in wb.sheetnames:
        continue
    w = wb[sh]
    h = [txt(x) for x in next(w.iter_rows(min_row=1, max_row=1, values_only=True))]
    # 销量列：优先「N月总销售数量」→「总销售数量」→ 兜底找含"销售数量"的第一列
    idx = None
    for cand in (f"{m}月总销售数量", "总销售数量", "月总销售数量"):
        if cand in h:
            idx = h.index(cand)
            break
    if idx is None:
        for i, v in enumerate(h):
            if "总销售数量" in v:
                idx = i
                break
    if idx is None:
        print(f"  [警告] {sh} 找不到销量列，表头前20={h[:20]}")
        continue
    for r in w.iter_rows(min_row=3, max_col=idx + 1, values_only=True):
        sku = txt(r[0])
        if not sku or sku == ".":
            continue
        v = num(r[idx])
        sales[sku] += v
        sales_by_month[sku][m] = v
print(f"\n=== 印尼日销表 1-10 月，SKU 数 {len(sales)} ===")

# ---------- 3. 交 34 个逐条对 ----------
manual34 = [x for x in rows if combos["A 在售 & 无现货 & 无在途"](x)]
manual34.sort(key=lambda x: x["sku"])
print(f"\n=== 「在售 + 无现货 + 无在途」全量 {len(manual34)} 个，逐条看今年销量 ===")
print(f"{'SKU':8s}{'评级':4s}{'状态':6s}{'Remarks':16s}{'InStock':>8s}{'InTransit':>10s}{'2026销量':>9s}{'10月':>6s}  归类")
has_sale, no_sale = [], []
for x in manual34:
    s26 = sales.get(x["sku"], 0.0)
    s10 = sales_by_month.get(x["sku"], {}).get(10, 0.0)
    if s26 > 0:
        has_sale.append(x)
        tag = "真断货(今年卖过)"
    else:
        no_sale.append(x)
        tag = "从未备货(今年零销量)"
    print(
        f"{x['sku']:8s}{x['rate']:4s}{x['state']:6s}{x['rem'][:14]:16s}"
        f"{x['ins']:8.0f}{x['tr']:10.0f}{s26:9.0f}{s10:6.0f}  {tag}"
    )
print(f"\n  其中今年有销量 = {len(has_sale)} 个 → {[x['sku'] for x in has_sale]}")
print(f"  其中今年零销量 = {len(no_sale)} 个")

# ---------- 4. 日销表 10 月页：无库存的 SKU ----------
print("\n=== 日销表 2026年10月：剩余+在途=0 的 SKU ===")
w = wb[sheet_of(10)]
h = [txt(x) for x in next(w.iter_rows(min_row=1, max_row=1, values_only=True))]
hmap = {v: i for i, v in enumerate(h) if v}
i_state = hmap.get("印尼状态")
i_ins = hmap.get("在库库存")
i_both = hmap.get("剩余库存+在途库存") or hmap.get("剩余库存 + 在途库存")
i_avg = hmap.get("近3个月日均")
i_oct = hmap.get("10月总销售数量")
i_prep = hmap.get("是否备货")
i_qty = hmap.get("备货数量")
print(f"  列位: 印尼状态={i_state} 在库={i_ins} 剩余+在途={i_both} 近3月日均={i_avg} 10月销量={i_oct} 是否备货={i_prep}")

oct_rows = []
for r in w.iter_rows(min_row=3, max_col=len(h), values_only=True):
    sku = txt(r[0])
    if not sku or sku == ".":
        continue
    oct_rows.append(
        {
            "sku": sku,
            "name": txt(r[1]) if len(r) > 2 else "",
            "state": txt(r[i_state]) if i_state is not None else "",
            "ins": num(r[i_ins]) if i_ins is not None else 0,
            "both": num(r[i_both]) if i_both is not None else 0,
            "avg3": num(r[i_avg]) if i_avg is not None else 0,
            "oct": num(r[i_oct]) if i_oct is not None else 0,
            "prep": txt(r[i_prep]) if i_prep is not None else "",
            "qty": r[i_qty] if i_qty is not None else None,
        }
    )
wb.close()

print(f"  10 月页共 {len(oct_rows)} 个 SKU")
zero = [x for x in oct_rows if x["ins"] <= 0 and x["both"] <= 0]
print(f"  剩余+在途=0 共 {len(zero)} 个")
print(f"    状态分布: {dict(Counter(x['state'] for x in zero).most_common())}")
print(f"    其中今年有销量: {sum(1 for x in zero if sales.get(x['sku'], 0) > 0)}")
print(f"    其中今年零销量: {sum(1 for x in zero if sales.get(x['sku'], 0) <= 0)}")
print(f"    其中 10 月有销量: {sum(1 for x in zero if x['oct'] > 0)}")

print("\n  --- 剩余+在途=0 且 10 月有销量 的明细 ---")
print(f"  {'SKU':8s}{'印尼状态':8s}{'在库':>5s}{'剩余+在途':>8s}{'近3月日均':>9s}{'10月':>5s}{'2026销量':>9s}{'是否备货':>8s}")
for x in sorted([x for x in zero if x["oct"] > 0], key=lambda y: -y["oct"]):
    print(
        f"  {x['sku']:8s}{x['state']:8s}{x['ins']:5.0f}{x['both']:8.0f}{x['avg3']:9.2f}"
        f"{x['oct']:5.0f}{sales.get(x['sku'], 0):9.0f}{x['prep']:>8s}"
    )

print("\n  --- 用户截图里那 9 个 ---")
user9 = ["G178", "G557", "G567", "G569", "G579", "G580", "G576", "G577", "G578"]
d = {x["sku"]: x for x in oct_rows}
prod = {x["sku"]: x for x in rows}
print(f"  {'SKU':7s}{'日销状态':8s}{'日销剩余+在途':>10s}{'商品资料状态':>10s}{'商品资料在途':>10s}{'2026销量':>9s}")
for s in user9:
    a, b = d.get(s), prod.get(s)
    print(
        f"  {s:7s}{(a['state'] if a else '-'):8s}{(a['both'] if a else -1):10.0f}"
        f"{(b['state'] if b else '-'):>10s}{(b['tr'] if b else -1):10.0f}{sales.get(s, 0):9.0f}"
    )
