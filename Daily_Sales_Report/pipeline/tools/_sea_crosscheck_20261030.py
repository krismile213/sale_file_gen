# -*- coding: utf-8 -*-
"""核对：印尼日销表（快照 2026-10-29）vs 霞姐备货表（东南亚商品资料及库存管理.xlsx）

核对五件事：SKU 覆盖 / 库存 / 在途 / 状态 / 评级。
只读，不改任何文件。
"""
import openpyxl
import collections

SALES = r"C:\Users\dell\Desktop\2026年印尼智慧物流销售表1-12月.xlsx"
STOCK = r"C:\Users\dell\Desktop\东南亚商品资料及库存管理.xlsx"


def num(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return 0.0


def txt(v):
    return str(v).strip() if v is not None else ""


# ---------------- 日销表
wb = openpyxl.load_workbook(SALES, read_only=True, data_only=True)
ws = wb["2026年10月"]
rows = list(ws.iter_rows(min_row=1, values_only=True))
hdr = rows[0]
pos = {txt(h): j for j, h in enumerate(hdr) if h is not None}
d1 = next(j for j, h in enumerate(hdr) if txt(h) == "1")
snap = txt(rows[2][4])[:10]

sales = {}
for r in rows[2:]:
    sku = txt(r[0])
    if not sku or sku == "合计":
        continue
    sales[sku] = dict(
        name=txt(r[1]), model=txt(r[2]),
        state=txt(r[9]),
        stock=num(r[pos["在库库存"]]),
        avail=num(r[pos["剩余库存+在途库存"]]),
        mtotal=num(r[pos["10月总销售数量"]]),
    )
# 评级列（31 日右边第 3 格，无表头）
RATING = {"A", "B", "C", "D", "E", "X", "XA", "XB", "XC", "XD", "XE"}
rc = None
for k in range(d1 + 31, d1 + 37):
    if txt(hdr[k]):
        continue
    vals = [txt(r[k]) for r in rows[2:] if k < len(r) and txt(r[k])]
    vals2 = [v for v in vals if v not in ("0", "0.0")]
    if len(vals2) >= 10 and all(v in RATING for v in vals2) and len(set(vals2)) >= 2:
        rc = k
        break
if rc is not None:
    for r in rows[2:]:
        sku = txt(r[0])
        if sku in sales and rc < len(r):
            sales[sku]["rating"] = txt(r[rc])
wb.close()

# ---------------- 备货表
wb = openpyxl.load_workbook(STOCK, read_only=True, data_only=True)
ws = wb["Indonesia stock"]
srows = list(ws.iter_rows(min_row=1, values_only=True))
stock = {}
for r in srows[2:]:
    sku = txt(r[0])
    if not sku:
        continue
    stock[sku] = dict(
        rating=txt(r[1]), state=txt(r[2]), remark=txt(r[3]),
        instock=num(r[6]), stock_transit=num(r[7]), transit=num(r[8]),
        weeks={txt(srows[0][j]): num(r[j]) for j in range(11, 17)},
    )
wb.close()

print("=" * 84)
print(f"日销表快照日：{snap}｜SKU {len(sales)}")
print(f"备货表 Indonesia stock：SKU {len(stock)}｜In stock 合计 {int(sum(v['instock'] for v in stock.values()))}")
print(f"日销表 在库库存 合计 {int(sum(v['stock'] for v in sales.values()))}"
      f"｜剩余+在途 合计 {int(sum(v['avail'] for v in sales.values()))}")
print(f"备货表 In stock+in transit 合计 {int(sum(v['stock_transit'] for v in stock.values()))}")
print("=" * 84)

# ---- 1. SKU 覆盖
both = set(sales) & set(stock)
only_s = sorted(set(sales) - set(stock))
only_k = sorted(set(stock) - set(sales))
print(f"\n① SKU 覆盖：交集 {len(both)}｜仅日销表 {len(only_s)}｜仅备货表 {len(only_k)}")
print(f"   仅日销表（备货表没列）：{only_s[:20]}{' …' if len(only_s) > 20 else ''}")
print(f"   仅备货表（日销表没列）：{only_k[:20]}{' …' if len(only_k) > 20 else ''}")

# ---- 2. 库存
diff_stock = [(s, sales[s]["stock"], stock[s]["instock"]) for s in sorted(both)
              if abs(sales[s]["stock"] - stock[s]["instock"]) > 0.001]
print(f"\n② 在库库存：{len(both)} 个可比，{len(diff_stock)} 个不一致")
for s, a, b in diff_stock[:15]:
    print(f"   {s:8s} 日销表={a:g}  备货表={b:g}  差{a - b:+g}")
if len(diff_stock) > 15:
    print(f"   … 另有 {len(diff_stock) - 15} 个")

# ---- 3. 在途
diff_tr = [(s, sales[s]["avail"] - sales[s]["stock"], stock[s]["transit"]) for s in sorted(both)
           if abs((sales[s]["avail"] - sales[s]["stock"]) - stock[s]["transit"]) > 0.001]
print(f"\n③ 在途：{len(diff_tr)} 个不一致（日销表 = 剩余+在途 − 在库；备货表 = In transit）")
for s, a, b in diff_tr[:12]:
    print(f"   {s:8s} 日销表={a:g}  备货表={b:g}  差{a - b:+g}")

# ---- 4. 状态
c = collections.Counter((sales[s]["state"], stock[s]["state"]) for s in both)
print("\n④ 状态交叉（日销表印尼状态 → 备货表状态）：")
for (a, b), n in sorted(c.items(), key=lambda x: -x[1]):
    print(f"   {a or '(空)':6s} → {b or '(空)':6s}  {n}")

# ---- 5. 评级
if rc is not None:
    c = collections.Counter((sales[s].get("rating", "") or "(空)", stock[s]["rating"] or "(空)")
                            for s in both)
    same = sum(n for (a, b), n in c.items() if a == b)
    print(f"\n⑤ 评级：{len(both)} 个可比，完全一致 {same} 个（{same * 100 / len(both):.1f}%）")
    for (a, b), n in sorted(c.items(), key=lambda x: -x[1]):
        if a != b:
            print(f"   日销表={a:4s} 备货表={b:4s}  {n}")
else:
    print("\n⑤ 评级：日销表评级列未定位到")

# ---- 6. 库存为 0 的分布
print("\n⑥ 零库存（In stock ≤ 0）SKU 的状态分布：")
z = [s for s in stock if stock[s]["instock"] <= 0]
print("   备货表零库存", len(z), dict(collections.Counter(stock[s]["state"] for s in z)))
z2 = [s for s in sales if sales[s]["stock"] <= 0]
print("   日销表零库存", len(z2), dict(collections.Counter(sales[s]["state"] for s in z2)))
