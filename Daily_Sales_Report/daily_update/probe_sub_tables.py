# -*- coding: utf-8 -*-
"""只读核对：3 个次表 vs 主表三国 sheet —— 确认次表归属国家与月份口径。"""
import os
from openpyxl import load_workbook

D = r"D:\Desktop\work-main\sale_file_gen\trunk\result\东南亚销量表\202609"
MAIN = os.path.join(D, "2026年东南亚销量表Southeast Asia Sales table.xlsx")


def num(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return 0.0


def read_rows(path, sheet):
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb[sheet]
    rows = list(ws.iter_rows(values_only=True))
    wb.close()
    return rows


def sub_summary(path):
    rows = read_rows(path, "Sheet1")
    head = [str(x) if x is not None else "" for x in rows[0]]
    data = [r for r in rows[1:] if r[0] is not None]
    # 找合计列 = 表头含 "Total" 的列
    ti = next(i for i, h in enumerate(head) if "Total" in h)
    ci = next(i for i, h in enumerate(head) if "Cumulative" in h)
    tot = sum(num(r[ti]) for r in data)
    cum = sum(num(r[ci]) for r in data)
    ch = sum(num(r[j]) for r in data for j in range(4, ti))
    return head, len(data), ch, tot, cum


print("=" * 78)
print("一、次表本体（当月合计列 / 渠道列合计 / SKU 数）")
print("=" * 78)
SUBS = {
    "Indonesia_202609_sales.xlsx": "印尼",
    "Philippines_202609_sales.xlsx": "菲律宾",
    "Thailand_202609_sales.xlsx": "泰国",
}
sub_data = {}
for f, cn in SUBS.items():
    head, n, ch, tot, cum = sub_summary(os.path.join(D, f))
    sub_data[f] = (head, n, ch, tot, cum)
    print(f"\n[{cn}] {f}")
    print(f"  表头: {head}")
    print(f"  SKU 行数={n}  渠道列合计={ch:.0f}  当月合计列求和={tot:.0f}  累计列求和={cum:.0f}")
    print(f"  渠道列 vs 合计列: {'一致' if abs(ch - tot) < 0.5 else '★不一致'}")

# ---- 2) 主表三国 sheet 的 8 月 / 9 月合计（数据行 r3+ 累加） ----
print("\n" + "=" * 78)
print("二、主表对应 sheet 的 8 月 / 9 月合计（数据行累加）")
print("=" * 78)
MAIN_BLOCKS = {
    "Philippines 菲律宾销量表": (51, 56),   # AY=1月块首 → 9月=BD(56), 8月=BC(55)
    "Thailand 泰国销量表": (43, 47),        # AQ → 9月=AU(47), 8月=AT(46)
    "Indonesia 印尼销量表": (35, 38),       # AI → 9月=AL(38), 8月=AK(37)
}
CN = {"Philippines 菲律宾销量表": "菲律宾", "Thailand 泰国销量表": "泰国", "Indonesia 印尼销量表": "印尼"}
main_vals = {}
for sn, (lo, hi) in MAIN_BLOCKS.items():
    rows = read_rows(MAIN, sn)
    # 行1/行2 确认列头
    head1 = [str(x) if x is not None else "" for x in rows[0]]
    head2 = [str(x) if x is not None else "" for x in rows[1]]
    tot_idx = hi - 1  # 0-based：9月合计列
    aug_idx = tot_idx - 1
    tot9 = sum(num(r[tot_idx]) for r in rows[2:] if r[0] is not None)
    tot8 = sum(num(r[aug_idx]) for r in rows[2:] if r[0] is not None)
    # 渠道列（块内除合计列）
    ch9 = sum(num(r[j]) for r in rows[2:] if r[0] is not None for j in range(lo - 1, tot_idx))
    main_vals[sn] = (tot8, tot9, ch9)
    print(f"\n[{CN[sn]}] {sn}")
    print(f"  9月块列头: {head1[lo-1:hi]}  行2: {head2[lo-1:hi]}")
    print(f"  9月合计列数据行求和={tot9:.0f}（渠道列合计={ch9:.0f}）  8月合计列数据行求和={tot8:.0f}")

# ---- 3) 三方比对 ----
print("\n" + "=" * 78)
print("三、次表 ↔ 主表 比对（判断次表是哪个月的数据）")
print("=" * 78)
PAIR = {"Philippines_202609_sales.xlsx": "Philippines 菲律宾销量表",
        "Thailand_202609_sales.xlsx": "Thailand 泰国销量表",
        "Indonesia_202609_sales.xlsx": "Indonesia 印尼销量表"}
for f, sn in PAIR.items():
    _, n, ch, tot, cum = sub_data[f]
    t8, t9, ch9 = main_vals[sn]
    verdict = []
    if abs(tot - t9) < 0.5:
        verdict.append("= 主表9月")
    if abs(tot - t8) < 0.5:
        verdict.append("= 主表8月")
    if abs(ch - ch9) < 0.5:
        verdict.append("渠道列=主表9月渠道")
    if abs(ch - (t9 - t8)) < 0.5:
        verdict.append("= 9月-8月增量?!")
    print(f"\n[{SUBS[f]}] {f}")
    print(f"  次表当月={tot:.0f}  主表8月={t8:.0f}  主表9月={t9:.0f}   → {'; '.join(verdict) or '★与8/9月都不等'}")
    print(f"  次表累计={cum:.0f}")
