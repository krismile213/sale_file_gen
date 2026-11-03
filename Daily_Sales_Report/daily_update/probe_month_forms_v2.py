# -*- coding: utf-8 -*-
"""probe_month_forms_v2.py — 月销总表传播点公式/缓存形态（修正 SKU 在 A 列）"""
from openpyxl import load_workbook

MON = r"D:\Desktop\work-main\sale_file_gen\菲律宾重算_202609\月销售总表.xlsx"

wf = load_workbook(MON, data_only=False)
wv = load_workbook(MON, data_only=True)
ws, wsv = wf["2026年"], wv["2026年"]

g554 = None
for r in range(4, ws.max_row + 1):
    if str(ws.cell(r, 1).value or "").strip().upper() == "G554":
        g554 = r
        break
print(f"2026年 G554 行号(A列) = {g554}")
if g554:
    for col, name in [(135, "EM菲"), (227, "HS东南亚"), (287, "KA总计"), (296, "KJ当年")]:
        print(f"  {name}: 公式={ws.cell(g554, col).value!r} 值={wsv.cell(g554, col).value!r}")
print(f"  行3: EM3 公式={ws.cell(3,135).value!r} 值={wsv.cell(3,135).value!r}")
print(f"       HS3 公式={ws.cell(3,227).value!r} 值={wsv.cell(3,227).value!r}")
print(f"       KA3 公式={ws.cell(3,287).value!r} 值={wsv.cell(3,287).value!r}")

yr, yrv = None, None
for sn in wf.sheetnames:
    if "销量情况" in sn:
        yr, yrv = wf[sn], wv[sn]
        print(f"年度 sheet: {sn!r}")
        break
for cell in ("G10", "AA10", "G14", "AA14", "J14", "AD14", "H10", "I10"):
    print(f"  {cell}: 公式={yr[cell].value!r} 值={yrv[cell].value!r}")
print(f"  A10={yr['A10'].value!r}｜A14={yr['A14'].value!r}｜A2={yr['A2'].value!r}｜B2={yr['B2'].value!r}")
print(f"  行10 G~J: {[(c, yrv[f'{c}10'].value) for c in ('E','F','G','H','I','J')]}")
print(f"  行10 U~AD: {[(c, yrv[f'{c}10'].value) for c in ('U','V','W','X','Y','Z','AA','AB','AC','AD')]}")
print(f"  行14 G~J: {[(c, yrv[f'{c}14'].value) for c in ('E','F','G','H','I','J')]}")
print(f"  行14 U~AD: {[(c, yrv[f'{c}14'].value) for c in ('U','V','W','X','Y','Z','AA','AB','AC','AD')]}")
wf.close()
wv.close()
