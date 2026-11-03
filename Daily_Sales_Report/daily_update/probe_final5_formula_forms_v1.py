# -*- coding: utf-8 -*-
"""probe_final5_formula_forms_v1.py — dump 传播点是公式还是静态值（决定改法）"""
from openpyxl import load_workbook

SEA = r"D:\Desktop\work-main\sale_file_gen\菲律宾重算_202609\2026年东南亚销量表Southeast Asia Sales table.xlsx"
MON = r"D:\Desktop\work-main\sale_file_gen\菲律宾重算_202609\月销售总表.xlsx"

print("=" * 80)
print("东南亚销量表 · 公式形态")
print("=" * 80)
wb = load_workbook(SEA, data_only=False)
ph = wb["Philippines 菲律宾销量表"]
print(f"  Philippines!G554(行19): BA19={ph.cell(19, 53).value!r} BD19={ph.cell(19, 56).value!r}")
print(f"  Philippines 行2: BA2={ph.cell(2, 53).value!r} BD2={ph.cell(2, 56).value!r}")
st = wb["SEA Total 东南亚总销量表"]
print(f"  SEA Total G554(行19): CI19={st.cell(19, 87).value!r} CJ19={st.cell(19, 88).value!r}")
print(f"  SEA Total 行2: CI2={st.cell(2, 87).value!r} CJ2={st.cell(2, 88).value!r}")
print(f"  SEA Total F2={st.cell(2, 6).value!r} G2={st.cell(2, 7).value!r}")
yr = wb["2023-2026年销量情况 Sales"]
print(f"  年度sheet AU2={yr.cell(2, 47).value!r} AU3(菲)={yr.cell(3, 47).value!r} AU4(泰)={yr.cell(4, 47).value!r}")
print(f"  年度sheet row1 前5列: {[yr.cell(1, c).value for c in range(1, 6)]}")
print(f"  年度sheet row2 前5列: {[yr.cell(2, c).value for c in range(1, 6)]}")
print(f"  年度sheet A3={yr.cell(3, 1).value!r} A4={yr.cell(4, 1).value!r}")
wb.close()

print()
print("=" * 80)
print("月销售总表 · 公式形态")
print("=" * 80)
wm = load_workbook(MON, data_only=False)
ws = wm["2026年"]
g554 = None
for r in range(4, ws.max_row + 1):
    if str(ws.cell(r, 297).value or "").strip().upper() == "G554":
        g554 = r
        break
print(f"  2026年 G554 行号 = {g554}")
print(f"  G554: EM={ws.cell(g554, 135).value!r} HS={ws.cell(g554, 227).value!r} "
      f"KA={ws.cell(g554, 287).value!r} KJ={ws.cell(g554, 296).value!r}")
print(f"  行3: EM3={ws.cell(3, 135).value!r} HS3={ws.cell(3, 227).value!r} KA3={ws.cell(3, 287).value!r}")
yr2 = None
for sn in wm.sheetnames:
    if "销量情况" in sn:
        yr2 = wm[sn]
        print(f"  年度sheet名: {sn!r}")
        break
for cell in ("G10", "AA10", "G14", "AA14", "J14", "AD14", "I14"):
    print(f"  {cell} = {yr2[cell].value!r}")
print(f"  A10={yr2['A10'].value!r} B10={yr2['B10'].value!r}")
print(f"  A14={yr2['A14'].value!r} B14={yr2['B14'].value!r}")
print(f"  A2={yr2['A2'].value!r} B2={yr2['B2'].value!r}")
wm.close()
