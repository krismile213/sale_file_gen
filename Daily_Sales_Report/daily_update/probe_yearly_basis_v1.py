# -*- coding: utf-8 -*-
"""probe_yearly_basis_v1.py — 确认月销年度 sheet 各比率格(同比/环比)的真实基数"""
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter as gcl

MON = r"D:\Desktop\work-main\sale_file_gen\菲律宾重算_202609\月销售总表.xlsx"
wv = load_workbook(MON, data_only=True)
yr = None
for sn in wv.sheetnames:
    if "销量情况" in sn:
        yr = wv[sn]
        break
print("sheet:", yr.title)
print("\n全球组 E~J（行2~14）：")
for r in range(2, 15):
    print(f"  行{r} A={yr.cell(r,1).value!r:>8}: " + " ".join(
        f"{gcl(c)}={yr.cell(r,c).value}" for c in range(5, 11)))
print("\n东南亚组 U~AD（行2~14）：")
for r in range(2, 15):
    print(f"  行{r} A={yr.cell(r,1).value!r:>8}: " + " ".join(
        f"{gcl(c)}={yr.cell(r,c).value}" for c in range(21, 31)))
wv.close()

# 反推基数
print("\n反推（改动后新值 G10=63199 AA10=11310 G14=595571 AA14=117019）：")
def inv(val, num):
    return num / (1 + val) if val is not None else None
print(f"  H10 基数 = 63199/(1+{-0.1401651723152696}) = {inv(-0.1401651723152696, 63199):.2f}（应=F10=73499）")
print(f"  I10 基数 = 63199/(1+{-0.12056609287374231}) = {inv(-0.12056609287374231, 63199):.2f}")
print(f"  AB10 基数 = 11310/(1+{-0.3087597041383948}) = {inv(-0.3087597041383948, 11310):.2f}（应=Z10=16359）")
print(f"  AC10 基数 = 11310/(1+{-0.1787348391313821}) = {inv(-0.1787348391313821, 11310):.2f}")
print(f"  I14 基数 = 595571/(1+{-0.28822437635720666}) = {inv(-0.28822437635720666, 595571):.2f}（应=F14=836737）")
print(f"  AC14 基数 = 117019/(1+{-0.3527104768226573}) = {inv(-0.3527104768226573, 117019):.2f}（应=Z14=180780）")
