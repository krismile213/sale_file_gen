# -*- coding: utf-8 -*-
"""终验②：普通模式一次性读列，避免 read_only 逐格 cell() 超时"""
import os
from openpyxl import load_workbook

ROOT = r"D:\Desktop\work-main\sale_file_gen"
SEA = os.path.join(ROOT, "菲律宾重算_202609", "2026年东南亚销量表Southeast Asia Sales table.xlsx")
M = os.path.join(ROOT, "菲律宾重算_202609", "月销售总表.xlsx")


def n0(v):
    if v is None:
        return 0
    if isinstance(v, str):
        s = v.strip()
        if s in ("00", ""):
            return 0
        try:
            return float(s)
        except ValueError:
            return 0
    return v if isinstance(v, (int, float)) else 0


wb = load_workbook(SEA, data_only=True)
ph = wb["Philippines 菲律宾销量表"]
sea = wb["SEA Total 东南亚总销量表"]
print("=== 东南亚销量表副本 ===")
print("  菲律宾 sheet 行2 BD(9月合计) =", n0(ph.cell(row=2, column=56).value))
print("  SEA Total 行2 CA(菲律宾渠道9月) =", n0(sea.cell(row=2, column=79).value))
print("  SEA Total 行2 CI(9月TOTAL) =", n0(sea.cell(row=2, column=87).value))
# 一次性取 BD 列（菲律宾）与 CI 列（SEA Total）全列值
bd_col = [c.value for c in ph["BD"]]
ci_col = [c.value for c in sea["CI"]]
bd_sum = sum(n0(v) for v in bd_col[2:])  # 行3起
ci_sum = sum(n0(v) for v in ci_col[2:])
print(f"  菲律宾 BD 数据行求和 = {bd_sum:.0f}（应=3329）")
print(f"  SEA Total CI 数据行求和 = {ci_sum:.0f}（应=11154）")
# G569 补行
for r in range(3, ph.max_row + 1):
    a = ph.cell(row=r, column=1).value
    if a and str(a).strip().upper() == "G569":
        print(f"  G569 补行在行{r}，BD(9月) =", n0(ph.cell(row=r, column=56).value))
        break
print("  菲律宾 max_row =", ph.max_row, " auto_filter =", ph.auto_filter.ref)
wb.close()

wb2 = load_workbook(M, data_only=False)
ws = wb2["2026年"]
print("\n=== 月销总表副本抽查（A列定位）===")
for sku in ("G376", "G531", "G569", "G165PRO"):
    for r in range(4, ws.max_row + 1):
        a = ws.cell(row=r, column=1).value
        if a and str(a).strip().upper() == sku:
            print(f"  行{r} {sku}: EM(菲9月)={ws.cell(row=r,column=143).value} HS(东南亚9月)={ws.cell(row=r,column=227).value} "
                  f"KA(总计9月)={ws.cell(row=r,column=287).value} KJ(当年)={ws.cell(row=r,column=296).value}")
            break
print("  行3 KA 公式 =", ws.cell(row=3, column=287).value)
print("  行3 KJ 公式 =", ws.cell(row=3, column=296).value)
wb2.close()
