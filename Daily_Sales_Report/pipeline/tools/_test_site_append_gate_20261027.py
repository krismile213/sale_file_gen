# -*- coding: utf-8 -*-
"""测试：站点表补行闸接进跑批（runner._site_append_rows）。

A 回归：当前正式簿 → 10 站应全部 0 补行（G598/K001 已于 2026-10-27 手工补过）。
B 功能：临时副本上把意大利的 G598 行名改成 G598X（制造缺行）→ 应补 1 行，
        且补完 apply_month_summary 后新行 10 月值 = it 渠道 G598 销量。
只动临时副本，不碰正式文件。
"""
import os, sys, shutil, tempfile
HERE = os.path.dirname(os.path.abspath(__file__))
CORE = os.path.abspath(os.path.join(HERE, "..", "core"))
sys.path.insert(0, CORE)
import config as C
import runner as R
import analysis as An
from openpyxl import load_workbook

SNAP = "20261027"
FNAME = "欧盟英国库存销售表.xlsx"
SRC = os.path.join(C.WORK_DIR, FNAME)
SITES = list(C.SITE_SHEET_CHANNELS.keys())

def strip_and_load(cp):
    st = cp + ".strip.xlsx"
    n = R.strip_pivot_records(cp, st)
    wb = load_workbook(st if n else cp, data_only=False)
    return wb, (st if n else cp)

print("=== A 回归：正式簿应 0 补行 ===")
wb, _p = strip_and_load(SRC)
n = R._site_append_rows(wb, SNAP)
print("补行总数:", n, "→", "OK" if n == 0 else "✗ 不应为 0")
wb.close()

print("\n=== B 功能：副本删 G598 行名 → 应补 1 行并拿到 10 月值 ===")
tmpd = tempfile.mkdtemp(prefix="sitegate_")
cp = os.path.join(tmpd, FNAME)
shutil.copy2(SRC, cp)
wb, wp = strip_and_load(cp)
ws = wb["意大利"]
g598_row = next((r for r in range(7, (ws.max_row or 7) + 1) if ws.cell(r, 1).value == "G598"), None)
print("意大利 G598 原行:", g598_row)
ws.cell(g598_row, 1).value = "G598X"          # 制造缺行（结构不动，只改行名）
n = R._site_append_rows(wb, SNAP)
print("补行总数:", n)
assert n == 1, "应恰好补 1 行"
newr = next((r for r in range(7, (ws.max_row or 7) + 1) if ws.cell(r, 1).value == "G598"), None)
last = max(r for r in range(7, (ws.max_row or 7) + 1) if ws.cell(r, 1).value)
print("新行: 行%d（末行 %d）" % (newr, last))
for col, tag in (("B","品名"),("D","状态"),("E","库存"),("AX","10月公式")):
    print("  %s(%s) = %s" % (tag, col, str(ws["%s%d" % (col, newr)].value)[:70]))
e6 = ws["E6"].value
print("  E6 = %s" % str(e6)[:60])
print("\n--- apply_month_summary（与跑批同序）---")
R.apply_month_summary(wb, SNAP, fname=FNAME)
wb.calculation.fullCalcOnLoad = True
out = os.path.join(tmpd, "out.xlsx")
wb.save(out); wb.close()
# 重开验证（openpyxl 保存后无缓存值，看公式与静态值）
wb2 = load_workbook(out, data_only=False)
ws2 = wb2["意大利"]
r2 = next((r for r in range(7, (ws2.max_row or 7) + 1) if ws2.cell(r, 1).value == "G598"), None)
print("复验: G598 行 %s（旧行 %s 现为 G598X）" % (r2, g598_row))
print("  B 品名 =", ws2["B%d" % r2].value)
print("  D 状态 =", ws2["D%d" % r2].value)
# 10 月列：意大利 10 月列 = AX（此前实测）
ax = ws2["AX%d" % r2].value
print("  AX 10月 = %s" % str(ax)[:70])
# 用 it 渠道核算期望值
dsn = "每日销售数据"
exp = (An.month_sales(load_workbook(SRC, read_only=True)["每日销售数据"], "2026-10", channel="Amazon.it") or {}).get("G598")
print("  期望值（it 渠道 G598 10月）= %s" % exp)
wb2.close()
print("\n副本:", out)
