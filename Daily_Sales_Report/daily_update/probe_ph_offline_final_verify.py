# -*- coding: utf-8 -*-
"""
手工单归线下 · 独立终验（2026-10-16）
不依赖改动探针的任何内部状态，全部从文件现状独立复核：
  1. zip 完整性（testzip）
  2. 主表公式格数量 = 441（XML 手术未碰其他 sheet 的证明之一）
  3. 主表 vs 备份：除 Philippines sheet 外**全簿逐格 0 差异**（XML 手术零副作用）
  4. 菲律宾 15 个手工 SKU：BA/BC（主表）、E/G（次表 sheet2）逐格值
  5. 行2 合计 + 数据行求和（六列全对）
  6. 恒等链：菲律宾 BD == SEA Total CA == 月销 EM；HS/KA/KJ 不变；印尼 AJ 未动
"""
import os, zipfile
import openpyxl

REPO = r"D:/Desktop/work-main/sale_file_gen"
MAIN = os.path.join(REPO, "菲律宾重算_202609", "2026年东南亚销量表Southeast Asia Sales table.xlsx")
SUB = os.path.join(REPO, "东南亚次表分发_202609", "Philippines_202609_sales.xlsx")
BAK = r"D:/dev-cache/sale_file_gen/backup/手工单归线下_20261016"
MANUAL = {"G376": 9, "G397": 10, "G401": 7, "G418": 15, "G427": 2, "G443": 9, "G445": 5,
          "G474": 12, "G484": 20, "G497": 8, "G530": 10, "G557": 7, "G567": 10, "G568": 10, "G569": 11}

ok = True


def check(cond, msg):
    global ok
    print(("  ✓ " if cond else "  ✗ ") + msg)
    if not cond:
        ok = False


print("== 1. zip 完整性 ==")
for p in (MAIN, SUB):
    with zipfile.ZipFile(p) as z:
        check(z.testzip() is None, f"{os.path.basename(p)} testzip OK")

print("== 2. 主表公式格数量（期望 441，证明未碰其他 sheet） ==")
wb_f = openpyxl.load_workbook(MAIN, data_only=False)
n_formula = sum(1 for sn in wb_f.sheetnames for row in wb_f[sn].iter_rows()
                for c in row if isinstance(c.value, str) and c.value.startswith("="))
check(n_formula == 441, f"公式格 = {n_formula}")

print("== 3. 主表 vs 备份：除 Philippines sheet 外逐格 0 差异 ==")
wb_new = openpyxl.load_workbook(MAIN, data_only=False)
wb_old = openpyxl.load_workbook(os.path.join(BAK, "2026年东南亚销量表Southeast Asia Sales table.xlsx"), data_only=False)
for sn in wb_old.sheetnames:
    if sn == "Philippines 菲律宾销量表":
        continue
    ws_n, ws_o = wb_new[sn], wb_old[sn]
    diff = 0
    for r in range(1, max(ws_n.max_row, ws_o.max_row) + 1):
        for c in range(1, max(ws_n.max_column, ws_o.max_column) + 1):
            if ws_n.cell(r, c).value != ws_o.cell(r, c).value:
                diff += 1
    check(diff == 0, f"sheet {sn!r} 差异格 = {diff}")

print("== 4. 菲律宾 sheet 内部：菲律宾 sheet 内其他区域未动（只许 9 月块 BA/BC + 行2 变） ==")
ws_n = wb_new["Philippines 菲律宾销量表"]
ws_o = wb_old["Philippines 菲律宾销量表"]
# 允许变化的格：9 月块 BA(53)/BC(55) 的全部行（含行2 合计）
changed = []
for r in range(1, max(ws_n.max_row, ws_o.max_row) + 1):
    for c in range(1, max(ws_n.max_column, ws_o.max_column) + 1):
        vn, vo = ws_n.cell(r, c).value, ws_o.cell(r, c).value
        if vn != vo:
            changed.append((r, c))
bad = [(r, c) for r, c in changed if not ((c in (53, 55)) and (r == 2 or r >= 3))]
check(not bad, f"菲律宾 sheet 变化格 {len(changed)} 个，全部位于 BA/BC 列（含行2）：{'是' if not bad else bad[:5]}")

print("== 5. 15 个手工 SKU 逐格值（主表 BA/BC、次表 E/G） ==")
wb_m = openpyxl.load_workbook(MAIN, data_only=True)
ws_m = wb_m["Philippines 菲律宾销量表"]
sku_row = {str(ws_m.cell(r, 1).value).strip(): r for r in range(3, ws_m.max_row + 1) if ws_m.cell(r, 1).value}
for sku, qty in MANUAL.items():
    r = sku_row[sku]
    bc = ws_m.cell(r, 55).value or 0
    ba = ws_m.cell(r, 53).value or 0
    check(bc == qty, f"主表 {sku} 行{r}: BC(线下)={bc}（应={qty}）BA(Shopee)={ba}")
wb_s = openpyxl.load_workbook(SUB, data_only=True)
ws_s = wb_s.worksheets[1]
sku_row2 = {str(ws_s.cell(r, 1).value).strip(): r for r in range(3, ws_s.max_row + 1) if ws_s.cell(r, 1).value}
for sku, qty in MANUAL.items():
    r = sku_row2[sku]
    g = ws_s.cell(r, 7).value or 0
    e = ws_s.cell(r, 5).value or 0
    check(g == qty, f"次表 {sku} 行{r}: G(Offline)={g}（应={qty}）E(Shopee)={e}")

print("== 6. 行2 + 数据行合计 ==")
for tag, ws, cols, exp in [("主表", ws_m, [51, 52, 53, 54, 55, 56], [168, 558, 1809, 649, 145, 3329]),
                           ("次表", ws_s, [3, 4, 5, 6, 7, 8], [168, 558, 1809, 649, 145, 3329])]:
    for c, e in zip(cols, exp):
        v2 = ws.cell(2, c).value
        s_data = sum(v for v in (ws.cell(r, c).value for r in range(3, ws.max_row + 1)) if isinstance(v, (int, float)))
        check(v2 == e and s_data == e, f"{tag} 列{c}: 行2={v2} 数据行和={s_data:.0f}（应={e}）")

print("== 7. 恒等链与国家维度不变 ==")
ws_st = wb_m["SEA Total 东南亚总销量表"]
check(ws_st.cell(2, 79).value == 3329, f"SEA Total 菲律宾 CA = {ws_st.cell(2, 79).value}")
check(ws_st.cell(2, 80).value == 3238, f"SEA Total 泰国 CB = {ws_st.cell(2, 80).value}（未动）")
check(ws_st.cell(2, 81).value == 4026, f"SEA Total 印尼 CC = {ws_st.cell(2, 81).value}（未动）")
ws_id = wb_m["Indonesia 印尼销量表"]
check(ws_id.cell(2, 36).value == 2775, f"印尼 AJ 线下 = {ws_id.cell(2, 36).value}（未动，含 YFDZ 200 件）")
wbm = openpyxl.load_workbook(os.path.join(REPO, "菲律宾重算_202609/月销售总表.xlsx"), data_only=True)
wsm = wbm["2026年"]
for col, name, exp_v in [(143, "EM 菲律宾", 3329), (227, "HS 东南亚", 11154), (287, "KA 总计", 63043)]:
    s = sum(v for v in (wsm.cell(r, col).value for r in range(4, wsm.max_row + 1)) if isinstance(v, (int, float)))
    check(s == exp_v, f"月销总表 {name} 数据行和 = {s}（应={exp_v}）")

print()
print("总结：" + ("全部通过 ✓" if ok else "存在失败项 ✗"))
