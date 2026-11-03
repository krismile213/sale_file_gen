# -*- coding: utf-8 -*-
"""
同步更新「菲律宾重算_202609」副本中所有引用 9 月销量 的地方：
  A. `2021-2026年销量情况` sheet：全球组(G/H/I/J) + 东南亚组(AA/AB/AC/AD)
     口径严格复刻 month_sale_file.py L1850-1891（任务③）
  B. `2026年` sheet 的 KJ(296)「当年总销量」列：按 KA(287) 的 delta 差值传播
     （不用 SUM 重算——实测有 50 行历史性 ±1~3 差异，重算会引入非预期 diff）
只改副本，不动原件。
"""
import os
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter as gcl

ROOT = r"D:\Desktop\work-main\sale_file_gen"
SRC = os.path.join(ROOT, "菲律宾重算_202609", "月销售总表.xlsx")
ORIG = os.path.join(ROOT, "trunk", "result", "数据归属周期", "202609", "月销售总表.xlsx")
SEA = os.path.join(ROOT, "菲律宾重算_202609", "2026年东南亚销量表Southeast Asia Sales table.xlsx")
SEA_ORIG = os.path.join(ROOT, "trunk", "result", "东南亚销量表", "202609",
                        "2026年东南亚销量表Southeast Asia Sales table.xlsx")

DELTA = 206  # 菲律宾重算净增件数


def n0(v):
    """数值化：'00'文本/None → 0"""
    if v is None:
        return 0
    if isinstance(v, str):
        s = v.strip()
        if s == "00" or s == "":
            return 0
        try:
            return float(s)
        except ValueError:
            return 0
    return v if isinstance(v, (int, float)) else 0


wb = load_workbook(SRC, data_only=False)
wsY = wb["2021-2026年销量情况"]
ws26 = wb["2026年"]

# ============ 0) 前置验证：副本 HS(227) 东南亚9月 是否已正确联动 ============
print("=== 0) 前置验证：副本 vs SEA Total CI(87) 恒等（抽 G376/G531/G569）===")
wb_sea = load_workbook(SEA, read_only=True, data_only=True)
sea_ws = wb_sea["SEA Total 东南亚总销量表"]
sea_ci = {}
for r in sea_ws.iter_rows(min_row=3, max_col=1):
    v = r[0].value
    if v is not None:
        sea_ci[str(v).strip().upper()] = r[0].row
wb_sea.close()
for sku in ("G376", "G531", "G569"):
    row26 = None
    for r in range(4, ws26.max_row + 1):
        a = ws26.cell(row=r, column=1).value
        if a and str(a).strip().upper() == sku:
            row26 = r
            break
    hs = ws26.cell(row=row26, column=227).value if row26 else None  # HS 东南亚9月
    sr = sea_ci.get(sku)
    print(f"  {sku}: 月销总表 HS(227)={hs!r}  (SEA Total 行 {sr})")

# ============ A. 年度 sheet ============
print("\n=== A. 2021-2026年销量情况 ===")
# 组定义：块首列、行14汇总列布局 col_y(当年)/col_y+1(同比)/col_y+2(环比或汇总同比)/col_y+3(历年总计)
# 全球 col_y=7(G)；东南亚 col_y=27(AA)
GROUPS = [("全球", 7), ("东南亚", 27)]
log = []
for gname, cy in GROUPS:
    r_cur, r_prev, r_sum = 10, 9, 14  # 9月行=10、8月行=9、汇总行=14
    old_cur = n0(wsY.cell(r_cur, cy).value)
    new_cur = old_cur + DELTA
    # 当月值
    wsY.cell(r_cur, cy).value = int(new_cur)
    # 同比增长率 = 当月 / 去年同月
    denom_y = n0(wsY.cell(r_cur, cy - 1).value)
    wsY.cell(r_cur, cy + 1).value = (new_cur / denom_y - 1) if denom_y else None
    # 今年环比 = 当月 / 今年上月
    denom_m = n0(wsY.cell(r_prev, cy).value)
    wsY.cell(r_cur, cy + 2).value = (new_cur / denom_m - 1) if denom_m else None
    # 汇总行：当年累计 = 1~12月之和
    y_sum = sum(n0(wsY.cell(r0, cy).value) for r0 in range(2, 14))
    wsY.cell(r_sum, cy).value = int(y_sum)
    # 汇总同比（模板在 +2 列，+1 留空）
    d14 = n0(wsY.cell(r_sum, cy - 1).value)
    wsY.cell(r_sum, cy + 2).value = (y_sum / d14 - 1) if d14 else None
    # 历年总计 = SUM(首年列..当年列)
    hist = sum(n0(wsY.cell(r_sum, c0).value) for c0 in range(2, cy + 1))
    wsY.cell(r_sum, cy + 3).value = int(hist)
    log.append((gname, old_cur, new_cur, y_sum, hist))
    print(f"  [{gname}] 9月: {old_cur:.0f} → {new_cur:.0f} | 当年累计: {y_sum:.0f} | 历年总计: {hist:.0f}")
    print(f"       同比={wsY.cell(r_cur, cy+1).value:.6f} 环比={wsY.cell(r_cur, cy+2).value:.6f} 汇总同比={wsY.cell(r_sum, cy+2).value:.6f}")

# ============ B. KJ 列差值传播 ============
print("\n=== B. 2026年 sheet KJ(296) 当年总销量 ===")
# 读原件 KA(287) 旧值算 delta
wb_o = load_workbook(ORIG, read_only=True, data_only=True)
wo = wb_o["2026年"]
orig_ka = {}
for r in wo.iter_rows(min_row=4, max_col=287):
    a = r[0].value
    if a is not None:
        orig_ka[str(a).strip().upper()] = r[286].value  # KA=287 → 索引286
wb_o.close()

changed = []
for r in range(4, ws26.max_row + 1):
    a = ws26.cell(row=r, column=1).value
    if a is None:
        continue
    s = str(a).strip().upper()
    ka_new = n0(ws26.cell(row=r, column=287).value)
    ka_old = n0(orig_ka.get(s))
    delta = ka_new - ka_old
    if abs(delta) < 1e-9:
        continue
    kj = ws26.cell(row=r, column=296).value
    kj_n = n0(kj)
    kj_new = kj_n + delta
    ws26.cell(row=r, column=296).value = int(kj_new)
    changed.append((s, r, kj_n, kj_new, ka_old, ka_new))
print(f"  需联动 KJ 的行：{len(changed)}")
for s, r, k0, k1, a0, a1 in changed[:12]:
    print(f"    行{r} {s}: KJ {k0:.0f}→{k1:.0f}  (KA {a0:.0f}→{a1:.0f})")
if len(changed) > 12:
    print(f"    … 其余 {len(changed)-12} 行略")
tot_kj_delta = sum(c[3] - c[2] for c in changed)
tot_ka_delta = sum(c[5] - c[4] for c in changed)
print(f"  KJ 合计增量={tot_kj_delta:.0f}  KA 合计增量={tot_ka_delta:.0f}  {'一致' if abs(tot_kj_delta-tot_ka_delta)<1e-9 else '★不一致'}")

wb.save(SRC)
print("\n已保存：", SRC)
