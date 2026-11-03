# -*- coding: utf-8 -*-
"""
修复 AD14（东南亚组历年总计写错）+ 全面验证。
修复：AD14 = SUM(V14:AA14)（组内年份列 22~27）
验证：
  1) 年度 sheet 各组恒等（G14/AA14=月列和、J14/AD14=组内年份汇总和、H/I=公式值）
  2) 年度 sheet 月度列 vs 2026年 sheet 对应组数据行求和（全球 vs KA、东南亚 vs HS(227)）
  3) 亚马逊组 Q14 / 北美组 AK14 陈旧性检查（只读记录）
  4) 副本 vs 原件全量 diff：年度 sheet 只应落预期格；KJ 列 50 行
"""
import os
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter as gcl

ROOT = r"D:\Desktop\work-main\sale_file_gen"
SRC = os.path.join(ROOT, "菲律宾重算_202609", "月销售总表.xlsx")
ORIG = os.path.join(ROOT, "trunk", "result", "数据归属周期", "202609", "月销售总表.xlsx")


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


wb = load_workbook(SRC, data_only=False)
wsY = wb["2021-2026年销量情况"]
ws26 = wb["2026年"]

# ===== 1) 修复 AD14 = SUM(V14:AA14) =====
ad_old = wsY.cell(14, 30).value
ad_new = int(sum(n0(wsY.cell(14, c).value) for c in range(22, 28)))  # V~AA
wsY.cell(14, 30).value = ad_new
print(f"[修复] AD14(东南亚历年总计): {ad_old} → {ad_new}")

# ===== 2) 年度 sheet 恒等验证 =====
print("\n=== 年度 sheet 恒等 ===")
for gname, cy in [("全球", 7), ("东南亚", 27)]:
    y_sum = sum(n0(wsY.cell(r, cy).value) for r in range(2, 14))
    v14 = n0(wsY.cell(14, cy).value)
    hist = sum(n0(wsY.cell(14, c).value) for c in range(cy - 5, cy + 1))
    v14h = n0(wsY.cell(14, cy + 3).value)
    print(f"  [{gname}] 当年累计 {v14:.0f} vs SUM(月列)={y_sum:.0f} {'✓' if abs(v14-y_sum)<1e-9 else '★'}")
    print(f"        历年总计 {v14h:.0f} vs SUM(组内年份)={hist:.0f} {'✓' if abs(v14h-hist)<1e-9 else '★'}")
    h10 = wsY.cell(10, cy + 1).value
    i10 = wsY.cell(10, cy + 2).value
    exp_h = n0(wsY.cell(10, cy).value) / n0(wsY.cell(10, cy - 1).value) - 1
    exp_i = n0(wsY.cell(10, cy).value) / n0(wsY.cell(9, cy).value) - 1
    print(f"        同比 H={h10:.6f} vs 期望 {exp_h:.6f} {'✓' if abs(h10-exp_h)<1e-9 else '★'}")
    print(f"        环比 I={i10:.6f} vs 期望 {exp_i:.6f} {'✓' if abs(i10-exp_i)<1e-9 else '★'}")

# ===== 3) 年度 sheet 月度列 vs 2026年 sheet 数据行求和 =====
print("\n=== 月度列交叉验证（年度 sheet vs 2026年 sheet 数据行求和）===")
# 2026年 sheet：A列SKU行4起；全球=KA(287)、东南亚=HS(227)
ka_sum, hs_sum = {}, {}
for r in range(4, ws26.max_row + 1):
    a = ws26.cell(row=r, column=1).value
    if a is None:
        continue
    ka_sum[r] = n0(ws26.cell(row=r, column=287).value)
    hs_sum[r] = n0(ws26.cell(row=r, column=227).value)
ka_tot = sum(ka_sum.values())
hs_tot = sum(hs_sum.values())
print(f"  全球 9月：年度G10={n0(wsY.cell(10,7).value):.0f} vs 2026年KA列数据行和={ka_tot:.0f} "
      f"{'✓' if abs(n0(wsY.cell(10,7).value)-ka_tot)<1e-9 else '★'}")
print(f"  东南亚 9月：年度AA10={n0(wsY.cell(10,27).value):.0f} vs 2026年HS(227)列数据行和={hs_tot:.0f} "
      f"{'✓' if abs(n0(wsY.cell(10,27).value)-hs_tot)<1e-9 else '★'}")
# 1~8月也交叉验（年度 AA2..AA9 vs HS 列历史月）——只验 8 月（东南亚组 EL 列? 不，2026年 sheet 只有当月列）
# 年度 sheet 其他月值是历次跑批写的，抽 8 月验证与 2026年 sheet 8月列无关（不同 sheet），跳过

# ===== 4) 亚马逊/北美组汇总陈旧性（只读） =====
print("\n=== 亚马逊组 Q14 / 北美组 AK14 陈旧性（只读，不改）===")
for gname, cy in [("亚马逊", 17), ("亚马逊北美", 37)]:
    m_sum = sum(n0(wsY.cell(r, cy).value) for r in range(2, 14))
    v14 = n0(wsY.cell(14, cy).value)
    # 1~7月、1~8月、1~9月分别是多少
    m7 = sum(n0(wsY.cell(r, cy).value) for r in range(2, 9))
    m8 = sum(n0(wsY.cell(r, cy).value) for r in range(2, 10))
    m9 = m_sum
    print(f"  [{gname}] Q/AK14={v14:.0f}  1~7月={m7:.0f} 1~8月={m8:.0f} 1~9月={m9:.0f}  "
          f"→ {'陈旧(漏月)' if abs(v14-m9)>1e-9 else '与1~9月一致'}")

# ===== 5) 副本 vs 原件全量 diff =====
print("\n=== 副本 vs 原件 diff（年度 sheet + KJ 列）===")
wb_o = load_workbook(ORIG, read_only=True, data_only=False)
wo_y = wb_o["2021-2026年销量情况"]
wo26 = wb_o["2026年"]
orig_y = {}
for r in wo_y.iter_rows(min_row=1, max_row=14, max_col=40):
    for c in r:
        if c.value is not None:
            orig_y[(c.row, c.column)] = c.value
diffs = []
for r in range(1, 15):
    for c in range(1, 41):
        ov = orig_y.get((r, c))
        nv = wsY.cell(r, c).value
        if ov is None and nv is None:
            continue
        if isinstance(ov, (int, float)) and isinstance(nv, (int, float)):
            if abs(ov - nv) > 1e-12:
                diffs.append((f"{gcl(c)}{r}", ov, nv))
        elif str(ov) != str(nv):
            diffs.append((f"{gcl(c)}{r}", ov, nv))
print(f"  年度 sheet 变化 {len(diffs)} 格：")
for coord, ov, nv in diffs:
    print(f"    {coord}: {ov} → {nv}")
wb_o.close()

# KJ 列 diff（原件 vs 副本）
wb_o2 = load_workbook(ORIG, read_only=True, data_only=False)
wo2 = wb_o2["2026年"]
orig_kj = {}
for r in wo2.iter_rows(min_row=4, max_col=296):
    a = r[0].value
    if a is not None:
        orig_kj[str(a).strip().upper()] = (r[0].row, r[295].value)
wb_o2.close()
kj_diffs = []
for r in range(4, ws26.max_row + 1):
    a = ws26.cell(row=r, column=1).value
    if a is None:
        continue
    s = str(a).strip().upper()
    nv = ws26.cell(row=r, column=296).value
    ov = orig_kj.get(s, (None, None))[1]
    if isinstance(ov, (int, float)) and isinstance(nv, (int, float)):
        if abs(ov - nv) > 1e-9:
            kj_diffs.append((s, r, ov, nv))
    elif ov != nv:
        kj_diffs.append((s, r, ov, nv))
print(f"\n  KJ 列变化 {len(kj_diffs)} 行，合计增量 {sum(d[3]-d[2] for d in kj_diffs if isinstance(d[2],(int,float)) and isinstance(d[3],(int,float))):.0f}")

wb.save(SRC)
print("\n已保存：", SRC)
