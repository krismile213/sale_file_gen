# -*- coding: utf-8 -*-
"""验证探针：副本 vs 原件逐格 diff + 副本内恒等 + 合计复算（证明"只动了菲律宾 9 月"）"""
import os
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter as gcl

BASE = r"D:\Desktop\work-main\sale_file_gen"
SEA_SRC = os.path.join(BASE, r"trunk\result\东南亚销量表\202609\2026年东南亚销量表Southeast Asia Sales table.xlsx")
MON_SRC = os.path.join(BASE, r"trunk\result\数据归属周期\202609\月销售总表.xlsx")
SEA_DST = os.path.join(BASE, r"菲律宾重算_202609\2026年东南亚销量表Southeast Asia Sales table.xlsx")
MON_DST = os.path.join(BASE, r"菲律宾重算_202609\月销售总表.xlsx")

def n0(v):
    if v is None: return 0.0
    if isinstance(v, str):
        s = v.strip()
        if s in ("", "00"): return 0.0
        try: return float(s)
        except Exception: return 0.0
    return float(v or 0)

def grid(path, data_only=True):
    wb = load_workbook(path, read_only=True, data_only=data_only)
    out = {}
    for sn in wb.sheetnames:
        ws = wb[sn]
        cells = {}
        for r in ws.iter_rows(values_only=False):
            for c in r:
                if c.value is not None:
                    cells[(c.row, c.column)] = c.value
        out[sn] = (ws.max_row, ws.max_column, cells)
    wb.close()
    return out

print("=" * 74)
print("① 逐格 diff：副本 vs 原件（东南亚销量表）")
print("=" * 74)
src = grid(SEA_SRC); dst = grid(SEA_DST)
total_diff = 0
for sn in src:
    if sn not in dst:
        print(f"  ★ sheet {sn} 在副本中缺失！"); continue
    mr, mc, sc = src[sn]; mr2, mc2, dc = dst[sn]
    diffs = []
    for k, v in sc.items():
        v2 = dc.get(k)
        if isinstance(v, (int, float)) and isinstance(v2, (int, float)):
            if abs(float(v) - float(v2)) > 1e-9: diffs.append((k, v, v2))
        elif v != v2:
            diffs.append((k, v, v2))
    for k in dc:
        if k not in sc: diffs.append((k, "<新格>", dc[k]))
    total_diff += len(diffs)
    if diffs:
        cols = sorted({k[1] for k, _, _ in diffs})
        print(f"  [{sn}] 变化 {len(diffs)} 格；涉及列: {[gcl(c) for c in cols]}")
        # 按列汇总
        bycol = {}
        for (r, c), a, b in diffs:
            bycol.setdefault(c, []).append((r, a, b))
        for c in sorted(bycol):
            items = bycol[c]
            print(f"     {gcl(c)}({c}): {len(items)} 格  样本: " +
                  ", ".join(f"r{r}:{a}→{b}" for r, a, b in items[:4]))
    else:
        print(f"  [{sn}] 无变化")
print(f"  合计变化格数 = {total_diff}")

print()
print("=" * 74)
print("② 逐格 diff：副本 vs 原件（月销售总表）")
print("=" * 74)
src2 = grid(MON_SRC); dst2 = grid(MON_DST)
total2 = 0
for sn in src2:
    if sn not in dst2:
        print(f"  ★ sheet {sn} 缺失！"); continue
    _, _, sc = src2[sn]; _, _, dc = dst2[sn]
    diffs = []
    for k, v in sc.items():
        v2 = dc.get(k)
        if isinstance(v, (int, float)) and isinstance(v2, (int, float)):
            if abs(float(v) - float(v2)) > 1e-9: diffs.append((k, v, v2))
        elif v != v2:
            diffs.append((k, v, v2))
    for k in dc:
        if k not in sc: diffs.append((k, "<新格>", dc[k]))
    total2 += len(diffs)
    if diffs:
        bycol = {}
        for (r, c), a, b in diffs: bycol.setdefault(c, []).append((r, a, b))
        print(f"  [{sn}] 变化 {len(diffs)} 格；涉及列: {[gcl(c) for c in sorted(bycol)]}")
        for c in sorted(bycol):
            items = bycol[c]
            print(f"     {gcl(c)}({c}): {len(items)} 格  样本: " + ", ".join(f"r{r}:{a}→{b}" for r, a, b in items[:4]))
    else:
        print(f"  [{sn}] 无变化")
print(f"  合计变化格数 = {total2}")

print()
print("=" * 74)
print("③ 副本内恒等 + 合计复算（数据行累加，绕开公式无缓存）")
print("=" * 74)
wb = load_workbook(SEA_DST, read_only=True, data_only=True)
ph_rows, ph_bd, ph_az, ph_ba, ph_bb = {}, {}, {}, {}, {}
for r in wb["Philippines 菲律宾销量表"].iter_rows(min_row=3, values_only=True):
    if r[0] is not None:
        k = str(r[0]).strip().upper(); ph_rows[k] = True
        ph_bd[k] = n0(r[55]); ph_az[k] = n0(r[51]); ph_ba[k] = n0(r[52]); ph_bb[k] = n0(r[53])
sea_ca, sea_ci = {}, {}
for r in wb["SEA Total 东南亚总销量表"].iter_rows(min_row=3, values_only=True):
    if r[0] is not None:
        k = str(r[0]).strip().upper(); sea_ca[k] = n0(r[78]); sea_ci[k] = n0(r[86])
wb.close()
bad1 = [k for k in ph_bd if k in sea_ca and abs(ph_bd[k] - sea_ca[k]) > 1e-9]
print(f"恒等1（菲律宾BD == SEA Total CA）：不一致 {len(bad1)} 行 {bad1[:5]}")

wb2 = load_workbook(MON_DST, read_only=True, data_only=True)
ms = wb2["2026年"]
m_em, m_hs, m_ka = {}, {}, {}
for r in ms.iter_rows(min_row=4, values_only=True):
    if r[0] is None: continue
    k = str(r[0]).strip().upper()
    m_em[k] = n0(r[142]); m_hs[k] = n0(r[226]); m_ka[k] = n0(r[286])
wb2.close()
bad2 = [k for k in ph_bd if k in m_em and abs(m_em[k] - ph_bd[k]) > 1e-9]
print(f"恒等A（月销EM == 菲律宾BD）：不一致 {len(bad2)} 行 {bad2[:5]}")
bad3 = [k for k in sea_ci if k in m_hs and abs(m_hs[k] - sea_ci[k]) > 1e-9]
print(f"恒等B（月销HS == SEA Total CI）：不一致 {len(bad3)} 行 {bad3[:5]}")

# 合计复算（数据行累加）
sum_az = sum(ph_az.values()); sum_ba = sum(ph_ba.values()); sum_bb = sum(ph_bb.values())
sum_bd = sum(ph_bd.values()); sum_ca = sum(sea_ca.values()); sum_ci = sum(sea_ci.values())
sum_em = sum(m_em.values()); sum_hs = sum(m_hs.values()); sum_ka = sum(m_ka.values())
print(f"\n菲律宾 sheet 数据行累加：AZ(跨境)={sum_az:.0f} BA(本土LZD)={sum_ba:.0f} BB?  —  见下")
print(f"  BD 合计列累加 = {sum_bd:.0f}（旧 3123，Δ={sum_bd-3123:+.0f}）")
print(f"  SEA Total CA 累加 = {sum_ca:.0f}  CI 累加 = {sum_ci:.0f}（旧 10948，Δ={sum_ci-10948:+.0f}）")
print(f"  月销总表 EM 累加 = {sum_em:.0f}  HS 累加 = {sum_hs:.0f}  KA 累加 = {sum_ka:.0f}")
# 行2 静态值核对
wb3 = load_workbook(SEA_DST, read_only=True, data_only=True)
r2 = [c.value for c in wb3["Philippines 菲律宾销量表"][2]]
print(f"\n菲律宾 sheet 行2：AY={r2[50]} AZ={r2[51]} BA={r2[52]} BB={r2[53]} BC={r2[54]} BD={r2[55]}")
r2s = [c.value for c in wb3["SEA Total 东南亚总销量表"][2]]
print(f"SEA Total 行2：CA(菲)={r2s[78]} CB(泰)={r2s[79]} CC(印尼)={r2s[80]} CI(合计)={r2s[86]}")
wb3.close()

# 抽查
print("\n抽查（新口径）：")
for k in ["G376", "G470", "G484", "G445", "G569", "G165", "G165PRO", "G175"]:
    if k in ph_bd:
        print(f"  {k:<9} 菲律宾BD={ph_bd[k]:>6.0f}  SEA_CA={sea_ca.get(k, 0):>6.0f}  月销EM={m_em.get(k, 0):>6.0f}  "
              f"月销HS={m_hs.get(k, 0):>6.0f}  月销KA={m_ka.get(k, 0):>6.0f}")
