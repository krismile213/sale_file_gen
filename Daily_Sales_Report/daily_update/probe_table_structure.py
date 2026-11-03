# -*- coding: utf-8 -*-
"""只读实测 v4：'00' 归一后重验恒等 + 探明数据行是公式还是静态值"""
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter as gcl

R = r"D:\Desktop\work-main\sale_file_gen\trunk\result\东南亚销量表\202609\2026年东南亚销量表Southeast Asia Sales table.xlsx"
M = r"D:\Desktop\work-main\sale_file_gen\trunk\result\数据归属周期\202609\月销售总表.xlsx"

def n0(v):
    if v is None: return 0
    if isinstance(v, str):
        s = v.strip()
        if s in ("", "00"): return 0
        try: return float(s)
        except Exception: return 0
    return float(v or 0)

wb = load_workbook(R, read_only=True, data_only=True)
ph_map, ca_map, ci_map = {}, {}, {}
for r in wb["Philippines 菲律宾销量表"].iter_rows(min_row=3, values_only=True):
    if r[0] is not None: ph_map[str(r[0]).strip().upper()] = r[55]
for r in wb["SEA Total 东南亚总销量表"].iter_rows(min_row=3, values_only=True):
    if r[0] is not None:
        k = str(r[0]).strip().upper(); ca_map[k] = r[78]; ci_map[k] = r[86]
wb.close()
badA = [k for k in ph_map if k in ca_map and n0(ph_map[k]) != n0(ca_map[k])]
print(f"恒等1（菲律宾BD == SEA Total CA，'00'归一后）：不一致 {len(badA)} 行", badA[:5])

GROUPS = [3, 15, 27, 39, 51, 63, 75, 87, 99, 111, 123, 135, 147, 159, 171, 183, 195, 207, 219]
wb2 = load_workbook(M, read_only=True, data_only=True)
ms = wb2["2026年"]
m_ph, m_sea, m_tot, m_grp9 = {}, {}, {}, {}
for r in ms.iter_rows(min_row=4, values_only=True):
    if r[0] is None: continue
    k = str(r[0]).strip().upper()
    m_ph[k] = r[142]; m_sea[k] = r[226]; m_tot[k] = r[286]
    m_grp9[k] = {c0: r[c0 + 7] for c0 in GROUPS}
wb2.close()
badB = [k for k in ph_map if k in m_ph and n0(m_ph[k]) != n0(ph_map[k])]
print(f"恒等A（月销菲律宾EM == 菲律宾sheet BD，'00'归一后）：不一致 {len(badB)} 行", badB[:8])
badC = [k for k in ci_map if k in m_sea and n0(m_sea[k]) != n0(ci_map[k])]
print(f"恒等B（月销东南亚HS == SEA Total CI，'00'归一后）：不一致 {len(badC)} 行", badC[:8])

# 恒等C：逐组打印 G445 的各国9月值，找出2倍来源
k = "G445"
print(f"\n{k} 各国组9月值：")
for c0 in GROUPS:
    print(f"   {gcl(c0)}({c0}): {m_grp9[k][c0]}")
print(f"   KA(287)总计 = {m_tot[k]}  各国之和 = {sum(n0(v) for v in m_grp9[k].values())}")

# 公式形态：G445 行的 EM / KA / HS 三格 + 行3合计行
wb3 = load_workbook(M, read_only=True, data_only=False)
ms3 = wb3["2026年"]
for r in ms3.iter_rows(min_row=4, max_row=4, values_only=True):
    if r[0] and str(r[0]).strip().upper() == "G445":
        print(f"\nG445 公式形态：EM(143)={r[142]!r}  HS(227)={r[226]!r}  KA(287)={r[286]!r}")
        print(f"   C(3)欧洲9月={r[10]!r}  EE(135)菲9月={r[142]!r}")
        break
# 行3（合计行）
for r in ms3.iter_rows(min_row=3, max_row=3, values_only=True):
    print(f"行3合计行：A={r[0]!r} C={r[2]!r} EM={r[142]!r} HS={r[226]!r} KA={r[286]!r}")
    break
# 行2（月标签行）
for r in ms3.iter_rows(min_row=2, max_row=2, values_only=True):
    print(f"行2：EE+8={r[142]!r} HK+8={r[226]!r} JS+8={r[286]!r}")
    break
wb3.close()
