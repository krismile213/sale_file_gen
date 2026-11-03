# -*- coding: utf-8 -*-
"""trunk 202609 成品对齐重算口径 3,485。v3：SKU 列修正为 A 列；副本有而 trunk 无的 SKU 行
（G569）在 trunk 末尾追加；其余同 v2（恒等式自证 + 幂等）。
"""
import os
from copy import copy
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

REPO = r'D:\Desktop\work-main\sale_file_gen'
TRUNK_MON = os.path.join(REPO, 'trunk', 'result', '数据归属周期', '202609', '月销售总表.xlsx')
TRUNK_SEA = os.path.join(REPO, 'trunk', 'result', '东南亚销量表', '202609', '2026年东南亚销量表Southeast Asia Sales table.xlsx')
TRUNK_PH = os.path.join(REPO, 'trunk', 'result', '东南亚销量表', '202609', 'Philippines_202609_sales.xlsx')
COPY_MON = os.path.join(REPO, '菲律宾重算_202609', '月销售总表.xlsx')
COPY_SEA = os.path.join(REPO, '菲律宾重算_202609', '2026年东南亚销量表Southeast Asia Sales table.xlsx')
COPY_PH = os.path.join(REPO, '菲律宾重算_202609', 'Philippines_202609_sales.xlsx')


def n0(v):
    if v is None:
        return 0
    if isinstance(v, str):
        s = v.strip()
        if s in ('00', ''):
            return 0
        try:
            return float(s)
        except ValueError:
            return 0
    return v if isinstance(v, (int, float)) else 0


def is_formula(v):
    return isinstance(v, str) and v.startswith('=')


def sku_index(ws, rmin, rmax, col=1):
    d = {}
    for r in range(rmin, rmax + 1):
        s = ws.cell(r, col).value
        if s:
            d[str(s).strip()] = r
    return d


def propagate(ws_t, ws_c, label, cols, rmin, rmax):
    """按 A 列 SKU 对齐传播静态值。返回 (格数, 净变化, 剩余差异, {仅副本有的SKU})。"""
    t_sku = sku_index(ws_t, rmin, rmax)
    c_sku = sku_index(ws_c, rmin, rmax)
    only_c = set(c_sku) - set(t_sku)
    row_map = {t_sku[s]: c_sku[s] for s in set(t_sku) & set(c_sku)}
    n, delta = 0, 0.0
    for c in cols:
        for rt in range(rmin, rmax + 1):
            rc = row_map.get(rt, rt)
            vt, vc = ws_t.cell(rt, c).value, ws_c.cell(rc, c).value
            if is_formula(vt) or is_formula(vc):
                continue
            if n0(vt) != n0(vc):
                delta += n0(vc) - n0(vt)
                ws_t.cell(rt, c).value = vc
                n += 1
    left = 0
    for c in cols:
        for rt in range(rmin, rmax + 1):
            rc = row_map.get(rt, rt)
            vt, vc = ws_t.cell(rt, c).value, ws_c.cell(rc, c).value
            if not is_formula(vt) and not is_formula(vc) and n0(vt) != n0(vc):
                left += 1
    print(f'  [{label}] 传播 {n} 格, 净变化 {delta:+g}, 剩余静态差异 {left}'
          + (f', 仅副本有 SKU: {sorted(only_c)}' if only_c else ''))
    return n, delta, left, only_c


def check(cond, msg):
    print(('  ✓ ' if cond else '  ✗ ') + msg)
    if not cond:
        raise AssertionError(msg)


def append_missing(ws_t, ws_c, only_c, cols, rmin, rmax, ref_row):
    """把副本有而 trunk 无的 SKU 行追加到 trunk 末尾（样式套参考行）。"""
    added = 0
    for s in sorted(only_c):
        rc = sku_index(ws_c, rmin, rmax)[s]
        rt = ws_t.max_row + 1
        for c in range(1, ws_c.max_column + 1):
            tc, sc = ws_t.cell(rt, c), ws_c.cell(rc, c)
            tc._style = copy(ws_t.cell(ref_row, c)._style)
            if not is_formula(sc.value):
                tc.value = sc.value
        # 参考行的月份块公式/其他列不动，只补 SKU/Model + 目标列
        print(f'    追加 {s} → trunk 行{rt}（样式套行{ref_row}）')
        added += 1
    return added


print('========== 1. 月销售总表 ==========')
wt = load_workbook(TRUNK_MON, data_only=False)
wc = load_workbook(COPY_MON, data_only=True)
m26_t, m26_c = wt['2026年'], wc['2026年']
maxr = m26_t.max_row
n1, d1, l1, oc1 = propagate(m26_t, m26_c, '月销[2026年] SKU行', [143, 227, 287, 296], 4, maxr)
check(l1 == 0 and not oc1, '月销[2026年] 传播后无剩余差异、无缺行')
em = sum(n0(m26_t.cell(r, 143).value) for r in range(4, maxr + 1))
hs = sum(n0(m26_t.cell(r, 227).value) for r in range(4, maxr + 1))
ka = sum(n0(m26_t.cell(r, 287).value) for r in range(4, maxr + 1))
check(abs(em - 3485) < 1e-9, f'月销 EM 列和 = {em:g}（应 3485）')
check(abs(hs - 11310) < 1e-9, f'月销 HS 列和 = {hs:g}（应 11310）')
check(abs(ka - 63199) < 1e-9, f'月销 KA 列和 = {ka:g}（应 63199）')
yr_t = next(wt[s] for s in wt.sheetnames if '销量情况' in s)
yr_c = next(wc[s] for s in wc.sheetnames if '销量情况' in s)
n2, d2, l2, oc2 = propagate(yr_t, yr_c, '月销[年度]', list(range(1, yr_t.max_column + 1)), 1, yr_t.max_row)
check(l2 == 0, '月销[年度] 传播后无剩余静态差异')
check(abs(n0(yr_t['G10'].value) - 63199) < 1e-9, f'年度 G10 = {n0(yr_t["G10"].value):g}（应 63199）')
check(abs(n0(yr_t['AA10'].value) - 11310) < 1e-9, f'年度 AA10 = {n0(yr_t["AA10"].value):g}（应 11310）')
wt.save(TRUNK_MON)
print('  已保存', TRUNK_MON)
wt.close(); wc.close()

print()
print('========== 2. 东南亚销量表 ==========')
st = load_workbook(TRUNK_SEA, data_only=False)
sc = load_workbook(COPY_SEA, data_only=True)
ph_t, ph_c = st['Philippines 菲律宾销量表'], sc['Philippines 菲律宾销量表']
n3, d3, l3, oc3 = propagate(ph_t, ph_c, 'SEA[Philippines] 9月块', list(range(51, 57)), 3, ph_t.max_row)
if oc3:
    print('  补行：')
    append_missing(ph_t, ph_c, oc3, list(range(51, 57)), 3, ph_c.max_row, ref_row=492)
# 重算补行后的剩余差异（补的行也应无差异）
l3b = sum(1 for c in range(51, 57) for r in range(3, ph_t.max_row + 1)
          if not is_formula(ph_t.cell(r, c).value) and not is_formula(ph_c.cell(r, c).value)
          and n0(ph_t.cell(r, c).value) != n0(ph_c.cell(r, c).value))
check(l3b == 0, 'SEA[Philippines] 补行后无剩余静态差异')
check(abs(n0(ph_t['BD2'].value) - 3485) < 1e-9, f'Philippines BD2 = {n0(ph_t["BD2"].value):g}（应 3485）')
for c, name in {51: '跨境Lazada', 52: '本土Lazada', 53: 'Shopee', 54: 'TikTok', 55: '线下'}.items():
    s = sum(n0(ph_t.cell(r, c).value) for r in range(3, ph_t.max_row + 1))
    check(abs(s - n0(ph_t.cell(2, c).value)) < 1e-6, f'Philippines {name} 数据行和 {s:g} = 行2 {n0(ph_t.cell(2, c).value):g}')
bd_sum = sum(n0(ph_t.cell(r, 56).value) for r in range(3, ph_t.max_row + 1))
check(abs(bd_sum - 3485) < 1e-9, f'Philippines BD 数据行和 {bd_sum:g} = 3485')
sea_t, sea_c = st['SEA Total 东南亚总销量表'], sc['SEA Total 东南亚总销量表']
n4, d4, l4, oc4 = propagate(sea_t, sea_c, 'SEA[SEA Total] 9月块', list(range(87, 96)), 3, sea_t.max_row)
check(l4 == 0 and not oc4, 'SEA[SEA Total] 传播后无剩余差异、无缺行')
check(abs(n0(sea_t['CI2'].value) - 11310) < 1e-9, f'SEA Total CI2 = {n0(sea_t["CI2"].value):g}（应 11310）')
check(abs(n0(sea_t['CJ2'].value) - 3485) < 1e-9, f'SEA Total CJ2(菲) = {n0(sea_t["CJ2"].value):g}（应 3485）')
nat_sum = sum(n0(sea_t.cell(2, 88 + i).value) for i in range(8))
check(abs(nat_sum - n0(sea_t['CI2'].value)) < 1e-6, f'SEA Total 8 国列和 {nat_sum:g} = 合计 {n0(sea_t["CI2"].value):g}')
yr2_t = next(st[s] for s in st.sheetnames if '销量情况' in s)
yr2_c = next(sc[s] for s in sc.sheetnames if '销量情况' in s)
n5, d5, l5, oc5 = propagate(yr2_t, yr2_c, 'SEA[年度] AU列', [47], 2, 10)
check(l5 == 0, 'SEA[年度] AU 列传播后无剩余差异')
check(abs(n0(yr2_t['AU2'].value) - 11310) < 1e-9, f'SEA 年度 AU2 = {n0(yr2_t["AU2"].value):g}（应 11310）')
check(abs(n0(yr2_t['AU3'].value) - 3485) < 1e-9, f'SEA 年度 AU3(菲) = {n0(yr2_t["AU3"].value):g}（应 3485）')
check(all(n0(yr2_t.cell(2, c).value) > 0 for c in range(39, 47)), 'SEA 年度 AM~AT（1~8月历史列）保持非零未动')
st.save(TRUNK_SEA)
print('  已保存', TRUNK_SEA)
st.close(); sc.close()

print()
print('========== 3. 菲律宾次表 ==========')
pt = load_workbook(TRUNK_PH, data_only=False)
pc = load_workbook(COPY_PH, data_only=True)
p1_t, p1_c = pt['Sheet1'], pc['Sheet1']
n6, d6, l6, oc6 = propagate(p1_t, p1_c, '菲次表[Sheet1] 渠道+Total', list(range(5, 11)), 2, p1_t.max_row)
if oc6:
    print(f'  ⚠️ 菲次表仅副本有 SKU: {sorted(oc6)}（需人工确认）')
    raise AssertionError('菲次表存在缺行，终止')
check(l6 == 0, '菲次表 Sheet1 传播后无剩余静态差异')
nk, dk = 0, 0.0
for r in range(2, p1_t.max_row + 1):
    if is_formula(p1_t.cell(r, 11).value):
        continue
    kt = n0(p1_t.cell(r, 11).value)
    jt = n0(p1_t.cell(r, 10).value)
    jc = n0(p1_c.cell(r, 10).value)
    if jc != jt:
        p1_t.cell(r, 11).value = kt + (jc - jt)
        nk += 1
        dk += jc - jt
print(f'  [菲次表 K列累计] 差值传播 {nk} 行, 净变化 {dk:+g}')
check(abs(dk - 156) < 1e-9, f'K 列净变化 {dk:+g}（应 +156）')
jt_sum = sum(n0(p1_t.cell(r, 10).value) for r in range(2, p1_t.max_row + 1))
check(abs(jt_sum - 3485) < 1e-9, f'菲次表 J 列和 = {jt_sum:g}（应 3485）')
pt.save(TRUNK_PH)
print('  已保存', TRUNK_PH)
pt.close(); pc.close()

print()
print('========== 完成，全部自证通过 ==========')
print(f'月销 {n1+n2} 格｜东南亚 {n3+n4+n5} 格（+补行 {len(oc3)}）｜菲次表 {n6} 格 + K 列 {nk} 行')
