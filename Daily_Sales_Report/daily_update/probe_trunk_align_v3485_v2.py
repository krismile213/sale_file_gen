# -*- coding: utf-8 -*-
"""把 trunk 202609 成品对齐到重算口径 3,485（菲律宾）/11,310（东南亚）/63,199（总计）。
v2：期望值改为「传播后恒等式自证」；幂等（重跑 0 传播时校验已对齐即通过）。
源 = 菲律宾重算_202609/ 副本；目标 = trunk/result。精准传播，不整表覆盖。
"""
import os
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


def propagate(ws_t, ws_c, label, cols, rmin, rmax, sku_col=None):
    """把 ws_c 的静态值搬到 ws_t（按 SKU 行对齐）。返回 (传播格数, 净变化, 剩余差异数)。"""
    row_map = None
    if sku_col:
        t_sku, c_sku = {}, {}
        for r in range(rmin, rmax + 1):
            s = ws_t.cell(r, sku_col).value
            if s:
                t_sku[str(s).strip()] = r
            s = ws_c.cell(r, sku_col).value
            if s:
                c_sku[str(s).strip()] = r
        row_map = {t_sku[s]: c_sku[s] for s in set(t_sku) & set(c_sku)}
        if set(t_sku) != set(c_sku):
            print(f'  [{label}] ⚠️ SKU 集合不一致: 仅trunk {len(set(t_sku)-set(c_sku))}, 仅副本 {len(set(c_sku)-set(t_sku))}')
    n, delta = 0, 0.0
    for c in cols:
        for rt in range(rmin, rmax + 1):
            rc = row_map.get(rt, rt) if row_map else rt
            vt, vc = ws_t.cell(rt, c).value, ws_c.cell(rc, c).value
            if is_formula(vt) or is_formula(vc):
                continue
            if n0(vt) != n0(vc):
                delta += n0(vc) - n0(vt)
                ws_t.cell(rt, c).value = vc
                n += 1
    # 剩余差异（应只剩公式格）
    left = 0
    for c in cols:
        for rt in range(rmin, rmax + 1):
            rc = row_map.get(rt, rt) if row_map else rt
            vt, vc = ws_t.cell(rt, c).value, ws_c.cell(rc, c).value
            if not is_formula(vt) and not is_formula(vc) and n0(vt) != n0(vc):
                left += 1
    print(f'  [{label}] 传播 {n} 格, 净变化 {delta:+g}, 剩余静态差异 {left}')
    return n, delta, left


def check(cond, msg):
    print(('  ✓ ' if cond else '  ✗ ') + msg)
    if not cond:
        raise AssertionError(msg)


print('========== 1. 月销售总表 ==========')
wt = load_workbook(TRUNK_MON, data_only=False)
wc = load_workbook(COPY_MON, data_only=True)
m26_t, m26_c = wt['2026年'], wc['2026年']
maxr = m26_t.max_row
n1, d1, l1 = propagate(m26_t, m26_c, '月销[2026年] SKU行', [143, 227, 287, 296], 4, maxr, sku_col=1)
check(l1 == 0, '月销[2026年] 传播后无剩余静态差异')
# 自证：EM 列数据行和 = 3485、HS = 11310、KA = 63199
em = sum(n0(m26_t.cell(r, 143).value) for r in range(4, maxr + 1))
hs = sum(n0(m26_t.cell(r, 227).value) for r in range(4, maxr + 1))
ka = sum(n0(m26_t.cell(r, 287).value) for r in range(4, maxr + 1))
check(abs(em - 3485) < 1e-9, f'月销 EM 列和 = {em:g}（应 3485）')
check(abs(hs - 11310) < 1e-9, f'月销 HS 列和 = {hs:g}（应 11310）')
check(abs(ka - 63199) < 1e-9, f'月销 KA 列和 = {ka:g}（应 63199）')
yr_t = next(wt[s] for s in wt.sheetnames if '销量情况' in s)
yr_c = next(wc[s] for s in wc.sheetnames if '销量情况' in s)
n2, d2, l2 = propagate(yr_t, yr_c, '月销[年度]', list(range(1, yr_t.max_column + 1)), 1, yr_t.max_row)
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
n3, d3, l3 = propagate(ph_t, ph_c, 'SEA[Philippines] 9月块', list(range(51, 57)), 2, ph_t.max_row, sku_col=2)
check(l3 == 0, 'SEA[Philippines] 传播后无剩余静态差异')
# 自证：BD2 = 3485；行2 = 数据行之和；各渠道列和
check(abs(n0(ph_t['BD2'].value) - 3485) < 1e-9, f'Philippines BD2 = {n0(ph_t["BD2"].value):g}（应 3485）')
ch = {51: '跨境Lazada', 52: '本土Lazada', 53: 'Shopee', 54: 'TikTok', 55: '线下'}
for c, name in ch.items():
    s = sum(n0(ph_t.cell(r, c).value) for r in range(3, ph_t.max_row + 1))
    check(abs(s - n0(ph_t.cell(2, c).value)) < 1e-6, f'Philippines {name} 数据行和 {s:g} = 行2 {n0(ph_t.cell(2, c).value):g}')
sea_t, sea_c = st['SEA Total 东南亚总销量表'], sc['SEA Total 东南亚总销量表']
n4, d4, l4 = propagate(sea_t, sea_c, 'SEA[SEA Total] 9月块', list(range(87, 96)), 2, sea_t.max_row, sku_col=2)
check(l4 == 0, 'SEA[SEA Total] 传播后无剩余静态差异')
check(abs(n0(sea_t['CI2'].value) - 11310) < 1e-9, f'SEA Total CI2 = {n0(sea_t["CI2"].value):g}（应 11310）')
check(abs(n0(sea_t['CJ2'].value) - 3485) < 1e-9, f'SEA Total CJ2(菲) = {n0(sea_t["CJ2"].value):g}（应 3485）')
# 国家列之和 = 合计列
nat_sum = sum(n0(sea_t.cell(2, 88 + i).value) for i in range(8))
check(abs(nat_sum - n0(sea_t['CI2'].value)) < 1e-6, f'SEA Total 8 国列和 {nat_sum:g} = 合计 {n0(sea_t["CI2"].value):g}')
yr2_t = next(st[s] for s in st.sheetnames if '销量情况' in s)
yr2_c = next(sc[s] for s in sc.sheetnames if '销量情况' in s)
n5, d5, l5 = propagate(yr2_t, yr2_c, 'SEA[年度] AU列', [47], 2, 10)
check(l5 == 0, 'SEA[年度] AU 列传播后无剩余差异')
check(abs(n0(yr2_t['AU2'].value) - 11310) < 1e-9, f'SEA 年度 AU2 = {n0(yr2_t["AU2"].value):g}（应 11310）')
check(abs(n0(yr2_t['AU3'].value) - 3485) < 1e-9, f'SEA 年度 AU3(菲) = {n0(yr2_t["AU3"].value):g}（应 3485）')
# AM~AT 历史列必须保持 trunk 原值（不被副本的 0 覆盖）
hist_ok = all(n0(yr2_t.cell(2, c).value) > 0 for c in range(39, 47))
check(hist_ok, 'SEA 年度 AM~AT（1~8月历史列）保持非零未动')
st.save(TRUNK_SEA)
print('  已保存', TRUNK_SEA)
st.close(); sc.close()

print()
print('========== 3. 菲律宾次表 ==========')
pt = load_workbook(TRUNK_PH, data_only=False)
pc = load_workbook(COPY_PH, data_only=True)
p1_t, p1_c = pt['Sheet1'], pc['Sheet1']
n6, d6, l6 = propagate(p1_t, p1_c, '菲次表[Sheet1] 渠道+Total', list(range(5, 11)), 2, p1_t.max_row, sku_col=1)
check(l6 == 0, '菲次表 Sheet1 传播后无剩余静态差异')
# K 列累计差值传播：K_new = K_old + (J_new - J_old)
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
# 自证：J 列和 = 3485；K 列和 = 旧 K 和 + 156
jt_sum = sum(n0(p1_t.cell(r, 10).value) for r in range(2, p1_t.max_row + 1))
check(abs(jt_sum - 3485) < 1e-9, f'菲次表 J 列和 = {jt_sum:g}（应 3485）')
kt_sum = sum(n0(p1_t.cell(r, 11).value) for r in range(2, p1_t.max_row + 1))
print(f'  ℹ 菲次表 K 列和 = {kt_sum:g}')
pt.save(TRUNK_PH)
print('  已保存', TRUNK_PH)
pt.close(); pc.close()

print()
print('========== 完成 ==========')
print(f'月销 {n1+n2} 格｜东南亚 {n3+n4+n5} 格｜菲次表 {n6} 格 + K 列 {nk} 行，全部自证通过')
