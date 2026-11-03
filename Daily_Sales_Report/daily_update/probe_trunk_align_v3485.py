# -*- coding: utf-8 -*-
"""把 trunk 202609 成品对齐到重算口径 3,485（菲律宾）/11,310（东南亚）/63,199（总计）。
精准传播（不整表覆盖），写前 assert 防呆、写后逐项复核。
源 = 菲律宾重算_202609/ 副本（78 项终验全绿版）；目标 = trunk/result。
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


log = []


def propagate(ws_t, ws_c, label, cols, rmin, rmax, expect_delta=None, sku_col=None):
    """把 ws_c 的静态值搬到 ws_t。cols=列号列表；行范围 rmin~rmax。
    expect_delta: 该区域净变化期望值（自证用）。
    sku_col: 若给列号，按 SKU 对齐（trunk 行 ↔ 副本行由 A 列 SKU 匹配）。"""
    # SKU 行映射
    row_map = None
    if sku_col:
        t_sku = {}
        for r in range(rmin, rmax + 1):
            s = ws_t.cell(r, sku_col).value
            if s:
                t_sku[str(s).strip()] = r
        c_sku = {}
        for r in range(rmin, rmax + 1):
            s = ws_c.cell(r, sku_col).value
            if s:
                c_sku[str(s).strip()] = r
        common = set(t_sku) & set(c_sku)
        row_map = {t_sku[s]: c_sku[s] for s in common}
        only_t = set(t_sku) - set(c_sku)
        only_c = set(c_sku) - set(t_sku)
        if only_t or only_c:
            print(f'  [{label}] ⚠️ SKU 不完全对齐: 仅trunk {len(only_t)} 个, 仅副本 {len(only_c)} 个')
    n = 0
    delta = 0.0
    for c in cols:
        for rt in range(rmin, rmax + 1):
            rc = row_map.get(rt, rt) if row_map else rt
            vt = ws_t.cell(rt, c).value
            vc = ws_c.cell(rc, c).value
            if is_formula(vt) or is_formula(vc):
                continue  # 公式格不碰
            if n0(vt) != n0(vc):
                old = n0(vt)
                ws_t.cell(rt, c).value = vc
                delta += n0(vc) - old
                n += 1
                if n <= 6:
                    log.append(f'  {label} {get_column_letter(c)}{rt}: {old:g} → {n0(vc):g}')
    print(f'  [{label}] 传播 {n} 格, 净变化 {delta:+g}')
    if expect_delta is not None and abs(delta - expect_delta) > 1e-9:
        raise AssertionError(f'{label} 净变化 {delta:+g} ≠ 期望 {expect_delta:+g}')
    return n, delta


print('========== 1. 月销售总表 ==========')
wt = load_workbook(TRUNK_MON, data_only=False)   # 保留公式
wc = load_workbook(COPY_MON, data_only=True)     # 读值
# 1a. [2026年] sheet：SKU 在 A 列，EM(143) 菲 9月 / HS(227) 东南亚 / KA(287) 总计 / KJ(296) 当年总销量
m26_t, m26_c = wt['2026年'], wc['2026年']
maxr = m26_t.max_row
n1, d1 = propagate(m26_t, m26_c, '月销[2026年] SKU行', [143, 227, 287, 296], 4, maxr,
                   expect_delta=156 * 4, sku_col=1)
# 1b. [2021-2026年销量情况] sheet：全部静态差异（diff 实测仅 12 格，都是预期传播点）
yr_t, yr_c = None, None
for sn in wt.sheetnames:
    if '销量情况' in sn:
        yr_t = wt[sn]
        break
for sn in wc.sheetnames:
    if '销量情况' in sn:
        yr_c = wc[sn]
        break
assert yr_t is not None and yr_c is not None, '找不到年度 sheet'
n2, d2 = propagate(yr_t, yr_c, '月销[年度]', list(range(1, yr_t.max_column + 1)), 1, yr_t.max_row)
print(f'    （年度 sheet 12 格: G10/AA10/G14/J14/AA14/AD14 各 +156 + 6 个比率格）')
wt.save(TRUNK_MON)
print('  已保存', TRUNK_MON)
wt.close(); wc.close()

print()
print('========== 2. 东南亚销量表 ==========')
st = load_workbook(TRUNK_SEA, data_only=False)
sc = load_workbook(COPY_SEA, data_only=True)
# 2a. Philippines sheet：9月块 AY(51)~BD(56)，行2 合计 + 数据行
ph_t = st['Philippines 菲律宾销量表']
ph_c = sc['Philippines 菲律宾销量表']
n3, d3 = propagate(ph_t, ph_c, 'SEA[Philippines] 9月块', list(range(51, 57)), 2, ph_t.max_row,
                   expect_delta=362 * 6, sku_col=2)
# 2b. SEA Total sheet：9月块 CI(87)~CQ(95)
sea_t = st['SEA Total 东南亚总销量表']
sea_c = sc['SEA Total 东南亚总销量表']
n4, d4 = propagate(sea_t, sea_c, 'SEA[SEA Total] 9月块', list(range(87, 96)), 2, sea_t.max_row,
                   expect_delta=362 * 9, sku_col=2)
# 2c. 年度 sheet：只 AU(47) 列（diff 实测仅 AU2/AU3 两格有差异；AM~AT 历史列保留 trunk）
yr2_t = None
for sn in st.sheetnames:
    if '销量情况' in sn:
        yr2_t = st[sn]
        break
yr2_c = None
for sn in sc.sheetnames:
    if '销量情况' in sn:
        yr2_c = sc[sn]
        break
n5, d5 = propagate(yr2_t, yr2_c, 'SEA[年度] AU列', [47], 2, 10, expect_delta=724)
print('    （AU2 total 10948→11310、AU3 菲律宾 3123→3485；AM~AT 历史列保留 trunk 未动）')
st.save(TRUNK_SEA)
print('  已保存', TRUNK_SEA)
st.close(); sc.close()

print()
print('========== 3. 菲律宾次表 ==========')
pt = load_workbook(TRUNK_PH, data_only=False)
pc = load_workbook(COPY_PH, data_only=True)
p1_t, p1_c = pt['Sheet1'], pc['Sheet1']
# 3a. 渠道列 E(5)~J(10)（5 渠道 + Total）按 SKU 行传播
n6, d6 = propagate(p1_t, p1_c, '菲次表[Sheet1] 渠道+Total', list(range(5, 11)), 2, p1_t.max_row,
                   expect_delta=156, sku_col=1)
# 3b. K(11) 列累计：差值传播 K_new = K_old + (J_new - J_old)
nk = 0
dK = 0.0
for r in range(2, p1_t.max_row + 1):
    kt = p1_t.cell(r, 11).value
    if is_formula(kt):
        continue
    kt_v = n0(kt)
    kc_v = n0(p1_c.cell(r, 11).value)
    jt_v = n0(p1_t.cell(r, 10).value)
    jc_v = n0(p1_c.cell(r, 10).value)
    dj = jc_v - jt_v
    if dj != 0:
        new_k = kt_v + dj
        p1_t.cell(r, 11).value = new_k
        nk += 1
        dK += dj
        if nk <= 4:
            log.append(f'  菲次表 K{r}: {kt_v:g} → {new_k:g}（ΔJ={dj:+g}）')
print(f'  [菲次表 K列累计] 差值传播 {nk} 行, 净变化 {dK:+g}')
assert abs(dK - 156) < 1e-9, f'K 列净变化 {dK:+g} ≠ +156'
pt.save(TRUNK_PH)
print('  已保存', TRUNK_PH)
pt.close(); pc.close()

print()
print('========== 传播明细（前 10 条） ==========')
for l in log[:10]:
    print(l)
print()
print(f'合计传播: 月销 {n1+n2} 格｜东南亚 {n3+n4+n5} 格｜菲次表 {n6} 格 + K 列 {nk} 行')
