# -*- coding: utf-8 -*-
"""sale_file_gen 销量双重验证（202608）
第一重：代码/数据层 —— 原始订单 → 中间长表 → 按脚本口径复算
第二重：文件层   —— 直接读成品 xlsx 单元格（公式缓存值 + openpyxl 独立求和）
输出：逐分组对撞表 + 恒等式检查 + 已知偏差定位
"""
import os
import math
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter as gcl

TRUNK = r'D:\Desktop\work-main\sale_file_gen\trunk'
BELONGING = '202609'          # ★ 换月只改这一处：归属周期，格式 YYYYMM
M = int(BELONGING[4:6])       # 归属月（列号 = 3+12k+7 自动推导）
TMP = os.path.join(TRUNK, 'sale_file_tmp')
RES = os.path.join(TRUNK, 'result', '数据归属周期', BELONGING, '月销售总表.xlsx')
SR = 4
S_MAIN, S_EUR, S_Y = f'{BELONGING[:4]}年', f'{BELONGING[:4]}年-欧洲各国每月销售情况', '2021-2026年销量情况'
YEAR = int(BELONGING[:4])


def n(v):
    if v is None:
        return 0.0
    if isinstance(v, float) and math.isnan(v):
        return 0.0
    s = str(v).strip()
    if s in ('', 'nan', 'NaN', 'None'):
        return 0.0
    try:
        return float(s)
    except ValueError:
        return 0.0


# ---------- 载入 ----------
df = pd.read_csv(os.path.join(TMP, 'total_sku_sale_df.csv'), dtype=str)
COL = [c for c in df.columns if '销量' in c][0]
df['_v'] = pd.to_numeric(df[COL], errors='coerce').fillna(0)
df['站点'] = df['站点'].fillna('')
df['平台'] = df['平台'].fillna('')
df['仓库'] = df['仓库'].fillna('')
df['实体'] = df['实体'].fillna('')

wbv = load_workbook(RES, data_only=True)
wbf = load_workbook(RES)
wsv, wsf = wbv[S_MAIN], wbf[S_MAIN]

LAST = 0
for r in range(SR, wsv.max_row + 1):
    if wsv.cell(r, 1).value not in (None, ''):
        LAST = r

# ---------- 分组定义：列块 + 口径（复刻脚本过滤条件） ----------
AMZ = {'欧洲': '欧洲', '英国': '英国', '美国': '美国', '加拿大': '加拿大', '日本': '日本',
       '澳洲': '澳大利亚', '阿联酋': '阿拉伯联合酋长国', '沙特': '沙特阿拉伯',
       '墨西哥': '墨西哥', '印度': '印度'}
SE_LOCAL = ['菲律宾', '泰国', '印尼']   # 注意：sheet 里的顺序是 菲律宾/泰国/印尼（脚本 month_sale_group 第 1111 行）
SE_SINGLE = {'新加坡': ('新加坡', None), '马来': ('马来西亚', 'Lazada'), '越南': ('越南', 'Lazada')}
CROSS_SITES = ['马来西亚', '越南']   # ★ 与主脚本 sum_shopee_cross_site 的 other_shopee_cross_site 一致（不含新加坡）

GROUPS = []
c0 = 3
for name, inv in AMZ.items():
    GROUPS.append((name, c0, lambda d, inv=inv: d[d['仓库'] == inv]))
    c0 += 12
GROUPS.append(('亚马逊', c0, lambda d: d[d['平台'] == '亚马逊'])); c0 += 12
for s in SE_LOCAL:
    GROUPS.append((s, c0, lambda d, s=s: d[d['站点'] == s])); c0 += 12
for name, (site, plat) in SE_SINGLE.items():
    if plat:
        GROUPS.append((name, c0, lambda d, site=site, plat=plat: d[(d['站点'] == site) & (d['平台'] == plat)]))
    else:
        GROUPS.append((name, c0, lambda d, site=site: d[d['站点'] == site]))
    c0 += 12
GROUPS.append(('虾皮境内', c0, lambda d: d[(d['平台'] == 'Shopee') & (d['站点'].isin(CROSS_SITES)) & (d['实体'] == '跨境')])); c0 += 12
GROUPS.append(('东南亚', c0, lambda d: d[d['站点'].isin(['菲律宾', '印尼', '泰国', '马来西亚', '新加坡', '越南'])])); c0 += 12
GROUPS.append(('京东', c0, lambda d: d[(d['平台'] == '京东') & (d['站点'] == '国内')])); c0 += 12
GROUPS.append(('eMAG', c0, lambda d: d[d['平台'] == 'EMAG'])); c0 += 12
GROUPS.append(('TEMU平台', c0, lambda d: d[d['平台'] == 'Temu'])); c0 += 12
GROUPS.append(('美客多', c0, lambda d: d[d['平台'] == '美客多'])); c0 += 12
GROUPS.append(('总计', c0, lambda d: d)); c0 += 12

print('=' * 100)
print('【第一重 · 代码/数据层】原始订单 → 中间长表 → 按脚本口径复算')
print('=' * 100)
amz = pd.read_csv(os.path.join(TMP, 'total_last30d_amazon_order_df.csv'), dtype=str, low_memory=False)
q = pd.to_numeric(amz['quantity'], errors='coerce')
print(f'亚马逊原始订单：{len(amz)} 行，quantity 合计 = {q.sum():.0f}')
print(f'  order-status 取值 = {amz["order-status"].value_counts().to_dict()}  （Cancelled 已在代码 193 行剔除）')
utc = pd.to_datetime(amz['purchase-date'], errors='coerce', utc=True, format='mixed')
loc = pd.to_datetime(amz['purchase-date-local'], errors='coerce', utc=True, format='mixed')
print(f'  purchase-date(UTC)   范围 {utc.min()} ~ {utc.max()}')
print(f'  purchase-date-local  范围 {loc.min()} ~ {loc.max()}')
print(f'  归一到本地日：{loc.dt.normalize().min().date()} ~ {loc.dt.normalize().max().date()}'
      f'（其中 {(loc.dt.normalize() < pd.Timestamp(f"{YEAR}-{M:02d}-01", tz="UTC")).sum()} 行本地落在上月'
      f'，代码不做二次日期过滤 → 计入 8 月）')
print(f'total_sku_sale_df：{df.shape}，{COL} 合计 = {df["_v"].sum():.0f}')

print()
print('=' * 100)
print('【第二重 · 文件层】成品单元格直读')
print('=' * 100)
print(f'{os.path.basename(RES)}  {S_MAIN}：数据行 {SR}~{LAST}（{LAST - SR + 1} 行），max_col={wsv.max_column}')

print()
print(f'【1】逐分组 {M} 月 · 四路对撞')
hdr = f'{"分组":<8}{"列":>5}{"①文件缓存值":>13}{"②文件独立求和":>14}{"③口径复算":>12}{"④CSV列和":>11}   结论'
print(hdr)
print('-' * len(hdr))
csvp = pd.read_csv(os.path.join(TMP, 'processed_month_sale_df.csv'), dtype=str)
res_tab = {}
NO_CACHE = False
for name, c0, cond in GROUPS:
    cm = c0 + (M - 1)                       # 月块首 = 1 月列；第 M 月 = 块首 + (M-1)
    v1 = n(wsv.cell(3, cm).value)           # 行3 合计（公式；脚本产物无缓存 → 通常为 0）
    v2 = sum(n(wsv.cell(r, cm).value) for r in range(SR, LAST + 1))
    v3 = cond(df)['_v'].sum()
    v4 = sum(n(x) for x in csvp.iloc[:, cm].tolist())
    if v1 == 0 and v2 != 0:
        NO_CACHE = True
    ok = (abs(v2 - v3) < 0.5) and (abs(v3 - v4) < 0.5) and (v1 == 0 or abs(v1 - v3) < 0.5)
    res_tab[name] = (cm, v1, v2, v3, v4, ok)
    print(f'{name:<8}{cm:>5}{v1:>13.0f}{v2:>14.0f}{v3:>12.0f}{v4:>11.0f}   {"一致" if ok else "★不一致"}')
if NO_CACHE:
    print('  注：①行3 为 0 —— 脚本产物公式无缓存值属正常；判据改用 ②文件独立求和 = ③口径复算 = ④CSV列和。')
print(f'  注：文件【{S_MAIN}】行 3 的合计公式（{gcl(res_tab["总计"][0])} 列）= {wsf.cell(3, res_tab["总计"][0]).value}（范围远超数据行 → 新增行自动纳入）')

print()
print('【2】恒等式检查（取 ②文件独立求和 = 各行数据累加）')
amz_g = [g for g in AMZ]
s10 = sum(res_tab[g][2] for g in amz_g)
sea7 = sum(res_tab[g][2] for g in ['菲律宾', '泰国', '印尼', '新加坡', '马来', '越南', '虾皮境内'])
tops = amz_g + ['菲律宾', '泰国', '印尼', '新加坡', '马来', '越南', '虾皮境内',
                '京东', 'eMAG', 'TEMU平台', '美客多']
s_top = sum(res_tab[g][2] for g in tops)
print(f'  ● 10 国和 {s10:.0f}  vs  亚马逊合计 {res_tab["亚马逊"][2]:.0f}   差 = {res_tab["亚马逊"][2] - s10:.0f}   {"✔" if abs(res_tab["亚马逊"][2]-s10)<0.5 else "✘"}')
print(f'  ● 东南亚 7 子列和 {sea7:.0f}  vs  东南亚合计 {res_tab["东南亚"][2]:.0f}   差 = {res_tab["东南亚"][2] - sea7:.0f}   {"✔" if abs(res_tab["东南亚"][2]-sea7)<0.5 else "✘  ← 已知重叠"}')
print(f'  ● 21 顶层分组和 {s_top:.0f}  vs  总计 {res_tab["总计"][2]:.0f}   差 = {res_tab["总计"][2] - s_top:.0f}   {"✔" if abs(res_tab["总计"][2]-s_top)<0.5 else "✘  ← 同一重叠"}')
print(f'  ● 长表全量 {df["_v"].sum():.0f}  vs  总计 {res_tab["总计"][2]:.0f}   差 = {res_tab["总计"][2] - df["_v"].sum():.0f}   {"✔" if abs(res_tab["总计"][2]-df["_v"].sum())<0.5 else "✘"}')

print()
print('【3】「虾皮境内」与「新加坡」口径去重核验')
sg = df[(df['站点'] == '新加坡') & (df['实体'] == '跨境') & (df['平台'] == 'Shopee')]
sg_shopee = sg['_v'].sum()
print(f'  「新加坡」列口径 = 站点==新加坡 全平台（含 Shopee 跨境） → {res_tab["新加坡"][2]:.0f}')
print(f'     其中 Shopee·跨境（不再计入「虾皮境内」）              → {sg_shopee:.0f}')
print(f'  「虾皮境内」列口径 = Shopee & 站点∈{{马来西亚,越南}} & 实体==跨境 → {res_tab["虾皮境内"][2]:.0f}')
print(f'  → 7 子列和 -「东南亚合计」= {sea7 - res_tab["东南亚"][2]:.0f}（0 = 无重复计数）')

print()
print('【4】' + S_EUR)
wev, wef = wbv[S_EUR], wbf[S_EUR]
ce0 = 3 + (M - 1) * 11
CTRY = ['比利时', '土耳其', '德国', '西班牙', '法国', '意大利', '荷兰', '波兰', '瑞典', '爱尔兰']
le = 0
for r in range(SR, wev.max_row + 1):
    if wev.cell(r, 1).value not in (None, ''):
        le = r
s11 = 0
for k, nm in enumerate(CTRY):
    # 行3 为 =SUM 公式（无缓存 → 0）；改为对数据行累加，取块内 10 个国家列
    v = sum(n(wev.cell(r, ce0 + k).value) for r in range(SR, le + 1))
    s11 += v
    print(f'     {nm:<5} {v:>8.0f}  ({wev.cell(1, ce0 + k).value})')
cls3 = sum(n(wev.cell(r, ce0 + 10).value) for r in range(SR, le + 1))
mcol = gcl(ce0 + 10)
print(f'     10 国和 {s11:.0f}  vs  {mcol}3 {cls3:.0f}  差 = {cls3 - s11:.0f}')
print(f'     {mcol}3 {cls3:.0f}  vs  {S_MAIN}!{gcl(res_tab["欧洲"][0])}3 {res_tab["欧洲"][2]:.0f}  差 = {res_tab["欧洲"][2] - cls3:.0f}  ← 已知缺行位')
print(f'     数据行 {SR}~{le}（{le - SR + 1} 行），max_row={wev.max_row}，合计列行内公式 = {wef.cell(SR, ce0 + 10).value}')

print()
print('【5】' + S_Y)
wyv, wyf = wbv[S_Y], wbf[S_Y]
cy = YEAR - 2019
tot = 0
for m in range(1, M + 1):
    v = n(wyv.cell(m + 1, cy).value)
    tot += v
    print(f'     {m:>2}月(行{m+1}) = {v:>8.0f}')
print(f'     1~{M}月累加 = {tot:.0f}   汇总 G14 = {n(wyv.cell(14, cy).value):.0f}   公式 = {wyf.cell(14, cy).value}')
print(f'     I14(汇总同比) = {n(wyv.cell(14, cy + 2).value)}   公式 = {wyf.cell(14, cy + 2).value}')
print(f'     J14(历年总计) = {n(wyv.cell(14, cy + 3).value):.0f}   公式 = {wyf.cell(14, cy + 3).value}')

print()
print('=' * 100)
ok_all = all(v[5] for v in res_tab.values())
print(f'【结论】24 个分组列，四路数据全部一致 = {ok_all}')
print('=' * 100)
