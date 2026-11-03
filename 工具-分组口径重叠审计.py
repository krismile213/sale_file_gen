# -*- coding: utf-8 -*-
"""
分组口径重叠审计（只读工具）
================================
用途：验证【2026年】sheet 的 24 个分组列之间**是否存在重复计数 / 漏算**，
      即"子列之和是否等于父列"、"兄弟列之间是否互相重叠"。
换月只需改下面的 BELONGING 一个值。
跑完看末尾【结论】。

用法：
  cd "D:/Desktop/work-main/sale_file_gen"
  PYTHONPATH="C:/Users/dell/AppData/Roaming/Python/Python314/site-packages" \
    /c/Python314/python.exe "工具-分组口径重叠审计.py"
"""
import os, math, itertools
import pandas as pd
from openpyxl import load_workbook

# ================= 换月只改这里 =================
BELONGING = '202609'          # 数据归属周期 YYYYMM
# ===============================================

TRUNK = r'D:\Desktop\work-main\sale_file_gen\trunk'
TMP = os.path.join(TRUNK, 'sale_file_tmp')
YEAR = int(BELONGING[:4])
M = int(BELONGING[4:])
SUF = f'{BELONGING}销量'
RES = os.path.join(TRUNK, 'result', '数据归属周期', BELONGING, '月销售总表.xlsx')
S_MAIN = f'{YEAR}年'

# ---- 24 个分组（顺序与 month_sale_file.py 的 month_sale_group 一致）----
AMZ_MAP = {"欧洲": "欧洲", "英国": "英国", "美国": "美国", "加拿大": "加拿大", "日本": "日本",
           "澳洲": "澳大利亚", "阿联酋": "阿拉伯联合酋长国", "沙特": "沙特阿拉伯",
           "墨西哥": "墨西哥", "印度": "印度"}
SE6 = ["菲律宾", "印尼", "泰国", "马来西亚", "新加坡", "越南"]
SE_CROSS_3 = ["新加坡", "马来西亚", "越南"]      # ← 当前脚本口径（虾皮境内）
SE_CROSS_FIXED = ["马来西亚", "越南"]            # ← 方案A（去掉新加坡）

ORDER = list(AMZ_MAP.keys()) + ['亚马逊'] + \
        ['菲律宾', '泰国', '印尼', '新加坡', '马来', '越南', '虾皮境内', '东南亚'] + \
        ['京东', 'eMAG', 'TEMU平台', '美客多', '总计']

df = pd.read_csv(os.path.join(TMP, 'total_sku_sale_df.csv'), dtype=str)
df[SUF] = pd.to_numeric(df[SUF], errors='coerce').fillna(0)
for k in ('实体', '平台', '仓库', '站点'):
    df[k] = df[k].fillna('')
df['SKU'] = df['sku'].fillna('').astype(str).str.strip()
df = df.reset_index(drop=True)

PRED = {}
for g, wh in AMZ_MAP.items():
    PRED[g] = ('仓库 == %s' % wh, lambda d, wh=wh: d['仓库'] == wh)
PRED['亚马逊'] = ('平台 == 亚马逊', lambda d: d['平台'] == '亚马逊')
for g in ("菲律宾", "印尼", "泰国"):
    PRED[g] = ('站点 == %s（全平台）' % g, lambda d, g=g: d['站点'] == g)
PRED['新加坡'] = ('站点 == 新加坡（全平台）', lambda d: d['站点'] == '新加坡')
PRED['马来'] = ('站点 == 马来西亚 & Lazada', lambda d: (d['站点'] == '马来西亚') & (d['平台'] == 'Lazada'))
PRED['越南'] = ('站点 == 越南 & Lazada', lambda d: (d['站点'] == '越南') & (d['平台'] == 'Lazada'))
PRED['虾皮境内'] = ('Shopee & 站点∈[马,越] & 跨境（新加坡另计）',
                 lambda d: (d['平台'] == 'Shopee') & (d['站点'].isin(SE_CROSS_FIXED)) & (d['实体'] == '跨境'))
PRED['东南亚'] = ('站点∈6 国（全平台）', lambda d: d['站点'].isin(SE6))
PRED['京东'] = ('平台 == 京东 & 站点 == 国内', lambda d: (d['平台'] == '京东') & (d['站点'] == '国内'))
PRED['eMAG'] = ('平台 == EMAG', lambda d: d['平台'] == 'EMAG')
PRED['TEMU平台'] = ('平台 == Temu', lambda d: d['平台'] == 'Temu')
PRED['美客多'] = ('平台 == 美客多', lambda d: d['平台'] == '美客多')
PRED['总计'] = ('全表', lambda d: pd.Series(True, index=d.index))

MASK = {}
print('=' * 92)
print(f'分组口径重叠审计 · 归属周期 {BELONGING}   长表行数 {len(df)}   {SUF} 合计 {df[SUF].sum():.0f}')
print('=' * 92)

# ---- 1) 逐分组：谓词复算 vs 成品行 3 ----
if not os.path.exists(RES):
    raise SystemExit(f'成品不存在：{RES}')
ws = load_workbook(RES, data_only=True)[S_MAIN]

def num(v):
    if v is None: return 0.0
    if isinstance(v, float) and math.isnan(v): return 0.0
    s = str(v).strip()
    if s in ('', 'nan', 'None'): return 0.0
    try: return float(s)
    except: return 0.0

COL0 = {g: 3 + i * 12 for i, g in enumerate(ORDER)}
# 行3 是 =SUM() 公式；脚本产物无缓存 → 直读为 0。改为对数据行(4~LAST)累加。
LAST = max((r for r in range(4, ws.max_row + 1) if ws.cell(r, 1).value not in (None, '')), default=4)
print('\n【1】逐分组：谓词复算 vs 成品列累加（行4~%d，行3 为无缓存公式）' % LAST)
print('%-10s %-34s %11s %11s %6s' % ('分组', '谓词', '复算', '文件累加', '一致'))
bad = []
for g in ORDER:
    expr, fn = PRED[g]
    MASK[g] = fn(df)
    calc = df.loc[MASK[g], SUF].sum()
    ex = sum(num(ws.cell(r, COL0[g] + M - 1).value) for r in range(4, LAST + 1))
    ok = abs(calc - ex) < 0.5
    if not ok: bad.append(g)
    print('%-10s %-34s %11.0f %11.0f %6s' % (g, expr, calc, ex, 'OK' if ok else 'X'))

# ---- 2) 兄弟重叠（同一父级的叶子列两两检测）----
print('\n【2】兄弟叶子列两两重叠（重叠＝同一笔订单被算进两列）')
LEAF_BLOCKS = {
    '亚马逊系': list(AMZ_MAP.keys()),
    '东南亚系': ['菲律宾', '泰国', '印尼', '新加坡', '马来', '越南', '虾皮境内'],
    '其它平台': ['京东', 'eMAG', 'TEMU平台', '美客多'],
}
overlaps = []
for bname, members in LEAF_BLOCKS.items():
    print(f'--- {bname} ---')
    hit = False
    for a, b in itertools.combinations(members, 2):
        ov = MASK[a] & MASK[b]
        n = int(ov.sum())
        if n == 0: continue
        hit = True
        overlaps.append((a, b, n, df.loc[ov, SUF].sum()))
        print('  ⚠ %s ∩ %s : %d 行 / %.0f 件' % (a, b, n, df.loc[ov, SUF].sum()))
        for idx, cnt in df.loc[ov].groupby(['站点', '平台', '实体', '仓库']).size().items():
            print('      %-6s %-8s %-8s %-12s %d 行' % (idx[0], idx[1], idx[2], idx[3], cnt))
    if not hit: print('  ✔ 无重叠')

# ---- 3) 父子：子列之和 vs 父列 ----
print('\n【3】父子：子列之和 vs 父列')
PARENTS = {'亚马逊': list(AMZ_MAP.keys()),
           '东南亚': ['菲律宾', '泰国', '印尼', '新加坡', '马来', '越南', '虾皮境内']}
for par, kids in PARENTS.items():
    s = df.loc[MASK[par], SUF].sum()
    k = sum(df.loc[MASK[c], SUF].sum() for c in kids)
    flag = 'OK' if abs(s - k) < 0.5 else '⚠ 不一致'
    print('  %-8s 父列 = %10.0f   子列和 = %10.0f   差 = %8.0f  %s' % (par, s, k, s - k, flag))

LEAVES = list(AMZ_MAP.keys()) + ['菲律宾', '泰国', '印尼', '新加坡', '马来', '越南', '虾皮境内'] + \
         ['京东', 'eMAG', 'TEMU平台', '美客多']
leaf_sum = sum(df.loc[MASK[g], SUF].sum() for g in LEAVES)
tot = df.loc[MASK['总计'], SUF].sum()
print('\n  21 个叶子之和 = %.0f   总计 = %.0f   差 = %.0f  %s'
      % (leaf_sum, tot, tot - leaf_sum, 'OK' if abs(tot - leaf_sum) < 0.5 else '⚠ 不等于总计'))

# ---- 4) 亚马逊系按仓库过滤是否越界 ----
print('\n【4】亚马逊系按「仓库」过滤 —— 是否有非亚马逊平台的行被算进去')
amz_wh = set(AMZ_MAP.values())
hitdf = df[df['仓库'].isin(amz_wh) & (df['平台'] != '亚马逊')]
print('  仓库命中亚马逊仓但平台≠亚马逊 = %d 行 %s' % (len(hitdf), '✔' if len(hitdf) == 0 else '⚠'))
if len(hitdf):
    print(hitdf.groupby(['仓库', '站点', '平台'])[SUF].agg(['count', 'sum']).to_string())

# ---- 5) 覆盖性 ----
print('\n【5】覆盖性')
union = pd.concat([MASK[g] for g in LEAVES], axis=1).any(axis=1)
print('  叶子并集覆盖 %d / %d 行 %s' % (int(union.sum()), len(df), '✔' if int(union.sum()) == len(df) else '⚠'))
gap = df.loc[~union]
if len(gap):
    print('  未被覆盖：')
    print(gap.groupby(['站点', '平台', '实体', '仓库'])[SUF].agg(['count', 'sum']).to_string())

# ---- 6) 源表目的地覆盖（虾皮跨境）----
print('\n【6】虾皮跨境目的地覆盖')
sites = sorted(df[(df['平台'] == 'Shopee') & (df['实体'] == '跨境')]['站点'].unique().tolist())
print('  长表里 Shopee∩跨境 的站点 =', sites)
print('  脚本纳入虾皮境内的是 =', SE_CROSS_FIXED, '→ 未纳入的 =',
      [s for s in sites if s not in SE_CROSS_FIXED] or '无')
print('  源表虾皮跨境含 8 个目的地（马来西亚/巴西/中国台湾/墨西哥/新加坡/越南/哥伦比亚/智利），')
print('  其中「中国台湾」历史月有量（1月16/3月10/4月8/5月5/6月7）—— 脚本无对应输入报告，属结构缺口。')

# ---- 结论 ----
print('\n' + '=' * 92)
print('【结论】')
print('  分组谓词与成品不一致的列 =', bad or '无 ✔')
print('  兄弟叶子重叠 =', ('%d 处：%s' % (len(overlaps), ', '.join('%s∩%s' % (a, b) for a, b, _, _ in overlaps))) if overlaps else '无 ✔')
print('  东南亚：子列和 = %.0f，父列 = %.0f，差 = %.0f %s'
      % (sum(df.loc[MASK[c], SUF].sum() for c in PARENTS['东南亚']),
         df.loc[MASK['东南亚'], SUF].sum(),
         df.loc[MASK['东南亚'], SUF].sum() - sum(df.loc[MASK[c], SUF].sum() for c in PARENTS['东南亚']),
         '✔' if abs(df.loc[MASK['东南亚'], SUF].sum() - sum(df.loc[MASK[c], SUF].sum() for c in PARENTS['东南亚'])) < 0.5 else '⚠'))
print('  21 叶子和 vs 总计：差 = %.0f %s' % (tot - leaf_sum, '✔' if abs(tot - leaf_sum) < 0.5 else '⚠'))
print('=' * 92)
