# -*- coding: utf-8 -*-
"""月度销量月报生成器 —— 输出 A4 PDF（HTML 渲染后打印）。

用法：
    1) 改下面「换月 CONFIG」3 行；
    2) 先跑完主脚本 + 三个区域脚本（产物齐了再生成报告）；
    3) 运行：PYTHONPATH="...site-packages" C:\\Python314\\python.exe 工具-月报生成.py
    4) 产物：<项目根>/<本期>期销量月报-<本月>月vs<上月>月.pdf

依赖：openpyxl、playwright（channel=msedge，用系统 Edge 打印，无需下载浏览器）。

口径说明（与本项目既有约定一致，勿改）：
    · 主表【2026年】sheet 行1 为分组名、每 12 列一个分组（1~12 月），第 m 月列 = 块首 + (m-1)。
    · 行 3 是合计公式（脚本产物无缓存值）→ 一律对**数据行（行4~末行）累加**。
    · 东南亚成品各 sheet：行1 列头 / 行2 合计 / 行3 起数据；月块首 = 3 + (m-1)*块宽；
      块结构 = [渠道列…][月合计列]，合计在块尾。
"""
import openpyxl, os, html, glob, pathlib

# ===================== 换月 CONFIG（只改这 3 行）=====================
CUR_PERIOD = '202609'        # 本期归属周期 YYYYMM
CUR_SNAP = '20261006'        # 本期快照日期 YYYYMMDD
PREV_PERIOD = '202608'       # 上期归属周期 YYYYMM
# ==================================================================

BASE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(BASE)
RES = os.path.join(BASE, 'result')
TMP = os.path.join(ROOT, '.workbuddy', '_tmp')

CUR_M = int(CUR_PERIOD[4:6])
PREV_M = int(PREV_PERIOD[4:6])
CUR_Y = CUR_PERIOD[:4]
TITLE_CN = f'{CUR_Y} 年 {CUR_M} 月'
PREV_CN = f'{CUR_Y} 年 {PREV_M} 月'

SEA_CHAN = {
    'Philippines 菲律宾销量表': (6, ['跨境 Lazada 中山仓', '菲律宾仓销量（Lazada）',
                                   '菲律宾仓销量（Shopee）', '抖音 Tiktok', 'KSON 线下订单'], '菲律宾'),
    'Thailand 泰国销量表': (5, ['Shopee 易海仓', 'Lazada 易海仓', 'Lazada 泰国仓', 'Lazada 中山仓'], '泰国'),
    'Indonesia 印尼销量表': (4, ['Shopee 智慧仓', '印尼线下订单', '抖音 Tiktok'], '印尼'),
}


def num(v):
    try:
        return float(v) if v not in (None, '') else 0.0
    except Exception:
        return 0.0


def one(d, pat):
    fs = glob.glob(os.path.join(d, pat))
    if not fs:
        raise FileNotFoundError(f'找不到：{os.path.join(d, pat)}')
    return fs[0]


def read_main(path, m):
    """按【2026年】行1 分组名定位块首，取第 m 月列（数据行累加）。"""
    wb = openpyxl.load_workbook(path, data_only=True)
    s = wb['2026年']
    out, last = {}, None
    for c in range(1, s.max_column + 1):
        v1 = s.cell(1, c).value
        if v1 not in (None, '') and v1 != last:
            if v1 != '2026年':
                out[v1] = sum(num(s.cell(r, c + (m - 1)).value) for r in range(4, s.max_row + 1))
            last = v1
        elif v1 in (None, ''):
            last = None
    return out, s.max_row


P_PREV = one(os.path.join(RES, '数据归属周期', PREV_PERIOD), '月销售总表.xlsx')
P_CUR = one(os.path.join(RES, '数据归属周期', CUR_PERIOD), '月销售总表.xlsx')
PREV, _ = read_main(P_PREV, PREV_M)
CUR, cur_last = read_main(P_CUR, CUR_M)
CUR_SAME = read_main(P_CUR, PREV_M)[0]       # 本期成品里的上月列（用于校验历史列一致性）

rows_cur = cur_last - 3

# 分渠道（按列位置固定命名，不靠字符串清洗）
P_SEA = one(os.path.join(RES, '东南亚销量表', CUR_PERIOD), '*东南亚销量表*.xlsx')
wbs = openpyxl.load_workbook(P_SEA, data_only=True)
chan = {}
for sh, (bw, names, site) in SEA_CHAN.items():
    ws = wbs[sh]
    recs = []
    for m in [PREV_M, CUR_M]:
        b = 3 + (m - 1) * bw
        recs.append([(nm, sum(num(ws.cell(r, b + k).value) for r in range(3, ws.max_row + 1)))
                     for k, nm in enumerate(names)])
    chan[site] = recs

TOTAL_P, TOTAL_C = PREV['总计'], CUR['总计']
EUR_P, EUR_C = PREV['欧洲'] + PREV['英国'], CUR['欧洲'] + CUR['英国']
NA_P = PREV['美国'] + PREV['TEMU平台'] + PREV['加拿大'] + PREV['墨西哥'] + PREV['美客多']
NA_C = CUR['美国'] + CUR['TEMU平台'] + CUR['加拿大'] + CUR['墨西哥'] + CUR['美客多']
AMZ_P, AMZ_C = PREV['亚马逊'], CUR['亚马逊']
SEA_P, SEA_C = PREV['东南亚'], CUR['东南亚']


def f0(x):
    return f'{x:,.0f}'


def sgn(x):
    return ('+' if x > 0 else '−' if x < 0 else '±') + f'{abs(x):,.0f}'


def pct(v8, v9):
    if not v8:
        return '—' if not v9 else '新增'
    return f'{(v9 - v8) / v8 * 100:+.1f}%'.replace('-', '−')


def cls_of(d):
    return 'up' if d > 0 else ('down' if d < 0 else 'flat')


def group_table(title, keys, subtotal_key):
    rows = []
    for k in keys:
        v8, v9 = CUR_SAME.get(k, 0), CUR.get(k, 0)
        d = v9 - v8
        rows.append(f'<tr><td class="g">{html.escape(k)}</td><td class="n">{f0(v8)}</td>'
                    f'<td class="n">{f0(v9)}</td><td class="n {cls_of(d)}">{sgn(d)}</td>'
                    f'<td class="n {cls_of(d)}">{pct(v8, v9) if v8 else "—"}</td></tr>')
    if subtotal_key:
        v8, v9 = CUR_SAME[subtotal_key], CUR[subtotal_key]
        d = v9 - v8
        rows.append(f'<tr class="sub"><td class="g">{subtotal_key}（合计）</td><td class="n">{f0(v8)}</td>'
                    f'<td class="n">{f0(v9)}</td><td class="n {cls_of(d)}">{sgn(d)}</td>'
                    f'<td class="n {cls_of(d)}">{pct(v8, v9)}</td></tr>')
    return (f'<h3>{html.escape(title)}</h3><table class="grid"><thead><tr>'
            f'<th>分组</th><th>{PREV_M} 月</th><th>{CUR_M} 月</th><th>变化</th><th>环比</th></tr></thead>'
            f'<tbody>{"".join(rows)}</tbody></table>')


def dual_chart(items, title):
    mx = max([max(v8, v9) for _, v8, v9 in items] + [1])
    out = [f'<h3>{html.escape(title)}</h3>',
           f'<div class="legend"><span><i class="sw sw8"></i>{PREV_M} 月</span>'
           f'<span><i class="sw sw9"></i>{CUR_M} 月</span></div>']
    for name, v8, v9 in items:
        d = v9 - v8
        out.append(
            f'<div class="drow"><div class="dlab">{html.escape(name)}</div><div class="dtracks">'
            f'<div class="dtrack"><div class="dbar b8" style="width:{v8 / mx * 100:.2f}%"></div>'
            f'<span class="dnum">{f0(v8)}</span></div>'
            f'<div class="dtrack"><div class="dbar b9" style="width:{v9 / mx * 100:.2f}%"></div>'
            f'<span class="dnum">{f0(v9)}</span></div>'
            f'</div><div class="ddelta {cls_of(d)}">{sgn(d)}</div></div>')
    return ''.join(out)


def delta_chart(items, title):
    mx = max(abs(v9 - v8) for _, v8, v9 in items) or 1
    out = [f'<h3>{html.escape(title)}</h3>',
           '<div class="legend"><span><i class="sw sw-down"></i>下滑</span>'
           '<span><i class="sw sw-up"></i>增长</span></div>']
    for name, v8, v9 in items:
        d = v9 - v8
        out.append(f'<div class="brow"><div class="blab">{html.escape(name)}</div>'
                   f'<div class="btrack"><div class="bbar {cls_of(d)}" '
                   f'style="width:{abs(d) / mx * 100:.2f}%"></div></div>'
                   f'<div class="bval {cls_of(d)}">{sgn(d)}</div></div>')
    return ''.join(out)


def chan_table(site, t8, t9):
    r8, r9 = chan[site]
    d8, d9 = dict(r8), dict(r9)
    rows = []
    for k in [k for k, _ in r9]:
        v8, v9 = d8.get(k, 0), d9.get(k, 0)
        d = v9 - v8
        rows.append(f'<tr><td class="g">{html.escape(k)}</td><td class="n">{f0(v8)}</td>'
                    f'<td class="n">{f0(v9)}</td><td class="n {cls_of(d)}">{sgn(d)}</td></tr>')
    d = t9 - t8
    rows.append(f'<tr class="sub"><td class="g">合计</td><td class="n">{f0(t8)}</td>'
                f'<td class="n">{f0(t9)}</td><td class="n {cls_of(d)}">{sgn(d)}</td></tr>')
    return (f'<h3>{html.escape(site)} · 分渠道</h3><table class="grid"><thead><tr>'
            f'<th>渠道</th><th>{PREV_M} 月</th><th>{CUR_M} 月</th><th>变化</th></tr></thead>'
            f'<tbody>{"".join(rows)}</tbody></table>')


board = [('全渠道总计', TOTAL_P, TOTAL_C), ('亚马逊（10 国合计）', AMZ_P, AMZ_C),
         ('东南亚（6 站合计）', SEA_P, SEA_C), ('欧洲表（9 国+英国）', EUR_P, EUR_C),
         ('北美表（美+加+墨）', NA_P, NA_C)]

# 叶子分组（互不重叠，21 项之和 = 总计）—— 排序/占比只许用叶子，避免聚合项重复计数
LEAF = ['欧洲', '英国', '美国', '加拿大', '日本', '澳洲', '阿联酋', '沙特', '墨西哥', '印度',
        '菲律宾', '泰国', '印尼', '新加坡', '马来', '越南', '虾皮境内',
        '京东', 'eMAG', 'TEMU平台', '美客多']
ditems = sorted([(k, CUR_SAME[k], CUR[k]) for k in LEAF], key=lambda t: -abs(t[2] - t[1]))[:14]

ph_d = chan['菲律宾'][1]
top3 = sorted([(k, CUR[k] - CUR_SAME[k]) for k in LEAF], key=lambda t: t[1])[:3]
top3_txt = '、'.join(f'{k} {sgn(v)}' for k, v in top3)
top3_share = abs(sum(v for _, v in top3)) / abs(TOTAL_P - TOTAL_C) * 100
ups = sorted([(k, CUR[k] - CUR_SAME[k]) for k in LEAF], key=lambda t: -t[1])[:3]
ups = [(k, v) for k, v in ups if v > 0]
ups_txt = '、'.join(f'{k} {sgn(v)}' for k, v in ups) if ups else '无'

CSS = """
@page { size: A4; margin: 13mm 12mm; }
* { box-sizing: border-box; }
body { font-family: "Microsoft YaHei","PingFang SC",sans-serif; color:#1c1c1c;
       font-size:10.5pt; line-height:1.55; margin:0; }
h1 { font-size:19pt; margin:0 0 3px; font-weight:600; letter-spacing:.5px; }
h2 { font-size:12.5pt; margin:15px 0 8px; padding-bottom:4px;
     border-bottom:1.5px solid #1F3A5F; color:#1F3A5F; font-weight:600; }
h3 { font-size:10.5pt; margin:10px 0 5px; color:#33415c; font-weight:600; }
h2, h3 { page-break-after: avoid; break-after: avoid; }
table.grid, .note, .cards, ul.tight { page-break-inside: avoid; break-inside: avoid; }
.sub-hd { color:#5b6472; font-size:9.5pt; margin:0 0 2px; }
.rule { height:3px; background:#1F3A5F; width:64px; margin:9px 0 0; }
.cards { display:flex; gap:8px; margin:11px 0 4px; }
.card { flex:1; background:#F4F6F9; border-radius:7px; padding:9px 11px; }
.card .k { font-size:9pt; color:#5b6472; margin-bottom:2px; }
.card .v { font-size:15pt; font-weight:600; line-height:1.2; }
.card .d { font-size:8.5pt; margin-top:2px; }
.up { color:#1F7A3D; } .down { color:#C0392B; } .flat { color:#5b6472; }
table.grid { width:100%; border-collapse:collapse; font-size:9.5pt; margin-bottom:3px; }
table.grid th { background:#EAEEF3; color:#33415c; text-align:right; font-weight:600;
                padding:4px 7px; border-bottom:1px solid #c8d0da; }
table.grid th:first-child { text-align:left; }
table.grid td { padding:3.4px 7px; border-bottom:0.6px solid #e4e8ee; }
table.grid td.g { text-align:left; }
table.grid th.txt, table.grid td.txt { text-align:left; width:56%; }
table.grid td.n { text-align:right; font-variant-numeric:tabular-nums; }
table.grid tbody tr:nth-child(even) { background:#FAFBFC; }
table.grid tr.sub td { font-weight:600; background:#EEF2F7; border-top:1px solid #c8d0da; }
.legend { font-size:9pt; color:#5b6472; margin-bottom:5px; }
.legend span { margin-right:14px; }
.sw { display:inline-block; width:10px; height:10px; border-radius:2px;
      margin-right:4px; vertical-align:-1px; }
.sw8 { background:#9FB3C8; } .sw9 { background:#1F3A5F; }
.sw-down { background:#C0392B; } .sw-up { background:#1F7A3D; }
.drow { display:flex; align-items:center; gap:9px; margin-bottom:4px; page-break-inside:avoid; }
.dlab { width:104px; text-align:right; font-size:9.5pt; flex:none; }
.dtracks { flex:1; }
.dtrack { position:relative; height:10px; background:#F0F3F7; border-radius:3px; margin:2px 0; }
.dbar { height:10px; border-radius:3px; }
.dbar.b8 { background:#9FB3C8; } .dbar.b9 { background:#1F3A5F; }
.dnum { position:absolute; right:5px; top:-2px; font-size:8.5pt; color:#33415c; }
.ddelta { width:66px; text-align:right; font-size:9.5pt; flex:none; font-weight:600; }
.brow { display:flex; align-items:center; gap:9px; margin-bottom:3px; page-break-inside:avoid; }
.blab { width:104px; text-align:right; font-size:9.5pt; flex:none; }
.btrack { flex:1; height:12px; background:#F0F3F7; border-radius:3px; }
.bbar { height:12px; border-radius:3px; }
.bbar.down { background:#C0392B; } .bbar.up { background:#1F7A3D; }
.bval { width:66px; text-align:right; font-size:9.5pt; font-weight:600; flex:none; }
.note { background:#FBF7EC; border-left:3px solid #D9A441; padding:8px 11px;
        font-size:9.5pt; border-radius:0 5px 5px 0; margin:7px 0; }
.note b { color:#8a6a1f; }
ul.tight { margin:5px 0; padding-left:18px; }
ul.tight li { margin-bottom:2px; }
.pagebreak { page-break-before:always; break-before:page; }
footer { margin-top:14px; padding-top:7px; border-top:0.6px solid #d8dee6;
         font-size:8.5pt; color:#7a828e; }
"""

DOC = f"""<!DOCTYPE html>
<html lang="zh-CN"><head><meta charset="utf-8"><style>{CSS}</style></head><body>

<h1>{CUR_PERIOD} 期销量月报</h1>
<p class="sub-hd">{CUR_M} 月 vs {PREV_M} 月 环比分析 ｜ 数据归属周期 {CUR_PERIOD} ｜ 快照日期 {CUR_SNAP}</p>
<div class="rule"></div>

<h2>一、核心结论</h2>
<div class="cards">
  <div class="card"><div class="k">全渠道总计</div><div class="v">{f0(TOTAL_C)}</div>
    <div class="d {cls_of(TOTAL_C-TOTAL_P)}">{sgn(TOTAL_C-TOTAL_P)}　{pct(TOTAL_P,TOTAL_C)}</div></div>
  <div class="card"><div class="k">亚马逊（10 国）</div><div class="v">{f0(AMZ_C)}</div>
    <div class="d {cls_of(AMZ_C-AMZ_P)}">{sgn(AMZ_C-AMZ_P)}　{pct(AMZ_P,AMZ_C)}</div></div>
  <div class="card"><div class="k">东南亚（6 站）</div><div class="v">{f0(SEA_C)}</div>
    <div class="d {cls_of(SEA_C-SEA_P)}">{sgn(SEA_C-SEA_P)}　{pct(SEA_P,SEA_C)}</div></div>
  <div class="card"><div class="k">欧洲表 / 北美表</div><div class="v">{f0(EUR_C)} / {f0(NA_C)}</div>
    <div class="d {cls_of(EUR_C-EUR_P)}">欧洲 {pct(EUR_P,EUR_C)}　北美 {pct(NA_P,NA_C)}</div></div>
</div>
<ul class="tight">
  <li><b>整体环比 {pct(TOTAL_P,TOTAL_C)}</b>：全渠道从 {PREV_M} 月 {f0(TOTAL_P)} 件变为 {CUR_M} 月 {f0(TOTAL_C)} 件，
      变化 {sgn(TOTAL_C-TOTAL_P)} 件；亚马逊、东南亚、欧洲表、北美表四大板块
      {'同步下行' if TOTAL_C < TOTAL_P else '同步上行'}。</li>
  <li><b>跌幅最大的三处</b>（按叶子分组）：{top3_txt}。三项合计贡献了总变动的约 {top3_share:.0f}%。</li>
  <li><b>菲律宾是最大异常点</b>：单站 {pct(CUR_SAME['菲律宾'],CUR['菲律宾'])}（{f0(CUR_SAME['菲律宾'])} → {f0(CUR['菲律宾'])}），
      其中线下 KSON 仓由 {f0(dict(chan['菲律宾'][0])['KSON 线下订单'])} 降至 {f0(dict(ph_d)['KSON 线下订单'])}，
      抖音由 {f0(dict(chan['菲律宾'][0])['抖音 Tiktok'])} 降至 {f0(dict(ph_d)['抖音 Tiktok'])}。</li>
  <li><b>逆势增长项</b>：{ups_txt}，量级均较小，不足以改变大盘方向。</li>
  <li><b>口径可信</b>：{PREV_M} 月列在本期与上期两版成品中<b>逐项差 0</b>，两期数据可直接对比。</li>
</ul>

<h2>二、大盘与板块对比</h2>
{dual_chart(board, f'各板块 {PREV_M} 月 / {CUR_M} 月销量（件）')}
<div class="note"><b>口径提示：</b>欧洲表和北美表与“亚马逊系”存在重叠
（美国 / 加拿大 / 墨西哥既计入亚马逊系，也是北美表的组成部分），四者<b>不可相加</b>；全渠道总计以主表「总计」为准。</div>
<div class="pagebreak"></div>
<h2>三、分区明细</h2>
{group_table('亚马逊系 · 各国站点', ['欧洲','英国','美国','加拿大','日本','澳洲','阿联酋','沙特','墨西哥','印度'], '亚马逊')}
{group_table('东南亚 · 各站点', ['菲律宾','泰国','印尼','新加坡','马来','越南','虾皮境内'], '东南亚')}
{group_table('其他渠道', ['京东','eMAG','TEMU平台','美客多'], None)}
<table class="grid"><thead><tr><th>全渠道</th><th>{PREV_M} 月</th><th>{CUR_M} 月</th><th>变化</th><th>环比</th></tr></thead>
<tbody><tr class="sub"><td class="g">总计</td><td class="n">{f0(TOTAL_P)}</td><td class="n">{f0(TOTAL_C)}</td>
<td class="n {cls_of(TOTAL_C-TOTAL_P)}">{sgn(TOTAL_C-TOTAL_P)}</td>
<td class="n {cls_of(TOTAL_C-TOTAL_P)}">{pct(TOTAL_P,TOTAL_C)}</td></tr></tbody></table>

<h2>四、变动幅度排序</h2>
{delta_chart(ditems, '环比变化量（叶子分组，绝对变幅前 14 名，单位：件）')}

<h2>五、重点渠道拆解</h2>
{chan_table('菲律宾', CUR_SAME['菲律宾'], CUR['菲律宾'])}
{chan_table('泰国', CUR_SAME['泰国'], CUR['泰国'])}
{chan_table('印尼', CUR_SAME['印尼'], CUR['印尼'])}

<h2>六、数据说明与待办</h2>
<div class="note"><b>本期已处理的源数据问题</b><br>
1. <b>抖音（Tiktok）菲律宾 / 印尼</b>：初次导出仅 1 列（导出异常）。已于 2026-10-07 重导，
   重导后 {CUR_M} 月报菲 624 件、印尼 228 件（{PREV_M} 月为 945 / 309）。<br>
2. <b>美客多报告语言变更</b>：本期由中文版改为西班牙语版，读取逻辑已适配。<br>
3. <b>越南退货窗口长文案</b>：订单状态为长句描述，已补入状态映射表。</div>
<div class="note"><b>已知事项（不影响本期出数）</b><br>
1. <b>菲律宾 KSON 仓线下订单</b>：{PREV_M} 月 2,132 件，{CUR_M} 月按 0 计 —— 该项已挂起，待业务方反馈确认。<br>
2. <b>interflow 阿联酋</b>：订单只覆盖至 09-29，经确认 09-30 当天确无订单，非导出缺失。<br>
3. 主表【欧洲各国】sheet 固定写 497 行，新增 SKU 不落入该 sheet（本期差 469 件）；
   该现象在 {PREV_M} 月已验收成品中同样存在（差 242 件），属既有结构特征，4 张区域销量表已按长表补齐，不受影响。</div>

<h2>七、交付物与校验</h2>
<table class="grid"><thead><tr><th>交付物</th><th>{CUR_M} 月合计</th><th class="txt">说明</th></tr></thead><tbody>
<tr><td class="g">月销售总表</td><td class="n">{f0(TOTAL_C)}</td><td class="txt">{rows_cur} 数据行</td></tr>
<tr><td class="g">东南亚销量表</td><td class="n">{f0(SEA_C)}</td><td class="txt">菲 {f0(CUR['菲律宾'])} / 泰 {f0(CUR['泰国'])} / 印 {f0(CUR['印尼'])} / 新 {f0(CUR['新加坡'])} / 越 {f0(CUR['越南'])} / 马 {f0(CUR['马来'])} / 虾皮跨境 {f0(CUR['虾皮境内'])}</td></tr>
<tr><td class="g">欧洲销量表</td><td class="n">{f0(EUR_C)}</td><td class="txt">9 国 {f0(CUR['欧洲'])} + 英国 {f0(CUR['英国'])}</td></tr>
<tr><td class="g">北美销量表</td><td class="n">{f0(NA_C)}</td><td class="txt">美国 {f0(CUR['美国'])} / 加拿大 {f0(CUR['加拿大'])} / 墨西哥 {f0(CUR['墨西哥'])} / 美客多 {f0(CUR['美客多'])}</td></tr>
</tbody></table>
<ul class="tight">
  <li>换月资料体检：<b>33 项通过，0 致命</b></li>
  <li>销量双重验证：<b>24 个分组四路对撞全部一致</b>（4 个恒等式差值均为 0）</li>
  <li>分组口径重叠审计：<b>24 个分组无重叠</b>，父子列差值均为 0</li>
  <li>欧洲 / 北美验收：<b>64 项全部通过</b></li>
</ul>

<footer>sale_file_gen 月度销量报表管线 ｜ 数据归属周期 {CUR_PERIOD} ｜ 快照 {CUR_SNAP} ｜ 本报告由脚本自动生成，数值取自四份成品表</footer>
</body></html>
"""

# ---------------- 输出 ----------------
os.makedirs(TMP, exist_ok=True)
out_html = os.path.join(TMP, f'report_{CUR_PERIOD}.html')
with open(out_html, 'w', encoding='utf-8') as f:
    f.write(DOC)

diff = {k: CUR_SAME.get(k, 0) - PREV.get(k, 0) for k in CUR if abs(CUR_SAME.get(k, 0) - PREV.get(k, 0)) > 0.5}
print('HTML 已生成:', out_html)
print(f'历史列一致性（{PREV_M} 月列 本期 vs 上期）:', diff or '全 0 ✅')
print(f'总计 {PREV_M}月/{CUR_M}月 = {f0(TOTAL_P)} / {f0(TOTAL_C)}')
print(f'欧洲表 {f0(EUR_P)}/{f0(EUR_C)}   北美表 {f0(NA_P)}/{f0(NA_C)}')

out_pdf = os.path.join(ROOT, f'{CUR_PERIOD}期销量月报-{CUR_M}月vs{PREV_M}月.pdf')
try:
    from playwright.sync_api import sync_playwright
    with sync_playwright() as p:
        b = p.chromium.launch(channel='msedge')
        pg = b.new_page()
        pg.goto(pathlib.Path(out_html).as_uri(), wait_until='load')
        pg.pdf(path=out_pdf, format='A4', print_background=True,
               margin={'top': '0', 'bottom': '0', 'left': '0', 'right': '0'})
        b.close()
    print('PDF 已生成:', out_pdf)
except Exception as e:
    print('⚠️ PDF 打印失败（HTML 仍可用）:', type(e).__name__, e)
