# -*- coding: utf-8 -*-
"""
欧洲销量表 8 月填报脚本（结构仿 southeast_asia_sale_file.py）

输入：
  1) 模板      data/数据归属周期/<归属月>/快照日期/<快照日>/欧洲/2026年欧洲销量表Europe Sales table.xlsx
  2) 渠道销量  sale_file_tmp/total_sku_sale_df.csv  ← month_sale_file.py 的中间产物
               （列：实体/平台/仓库/站点/sku/asin/<归属月>销量；已含状态过滤与 SKU 归集）
               注：东南亚脚本读 se_a_sku_sale_df.csv；欧/美用全渠道总表即可，主脚本 0 改动

输出：
  result/欧洲销量表/<归属月>/2026年欧洲销量表Europe Sales table.xlsx
  result/欧洲销量表/<归属月>/<Country>_<归属月>_sales.xlsx（10 份纯英文表头交付件）

口径（与月销总表同源）：
  欧洲 10 国 = 比利时/德国/西班牙/法国/意大利/荷兰/波兰/瑞典/爱尔兰 + 英国
  每国 = 亚马逊该国站点；8 月 10 国之和 = 月销表「欧洲」28315 + 「英国」8018 = 36333
"""
import os
import datetime
from copy import copy

import pandas as pd
import openpyxl
from openpyxl.utils import get_column_letter as gl
from openpyxl.comments import Comment
from openpyxl.formatting.rule import Rule
from openpyxl.styles import Font, PatternFill
from openpyxl.styles.differential import DifferentialStyle

from wps_fidelity import apply_fidelity, check_fidelity

# ==================== 月度参数（每月只改这三行） ====================
belonging_date = "202610"      # 数据归属周期 YYYYMM
snapshot_date = "20261103"     # 快照日期 YYYYMMDD（11-03 统一时点：欧盟报告=英国版+欧盟独有合并）
month_offset = 10              # 当月月份 1~12
AUTO_OPEN = False              # 跑完是否自动打开成品
# ===================================================================

cur = datetime.datetime.strptime(belonging_date, "%Y%m").date()
YEAR = cur.year
Q = f"{cur.strftime('%Y%m')}销量"
ABBR = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"][month_offset - 1]

trunk_dir = os.path.dirname(os.path.abspath(__file__))
tmp_dir = os.path.join(trunk_dir, "sale_file_tmp")
snap_dir = os.path.join(trunk_dir, "data", "数据归属周期", belonging_date, "快照日期", snapshot_date)
template_path = os.path.join(snap_dir, "欧洲", f"{YEAR}年欧洲销量表Europe Sales table.xlsx")
result_dir = os.path.join(trunk_dir, "result", "欧洲销量表", belonging_date)
result_path = os.path.join(result_dir, os.path.basename(template_path))

TOTAL_SHEET = "EU Total 欧洲总销量表"
HIST_SHEET = f"2025-{YEAR}年销量情况 Sales"
BASE = 6                                        # Total sheet 基础列 A~F
MONTHS_ALL = [(y, m) for y in (2025, YEAR) for m in range(1, 13)]

# ---------------- SKU 基础信息（补行位时用：型号/类型/品牌/得分） ----------------
# 模板的 SKU 行集合取自月销总表「欧洲各国」sheet，而该 sheet 有 13 个主表 SKU 缺行位
# （即那个 242 件缺口涉及的 SKU）。这 6 个 SKU 长表里有销量却没行位 → 运行时补行，
# 补行时基础列不能留空，统一从月销总表【{YEAR}年】取（291/292/294 = 类型/品牌/得分）。
main_path = os.path.join(trunk_dir, "result", "数据归属周期", belonging_date, "月销售总表.xlsx")
SKU_INFO = {}
if os.path.exists(main_path):
    _mw = openpyxl.load_workbook(main_path, data_only=True)[f"{YEAR}年"]
    for _r in range(4, _mw.max_row + 1):
        _s = _mw.cell(_r, 1).value
        if _s:
            SKU_INFO[str(_s).strip()] = (_mw.cell(_r, 2).value,    # 型号
                                         _mw.cell(_r, 291).value,  # 类型
                                         _mw.cell(_r, 292).value,  # 品牌
                                         _mw.cell(_r, 294).value)  # 销量得分

# (sheet 名, 中文, 英文, [(渠道英文名, [(实体, 平台, 仓库, 站点)…])…])
COUNTRIES = [
    ("Belgium 比利时销量表", "比利时", "Belgium",
     [("Amazon", [("巨川", "亚马逊", "欧洲", "比利时")])]),
    ("Germany 德国销量表", "德国", "Germany",
     [("Amazon", [("巨川", "亚马逊", "欧洲", "德国")])]),
    ("Spain 西班牙销量表", "西班牙", "Spain",
     [("Amazon", [("巨川", "亚马逊", "欧洲", "西班牙")])]),
    ("France 法国销量表", "法国", "France",
     [("Amazon", [("巨川", "亚马逊", "欧洲", "法国")])]),
    ("Italy 意大利销量表", "意大利", "Italy",
     [("Amazon", [("巨川", "亚马逊", "欧洲", "意大利")])]),
    ("Netherlands 荷兰销量表", "荷兰", "Netherlands",
     [("Amazon", [("巨川", "亚马逊", "欧洲", "荷兰")])]),
    ("Poland 波兰销量表", "波兰", "Poland",
     [("Amazon", [("巨川", "亚马逊", "欧洲", "波兰")])]),
    ("Sweden 瑞典销量表", "瑞典", "Sweden",
     [("Amazon", [("巨川", "亚马逊", "欧洲", "瑞典")])]),
    ("Ireland 爱尔兰销量表", "爱尔兰", "Ireland",
     [("Amazon", [("巨川", "亚马逊", "欧洲", "爱尔兰")])]),
    ("United Kingdom 英国销量表", "英国", "United Kingdom",
     [("Amazon", [("巨川", "亚马逊", "英国", "英国")])]),
]

EXPECT = {          # ★ 202610 实测值（10 月，20261103 统一时点快照）；换月先置空 {} 跑通，再把实测值填回
    "比利时": 877, "德国": 6923, "西班牙": 5237, "法国": 4094, "意大利": 4828,
    "荷兰": 777, "波兰": 292, "瑞典": 610, "爱尔兰": 943, "英国": 6496,
}
EXPECT_TOTAL = 31077  # ★ = 主表「欧洲」24581 + 「英国」6496（0 = 不硬断言）

# ---------------- 读渠道销量 ----------------
df = pd.read_csv(os.path.join(tmp_dir, "total_sku_sale_df.csv"), dtype=str)
df.columns = [c.replace("Unnamed: 0", "_idx") for c in df.columns]
assert Q in df.columns, f"中间产物里没有列 {Q}（列：{list(df.columns)}）"
df[Q] = pd.to_numeric(df[Q], errors="coerce").fillna(0)


def fetch(conds):
    """按 (实体,平台,仓库,站点) 过滤（多条取并集）→ {sku: 数量}"""
    mask = None
    for e, p, w, s in conds:
        m = (df["实体"] == e) & (df["平台"] == p) & (df["仓库"] == w) & (df["站点"] == s)
        mask = m if mask is None else (mask | m)
    sub = df[mask].copy()
    sub["sku"] = sub["sku"].astype(str).str.strip()
    return sub.groupby("sku")[Q].sum().astype(int).to_dict()


def add_into(acc, d):
    for k, v in d.items():
        acc[k] = acc.get(k, 0) + v
    return acc


# ---------------- Excel 样式助手 ----------------
def snap(src, dst):
    dst.font = copy(src.font)
    dst.fill = copy(src.fill)
    dst.border = copy(src.border)
    dst.alignment = copy(src.alignment)
    dst.number_format = src.number_format
    dst.protection = copy(src.protection)


def putv(ws, r, c, v, ref_c):
    """写值 + 套用上月块同位置列的样式（当月块在模板里是空的，没有样式）"""
    cell = ws.cell(r, c, v)
    snap(ws.cell(r, ref_c), cell)
    return cell


def ensure_rows(ws, need, base_cols=2):
    """模板缺行的 SKU → 追加行（参考行取最后一条数据行，不取第 2 行合计行）

    base_cols —— 该 sheet 的**基础列数**，各 sheet 不同：
        国家 sheet = 2（A=SKU、B=型号，C 列起就是月块）
        Total 表  = 6（A~F）
    补行只能回填基础列范围内的格子。若一律按 Total 表的 6 列去补国家 sheet，
    会把 C 列（1 月渠道列）覆盖成"类型"文本，后面算累计时直接 TypeError。
    """
    exist = {}
    for r in range(3, ws.max_row + 1):
        s = ws.cell(r, 1).value
        if s:
            exist[str(s).strip()] = r
    add = sorted({str(s).strip() for s in need} - set(exist))
    last = ws.max_row
    for i, s in enumerate(add):
        r = last + 1 + i
        for c in range(1, ws.max_column + 1):
            snap(ws.cell(last, c), ws.cell(r, c))
        ws.row_dimensions[r].height = ws.row_dimensions[last].height or 30
        ws.cell(r, 1, s)
        info = SKU_INFO.get(s)                      # 补行时基础列一起带上（型号/类型/品牌/得分）
        if info:
            for c, v in zip(range(2, base_cols + 1), info):
                if v is not None:
                    ws.cell(r, c, v)
        exist[s] = r
    return exist, add


def apply_outline(wb, blocks, month_offset):
    """列分组（分级显示）：**只有当月展开渠道列，其余 11 个月折叠成「月合计」一列**。
    这是东南亚模板的观感——打开就是一张窄表，SKU / 产品型号永远在视野里，不用横滚。
    blocks = [(sheet 名, 月块首列, 块宽, 渠道列数)]

    同时清掉 sheetView 上的 topLeftCell：它和 <pane state="frozen"> 冲突会导致
    WPS 里冻结窗格失效——表现就是横滚到 8 月块后「SKU / 产品型号两列不见了」。
    """
    for name, base, bw, nch in blocks:
        ws = wb[name]
        for m in range(1, 13):
            c0 = base + (m - 1) * bw
            for k in range(nch):                                  # 渠道列：只有当月露出来
                cd = ws.column_dimensions[gl(c0 + k)]
                cd.outlineLevel = 1
                cd.hidden = (m != month_offset)
            cd = ws.column_dimensions[gl(c0 + nch)]               # 月合计列：永远可见 + 挂折叠标记
            cd.outlineLevel = 0
            cd.hidden = False
            cd.collapsed = True
        ws.sheet_properties.outlinePr.summaryRight = True
        ws.sheet_view.topLeftCell = None


# 浅红底 + 深红字（与东南亚模板的 dxf 同款）
DXF_RED = DifferentialStyle(font=Font(color="FF9C0006"),
                            fill=PatternFill(bgColor="FFFFC7CE"))
ROWS_ALL = 1048576          # 条件格式铺满整列（东南亚就是写到 1048576，插行也不会漏）


def add_formats(wb, blocks):
    """条件格式（照抄东南亚）：A 列重复 SKU 标红 + 渠道列 0 值着色"""
    for name, base, bw, nch in blocks:
        ws = wb[name]
        ws.conditional_formatting.add(f"A1:A{ROWS_ALL}", Rule(type="duplicateValues", dxf=DXF_RED))
        for m in range(1, 13):
            c0 = base + (m - 1) * bw
            if nch:
                ws.conditional_formatting.add(
                    f"{gl(c0)}3:{gl(c0 + nch - 1)}{ROWS_ALL}",
                    Rule(type="cellIs", operator="equal", formula=["0"], dxf=DXF_RED))


def add_notes(wb, notes):
    """给渠道列/国家列表头加批注，说明这一列的数是从哪来的（实体/平台/仓库/站点）"""
    for sheet, items in notes.items():
        ws = wb[sheet]
        for addr, text in items.items():
            ws[addr].comment = Comment(text, "sale_file_gen")


# ---------------- 开填 ----------------
wb = openpyxl.load_workbook(template_path)
log = []
per_country = {}

# [1] 各国 sheet：当月块 = [渠道列…][月合计列]
for sheet, cn, en, chdefs in COUNTRIES:
    ws = wb[sheet]
    nch, bw = len(chdefs), len(chdefs) + 1
    c0 = 3 + (month_offset - 1) * bw
    ch_data = [(nm, fetch(cd)) for nm, cd in chdefs]
    need = set()
    for _nm, d in ch_data:
        need |= set(d)
    _exist, add = ensure_rows(ws, need)
    tot = 0
    for r in range(3, ws.max_row + 1):
        s = ws.cell(r, 1).value
        if not s:
            continue
        s = str(s).strip()
        rowtot = 0
        for k, (_nm, d) in enumerate(ch_data):
            v = int(d.get(s, 0))
            putv(ws, r, c0 + k, v, c0 - bw + k)
            rowtot += v
        putv(ws, r, c0 + nch, rowtot, c0 - bw + nch)
        tot += rowtot
    for k, (_nm, d) in enumerate(ch_data):
        putv(ws, 2, c0 + k, int(sum(d.values())), c0 - bw + k)
    putv(ws, 2, c0 + nch, tot, c0 - bw + nch)
    acc = {}
    for _nm, d in ch_data:
        add_into(acc, d)
    per_country[cn] = acc
    log.append((cn, tot, len(add)))

# [2] Total sheet：当月块 = 各国列 + 月合计；年累计 / 区域累计按"当月块重算"（幂等）
ws = wb[TOTAL_SHEET]
n = len(COUNTRIES)
bw = n + 1
c0 = BASE + n + 1 + (month_offset - 1) * bw
allneed = set()
for d in per_country.values():
    allneed |= set(d)
_e, _add_t = ensure_rows(ws, allneed, base_cols=6)      # Total 表基础列 A~F
area_tot = [0] * n
gtot = 0
for r in range(3, ws.max_row + 1):
    s = ws.cell(r, 1).value
    if not s:
        continue
    s = str(s).strip()
    tot = 0
    for j, (_sh, cn, _en, _cd) in enumerate(COUNTRIES):
        v = int(per_country[cn].get(s, 0))
        putv(ws, r, c0 + j, v, c0 - bw + j)
        tot += v
    putv(ws, r, c0 + n, tot, c0 - bw + n)
    gtot += tot
    for j in range(n):
        acc = sum(ws.cell(r, BASE + n + 1 + (m - 1) * bw + j).value or 0
                  for m in range(1, month_offset + 1))
        ws.cell(r, BASE + 1 + j).value = acc
        area_tot[j] += acc
    ws.cell(r, BASE).value = sum(ws.cell(r, BASE + 1 + j).value or 0 for j in range(n))
for j in range(n):
    ws.cell(2, BASE + 1 + j).value = area_tot[j]
ws.cell(2, BASE).value = sum(area_tot)
for j in range(n):
    ws.cell(2, c0 + j).value = int(sum(per_country[COUNTRIES[j][1]].values()))
ws.cell(2, c0 + n).value = gtot

META = {}                                  # SKU → (类型, 品牌)，交付件用
for r in range(3, ws.max_row + 1):
    s = ws.cell(r, 1).value
    if s:
        META[str(s).strip()] = (ws.cell(r, 3).value, ws.cell(r, 4).value)

# [3] 历史表补当月列
ws = wb[HIST_SHEET]
col = 2 + MONTHS_ALL.index((YEAR, month_offset))
for i, (_sh, cn, _en, _cd) in enumerate(COUNTRIES):
    ws.cell(3 + i, col).value = sum(per_country[cn].values())
ws.cell(2, col).value = sum(ws.cell(3 + i, col).value or 0 for i in range(n))

# [4] 交付件（纯英文表头）
os.makedirs(result_dir, exist_ok=True)
for sheet, cn, en, chdefs in COUNTRIES:
    ws_src = wb[sheet]
    nch, bw = len(chdefs), len(chdefs) + 1
    ch_data = [(nm, fetch(cd)) for nm, cd in chdefs]
    rows = []
    for r in range(3, ws_src.max_row + 1):
        s = ws_src.cell(r, 1).value
        if not s:
            continue
        s = str(s).strip()
        vals = [int(d.get(s, 0)) for _nm, d in ch_data]
        cum = sum(ws_src.cell(r, 3 + (m - 1) * bw + bw - 1).value or 0
                  for m in range(1, month_offset + 1))
        typ, brand = META.get(s, (None, None))
        rows.append((s, ws_src.cell(r, 2).value, typ, brand, vals, sum(vals), cum))
    rows.sort(key=lambda x: -x[5])
    wb2 = openpyxl.Workbook()
    ws2 = wb2.active
    hdr = ["SKU", "Model", "Type", "Brand"] + [nm for nm, _ in chdefs] + \
          [f"{ABBR} {YEAR} Total", f"Jan-{ABBR} {YEAR} Cumulative"]
    ws2.append(hdr)
    for c in range(1, len(hdr) + 1):
        ws2.cell(1, c).font = openpyxl.styles.Font(bold=True)
    for sku, model, typ, brand, vals, tot, cum in rows:
        ws2.append([sku, model, typ, brand] + vals + [tot, cum])
    ws2.column_dimensions["A"].width = 13
    ws2.column_dimensions["B"].width = 66
    for c in ("C", "D"):
        ws2.column_dimensions[c].width = 13
    for c in range(5, len(hdr) + 1):
        ws2.column_dimensions[gl(c)].width = 18
    ws2.freeze_panes = "C2"
    wb2.save(os.path.join(result_dir, f"{en.replace(' ', '')}_{belonging_date}_sales.xlsx"))

# [5] 列分组（只展开当月）+ 条件格式 + 表头批注（东南亚同款观感）
BLOCKS = [(sh, 3, len(cd) + 1, len(cd)) for sh, _cn, _en, cd in COUNTRIES] \
         + [(TOTAL_SHEET, BASE + n + 1, n + 1, n)]
apply_outline(wb, BLOCKS, month_offset)
add_formats(wb, BLOCKS)

SPARK_COL = gl(2 + len(MONTHS_ALL) + 1)     # 历史表：国名 1 列 + 24 个月 → 迷你图落在第 26 列 Z
wb[HIST_SHEET].column_dimensions[SPARK_COL].width = 49.625

_notes = {}
for _sh, cn, _en, cd in COUNTRIES:
    _b = 3 + (month_offset - 1) * (len(cd) + 1)                     # 各国 sheet 当月块首列
    _notes[_sh] = {f"{gl(_b + k)}1": f"{cn}：{nm}（{' / '.join('·'.join(t) for t in cond)}）"
                   for k, (nm, cond) in enumerate(cd)}
_tc0 = BASE + n + 1 + (month_offset - 1) * (n + 1)                  # Total 表当月块首列
_notes[TOTAL_SHEET] = {f"{gl(_tc0 + j)}1": f"{cn}：{' + '.join(nm for nm, _ in cd)}"
                       for j, (_sh, cn, _en, cd) in enumerate(COUNTRIES)}
add_notes(wb, _notes)

wb.calculation.fullCalcOnLoad = True    # A1 是跨表引用公式，新文件无缓存值 → 让 Excel/WPS 打开即重算
wb.save(result_path)

# [6] 观感保真：迷你图 + WPS 专有部件（阅读模式黄十字）+ 修冻结窗格冲突
apply_fidelity(result_path, spark_specs=[
    dict(sheet=HIST_SHEET, col=SPARK_COL, r1=3, r2=2 + n, c1="B", c2=gl(1 + len(MONTHS_ALL))),
])

# [5] 校验
print("=" * 64)
for cn, tot, na in log:
    exp = EXPECT.get(cn)
    flag = "✅" if (exp is None or tot == exp) else f"❌ 期望 {exp}"
    print(f"  {cn:<6} {month_offset}月 {tot:>6}  新增行 {na}  {flag}")
print("-" * 64)
print(f"  {n} 国合计 {gtot}  | 期望 {EXPECT_TOTAL} {'✅' if gtot == EXPECT_TOTAL else '❌'}")
assert EXPECT_TOTAL == 0 or gtot == EXPECT_TOTAL, f"欧洲 {month_offset} 月合计 {gtot} ≠ {EXPECT_TOTAL}"
print(f"✅ 已写入 {result_path}")
print(f"✅ 交付件 {n} 份 → {result_dir}")
_lines, _ok = check_fidelity(result_path)
print("-" * 64)
print("  观感保真自检：" + ("全部通过" if _ok else "有未通过项 ⚠️"))
for _l in _lines:
    print("   ", _l)
if AUTO_OPEN:
    os.startfile(result_path)
