# -*- coding: utf-8 -*-
"""
东南亚销量表 8 月填报脚本（仿 month_sale_file.py 的写法）

输入：
  1) 模板      data/数据归属周期/<归属月>/快照日期/<快照日>/东南亚/2026年东南亚销量表Southeast Asia Sales table.xlsx
  2) 渠道销量  sale_file_tmp/se_a_sku_sale_df.csv  ← month_sale_file.py 的中间产物
               （列：实体/平台/仓库/站点/sku/<归属月>销量；已含状态过滤与 SKU 归集，与月销总表同源）
  3) 跨境虾皮  data/.../东南亚/跨境/Shopee跨境{马来西亚,越南}站点*订单报告.xlsx
               （用于「虾皮跨境」sheet 的 8 国拆分；新加坡站的订单归 Singapore sheet，不在此重复计入）
  4) 商品资料  商品资料.xlsx（补新行位的「型号」，以及交付文件的产品属性）

输出：
  result/东南亚销量表/<归属月>/2026年东南亚销量表Southeast Asia Sales table.xlsx（不覆盖模板）
  result/东南亚销量表/<归属月>/{Philippines,Indonesia,Thailand}_<归属月>_sales.xlsx（纯英文表头交付件）

口径要点（详见《东南亚销量表填报-需求对齐.md》）：
  * 月度块布局 = [渠道列 …][月合计列]，**月合计列在块尾**；历史已填月的块里**没有空格**（无销量写 0）
  * 东南亚销量表口径与月销总表**同源**：8 月 8 列之和 = 月销总表「东南亚合计」
  * 「虾皮跨境」sheet 的「新加坡销量」列 = **0**（SG 的虾皮订单记在 Singapore sheet 的「虾皮销量」列）
  * 表头**中英双语**（9.16 起）：主表保存前把「只有中文」的表头补成「中文 + English」，
    并把表头行高撑到能放下（见 §5.5）。幂等，反复跑不会叠加。
"""
import os
import re
import glob
import datetime
from copy import copy

import numpy as np
import pandas as pd
import openpyxl
from openpyxl.utils import get_column_letter

# ==================== 月度参数（每月只改这三行） ====================
belonging_date = "202610"      # 数据归属周期 YYYYMM
snapshot_date = "20261102"     # 快照日期 YYYYMMDD
month_offset = 10              # 当月月份 1~12
# ===================================================================

cur_acct_cycle_date = datetime.datetime.strptime(belonging_date, "%Y%m").date()
month_col_prefix = cur_acct_cycle_date.strftime("%Y%m")

trunk_dir = os.path.dirname(os.path.abspath(__file__))
tmp_dir = os.path.join(trunk_dir, "sale_file_tmp")
_snap_dir = os.path.join(trunk_dir, "data", "数据归属周期", belonging_date, "快照日期", snapshot_date)
# 快照布局有两种：旧版把东南亚资料放在 <快照日>/东南亚/ 子目录下；新版（20260915 起）四个总表直接平铺在 <快照日>/ 下
se_a_data_dir = os.path.join(_snap_dir, "东南亚") if os.path.isdir(os.path.join(_snap_dir, "东南亚")) else _snap_dir
cross_data_dir = os.path.join(se_a_data_dir, "跨境")
template_path = os.path.join(se_a_data_dir, "2026年东南亚销量表Southeast Asia Sales table.xlsx")
product_info_path = os.path.join(trunk_dir, "商品资料.xlsx")
result_dir = os.path.join(trunk_dir, "result", "东南亚销量表", belonging_date)

# ---------------- 渠道定义：每列对应 (实体, 平台, 站点) 过滤条件 ----------------
# None = 该渠道 data 里无数据源，恒填 0
PH_CHANNELS = [
    ("中山仓lazada", ("跨境", "Lazada", "菲律宾")),
    ("菲仓lazada", ("本土", "Lazada", "菲律宾")),
    ("菲仓shopee", ("本土", "Shopee", "菲律宾")),
    ("Tiktok", ("本土", "Tiktok", "菲律宾")),
    ("线下订单", ("本土", "KSON", "菲律宾")),
]
TH_CHANNELS = [
    ("shopee易海", ("本土", "Shopee", "泰国")),
    ("LAZADA易海", ("本土", "Lazada", "泰国")),
    ("LAZADA泰国仓", None),                     # 无数据源，恒 0
    ("LAZADA中山仓", ("跨境", "Lazada", "泰国")),
]
ID_CHANNELS = [
    ("shopee智慧仓", ("本土", "Shopee", "印尼")),
    ("印尼线下", ("本土", "线下", "印尼")),
    ("Tiktok", ("本土", "Tiktok", "印尼")),
]
SG_CHANNELS = [
    ("LAZADA", ("跨境", "Lazada", "新加坡")),
    ("虾皮", ("跨境", "Shopee", "新加坡")),
]
VN_CHANNELS = [("越南销量", ("跨境", "Lazada", "越南"))]
MYLZ_CHANNELS = [("马来Lazada", ("跨境", "Lazada", "马来西亚"))]
# 「虾皮跨境」sheet 的 8 国：由跨境 Shopee 报告的 Country 列拆分，单独处理
MSHOPEE_COUNTRIES = ["马来西亚", "巴西", "台湾", "墨西哥", "新加坡", "越南", "哥伦比亚", "智利"]

# sheet → dict(基础列, 渠道列表, 是否带月合计列)
SHEET_CONF = {
    "Philippines 菲律宾销量表": dict(base_col=3, channels=PH_CHANNELS, has_total=True),
    "Thailand 泰国销量表": dict(base_col=3, channels=TH_CHANNELS, has_total=True),
    "Indonesia 印尼销量表": dict(base_col=3, channels=ID_CHANNELS, has_total=True),
    "Singapore 新加坡销量表": dict(base_col=3, channels=SG_CHANNELS, has_total=True),
    "Vietnam Lazada 越南销销量表": dict(base_col=3, channels=VN_CHANNELS, has_total=False),
    "Malaysia Lazada 马来西亚销量表": dict(base_col=3, channels=MYLZ_CHANNELS, has_total=False),
}   # 「团购经销商」无渠道列（单列、恒 0），单独处理
MSHOPEE_SHEET = "Malaysia Shopee 虾皮跨境销量表"
SEA_TOTAL_SHEET = "SEA Total 东南亚总销量表"
# SEA Total 的 8 个口径列（顺序与模板一致）
SEA_TOTAL_COLS = ["菲律宾", "泰国", "印尼", "新加坡", "越南", "马来Lazada", "虾皮跨境", "团购"]


def block_start(base_col, width, month):
    """月度块首列（1-based）：块 = [渠道列…][月合计列]"""
    return base_col + (month - 1) * width


def copy_row_style(ws, src_row, dst_row):
    """把 src_row 的整行样式（含行高）复制到 dst_row"""
    for c in range(1, ws.max_column + 1):
        ws.cell(dst_row, c)._style = copy(ws.cell(src_row, c)._style)
    if ws.row_dimensions[src_row].height is not None:
        ws.row_dimensions[dst_row].height = ws.row_dimensions[src_row].height


def get_data_rows(ws):
    """返回 (sku → 行号) 映射 与 最后一个数据行行号；A 列为空的行视为结束"""
    sku_row_map = {}
    last = 2
    for r in range(3, ws.max_row + 1):
        v = ws.cell(r, 1).value
        if v is None or str(v).strip() == "":
            continue
        key = str(v).strip()
        if key not in sku_row_map:
            sku_row_map[key] = r
        last = r
    return sku_row_map, last


def read_shopee_cross_by_country():
    """读跨境 Shopee 报告（取马来西亚站 + 越南站 + 台湾站），按 Country 拆国家 → {国家列名: {sku: 数量}}
    新加坡站的订单归 Singapore sheet，不在此处计入（避免 119 双计）。
    台湾站 2026-11-02 起纳入（并入月销「虾皮境内」，本表单列在「中国台湾销量」列）。"""
    result = {c: {} for c in MSHOPEE_COUNTRIES}
    code_to_col = {
        "MY": "马来西亚", "BN": "马来西亚",   # 文莱模板无列，并入马来西亚
        "BR": "巴西", "TW": "台湾", "MX": "墨西哥",
        "VN": "越南", "CO": "哥伦比亚", "CL": "智利",
        "SG": None,                            # 新加坡站不在本表口径内
    }
    files = sorted(glob.glob(os.path.join(cross_data_dir, "Shopee跨境*站点*订单报告.xlsx")))
    # 2026-11-02 修正（二）：契约扩为「马来西亚站 + 越南站 + **台湾站**」。
    #   · 2026-11-02 第一次修正曾把台湾一起排除，理由是「月销总表无台湾列」——**该结论有误**：
    #     月销总表的「虾皮境内」= Shopee 跨境站点合计，历史上一直含台湾（1~2 月 543 = 马来 527 + 台湾 16
    #     可逐格验证）；且本表《虾皮跨境》sheet 本来就有「中国台湾销量 Taiwan, China Sales」列（1~7 月
    #     依次 16/3/10/10/8/5/7 件，8~9 月因未导出台湾报告才是 0）。
    #   · 墨西哥 / 巴西 仍是**目的地导出**且本期是 0 列空壳（由下面的空表判断跳过），不参与。
    #   · 用户 2026-11-02 拍板：台湾并入「虾皮境内」，不新增列组。
    _CROSS_SITES = ("马来西亚", "越南", "台湾")
    _all_files = files
    files = [f for f in files if any(s in os.path.basename(f) for s in _CROSS_SITES)]
    _excluded = [os.path.basename(f) for f in _all_files if f not in files and "新加坡" not in os.path.basename(f)]
    if _excluded:
        print(f"   ⚠️ 目录里有非本表目的地/非本表站点的跨境 Shopee 报告，按契约排除：{_excluded}")
    files = [f for f in files if "新加坡" not in os.path.basename(f)]
    if not files:
        # 新版快照（20260915 起）不再单独给跨境 Shopee 报告 → 「虾皮跨境」sheet 直接取中间产物里的跨境 Shopee 站点
        sites = ["马来西亚", "越南", "台湾"]   # 2026-11-02：台湾并入中间产物后走同一条兜底
        for _s in sites:
            _m = series_to_dict(("跨境", "Shopee", _s))
            if _m:
                for _sku, _v in _m.items():
                    result[_s][_sku] = result[_s].get(_sku, 0) + _v
        print(f"    ⚠️ 未找到跨境 Shopee 报告，改用中间产物跨境 Shopee 站点：{sites}")
        return result
    for f in files:
        df = pd.read_excel(f)
        # 2026-11-02 新增：本期多了「巴西/墨西哥站点」两份**空表**（0 列 0 行，5,640B 的空壳导出）。
        # 不加这道判断，下面的 next(...) 会抛 StopIteration，整张表都出不来。
        if df.shape[1] == 0 or len(df) == 0:
            print(f"  跳过 {os.path.basename(f)} —— 空表（0 列 / 0 行），该国本月按 0 计")
            continue
        df.columns = [str(c).strip() for c in df.columns]
        sku_col = next(c for c in df.columns if c in ("SKU Reference No.", "商品规格货号"))
        # 数量列名随导出语言/站点变：马来西亚·新加坡=Quantity(英)，越南=数量(中)，台湾=Quantity per Item。
        qty_col = next(c for c in df.columns if c in ("Quantity", "数量", "Quantity per Item"))
        st_col = next(c for c in df.columns if c in ("Order Status", "订单状态"))
        cty_col = next(c for c in df.columns if c in ("Country", "国家"))
        # 与 month_sale_file.py 同口径：剔除取消单（「退货窗口期」保留）
        # ★ 2026-11-02：补「不成立」——台湾站（中文表头）取消单的状态就是「不成立」
        #   （买家取消/付款失败，导出里「取消原因」有值）。不补这条，202610 台湾会被算成 17 件，
        #   真销量只有 6 件（已完成 5 + 已送達 1）。
        raw_status = df[st_col].astype(str)
        keep = ~raw_status.str.contains("Cancel|取消|ยกเลิก|Đã hủy|Batal|不成立", regex=True, case=False)
        d = df[keep].copy()
        d["_sku"] = d[sku_col].astype(str).str.strip()
        d["_cty"] = d[cty_col].astype(str).str.strip()
        print(f"  读取 {os.path.basename(f)}  行 {len(df)} → 过滤后 {len(d)}  件 {d[qty_col].sum():.0f}")
        for cty, grp in d.groupby("_cty"):
            col = code_to_col.get(cty)
            if col is None:
                print(f"    ⚠️ 国家码 {cty} 未在本表列内，跳过（{grp[qty_col].sum():.0f} 件）")
                continue
            g = grp.groupby("_sku")[qty_col].sum()
            for sku, v in g.items():
                result[col][sku] = result[col].get(sku, 0) + v
    return result


print("=" * 70)
print(f"东南亚销量表填报  归属月={belonging_date}  快照日={snapshot_date}  月份={month_offset}")
print("=" * 70)

# ---------------- 1. 读渠道级销量（month_sale_file.py 的中间产物） ----------------
channel_csv = os.path.join(tmp_dir, "se_a_sku_sale_df.csv")
assert os.path.exists(channel_csv), f"缺少中间产物 {channel_csv}，请先跑 month_sale_file.py"
sale_df = pd.read_csv(channel_csv)
sale_col = f"{month_col_prefix}销量"
assert sale_col in sale_df.columns, f"{channel_csv} 里没有列 {sale_col}"
sale_df["实体"] = sale_df["实体"].astype(str).str.strip()
sale_df["平台"] = sale_df["平台"].astype(str).str.strip()
sale_df["站点"] = sale_df["站点"].astype(str).str.strip()
sale_df["sku"] = sale_df["sku"].astype(str).str.strip()

# 2026-11-02：把「台湾」纳入本表的站点集合——它是 Shopee 跨境站点，销量落在《虾皮跨境》sheet 的
# 「中国台湾销量」列、月销总表并进「虾皮境内」列。不纳入这张清单，se_total 会少 6 件，
# 第 [5] 步「各 sheet 之和 == 中间产物 6 站合计」必然报错（校验会挡住，不会静默出错）。
SE_A_SITES = ["菲律宾", "泰国", "印尼", "新加坡", "马来西亚", "越南", "台湾"]
se_df = sale_df[sale_df["站点"].isin(SE_A_SITES)].copy()
se_total = float(se_df[sale_col].sum())
print(f"\n[1] 渠道级销量：{len(sale_df)} 行，东南亚 6 站合计 = {se_total:.0f}")

# (实体,平台,站点) → {sku: 数量}
chan_maps = {}
for key, grp in se_df.groupby(["实体", "平台", "站点"]):
    chan_maps[key] = grp.groupby("sku")[sale_col].sum().to_dict()


def series_to_dict(key):
    return chan_maps.get(key, {})


# ---------------- 2. 虾皮跨境 8 国拆分 ----------------
print("\n[2] 读取跨境 Shopee 报告（按 Country 拆国家）")
ms_shopee = read_shopee_cross_by_country()
ms_shopee_check = {k: float(sum(v.values())) for k, v in ms_shopee.items()}
print(f"    虾皮跨境 8 国合计 = {sum(ms_shopee_check.values()):.0f}  {ms_shopee_check}")
# 与渠道级中间产物对撞：虾皮跨境 = 跨境 Shopee 马来西亚站 + 越南站（新加坡站归 Singapore sheet）
for _cty, _site in (("马来西亚", "马来西亚"), ("越南", "越南"), ("台湾", "台湾")):
    _a = ms_shopee_check[_cty]
    _b = float(sum(series_to_dict(("跨境", "Shopee", _site)).values()))
    assert abs(_a - _b) < 1e-6, f"虾皮跨境 {_cty} {_a} ≠ 中间产物 {_b}"
    print(f"    ✅ {_cty}: 按 Country 拆分 {_a:.0f} = 中间产物跨境 Shopee {_site} {_b:.0f}")

# ---------------- 3. 打开模板副本，逐 sheet 写 {month_offset} 月块 ----------------
assert os.path.exists(template_path), f"找不到模板 {template_path}"
os.makedirs(result_dir, exist_ok=True)
result_path = os.path.join(result_dir, os.path.basename(template_path))
wb = openpyxl.load_workbook(template_path, data_only=False)

# 商品资料（补新行位的型号 + 交付文件属性）
product_info_df = pd.read_excel(product_info_path, sheet_name="商品资料", skiprows=1)
product_info_df["SKU"] = product_info_df["SKU"].astype(str).str.strip()
for _c in ("型号", "类型", "品牌"):
    if _c not in product_info_df.columns:
        product_info_df[_c] = None
model_map = product_info_df.set_index("SKU")["型号"].to_dict()

SHEET_HANDLES = {nm: wb[nm] for nm in wb.sheetnames}
# 新版模板（20260915 起）sheet 名可能省掉中英文后缀，做一次别名映射
_SHEET_ALIASES = {
    "SEA Total 东南亚总销量表": ["SEA Total", "SEA Total 东南亚总销量"],
    "Philippines 菲律宾销量表": ["Philippines", "Philippines 菲律宾"],
    "Thailand 泰国销量表": ["Thailand", "Thailand 泰国"],
    "Indonesia 印尼销量表": ["Indonesia", "Indonesia 印尼"],
    "Singapore 新加坡销量表": ["Singapore", "Singapore 新加坡"],
    "Vietnam Lazada 越南销销量表": ["Vietnam Lazada", "Vietnam"],
    "Malaysia Lazada 马来西亚销量表": ["Malaysia Lazada"],
    "Malaysia Shopee 虾皮跨境销量表": ["Malaysia Shopee", "虾皮跨境销量表"],
    "Group buying distributor团购经销商": ["Group buying distributor", "团购经销商"],
}
_CONF_KEYS = list(SHEET_CONF.keys()) + [MSHOPEE_SHEET, "Group buying distributor团购经销商"]
for _std in _CONF_KEYS:
    if _std in SHEET_HANDLES:
        continue
    for _alt in _SHEET_ALIASES.get(_std, []):
        _hit = next((nm for nm in wb.sheetnames if _alt in nm), None)
        if _hit:
            SHEET_HANDLES[_std] = wb[_hit]
            print(f"    ⚠️ sheet 名别名：{_std} → {_hit}")
            break
    assert _std in SHEET_HANDLES, f"模板缺少 sheet：{_std}"

appended_log = {}
sheet_sku_values = {}      # sheet → {sku: {列序号: 值}}
sheet_month_totals = {}

for sheet_name, conf in SHEET_CONF.items():
    ws = SHEET_HANDLES[sheet_name]
    ch = conf["channels"]
    width = len(ch) + (1 if conf["has_total"] else 0)
    start = block_start(conf["base_col"], width, month_offset)
    sku_row_map, last_row = get_data_rows(ws)

    # 本 sheet 的每列值
    col_maps = []
    for name, key in ch:
        col_maps.append(series_to_dict(key) if key else {})
    # 本 sheet 有销量的 SKU（用于补行位）
    sold_skus = set()
    for m in col_maps:
        sold_skus |= set(m.keys())
    missing = sorted(s for s in sold_skus if s not in sku_row_map)

    # 补行位：复制最后一行样式
    for s in missing:
        last_row += 1
        ws.cell(last_row, 1).value = s
        ws.cell(last_row, 2).value = model_map.get(s)
        copy_row_style(ws, last_row - 1, last_row)
        sku_row_map[s] = last_row
    if missing:
        appended_log[sheet_name] = missing
        print(f"    {sheet_name}: 补 {len(missing)} 个行位 {missing}")

    # 写值：历史已填块是「无销量写 0」，{month_offset} 月块保持同一风格 → 全行写满
    row_vals = {s: [float(m.get(s, 0)) for m in col_maps] for s in sku_row_map}
    for s, vals in row_vals.items():
        r = sku_row_map[s]
        for i, v in enumerate(vals):
            ws.cell(r, start + i).value = v
        if conf["has_total"]:
            ws.cell(r, start + len(ch)).value = float(sum(vals))

    # r2 合计行
    data_rows = sorted(sku_row_map.values())
    for i in range(width):
        c = start + i
        ws.cell(2, c).value = float(sum(ws.cell(r, c).value or 0 for r in data_rows))

    sheet_sku_values[sheet_name] = {s: {start + i: v for i, v in enumerate(vals)} for s, vals in row_vals.items()}
    # 月合计 = 渠道列之和（不含块尾的合计列本身）
    sheet_month_totals[sheet_name] = float(sum(ws.cell(2, start + i).value or 0 for i in range(len(ch))))
    print(f"[3] {sheet_name:<40} {month_offset}月块 {get_column_letter(start)}~{get_column_letter(start + width - 1)}"
          f"  渠道 {len(ch)} 列  合计 = {sheet_month_totals[sheet_name]:.0f}")

# ---- 虾皮跨境 sheet（8 国拆分）----
ws = SHEET_HANDLES[MSHOPEE_SHEET]
width = len(MSHOPEE_COUNTRIES) + 1
start = block_start(3, width, month_offset)
sku_row_map, last_row = get_data_rows(ws)
sold_skus = set()
for c in MSHOPEE_COUNTRIES:
    sold_skus |= set(ms_shopee[c].keys())
missing = sorted(s for s in sold_skus if s not in sku_row_map)
for s in missing:
    last_row += 1
    ws.cell(last_row, 1).value = s
    ws.cell(last_row, 2).value = model_map.get(s)
    copy_row_style(ws, last_row - 1, last_row)
    sku_row_map[s] = last_row
if missing:
    appended_log[MSHOPEE_SHEET] = missing
    print(f"    {MSHOPEE_SHEET}: 补 {len(missing)} 个行位 {missing}")

ms_row_vals = {}
for s, r in sku_row_map.items():
    vals = [float(ms_shopee[c].get(s, 0)) for c in MSHOPEE_COUNTRIES]
    ms_row_vals[s] = vals
    for i, v in enumerate(vals):
        ws.cell(r, start + i).value = v
    ws.cell(r, start + len(MSHOPEE_COUNTRIES)).value = float(sum(vals))
data_rows = sorted(sku_row_map.values())
for i in range(width):
    c = start + i
    ws.cell(2, c).value = float(sum(ws.cell(r, c).value or 0 for r in data_rows))
sheet_month_totals[MSHOPEE_SHEET] = float(sum(ws.cell(2, start + i).value or 0 for i in range(len(MSHOPEE_COUNTRIES))))
print(f"[3] {MSHOPEE_SHEET:<40} {month_offset}月块 {get_column_letter(start)}~{get_column_letter(start + width - 1)}"
      f"  8 国  合计 = {sheet_month_totals[MSHOPEE_SHEET]:.0f}")

# ---- 团购 sheet（无数据源，写 0）----
ws = SHEET_HANDLES["Group buying distributor团购经销商"]
start = block_start(3, 1, month_offset)
sku_row_map, last_row = get_data_rows(ws)
for s, r in sku_row_map.items():
    ws.cell(r, start).value = 0
data_rows = sorted(sku_row_map.values())
ws.cell(2, start).value = 0.0
sheet_month_totals["Group buying distributor团购经销商"] = 0.0
print(f"[3] {'Group buying distributor团购经销商':<40} {month_offset}月块 {get_column_letter(start)}  合计 = 0")

# ---------------- 4. SEA Total（8 国/渠道汇总）----------------
# 每个口径列 = 对应 sheet 的「每 SKU 合计」
per_country = {
    "菲律宾": {s: sum(v.values()) for s, v in sheet_sku_values["Philippines 菲律宾销量表"].items()},
    "泰国": {s: sum(v.values()) for s, v in sheet_sku_values["Thailand 泰国销量表"].items()},
    "印尼": {s: sum(v.values()) for s, v in sheet_sku_values["Indonesia 印尼销量表"].items()},
    "新加坡": {s: sum(v.values()) for s, v in sheet_sku_values["Singapore 新加坡销量表"].items()},
    "越南": {s: sum(v.values()) for s, v in sheet_sku_values["Vietnam Lazada 越南销销量表"].items()},
    "马来Lazada": {s: sum(v.values()) for s, v in sheet_sku_values["Malaysia Lazada 马来西亚销量表"].items()},
    "虾皮跨境": {s: sum(v) for s, v in ms_row_vals.items()},
    "团购": {},
}

ws = SHEET_HANDLES[SEA_TOTAL_SHEET]
# 2026-10-16 修订：SEA Total 月块 base 原误用 7（G 列，年度区"菲律宾"列），且把合计列当块尾。
# 模板真实结构：月块从 O(15) 起、块宽 9、**合计列在块首**，后跟 8 个国家列
# （[合计, 菲, 泰, 印尼, 新, 越, 马来Lazada, 虾皮跨境, 团购]，表头见 O1/P1/…）。
# 旧写法 7+(m-1)*9 使国家列落到 m-1 月块、当月块国家列恒 0（合计列因 +8 巧合撞对）。
start = block_start(15, 9, month_offset)
sku_row_map, last_row = get_data_rows(ws)
sold_skus = set()
for c in SEA_TOTAL_COLS:
    sold_skus |= set(per_country[c].keys())
missing = sorted(s for s in sold_skus if s not in sku_row_map)
for s in missing:
    last_row += 1
    ws.cell(last_row, 1).value = s
    ws.cell(last_row, 2).value = model_map.get(s)
    copy_row_style(ws, last_row - 1, last_row)
    sku_row_map[s] = last_row
if missing:
    appended_log[SEA_TOTAL_SHEET] = missing
    print(f"    {SEA_TOTAL_SHEET}: 补 {len(missing)} 个行位 {missing}")

for s, r in sku_row_map.items():
    vals = [float(per_country[c].get(s, 0)) for c in SEA_TOTAL_COLS]
    ws.cell(r, start).value = float(sum(vals))          # 2026-10-16：合计列在块首
    for i, v in enumerate(vals):
        ws.cell(r, start + 1 + i).value = v             # 2026-10-16：8 国家列紧随其后
data_rows = sorted(sku_row_map.values())
for i in range(9):
    c = start + i
    ws.cell(2, c).value = float(sum(ws.cell(r, c).value or 0 for r in data_rows))
sea_total_month = float(sum(ws.cell(2, start + 1 + i).value or 0 for i in range(8)))  # 2026-10-16：8 国和（不含合计列）
print(f"[4] {SEA_TOTAL_SHEET:<40} {month_offset}月块 {get_column_letter(start)}~{get_column_letter(start + 8)}"
      f"  8 国  合计 = {sea_total_month:.0f}")

# ---------------- 5. 校验 ----------------
print("\n[5] 校验")
errs = []
# 按口径列名整理（顺序与 SEA Total 一致）
country_totals = {
    "菲律宾": sheet_month_totals["Philippines 菲律宾销量表"],
    "泰国": sheet_month_totals["Thailand 泰国销量表"],
    "印尼": sheet_month_totals["Indonesia 印尼销量表"],
    "新加坡": sheet_month_totals["Singapore 新加坡销量表"],
    "越南": sheet_month_totals["Vietnam Lazada 越南销销量表"],
    "马来Lazada": sheet_month_totals["Malaysia Lazada 马来西亚销量表"],
    "虾皮跨境": sheet_month_totals[MSHOPEE_SHEET],
    "团购": sheet_month_totals["Group buying distributor团购经销商"],
}
# 5.1 各 sheet 月合计之和 = 中间产物 6 站合计 = SEA Total {month_offset} 月块和
sum_sheets = float(sum(country_totals.values()))
print(f"    8 个 sheet 月合计之和         = {sum_sheets:.0f}")
print(f"    中间产物 东南亚 6 站合计      = {se_total:.0f}")
print(f"    SEA Total {month_offset} 月块和            = {sea_total_month:.0f}")
if abs(sum_sheets - se_total) > 1e-6:
    errs.append(f"各 sheet 之和 {sum_sheets} ≠ 中间产物合计 {se_total}")
if abs(sea_total_month - se_total) > 1e-6:
    errs.append(f"SEA Total {month_offset} 月块和 {sea_total_month} ≠ 中间产物合计 {se_total}")

# 5.2 与月销总表成品逐国对照
mst_path = os.path.join(trunk_dir, "result", "数据归属周期", belonging_date, "月销售总表.xlsx")
if os.path.exists(mst_path):
    mst = pd.read_excel(mst_path, sheet_name=f"{cur_acct_cycle_date.year}年", header=None)
    MST_GRP = ["欧洲", "英国", "美国", "加拿大", "日本", "澳洲", "阿联酋", "沙特", "墨西哥", "印度", "亚马逊合计",
               "菲律宾", "泰国", "印尼", "新加坡", "马来", "越南", "虾皮境内", "东南亚合计"]
    def mst_val(group, month):
        gi = MST_GRP.index(group)
        return pd.to_numeric(mst.iloc[3:, 3 + gi * 12 + (month - 1) - 1], errors="coerce").sum()
    pairs = [("菲律宾", "菲律宾"), ("泰国", "泰国"), ("印尼", "印尼"), ("新加坡", "新加坡"),
             ("马来Lazada", "马来"), ("越南", "越南"), ("虾皮跨境", "虾皮境内")]
    print(f"    {'口径列':<10}{'本表':>10}{'月销总表':>10}{'差额':>8}")
    for a, b in pairs:
        va, vb = country_totals[a], mst_val(b, month_offset)
        print(f"    {a:<10}{va:>10.0f}{vb:>10.0f}{va - vb:>8.0f}")
        if abs(va - vb) > 1e-6:
            errs.append(f"{a} 本表 {va} ≠ 月销总表 {b} {vb}")
    va, vb = sea_total_month, mst_val("东南亚合计", month_offset)
    print(f"    {'合计':<10}{va:>10.0f}{vb:>10.0f}{va - vb:>8.0f}")
    if abs(va - vb) > 1e-6:
        errs.append(f"本表合计 {va} ≠ 月销总表东南亚合计 {vb}")
else:
    print(f"    ⚠️ 未找到月销总表成品 {mst_path}，跳过逐国对照")

if errs:
    for e in errs:
        print(f"    ❌ {e}")
    raise AssertionError("校验未通过：" + "；".join(errs))
print("    ✅ 校验通过")

# ---------------- 5.5 表头中英双语化 ----------------
# 起因：模板里相当一部分表头是**纯中文**（尤其泰/印尼/新/越/马/虾皮跨境/团购这几个 sheet 的
# 月度列与渠道列），而这份表要发给对应国家的同事看，看不懂中文。
# 做法：保存前把「只有中文」的表头补成「中文 + English」，风格与模板里**本已双语**的部分保持一致
#       （如菲律宾 sheet 的「1月菲律宾总销量 PHILIPINES January TOTAL SALES」）。
# ⚠️ 幂等：每条规则都要求中文串**以固定中文结尾**（如 `…总销量$` / `…销量$`），
#    翻完的串尾部变成英文，下轮不会再匹配 → 反复跑不会叠加。
# ⚠️ 只翻「已登记」的串；遇到没见过的中文表头**保持原样**（宁可漏翻，也不乱翻）。
_MONTH_EN = dict(enumerate(
    ["January", "February", "March", "April", "May", "June",
     "July", "August", "September", "October", "November", "December"], start=1))

# 「<中文主体>销量 / <中文主体>总销量」里主体的英文
_SITE_EN = {
    "菲律宾": "Philippines", "泰国": "Thailand", "印尼": "Indonesia",
    "新加坡": "Singapore", "越南": "Vietnam", "马来": "Malaysia",
    "马来西亚": "Malaysia", "巴西": "Brazil", "墨西哥": "Mexico",
    "哥伦比亚": "Colombia", "智利": "Chile",
    "台湾": "Taiwan, China",                       # 中国台湾
    "虾皮": "Shopee", "LAZADA": "LAZADA",
    "虾皮跨境": "Shopee Cross-border", "团购经销商": "Group Buy Distributor",
}
# 不随月份变化的整串（渠道列 / 静态列）。键必须与单元格里 **strip 之后** 的原文完全一致。
_HEADER_EXACT = {
    "类型": "类型 Type",
    "品牌": "品牌 Brand",
    "数量合计\n产品型号": "数量合计 Total Qty\n产品型号 Model",
    # 泰国渠道列
    "shopee泰国销量（易海）": "shopee泰国销量（易海） Shopee (Yihai)",
    "LAZADA销量（易海）": "LAZADA销量（易海） LAZADA (Yihai)",
    "LAZADA泰国仓销量": "LAZADA泰国仓销量 LAZADA Thailand Warehouse",
    "LAZADA中山仓销量": "LAZADA中山仓销量 LAZADA Zhongshan Warehouse",
    # 印尼渠道列
    "shopee销量（智慧仓）": "shopee销量（智慧仓） Shopee (Smart Warehouse)",
    # 新加坡渠道列
    "LAZADA销量": "LAZADA销量 LAZADA Sales",
    "虾皮销量": "虾皮销量 Shopee Sales",
    # 虾皮跨境 8 国列
    "马来西亚销量": "马来西亚销量 Malaysia Sales",
    "巴西销量": "巴西销量 Brazil Sales",
    "台湾销量": "中国台湾销量 Taiwan, China Sales",
    "墨西哥销量": "墨西哥销量 Mexico Sales",
    "新加坡销量": "新加坡销量 Singapore Sales",
    "越南销量": "越南销量 Vietnam Sales",
    "哥伦比亚": "哥伦比亚销量 Colombia Sales",
    "智利销量": "智利销量 Chile Sales",
}
_YEAR_RE = re.compile(r"^(\d{4})年$")                       # 2026年
_YEAR_TOTAL_RE = re.compile(r"^(\d{4})年总销量$")            # 2026年总销量
_SCORE_RE = re.compile(r"^(\d{1,2})月销量得分$")             # 7月销量得分
# 「[前缀]N月[站点]总销量」：前缀空（泰/印尼/新）或「虾皮跨境」
_TOTAL_RE = re.compile(r"^(?P<pre>[^0-9]*)(?P<m>\d{1,2})月(?P<mid>[^0-9]*)总销量$")
# 「N月<站点>销量」：越南 / 马来 / 团购经销商
_SALES_RE = re.compile(r"^(?P<m>\d{1,2})月(?P<site>[^0-9]*)销量$")


def i18n_header(text):
    """把「纯中文」表头翻成「中文 + English」；公式 / 纯数字 / 未登记 / 已翻过的 → 原样返回。"""
    if not isinstance(text, str):
        return text
    core = text.strip()
    if not core or core.startswith("="):
        return text                       # 公式不动
    head = text[:len(text) - len(text.lstrip())]      # 保留原前导空白（B2 有很多空格）
    tail = text[len(text.rstrip()):]
    m = _YEAR_RE.match(core)
    if m:
        return f"{head}{core} Year {m.group(1)}{tail}"
    m = _YEAR_TOTAL_RE.match(core)
    if m:
        return f"{head}{core} Total Sales {m.group(1)}{tail}"
    m = _SCORE_RE.match(core)
    if m:
        return f"{head}{core} {_MONTH_EN[int(m.group(1))]} Sales Score{tail}"
    if core in _HEADER_EXACT:
        return f"{head}{_HEADER_EXACT[core]}{tail}"
    m = _TOTAL_RE.match(core)
    if m:
        en = _MONTH_EN[int(m.group("m"))]
        if m.group("mid"):
            en += " " + _SITE_EN.get(m.group("mid"), m.group("mid"))
        en += " Total Sales"
        if m.group("pre"):
            en = _SITE_EN.get(m.group("pre"), m.group("pre")) + " " + en
        return f"{head}{core} {en}{tail}"
    m = _SALES_RE.match(core)
    if m:
        site = m.group("site")
        en = f"{_MONTH_EN[int(m.group('m'))]} {_SITE_EN.get(site, site)} Sales"
        return f"{head}{core} {en}{tail}"
    return text                           # 未登记的中文表头 → 保持原样


def fit_header_row_height(ws, row=1, line_h=13.5, extra=6.0):
    """按「换行后最长表头」把表头行高撑起来（只增不减）。"""
    need = 1
    for c in range(1, ws.max_column + 1):
        v = ws.cell(row, c).value
        if not isinstance(v, str) or v.startswith("="):
            continue
        avail = max((ws.column_dimensions[get_column_letter(c)].width or 8.43) - 1.0, 4.0)
        lines = 0
        for seg in v.split("\n"):
            w = sum(2 if ord(ch) > 0x2E7F else 1 for ch in seg)     # 全角/CJK 按 2 个字符宽算
            lines += max(1, int(np.ceil(w / avail)))
        need = max(need, lines)
    h = need * line_h + extra
    ws.row_dimensions[row].height = max(h, ws.row_dimensions[row].height or 0)


# sheet 改名：唯一一个纯中文 sheet 名（已确认全表无公式引用它，改名安全）
_RENAME_SHEET = {"2023-2026年销量情况": "2023-2026年销量情况 Sales"}

print("\n[5.5] 表头中英双语化")
i18n_total = 0
for sn in list(wb.sheetnames):
    ws = wb[sn]
    n = 0
    for r in (1, 2):                              # 表头只在前两行（第 3 行起就是数据）
        for c in range(1, ws.max_column + 1):
            cell = ws.cell(r, c)
            new = i18n_header(cell.value)
            if new != cell.value:
                cell.value = new
                n += 1
    if n:
        fit_header_row_height(ws, 1)              # 文字变长了 → 行高同步撑开
    i18n_total += n
    print(f"    {sn:<36} 翻译表头 {n:>3} 个")
for old, new in _RENAME_SHEET.items():
    if old in wb.sheetnames:
        wb[old].title = new
        print(f"    sheet 改名：{old!r} → {new!r}")
print(f"    合计翻译 {i18n_total} 个表头单元格")

# ---------------- 6. 保存主表 ----------------
try:
    wb.save(result_path)
    print(f"\n[6] 已保存 {result_path}")
except PermissionError:
    raise PermissionError(f"保存失败：{result_path} 正被 Excel/WPS 打开，请先关闭再重跑")

# ---------------- 7. 三个交付文件（纯英文表头）----------------
print("\n[7] 生成交付文件")
DELIVERABLE = {
    "Philippines": ("Philippines 菲律宾销量表", ["菲律宾"], PH_CHANNELS,
                    ["Zhongshan Warehouse (Lazada)", "Philippine Warehouse (Lazada)",
                     "Philippine Warehouse (Shopee)", "TikTok", "Offline Orders"]),
    "Indonesia": ("Indonesia 印尼销量表", ["印尼"], ID_CHANNELS,
                  ["Shopee (Smart Warehouse)", "Indonesia Offline Orders", "TikTok"]),
    "Thailand": ("Thailand 泰国销量表", ["泰国"], TH_CHANNELS,
                 ["Shopee (Yihai)", "LAZADA (Yihai)", "LAZADA Thailand Warehouse", "LAZADA Zhongshan Warehouse"]),
}
type_map = product_info_df.set_index("SKU")["类型"].to_dict()
brand_map = product_info_df.set_index("SKU")["品牌"].to_dict()
for prefix, (sheet_name, countries, ch, en_headers) in DELIVERABLE.items():
    ws = SHEET_HANDLES[sheet_name]
    conf = SHEET_CONF[sheet_name]
    width = len(ch) + 1
    start = block_start(conf["base_col"], width, month_offset)
    rows = []
    sku_row_map, _ = get_data_rows(ws)
    for s, r in sorted(sku_row_map.items(), key=lambda kv: kv[1]):
        vals = [ws.cell(r, start + i).value or 0 for i in range(len(ch))]
        ytd = 0.0
        for m in range(1, month_offset + 1):
            v = ws.cell(r, block_start(conf["base_col"], width, m) + len(ch)).value
            ytd += v if isinstance(v, (int, float)) else 0
        rows.append([s, model_map.get(s), type_map.get(s), brand_map.get(s)] + vals + [sum(vals), ytd])
    # 2026-10-16 修复：次表表头月份原来写死 Aug/Jan-Aug（2026-09-15 埋的坑），换月重跑不跟随归属月。
    # 改用归属月缩写（strftime("%b")：9月=Sep、12月=Dec；1 月时当月=累计，累计列特判为 "Jan <年> Cumulative"）。
    _mon_abbr = cur_acct_cycle_date.strftime("%b")
    _cum_hdr = (f"Jan {cur_acct_cycle_date.year} Cumulative"
                if cur_acct_cycle_date.month == 1
                else f"Jan-{_mon_abbr} {cur_acct_cycle_date.year} Cumulative")
    out = pd.DataFrame(rows, columns=["SKU", "Model", "Type", "Brand"] + en_headers
                       + [f"{_mon_abbr} {cur_acct_cycle_date.year} Total", _cum_hdr])
    out = out[out["SKU"].astype(str).str.match(r"^[GK]\d", na=False)].copy()
    out_path = os.path.join(result_dir, f"{prefix}_{belonging_date}_sales.xlsx")
    out.to_excel(out_path, index=False)
    print(f"    {os.path.basename(out_path):<40} {len(out)} 行  {_mon_abbr} 合计 = {out.iloc[:, 4 + len(ch)].sum():.0f}")


# ---------------- 8. 保真修复：补回 openpyxl 会丢掉的 x14 扩展 ----------------
# openpyxl 不保留 x14 扩展。模板 sheet1「2023-2026年销量情况」的 <extLst> 里有【两块】：
#   ① uri={78C0D931…} x14:conditionalFormattings —— 10 条红底规则（各月合计 ≠ 上月值为红）
#   ② uri={05C60535…} x14:sparklineGroups     —— 8 条迷你图（数据区 B3:AW3 … B10:AW10，显示在 AX3:AX10）
# 这个 sheet 我们不改（内容由 SEA Total 经 XLOOKUP 驱动）→ 整块 <extLst> 原样搬过来即可。
# ⚠️ 两个必须做对，否则 WPS 打开后再保存会整块丢掉（9.15 实测丢过 22 条迷你图）：
#    (1) 必须是完整的 <extLst>…</extLst> 外壳，<ext> 不能直接挂在 <worksheet> 下（违反 schema）；
#    (2) worksheet 根上必须声明 xr2 等前缀并写入 mc:Ignorable，否则扩展被当成"未知内容"规范化掉。
EXT_NS = {
    "xmlns:mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "xmlns:x14ac": "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac",
    "xmlns:xr": "http://schemas.microsoft.com/office/spreadsheetml/2014/revision",
    "xmlns:xr2": "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2",
    "xmlns:xr3": "http://schemas.microsoft.com/office/spreadsheetml/2016/revision3",
}
EXT_IGNORABLE = ["x14ac", "xr", "xr2", "xr3"]


def patch_sheet_exts(xml, ext_lst):
    """给 sheet XML 补根命名空间声明 + 注入完整 <extLst>（幂等）"""
    if "<extLst>" in xml and "{05C60535" in xml:
        return xml
    xml = re.sub(r"<extLst\s*/>", "", xml)          # 去掉 openpyxl 写的空 <extLst/>
    root = re.search(r"<worksheet\b[^>]*>", xml).group(0)
    new_root = root
    add = "".join(f' {k}="{v}"' for k, v in EXT_NS.items() if f"{k}=" not in new_root)
    if add:
        new_root = new_root[:-1] + add + ">"
    m = re.search(r'mc:Ignorable="([^"]*)"', new_root)
    if m:
        have = m.group(1).split()
        merged = have + [w for w in EXT_IGNORABLE if w not in have]
        new_root = new_root.replace(m.group(0), f'mc:Ignorable="{" ".join(merged)}"')
    else:
        new_root = new_root[:-1] + f' mc:Ignorable="{" ".join(EXT_IGNORABLE)}">'
    return xml.replace(root, new_root, 1).replace("</worksheet>", ext_lst + "</worksheet>")


def _replace_file(tmp, dst):
    """把临时文件替换成目标文件。
    优先 os.replace（原子替换）；但目标被 WPS / 预览器打开时，Windows 会拒绝对它做"删除+改名"
    （PermissionError: [WinError 5] 拒绝访问）→ 此时退化成原地覆写（保留对方已打开的句柄）。"""
    try:
        os.replace(tmp, dst)
        return "replace"
    except PermissionError:
        with open(tmp, "rb") as f:
            data = f.read()
        with open(dst, "wb") as f:
            f.write(data)
        os.remove(tmp)
        print("    ⚠️ 目标文件被其他程序占用，已退化为原地覆写（建议关掉 WPS/预览再跑一次）")
        return "overwrite"


def restore_sparklines(result_path, template_path, sheet_xml_name="xl/worksheets/sheet1.xml"):
    import zipfile
    with zipfile.ZipFile(template_path) as tz:
        tpl_xml = tz.read(sheet_xml_name).decode("utf-8")
    m = re.search(r"<extLst>.*?</extLst>", tpl_xml, re.S)
    if not m:
        return False
    ext_lst = m.group(0)
    with zipfile.ZipFile(result_path) as zr:
        tmp = result_path + ".tmp"
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zw:
            for it in zr.infolist():
                data = zr.read(it.filename)
                if it.filename == sheet_xml_name:
                    data = patch_sheet_exts(data.decode("utf-8"), ext_lst).encode("utf-8")
                elif it.filename == "xl/workbook.xml":
                    # 打开时强制重算：SEA Total 变了，「2023-2026年销量情况」的 XLOOKUP 缓存值必须刷新
                    t = data.decode("utf-8")
                    if "<calcPr" in t:
                        t = re.sub(r'<calcPr([^>]*?)/>', lambda mm: f'<calcPr{mm.group(1)} fullCalcOnLoad="1"/>'
                                   if "fullCalcOnLoad" not in mm.group(1) else mm.group(0), t, count=1)
                    else:
                        t = t.replace("</workbook>", '<calcPr fullCalcOnLoad="1"/></workbook>')
                    data = t.encode("utf-8")
                zw.writestr(it, data)
    _replace_file(tmp, result_path)
    return True


if restore_sparklines(result_path, template_path):
    print("\n[8] ✅ 已补回 sheet1 的 x14 扩展（8 条迷你图 + 10 条红底条件格式），并补齐根命名空间")
else:
    print("\n[8] ⚠️ 模板里没找到 <extLst>，跳过")


# ---------------- 9. 保真修复：搬回 openpyxl 会丢掉的 WPS 专有部件 ----------------
# 下面这些部件都不是"单元格数据"，但决定打开后的观感，openpyxl 一个都不会保留：
#   docProps/custom.xml → ★ KSOReadingLayout=true = WPS「阅读模式」：点任一单元格，整行整列高亮成黄色十字
#   customXml/item1.xml → WPS pixelatorList（每表显示微调）
#   customXml/item2.xml → WPS woSheetProps（interlineOnOff / cellprotection 等每表属性）
#   xl/printerSettings/*.bin + sheet2/5/6 的 rels → 各表打印设置
#   docProps/thumbnail.emf → 文件缩略图
# 做法：整包重建，把这些部件从模板原样搬过来，并在 [Content_Types].xml / _rels/.rels /
#       xl/_rels/workbook.xml.rels 里补齐登记（否则部件是"孤儿"，WPS/Excel 会忽略）。
WPS_PARTS = [
    # (部件路径, 需要的 Content_Types Override 类型；None = 用 xml 默认，无需 Override)
    ("docProps/custom.xml", "application/vnd.openxmlformats-officedocument.custom-properties+xml"),
    ("customXml/itemProps1.xml", "application/vnd.openxmlformats-officedocument.customXmlProperties+xml"),
    ("customXml/itemProps2.xml", "application/vnd.openxmlformats-officedocument.customXmlProperties+xml"),
    ("customXml/item1.xml", None),
    ("customXml/item2.xml", None),
    ("customXml/_rels/item1.xml.rels", None),
    ("customXml/_rels/item2.xml.rels", None),
    ("docProps/thumbnail.emf", None),
    ("xl/printerSettings/printerSettings1.bin", None),
    ("xl/printerSettings/printerSettings2.bin", None),
    ("xl/worksheets/_rels/sheet2.xml.rels", None),
    ("xl/worksheets/_rels/sheet5.xml.rels", None),
    ("xl/worksheets/_rels/sheet6.xml.rels", None),
]
CT_DEFAULTS = [
    ("bin", "application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings"),
    ("emf", "image/x-emf"),
]
PKG_RELS = [   # _rels/.rels 里要补的关系
    ("http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties",
     "docProps/custom.xml"),
    ("http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail",
     "docProps/thumbnail.emf"),
]
WB_CUSTOM_XML = ["../customXml/item1.xml", "../customXml/item2.xml"]   # workbook.xml.rels 里要补的关系
CUSTOMXML_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml"


def _ensure_default(ct_xml, ext, ctype):
    if f'Extension="{ext}"' in ct_xml:
        return ct_xml
    i = ct_xml.index(">", ct_xml.index("<Types")) + 1
    return ct_xml[:i] + f'<Default Extension="{ext}" ContentType="{ctype}"/>' + ct_xml[i:]


def _ensure_override(ct_xml, part, ctype):
    if f'PartName="/{part}"' in ct_xml:
        return ct_xml
    return ct_xml.replace("</Types>", f'<Override PartName="/{part}" ContentType="{ctype}"/></Types>')


def _ensure_rel(rels_xml, rtype, target):
    """在 Relationships 里补一条关系（已存在同 Target 则跳过，保证幂等）"""
    if f'Target="{target}"' in rels_xml:
        return rels_xml
    used = {int(m) for m in re.findall(r'Id="rId(\d+)"', rels_xml)}
    n = 0
    while n in used:
        n += 1
    rel = f'<Relationship Id="rId{n}" Type="{rtype}" Target="{target}"/>'
    return rels_xml.replace("</Relationships>", rel + "</Relationships>")


def _prune_dangling_rels(part, data, present, blobs):
    """删掉 .rels 里 Target 指向「输出包里并不存在」的关系条目。返回 (新 bytes, 是否已空)。

    悬空关系的来源：openpyxl **不保留某些图形对象**（如 <xdr:cxnSp> 直线连接符），
    保存时会连带丢掉对应的 drawing 部件；而我们又把模板的 rels 原样写回、仍指向它
    → 之后再用 openpyxl 打开就会 `KeyError: xl/drawings/drawing1.xml`。
    ⚠️ 9.16 实测：新加坡 sheet（sheet6）模板里就有一条 hidden、零宽的「直接连接符 1」，
       属无用装饰，删掉它的关系即可（对 WPS/Excel 打开无影响）。
    只有「该 rId 在引用方 XML 里已不被引用」时才删；仍被引用则保留并告警，避免误删。
    """
    import posixpath
    xml = data.decode("utf-8")
    base = posixpath.dirname(posixpath.dirname(part))       # xl/worksheets/_rels/x.rels → xl/worksheets
    owner = part[:-len(".rels")]                            # xl/worksheets/sheet6.xml
    owner_xml = blobs.get(owner, b"").decode("utf-8", "ignore")
    keep = []
    for m in re.finditer(r"<Relationship\b[^>]*?/>", xml):
        tag = m.group(0)
        if 'TargetMode="External"' in tag:                  # 外链不校验
            keep.append(tag)
            continue
        t = re.search(r'Target="([^"]+)"', tag)
        if not t:
            keep.append(tag)
            continue
        tgt = t.group(1)
        p = tgt.lstrip("/") if tgt.startswith("/") else posixpath.normpath(posixpath.join(base, tgt))
        if p in present:
            keep.append(tag)
            continue
        rid_m = re.search(r'Id="([^"]+)"', tag)
        rid = rid_m.group(1) if rid_m else "?"
        if rid != "?" and f'r:id="{rid}"' in owner_xml:
            print(f"    ⚠️ {part}: 关系 {rid} 指向缺失部件 {p}，但 {owner} 仍在引用 → 保留（需人工检查）")
            keep.append(tag)
        else:
            print(f"    🧹 {part}: 删除悬空关系 {rid} → {p}（引用方已无此引用）")
    if not keep:
        return b"", True
    open_tag = re.search(r"<Relationships\b[^>]*?/?>", xml)
    _tag = open_tag.group(0)
    head = xml[:open_tag.start()] + (_tag[:-2] + ">" if _tag.endswith("/>") else _tag)
    return (head + "".join(keep) + "</Relationships>").encode("utf-8"), False


def restore_wps_parts(result_path, template_path):
    """把模板的 WPS 专有部件按原样写回输出包（含 docProps/custom.xml 的「阅读模式」开关）。
    ⚠️ docProps/custom.xml 是【覆盖】不是【新增】：openpyxl 自己也会写一个同名的
       （只保留 KSOProductBuildVer/ICV/KSOReadingLayout，写法是 <vt:bool>1</vt:bool>），
       直接追加会得到两份同名部件 → zipfile 读到的还是 openpyxl 那份，覆盖才生效且不重复。"""
    import zipfile
    with zipfile.ZipFile(template_path) as tz:
        tpl_names = set(tz.namelist())
        parts = {p: tz.read(p) for p, _ in WPS_PARTS if p in tpl_names}
        missing = [p for p, _ in WPS_PARTS if p not in tpl_names]
        # 模板的 sheetId：openpyxl 会重编号（销量情况 10→1、团购 11→10），
        # 而 WPS 的 customXml 是按 sheetId 索引每张表的 → 按表名还原成模板值
        tpl_sheetid = {}
        for m in re.finditer(r"<sheet\b[^>]*?/?>", tz.read("xl/workbook.xml").decode("utf-8")):
            nm = re.search(r'name="([^"]*)"', m.group(0))
            sid = re.search(r'sheetId="([^"]*)"', m.group(0))
            if nm and sid:
                tpl_sheetid[nm.group(1)] = sid.group(1)
    with zipfile.ZipFile(result_path) as zr:
        items = [(it, zr.read(it.filename)) for it in zr.infolist()]

    def _fix_sheetid(mm):
        tag = mm.group(0)
        nm = re.search(r'name="([^"]*)"', tag)
        if nm and nm.group(1) in tpl_sheetid:
            return re.sub(r'sheetId="[^"]*"', f'sheetId="{tpl_sheetid[nm.group(1)]}"', tag, count=1)
        return tag

    blobs = {it.filename: data for it, data in items}
    present = set(blobs) | set(parts)          # 输出包里最终会存在的部件全集
    written, skipped = set(), set()
    tmp = result_path + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zw:
        for it, data in items:
            fn = it.filename
            if fn in parts:                    # 覆盖 openpyxl 写的同名部件
                data = parts[fn]
                if fn.endswith(".rels"):       # 剪掉指向「已不存在部件」的悬空关系
                    data, emptied = _prune_dangling_rels(fn, data, present, blobs)
                    if emptied:
                        skipped.add(fn)
                        continue
                written.add(fn)
            if fn == "[Content_Types].xml":
                t = data.decode("utf-8")
                for ext, ctype in CT_DEFAULTS:
                    t = _ensure_default(t, ext, ctype)
                for p, ctype in WPS_PARTS:
                    if ctype and p in parts:
                        t = _ensure_override(t, p, ctype)
                data = t.encode("utf-8")
            elif fn == "_rels/.rels":
                t = data.decode("utf-8")
                for rtype, tgt in PKG_RELS:
                    if tgt in parts:
                        t = _ensure_rel(t, rtype, tgt)
                data = t.encode("utf-8")
            elif fn == "xl/_rels/workbook.xml.rels":
                t = data.decode("utf-8")
                for tgt in WB_CUSTOM_XML:
                    if tgt.lstrip("./") in parts:
                        t = _ensure_rel(t, CUSTOMXML_REL_TYPE, tgt)
                data = t.encode("utf-8")
            elif fn == "xl/workbook.xml" and tpl_sheetid:
                data = re.sub(r"<sheet\b[^>]*?/?>", _fix_sheetid, data.decode("utf-8")).encode("utf-8")
            zw.writestr(it, data)
        for p, _ in WPS_PARTS:            # 只有 openpyxl 没写的部件才追加（[Content_Types].xml 仍在最前）
            if p in parts and p not in written and p not in skipped:
                data = parts[p]
                if p.endswith(".rels"):
                    data, emptied = _prune_dangling_rels(p, data, present, blobs)
                    if emptied:
                        skipped.add(p)
                        continue
                zw.writestr(p, data)
                written.add(p)
    _replace_file(tmp, result_path)
    return sorted((written | {p for p, _ in WPS_PARTS if p in parts}) - skipped), missing


_parts, _miss = restore_wps_parts(result_path, template_path)
print(f"\n[9] ✅ 已写回 {len(_parts)} 个 WPS 专有部件（docProps/custom.xml 含「阅读模式」黄色十字开关；"
      f"sheetId 已还原为模板值）")
print(f"    {_parts}")
if _miss:
    print(f"    ⚠️ 模板里没有、已跳过：{_miss}")


# ---------------- 10. 列分组展开态：把模板「当月展开」推进到本月 ----------------
# 模板是 7 月保存的 → 每个 sheet 的列分组里"展开的那一块"正好是第 7 个月块（打开就是 7 月，
# 得手动收 7 月、展 8 月，很麻烦）。这里按同一规则把展开块推进到 month_offset 月。
# 布局 = (月份块起始列, 块宽)；块内前 宽−1 列是明细列、最后一列是"月合计列"（可见、不参与分级）。
BLOCK_LAYOUT = {
    "sheet2.xml": (7, 9),    # SEA Total 东南亚总销量表：8 分类 + 1 合计
    "sheet3.xml": (3, 6),    # Philippines 菲律宾销量表：5 渠道 + 1 合计
    "sheet4.xml": (3, 5),    # Thailand 泰国销量表：4 渠道 + 1 合计
    "sheet5.xml": (3, 4),    # Indonesia 印尼销量表：3 渠道 + 1 合计
    "sheet6.xml": (3, 3),    # Singapore 新加坡销量表：2 渠道 + 1 合计
    "sheet9.xml": (3, 9),    # Malaysia Shopee 虾皮跨境销量表：8 渠道 + 1 合计
}
# 每月只有 1 列、没有"月合计列"的 sheet → 分组只分"已过月份 / 未来月份"两段（首列 1 月不分组）
PLAIN_LAYOUT = {
    "sheet7.xml": (3, 14),   # Vietnam Lazada 越南销销量表
    "sheet8.xml": (3, 14),   # Malaysia Lazada 马来西亚销量表
    "sheet10.xml": (3, 14),  # Group buying distributor 团购经销商
}
COL_ATTR_ORDER = ["min", "max", "width", "style", "hidden", "bestFit", "customWidth",
                  "outlineLevel", "collapsed"]


def _col_xml(a):
    keys = [k for k in COL_ATTR_ORDER if k in a] + [k for k in sorted(a) if k not in COL_ATTR_ORDER]
    return "<col " + " ".join(f'{k}="{a[k]}"' for k in keys) + "/>"


def _split_entries(entries, points):
    """在指定列号处切断 <col> 条目，保证切完每条只承担一个角色"""
    out = []
    for a in entries:
        lo, hi = int(a["min"]), int(a["max"])
        cur = lo
        for p in sorted(x for x in points if lo < x <= hi) + [hi + 1]:
            b = dict(a)
            b["min"], b["max"] = str(cur), str(p - 1)
            out.append(b)
            cur = p
    return out


def _merge_entries(entries):
    """把属性完全相同的相邻条目并回一条（尽量贴近模板的写法）"""
    out = []
    for a in entries:
        if out and int(out[-1]["max"]) + 1 == int(a["min"]):
            prev = {k: v for k, v in out[-1].items() if k not in ("min", "max")}
            cur = {k: v for k, v in a.items() if k not in ("min", "max")}
            if prev == cur:
                out[-1]["max"] = a["max"]
                continue
        out.append(dict(a))
    return out


def _wrap_cols(xml, cols_match, entries):
    return xml.replace(cols_match, "<cols>" + "".join(_col_xml(a) for a in entries) + "</cols>", 1)


def apply_month_outline(xml, month_offset, block=None, plain=None):
    """重排 <cols>：本月块展开、已过月份折叠只留月合计列、未来月份整块隐藏。
    返回 (新 xml, 本月块首列号, 本月合计列号)。"""
    m = re.search(r"<cols>.*?</cols>", xml, re.S)
    if not m:
        return xml, None, None
    entries = []
    for cm in re.finditer(r"<col\b([^>]*?)/>", m.group(0)):
        entries.append(dict(re.findall(r'(\w+)="([^"]*)"', cm.group(1))))

    if block:
        start, width = block
        end = start + width * 12 - 1
        points = []
        for k in range(1, 13):
            points += [start + width * (k - 1), start + width * k - 1]
        new = []
        for a in _split_entries(entries, points):
            lo = int(a["min"])
            if lo < start or lo > end:
                new.append(a)
                continue
            off = lo - start
            k, pos = off // width + 1, off % width
            if pos < width - 1:                       # 明细列
                a["outlineLevel"] = "1"
                a.pop("collapsed", None)
                if k == month_offset:
                    a.pop("hidden", None)
                else:
                    a["hidden"] = "1"
            else:                                     # 月合计列：始终可见、自身不分级
                a.pop("outlineLevel", None)
                if k < month_offset:
                    a.pop("hidden", None)
                    a["collapsed"] = "1"
                elif k == month_offset:
                    a.pop("hidden", None)
                    a.pop("collapsed", None)
                else:
                    a["hidden"] = "1"
                    a["collapsed"] = "1"
            new.append(a)
        new = _merge_entries(new)
        return (_wrap_cols(xml, m.group(0), new),
                start + width * (month_offset - 1), start + width * month_offset - 1)

    if plain:
        first, last = plain
        vis_from = first + 1                          # 首列（1 月）模板本身不分组，保持不动
        vis_to = max(vis_from, 2 + month_offset)      # 已过月份（含本月）
        hid_from = vis_to + 1                         # 未来月份
        new = []
        for a in _split_entries(entries, [vis_from, hid_from]):
            lo = int(a["min"])
            if vis_from <= lo <= vis_to:
                a["outlineLevel"] = "1"
                a.pop("hidden", None)
                a.pop("collapsed", None)
            elif hid_from <= lo <= last:
                a["outlineLevel"] = "1"
                a["hidden"] = "1"
                a.pop("collapsed", None)
            new.append(a)
        new = _merge_entries(new)
        aug = 2 + month_offset
        return _wrap_cols(xml, m.group(0), new), aug, aug

    return xml, None, None


def ensure_outline_level_col(xml):
    """补回 sheetFormatPr 的 outlineLevelCol（openpyxl 会丢，丢了分组数不显示）"""
    m = re.search(r"<sheetFormatPr\b[^>]*/>", xml)
    if not m or "outlineLevelCol" in m.group(0):
        return xml
    return xml.replace(m.group(0), m.group(0)[:-2] + ' outlineLevelCol="1"/>', 1)


def apply_month_view(xml, first_col, sum_col):
    """冻结窗格的滚动起点移到本月块首列；活动单元格放到本月合计列第 2 行（打开即停在当月）。
    ⚠️ openpyxl 会保留模板 sheetView 里的 <selection pane="bottomRight">，必须【替换】而不是追加，
       否则同一个 pane 会出现两个 selection（不合规范）。"""
    if not first_col:
        return xml
    pm = re.search(r"<pane\b[^>]*/>", xml)
    if not pm:
        return xml
    pane = pm.group(0)
    ys = re.search(r'ySplit="(\d+)"', pane)
    row = (int(ys.group(1)) + 1) if ys else 1
    new_pane = re.sub(r'topLeftCell="[^"]*"', f'topLeftCell="{get_column_letter(first_col)}{row}"', pane)
    sel = (f'<selection pane="bottomRight" activeCell="{get_column_letter(sum_col)}2" '
           f'sqref="{get_column_letter(sum_col)}2"/>')
    xml = re.sub(r'<selection\b[^>]*pane="bottomRight"[^>]*/>', "", xml)   # 先删掉模板留下的旧选区
    return xml.replace(pane, new_pane + sel, 1)


def apply_outline_and_view(result_path, month_offset):
    import zipfile
    done = []
    with zipfile.ZipFile(result_path) as zr:
        items = [(it, zr.read(it.filename)) for it in zr.infolist()]
    tmp = result_path + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zw:
        for it, data in items:
            mm = re.fullmatch(r"xl/worksheets/(sheet\d+\.xml)", it.filename)
            if mm and (mm.group(1) in BLOCK_LAYOUT or mm.group(1) in PLAIN_LAYOUT):
                sh = mm.group(1)
                x = data.decode("utf-8")
                x, c1, c2 = apply_month_outline(x, month_offset, BLOCK_LAYOUT.get(sh), PLAIN_LAYOUT.get(sh))
                if c1:
                    x = ensure_outline_level_col(x)
                    x = apply_month_view(x, c1, c2)
                    done.append((sh, get_column_letter(c1), get_column_letter(c2)))
                data = x.encode("utf-8")
            zw.writestr(it, data)
    _replace_file(tmp, result_path)
    return done


_out = apply_outline_and_view(result_path, month_offset)
print(f"\n[10] ✅ 列分组展开态已推进到 {month_offset} 月（原模板锁在 7 月），共 {len(_out)} 个 sheet")
for _sh, _c1, _c2 in _out:
    print(f"     {_sh:<14} 打开定位到 {_c1} 列，选中 {_c2}2")

print("\n补行位汇总：")
for k, v in appended_log.items():
    print(f"  {k}: {len(v)} 个 → {v}")
print("\n完成。")
