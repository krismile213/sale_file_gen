# -*- coding: utf-8 -*-
"""
东南亚三国（菲律宾/泰国/印尼）9 月月销 同步核对探针（只读 + import 生产口径）
五层对账：
  L1 源报告（import month_sale_file 生产 getter 重算）
  L2 中间产物 se_a_sku_sale_df.csv
  L3 原成品 东南亚销量表（trunk/result）
  L4 重算副本 菲律宾重算_202609/2026年东南亚销量表...
  L5 月销售总表（原件 + 副本）五组 9 月列
"""
import os, sys, io, contextlib

TRUNK = r"D:\Desktop\work-main\sale_file_gen\trunk"
ROOT = r"D:\Desktop\work-main\sale_file_gen"
os.chdir(TRUNK)
sys.path.insert(0, TRUNK)

import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter as gcl

BD, SD = "202609", "20261006"
SEA_ORIG = os.path.join(TRUNK, "result", "东南亚销量表", "202609",
                        "2026年东南亚销量表Southeast Asia Sales table.xlsx")
SEA_COPY = os.path.join(ROOT, "菲律宾重算_202609",
                        "2026年东南亚销量表Southeast Asia Sales table.xlsx")
MS_ORIG = os.path.join(TRUNK, "result", "数据归属周期", "202609", "月销售总表.xlsx")
MS_COPY = os.path.join(ROOT, "菲律宾重算_202609", "月销售总表.xlsx")
CH_CSV = os.path.join(TRUNK, "sale_file_tmp", "se_a_sku_sale_df.csv")


def n0(v):
    """归一：'00' 文本/None → 0"""
    if v is None:
        return 0
    if isinstance(v, str):
        s = v.strip()
        if s in ("", "00", "0"):
            return 0
        try:
            return float(s)
        except ValueError:
            return 0
    return float(v)


# ---------- L1: 源报告重算（生产口径） ----------
import month_sale_file as msf

JOBS = [
    # (国家, 渠道显示名, getter, entity, sale_group)
    ("菲律宾", "跨境Lazada(中山仓)", msf.get_belonging_lazada_order_df, "跨境", "菲律宾"),
    ("菲律宾", "本土Lazada(菲仓)", msf.get_belonging_lazada_order_df, "本土", "菲律宾"),
    ("菲律宾", "本土Shopee(菲仓)", msf.get_belonging_shopee_order_df, "本土", "菲律宾"),
    ("菲律宾", "本土Tiktok", msf.get_belonging_tkshop_order_df, "本土", "菲律宾"),
    ("菲律宾", "线下KSON", msf.get_belonging_kson_offline_order_df, "本土", "菲律宾"),
    ("泰国", "本土Shopee(易海)", msf.get_belonging_shopee_order_df, "本土", "泰国"),
    ("泰国", "本土Lazada(易海)", msf.get_belonging_lazada_order_df, "本土", "泰国"),
    ("泰国", "跨境Lazada(中山仓)", msf.get_belonging_lazada_order_df, "跨境", "泰国"),
    ("印尼", "本土Shopee(智慧仓)", msf.get_belonging_shopee_order_df, "本土", "印尼"),
    ("印尼", "本土Tiktok", msf.get_belonging_tkshop_order_df, "本土", "印尼"),
    ("印尼", "线下(印尼海外仓)", msf.get_belonging_indonesia_warehouse_offline_order_df, "本土", "印尼"),
]

src = {}          # (国家, 渠道) -> 件数
src_rows = {}     # (国家, 渠道) -> 行数
src_sku = {}      # (国家, 渠道) -> {sku: qty}
print("=" * 100)
print("L1 源报告重算（import month_sale_file 生产 getter，快照 20261006）")
print("=" * 100)
for country, label, fn, ent, grp in JOBS:
    buf = io.StringIO()
    try:
        with contextlib.redirect_stdout(buf):
            df = fn(ent, grp, BD, SD)
        q = pd.to_numeric(df["数量"], errors="coerce").fillna(0).sum()
        sk = df.assign(_q=pd.to_numeric(df["数量"], errors="coerce").fillna(0)) \
               .groupby(df["sku"].astype(str).str.strip().str.upper())["_q"].sum()
        src[(country, label)] = float(q)
        src_rows[(country, label)] = len(df)
        src_sku[(country, label)] = sk.to_dict()
        print(f"  {country} | {label:<22} 行数={len(df):>5}  件数={q:>8.0f}")
    except Exception as e:
        print(f"  {country} | {label:<22} ★失败: {e}")
        src[(country, label)] = None

# ---------- L2: 中间产物 ----------
print()
print("=" * 100)
print("L2 中间产物 se_a_sku_sale_df.csv 按 (站点,平台,实体,仓库) 聚合")
print("=" * 100)
ch = pd.read_csv(CH_CSV, low_memory=False)
print("  总行数:", len(ch), " 列:", list(ch.columns)[:12])
mid = {}
for (site, plat, ent, wh), g in ch.groupby([ch["站点"], ch["平台"], ch["实体"], ch["仓库"]]):
    q = pd.to_numeric(g["202609销量"], errors="coerce").fillna(0).sum()
    mid[(str(site), str(plat), str(ent), str(wh))] = float(q)
for k in sorted(mid, key=lambda x: (x[0], x[1])):
    if k[0] in ("菲律宾", "泰国", "印尼"):
        print(f"  {k[0]} | {k[1]:<8} | {k[2]} | {k[3]:<10} = {mid[k]:>8.0f}")

# ---------- L3/L4: 东南亚销量表 9 月块 ----------
SHEET_BLOCKS = {
    "Philippines 菲律宾销量表": (51, [("AY", "跨境Lazada(中山仓)"), ("AZ", "本土Lazada(菲仓)"),
                                  ("BA", "本土Shopee(菲仓)"), ("BB", "本土Tiktok"), ("BC", "线下KSON")], "BD"),
    "Thailand 泰国销量表": (43, [("AQ", "本土Shopee(易海)"), ("AR", "本土Lazada(易海)"),
                              ("AS", "LAZADA泰国仓(恒0)"), ("AT", "跨境Lazada(中山仓)")], "AU"),
    "Indonesia 印尼销量表": (35, [("AI", "本土Shopee(智慧仓)"), ("AJ", "线下(印尼海外仓)"),
                               ("AK", "本土Tiktok")], "AL"),
}


def read_sea(path, tag):
    wb = load_workbook(path, read_only=True, data_only=True)
    out = {}
    for sn, (base, chans, tot) in SHEET_BLOCKS.items():
        ws = wb[sn]
        r2 = {c.column: c.value for c in ws[2]}
        vals = {label: n0(r2.get(base + i)) for i, (_, label) in enumerate(chans)}
        vals["合计"] = n0(r2.get(base + len(chans)))
        out[sn] = vals
    # SEA Total
    ws = wb["SEA Total 东南亚总销量表"]
    r2 = {c.column: c.value for c in ws[2]}
    sea_cols = {"菲律宾": 79, "泰国": 80, "印尼": 81, "新加坡": 82, "越南": 83,
                "马来Lazada": 84, "虾皮跨境": 85, "团购": 86, "合计": 87}
    out["SEA Total"] = {k: n0(r2.get(v)) for k, v in sea_cols.items()}
    wb.close()
    print(f"\n----- {tag} -----")
    for sn, vals in out.items():
        print(f"  {sn}:")
        for k, v in vals.items():
            print(f"     {k:<22} = {v:>8.0f}")
    return out


print()
print("=" * 100)
print("L3 原成品 东南亚销量表 9 月块（行2 合计行）")
print("=" * 100)
sea_orig = read_sea(SEA_ORIG, "原件 trunk/result/东南亚销量表/202609")

print()
print("=" * 100)
print("L4 重算副本 菲律宾重算_202609")
print("=" * 100)
sea_copy = read_sea(SEA_COPY, "副本 菲律宾重算_202609")

# ---------- L5: 月销售总表 ----------
MS_GROUPS = {"菲律宾": 135, "泰国": 147, "印尼": 159, "新加坡": 171, "马来": 183,
             "越南": 195, "虾皮境内": 207, "东南亚": 219, "京东": 231, "eMAG": 243,
             "TEMU平台": 255, "美客多": 267, "总计": 279}


def read_ms(path, tag):
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb["2026年"]
    # 行2 月标签定位（确认组首+8 是 9月）
    r2 = {c.column: c.value for c in ws[2]}
    # SKU 列：最后一列
    sku_col = ws.max_column
    rows = {}
    for r in ws.iter_rows(min_row=4):
        sku = r[sku_col - 1].value if len(r) >= sku_col else None
        if sku is None:
            continue
        rows[str(sku).strip().upper()] = r[0].row if False else None
    # 重新按行号聚合（iter_rows 只能顺序扫，这里扫两遍太慢 → 用 values_only）
    sums = {}
    detail = {}
    for r in ws.iter_rows(min_row=4, values_only=True):
        sku = r[sku_col - 1]
        if sku is None:
            continue
        rr = None
        for gname, gcol in MS_GROUPS.items():
            c9 = gcol + 8
            v = n0(r[c9 - 1])
            sums[gname] = sums.get(gname, 0.0) + v
            detail.setdefault(gname, {})[str(sku).strip().upper()] = v
    wb.close()
    print(f"\n----- {tag} 月销售总表 2026年 sheet 9 月列（数据行 r4+ 求和）-----")
    for g in ["菲律宾", "泰国", "印尼", "新加坡", "马来", "越南", "虾皮境内", "东南亚", "eMAG", "美客多", "总计"]:
        if g in sums:
            print(f"  {g:<10} 9月列={gcl(MS_GROUPS[g]+8):<4} 合计={sums[g]:>10.0f}")
    return sums, detail


print()
print("=" * 100)
print("L5 月销售总表 9 月列")
print("=" * 100)
ms_orig_sums, ms_orig_detail = read_ms(MS_ORIG, "原件")
ms_copy_sums, ms_copy_detail = read_ms(MS_COPY, "副本")

# ---------- 对账输出 ----------
print()
print("=" * 100)
print("对账一：L1 源报告 vs L3 原成品表（按渠道）")
print("=" * 100)
COUNTRY_SHEET = {"菲律宾": "Philippines 菲律宾销量表", "泰国": "Thailand 泰国销量表",
                 "印尼": "Indonesia 印尼销量表"}
ok1 = True
for country in ["菲律宾", "泰国", "印尼"]:
    print(f"  【{country}】")
    sn = COUNTRY_SHEET[country]
    for label in [l for (c, l) in src if c == country]:
        s = src.get((country, label))
        w = sea_orig[sn].get(label)
        if s is None:
            flag = "★源读取失败"
            ok1 = False
        elif abs(s - w) < 1e-6:
            flag = "✓"
        else:
            flag = f"★差 {s - w:+.0f}"
            ok1 = False
        print(f"    {label:<22} 源={s if s is None else f'{s:>8.0f}'}  表={w:>8.0f}  {flag}")
    st = sum(v for (c, l), v in src.items() if c == country and v is not None)
    wt = sea_orig[sn]["合计"]
    print(f"    {'合计':<22} 源={st:>8.0f}  表={wt:>8.0f}  {'✓' if abs(st-wt)<1e-6 else '★差 %+.0f' % (st-wt)}")

print()
print("=" * 100)
print("对账二：L1 源报告 vs L2 中间产物（SKU 级，逐渠道）")
print("=" * 100)
MAP = {  # (国家,渠道label) -> (站点,平台,实体,仓库)
    ("菲律宾", "跨境Lazada(中山仓)"): ("菲律宾", "Lazada", "跨境", "旭艺仓"),
    ("菲律宾", "本土Lazada(菲仓)"): ("菲律宾", "Lazada", "本土", "菲律宾"),
    ("菲律宾", "本土Shopee(菲仓)"): ("菲律宾", "Shopee", "本土", "菲律宾"),
    ("菲律宾", "本土Tiktok"): ("菲律宾", "Tiktok", "本土", "菲律宾"),
    ("菲律宾", "线下KSON"): ("菲律宾", "KSON", "本土", "KSON"),
    ("泰国", "本土Shopee(易海)"): ("泰国", "Shopee", "本土", "泰国"),
    ("泰国", "本土Lazada(易海)"): ("泰国", "Lazada", "本土", "泰国"),
    ("泰国", "跨境Lazada(中山仓)"): ("泰国", "Lazada", "跨境", "旭艺仓"),
    ("印尼", "本土Shopee(智慧仓)"): ("印尼", "Shopee", "本土", "印尼"),
    ("印尼", "本土Tiktok"): ("印尼", "Tiktok", "本土", "印尼"),
    ("印尼", "线下(印尼海外仓)"): ("印尼", "线下", "本土", "印尼海外仓"),
}
ok2 = True
for key, mk in MAP.items():
    s_sku = src_sku.get(key, {})
    m_sku = {}
    sub = ch[(ch["站点"] == mk[0]) & (ch["平台"] == mk[1]) & (ch["实体"] == mk[2]) & (ch["仓库"] == mk[3])]
    for _, row in sub.iterrows():
        k = str(row["sku"]).strip().upper()
        m_sku[k] = m_sku.get(k, 0) + n0(row["202609销量"])
    diffs = []
    for k in set(s_sku) | set(m_sku):
        a, b = s_sku.get(k, 0), m_sku.get(k, 0)
        if abs(a - b) > 1e-6:
            diffs.append((k, a, b))
    status = "✓" if not diffs else f"★{len(diffs)} 个 SKU 不一致"
    if diffs:
        ok2 = False
    print(f"  {key[0]} | {key[1]:<22} SKU数 源={len(s_sku):>3} 中={len(m_sku):>3}  {status}")
    for k, a, b in sorted(diffs)[:6]:
        print(f"      {k}: 源={a:.0f} 中={b:.0f}")

print()
print("=" * 100)
print("对账三：L3/L4 东南亚销量表 vs L5 月销售总表（原件 & 副本）")
print("=" * 100)
pairs = [("菲律宾", "Philippines 菲律宾销量表", "菲律宾"),
         ("泰国", "Thailand 泰国销量表", "泰国"),
         ("印尼", "Indonesia 印尼销量表", "印尼")]
for tag, sea, ms_s in [("原件", sea_orig, ms_orig_sums), ("副本", sea_copy, ms_copy_sums)]:
    print(f"  【{tag}】")
    for cname, sn, gname in pairs:
        a = sea[sn]["合计"]
        b = ms_s.get(gname)
        print(f"    {cname}: 东南亚表={a:>8.0f}  月销总表={b:>8.0f}  {'✓' if abs(a-b)<1e-6 else '★差 %+.0f' % (a-b)}")
    a = sea["SEA Total"]["合计"]
    b = ms_s.get("东南亚")
    print(f"    东南亚合计: 东南亚表={a:>8.0f}  月销总表={b:>8.0f}  {'✓' if abs(a-b)<1e-6 else '★差 %+.0f' % (a-b)}")

print()
print("=" * 100)
print("对账四：菲律宾 新旧口径（副本 vs 原件）")
print("=" * 100)
ph_o, ph_c = sea_orig["Philippines 菲律宾销量表"], sea_copy["Philippines 菲律宾销量表"]
for label in ["跨境Lazada(中山仓)", "本土Lazada(菲仓)", "本土Shopee(菲仓)", "本土Tiktok", "线下KSON", "合计"]:
    a, b = ph_o.get(label, 0), ph_c.get(label, 0)
    print(f"  {label:<22} 原件={a:>8.0f}  副本={b:>8.0f}  Δ={b-a:+.0f}")
print(f"  月销总表菲律宾列: 原件={ms_orig_sums.get('菲律宾'):.0f}  副本={ms_copy_sums.get('菲律宾'):.0f}  Δ={ms_copy_sums.get('菲律宾',0)-ms_orig_sums.get('菲律宾',0):+.0f}")
print(f"  月销总表东南亚列: 原件={ms_orig_sums.get('东南亚'):.0f}  副本={ms_copy_sums.get('东南亚'):.0f}  Δ={ms_copy_sums.get('东南亚',0)-ms_orig_sums.get('东南亚',0):+.0f}")
print(f"  月销总表总计列:   原件={ms_orig_sums.get('总计'):.0f}  副本={ms_copy_sums.get('总计'):.0f}  Δ={ms_copy_sums.get('总计',0)-ms_orig_sums.get('总计',0):+.0f}")

print()
print("=" * 100)
print("总结论")
print("=" * 100)
print(f"  对账一（源 vs 原成品表，按渠道）: {'全部一致 ✓' if ok1 else '存在差异 ★'}")
print(f"  对账二（源 vs 中间产物，SKU级）: {'全部一致 ✓' if ok2 else '存在差异 ★'}")
