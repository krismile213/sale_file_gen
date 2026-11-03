# -*- coding: utf-8 -*-
"""只读复核：2026年9月 东南亚数据全链路（中间产物 → 主簿 → 单国簿 → 月销总表 → 原始报告）"""
import os, re, datetime
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter as gcl

R = r"D:\Desktop\work-main\sale_file_gen"
TMP = os.path.join(R, "trunk", "sale_file_tmp")
SNAP = os.path.join(R, "trunk", "data", "数据归属周期", "202609", "快照日期", "20261006")
SNAP_SEA = os.path.join(SNAP, "东南亚")
RESULT_SEA = os.path.join(R, "trunk", "result", "东南亚销量表", "202609")
MONTH_TOTAL = os.path.join(R, "trunk", "result", "数据归属周期", "202609", "月销售总表.xlsx")
EXP_SEA, EXP_ID, EXP_TOTAL = 10948, 4026, 62837
fails = []

def check(name, got, exp, extra=""):
    ok = (got == exp)
    print(f"{'PASS' if ok else 'FAIL'}  {name}: {got}" + ("" if ok else f"  (期望 {exp})") + (f"  {extra}" if extra else ""))
    if not ok:
        fails.append(name)

def num(v):
    return v if isinstance(v, (int, float)) else 0

# ============ A. 中间产物 ============
print("=" * 70)
print("A. 中间产物（10-09 16:58 最终跑）")
sea = pd.read_csv(os.path.join(TMP, "se_a_sku_sale_df.csv"), index_col=0)
check("A1 se_a_sku_sale_df 东南亚9月合计", int(sea["202609销量"].sum()), EXP_SEA)
print("   实体/平台/仓库/站点 分解：")
print(sea.groupby(["实体", "平台", "仓库", "站点"])["202609销量"].sum().to_string())
ind_mask = sea["站点"].astype(str).str.contains("印尼")
check("A2 印尼9月合计", int(sea[ind_mask]["202609销量"].sum()), EXP_ID)
offline = sea[ind_mask & (sea["实体"].astype(str).str.contains("本土"))]
print(f"   印尼本土(线下) = {int(offline['202609销量'].sum())}（期望 2775）")
summ = pd.read_csv(os.path.join(TMP, "total_southeast_asia_summary.csv"), index_col=0)
check("A3 SKU汇总表 合计", int(summ["202609销量"].sum()), EXP_SEA)
det = sea.groupby("sku")["202609销量"].sum()
m = det.to_frame("明细").join(summ.set_index("SKU")["202609销量"].rename("汇总"))
bad = m[(m["明细"] != m["汇总"].fillna(0))]
check("A4 明细长表 vs SKU汇总 不一致SKU数", len(bad), 0)
if len(bad):
    print(bad.head(10).to_string())

# ============ B. 主簿 ============
print("=" * 70)
print("B. 主簿 2026年东南亚销量表（10-09 16:59）")
main_path = os.path.join(RESULT_SEA, "2026年东南亚销量表Southeast Asia Sales table.xlsx")
wb = load_workbook(main_path, read_only=True, data_only=True)
sheet_tot = {}   # sheet -> (row2值, 数据行求和, 块渠道列范围, 合计列)
for sn in wb.sheetnames:
    ws = wb[sn]
    if sn.startswith("2023"):
        continue
    # 找行1含"9月"的列 = 当月合计列
    hit = None
    for c in ws[1]:
        if c.value is not None and "9月" in str(c.value):
            hit = c.column
            break
    if hit is None:
        print(f"   [skip] {sn}: 行1无9月标签")
        continue
    # 渠道列 = 合计列往前推到上一个"月"标签/表头起点：用块宽映射
    BW = {"Philippines 菲律宾销量表": 6, "Thailand 泰国销量表": 5, "Indonesia 印尼销量表": 4,
          "Singapore 新加坡销量表": 3, "Vietnam Lazada 越南销销量表": 1, "Malaysia Lazada 马来西亚销量表": 1,
          "Malaysia Shopee 虾皮跨境销量表": 9, "Group buying distributor团购经销商": 1,
          "SEA Total 东南亚总销量表": 9}
    bw = BW.get(sn, 9)
    ch_lo, ch_hi = (hit, hit) if bw == 1 else (hit - (bw - 1), hit - 1)
    # 数据行求和（行3起；渠道列）
    ds = 0
    for row in ws.iter_rows(min_row=3, max_col=hit):
        for i, c in enumerate(row, start=1):
            if ch_lo <= i <= ch_hi and isinstance(c.value, (int, float)):
                ds += c.value
    # 行2 合计行缓存值
    r2v = None
    for i, c in enumerate(ws[2], start=1):
        if i == hit:
            r2v = c.value
    sheet_tot[sn] = (r2v, ds, ch_lo, ch_hi)
    print(f"   {sn:<42} 9月合计列={gcl(hit)}({hit}) 渠道列={gcl(ch_lo)}..{gcl(ch_hi)}  行2值={r2v}  数据行求和={ds}  {'一致' if num(r2v)==ds else '★不一致'}")
# SEA Total 9月
sea_tot = sheet_tot.get("SEA Total 东南亚总销量表")
check("B1 SEA Total 9月合计", int(sea_tot[1]), EXP_SEA)
# 各国 sheet 9月 合计
country_sheets = [s for s in sheet_tot if s != "SEA Total 东南亚总销量表"]
s = sum(int(v[1]) for v in sheet_tot.items() if v[0] != "SEA Total 东南亚总销量表")
check("B2 各国sheet 9月求和", s, EXP_SEA, f"（{len(country_sheets)} 张）")
wb.close()

# ============ C. SEA Total 分渠道 vs 中间产物 ============
print("=" * 70)
print("C. SEA Total 9月 分渠道 vs 中间产物")
wb = load_workbook(main_path, read_only=True, data_only=True)
ws = wb["SEA Total 东南亚总销量表"]
tot_col = None
for c in ws[1]:
    if c.value is not None and "9月" in str(c.value):
        tot_col = c.column
ch_lo = tot_col - 8
# 行2 = 各渠道合计
r2 = {c.column: c.value for c in ws[2] if ch_lo <= c.column <= tot_col}
wb.close()
print("   SEA Total 9月 行2 各渠道值：", {gcl(k): v for k, v in r2.items()})
# 中间产物口径映射
def site_sum(mask):
    return int(sea[mask]["202609销量"].sum())
site = sea["站点"].astype(str); plat = sea["平台"].astype(str); ent = sea["实体"].astype(str)
mapping = {
    "菲律宾": site_sum(site.str.contains("菲律宾")),
    "泰国": site_sum(site.str.contains("泰国")),
    "印尼": site_sum(site.str.contains("印尼")),
    "新加坡": site_sum(site.str.contains("新加坡")),
    "越南": site_sum(site.str.contains("越南")),
    "马来Lazada": site_sum(site.str.contains("马来") & plat.str.contains("Lazada")),
    "虾皮跨境": site_sum(plat.str.contains("Shopee") & ent.str.contains("跨境")),
    "团购": site_sum(plat.str.contains("团购")),
}
for i, (k, v) in enumerate(mapping.items()):
    book_v = num(r2.get(ch_lo + i))
    check(f"C-{k}", book_v, v, f"（簿 {book_v} vs 中间产物 {v}）")
wb.close()

# ============ D. 单国交付簿 ============
print("=" * 70)
print("D. 单国交付簿（Philippines/Indonesia/Thailand_202609_sales.xlsx）")
for fn in ["Philippines_202609_sales.xlsx", "Indonesia_202609_sales.xlsx", "Thailand_202609_sales.xlsx"]:
    p = os.path.join(RESULT_SEA, fn)
    wb2 = load_workbook(p, read_only=True, data_only=True)
    for sn in wb2.sheetnames:
        ws = wb2[sn]
        hits = [(gcl(c.column), str(c.value)[:40]) for c in ws[1] if c.value is not None and ("9月" in str(c.value) or "Sep" in str(c.value))]
        print(f"   {fn} [{sn}] max_row={ws.max_row} max_col={ws.max_column} 行1含9月/Sep: {hits}")
        if hits:
            col_letter, label = hits[0]
            from openpyxl.utils import column_index_from_string as cifs
            col = cifs(col_letter)
            r2v = next((c.value for c in ws[2] if c.column == col), None)
            ds = 0
            for row in ws.iter_rows(min_row=3, max_col=col):
                for c in row:
                    if isinstance(c.value, (int, float)) and c.column < col:
                        ds += c.value
            print(f"      → {col_letter}列 行2值={r2v} 该列左侧数据行求和={ds}")
    wb2.close()

# ============ E. 月销总表 ============
print("=" * 70)
print("E. 月销总表（result/数据归属周期/202609/月销售总表.xlsx, 10-09 16:58）")
wbm = load_workbook(MONTH_TOTAL, read_only=True, data_only=True)
ws = wbm["2026年"]
r1 = {c.column: str(c.value) for c in ws[1] if c.value is not None}
r2 = {c.column: str(c.value) for c in ws[2] if c.value is not None}
# 找"东南亚"组：行1 含"东南亚"的列起点；9月 = 行2 含"9月"
sea_cols = {k: v for k, v in r1.items() if "东南亚" in v or "SEA" in v.upper()}
print("   行1 东南亚相关列：", {gcl(k): v[:20] for k, v in sea_cols.items()})
sep_cols = {k: v for k, v in r2.items() if v.strip() == "9月"}
print("   行2 '9月' 标签列：", sorted(gcl(k) for k in sep_cols))
# 取东南亚组块（连续含东南亚标签的列 ~ 其后 12 列中找9月）
if sea_cols:
    g0 = min(sea_cols)
    grp = [g0 + i for i in range(12)]
    print(f"   东南亚组列范围: {gcl(g0)}..{gcl(g0+11)}  行1标签: {[r1.get(g0+i,'')[:12] for i in range(12)]}")
    # 9月列 = 组首+8
    c9 = g0 + 8
    print(f"   → 9月列 = {gcl(c9)}  行2标签={r2.get(c9)}  行1标签={r1.get(c9)}")
    # 数据行求和（行4起）
    ds = 0
    for row in ws.iter_rows(min_row=4, max_col=c9):
        for c in row:
            if c.column == c9 and isinstance(c.value, (int, float)):
                ds += c.value
    print(f"   {gcl(c9)} 数据行求和 = {ds}")
wbm.close()

# ============ F. 原始报告 ============
print("=" * 70)
print("F. 原始报告与快照")
# F1 快照双位置检查
root_files = [f for f in os.listdir(SNAP) if os.path.isfile(os.path.join(SNAP, f))]
sub_files = [f for f in os.listdir(SNAP_SEA) if os.path.isfile(os.path.join(SNAP_SEA, f))]
cross_dir = os.path.join(SNAP_SEA, "跨境")
cross_files = [f for f in os.listdir(cross_dir) if os.path.isfile(os.path.join(cross_dir, f))] if os.path.isdir(cross_dir) else []
dup = set(root_files) & set(sub_files)
print(f"   F1 快照根目录文件 {len(root_files)} 个；东南亚/ {len(sub_files)} 个；跨境/ {len(cross_files)} 个；同名跨位置: {dup or '无'}")
# F2 印尼海外仓深查
p = os.path.join(SNAP_SEA, "印尼海外仓202609订单报告.xlsx")
df = pd.read_excel(p)
oid, qty = "订单号\nOrder No.", "商品数量\nQuantity"
d = df[[oid, qty, "订单状态\nStatus", "订单来源\nOrder Source", "订单时间\nOrder Time", "商品编码\nGoods No"]].copy()
d[oid] = d[oid].astype(str).str.strip()
d = d[d[oid].str.startswith("YFDZ") & (d["订单状态\nStatus"].astype(str).str.startswith("OUTER"))]
check("F2 印尼海外仓 YFDZ+OUTER 行数", len(d), 219)
check("F3 印尼海外仓 YFDZ+OUTER 件数", int(d[qty].sum()), 2775)
dupkey = d.duplicated(subset=[oid, "商品编码\nGoods No"]).sum()
check("F4 (订单号,商品编码) 重复数", int(dupkey), 0)
dt = pd.to_datetime(d["订单时间\nOrder Time"], errors="coerce")
out = dt[(dt < "2026-09-01") | (dt > "2026-09-30 23:59")].shape[0]
check("F5 订单时间出9月界行数", int(out), 0)
two = d[d[oid].isin(["YFDZ-20260923-005", "YFDZ-20260923-010"])]
print(f"   F6 两笔争议单: {two[[oid,'商品编码\nGoods No',qty]].to_dict('records')}")
# F7 新菲律宾报告开关
src = open(os.path.join(R, "trunk", "month_sale_file.py"), encoding="utf-8").read()
sw = re.search(r"USE_PH_NEW_ORDER_REPORT\s*=\s*(True|False)", src)
print(f"   F7 主脚本 USE_PH_NEW_ORDER_REPORT = {sw.group(1) if sw else '??'}（False=菲律宾仍走旧三报告，145件手工单不计入9月）")

print("=" * 70)
print(("全部通过 ✅" if not fails else f"未过项：{fails}"))
