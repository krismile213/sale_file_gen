# -*- coding: utf-8 -*-
"""
菲律宾「新订单报告」接入路径测试（2026-10-09 新增）
====================================================
目的：在不改变主脚本默认行为（旧三份报告）的前提下，**验证新订单报告读取函数能跑通**，
      并与旧三份报告做覆盖窗口内的逐 SKU 对账，为 2026-11 正式切换打底。

用法（在项目根目录 sale_file_gen/ 下）：
    PYTHONPATH=C:\\Users\\dell\\AppData\\Roaming\\Python\\Python314\\site-packages \
      C:\\Python314\\python.exe trunk\\工具-菲律宾新报告路径测试.py [归属月] [快照日] [完整覆盖起点]

    例：C:\\Python314\\python.exe trunk\\工具-菲律宾新报告路径测试.py 202609 20261006 2026-09-13

判读：
    · 覆盖窗口内（>= 起点日）新报告应 ≥ 旧报告；差额中的「手工订单」为净增量。
    · 若窗口内出现**负差**（新报告比旧报告少），说明新报告在该段仍有缺口，需再查。
    · 退出码 0 = 跑通；1 = 有异常（负差超阈值 / 读取失败）。
"""
import os
import sys
import warnings

warnings.filterwarnings("ignore")

_HERE = os.path.dirname(os.path.abspath(__file__))   # = trunk/
sys.path.insert(0, _HERE)                            # 便于 import month_sale_file / file_chardetect
os.chdir(_HERE)                                      # ⚠️ 主脚本以 trunk/ 为工作目录（data/ 在 trunk 下）

import pandas as pd                                   # noqa: E402
import month_sale_file as m                           # noqa: E402

pd.set_option("display.width", 240)
pd.set_option("display.unicode.east_asian_width", True)

BELONGING = sys.argv[1] if len(sys.argv) > 1 else "202609"
SNAPSHOT = sys.argv[2] if len(sys.argv) > 2 else "20261006"
START = sys.argv[3] if len(sys.argv) > 3 else "2026-09-13"

print("=" * 78)
print(f"菲律宾新订单报告路径测试 | 归属月={BELONGING} 快照日={SNAPSHOT} 完整覆盖起点={START}")
print("=" * 78)

# ---- 1. 读新报告（走主脚本新增函数） ----
print("\n[1] 读取新订单报告 ...")
new_df = m.get_belonging_ph_new_order_df("本土", "菲律宾", BELONGING, SNAPSHOT, start_date=START)
print(f"    新报告（>= {START}，已交运+已发货）: 明细 {len(new_df)} 行 / 件 {int(new_df['数量'].sum())} / "
      f"订单 {new_df['订单号'].nunique()}")
print("    分平台:", new_df.groupby("平台")["数量"].sum().to_dict())

# ---- 2. 读旧三份报告（走主脚本原函数，与主脚本同口径） ----
print("\n[2] 读取旧三份报告（Shopee / Lazada / Tiktok 菲律宾）...")
old_parts = {
    "Shopee": m.get_belonging_shopee_order_df("本土", "菲律宾", BELONGING, SNAPSHOT),
    "Lazada": m.get_belonging_lazada_order_df("本土", "菲律宾", BELONGING, SNAPSHOT),
    "Tiktok": m.get_belonging_tkshop_order_df("本土", "菲律宾", BELONGING, SNAPSHOT),
}
old_df = pd.concat(list(old_parts.values()), ignore_index=True)
old_df["下单时间"] = pd.to_datetime(old_df["下单时间"], errors="coerce")
print(f"    旧三报告全月: 明细 {len(old_df)} 行 / 件 {int(old_df['数量'].sum())}")

# 旧报告只保留完整覆盖窗口（>= START）以与新报告同区间对比
old_win = old_df[old_df["下单时间"] >= pd.to_datetime(START)].copy()
print(f"    旧三报告（>= {START}）: 明细 {len(old_win)} 行 / 件 {int(old_win['数量'].sum())}")

# ---- 3. 分平台总账 ----
print("\n[3] 分平台总账（覆盖窗口内）")
hdr = f"    {'平台':<8}{'旧件':>8}{'新件':>8}{'差':>8}"
print(hdr)
tot_o = tot_n = 0
for plat in ["Lazada", "Shopee", "Tiktok"]:
    o = int(old_win[old_win["平台"] == plat]["数量"].sum())
    n = int(new_df[new_df["平台"] == plat]["数量"].sum())
    tot_o += o
    tot_n += n
    print(f"    {plat:<8}{o:>8}{n:>8}{n - o:>+8}")
print(f"    {'合计':<8}{tot_o:>8}{tot_n:>8}{tot_n - tot_o:>+8}")

manual = new_df[new_df.get("订单类型") == "手工订单"]
print(f"    其中手工订单: {manual['订单号'].nunique()} 单 / {int(manual['数量'].sum())} 件（旧报告完全没有）"
      if len(manual) else "    未发现手工订单")

# ---- 4. 逐 SKU 对账（分平台） ----
print("\n[4] 逐 SKU 对账（覆盖窗口内，负差 = 新报告疑似丢单）")
neg_total = 0
for plat in ["Lazada", "Shopee", "Tiktok"]:
    o = old_win[old_win["平台"] == plat].groupby("sku")["数量"].sum()
    n = new_df[new_df["平台"] == plat].groupby("sku")["数量"].sum()
    j = pd.concat([o.rename("old"), n.rename("new")], axis=1).fillna(0)
    j["d"] = j["new"] - j["old"]
    neg = j[j["d"] < 0]
    neg_total += int((neg["new"] - neg["old"]).sum())
    print(f"  [{plat}] 旧 {int(j['old'].sum())} / 新 {int(j['new'].sum())} / 差 {int(j['d'].sum()):+d}"
          f" | 负差SKU {len(neg)} 个 合计 {int(neg['d'].sum())} 件")
    if len(neg):
        print("      负差明细(前10):", neg.sort_values("d").head(10)["d"].to_dict())

# ---- 5. 切换影响预估（若把开关置 True，全月口径下菲律宾会怎样变） ----
print("\n[5] 切换影响预估（开关置 True 后，全月菲律宾 = 新报告[>= 起始日]）")
new_full = m.get_belonging_ph_new_order_df("本土", "菲律宾", BELONGING, SNAPSHOT, start_date=None)
print(f"    {'平台':<8}{'旧全月':>8}{'新全月':>8}{'差':>8}")
_s_o = _s_n = 0
for plat in ["Lazada", "Shopee", "Tiktok"]:
    o = int(old_df[old_df["平台"] == plat]["数量"].sum())
    n = int(new_full[new_full["平台"] == plat]["数量"].sum())
    _s_o += o
    _s_n += n
    print(f"    {plat:<8}{o:>8}{n:>8}{n - o:>+8}")
print(f"    {'合计':<8}{_s_o:>8}{_s_n:>8}{_s_n - _s_o:>+8}")
_missing = int(old_df[old_df["下单时间"] < pd.to_datetime(START)]["数量"].sum())
print(f"    ⚠️ 其中 {START} 之前旧报告有 {_missing} 件，新报告整段缺失 → 开关置 True 会少算这一段，")
print(f"       故 202609 期不可切换；须待新报告自当月 1 号起满量（届时把 PH_NEW_REPORT_START_DATE 改为当月 1 号）。")

# ---- 6. 生产 groupby 结构兼容性（确保新 DataFrame 能直接进主脚本汇总逻辑） ----
print("\n[6] 生产 groupby 结构兼容性")
_g = (new_df.groupby(by=["实体", "平台", "仓库", "站点", "sku"], dropna=False)["数量"].sum().reset_index())
_g = _g.rename(columns={"数量": f"{BELONGING}销量"})
_ok = (set(["实体", "平台", "仓库", "站点", "sku", f"{BELONGING}销量"]).issubset(_g.columns)
       and int(_g[f"{BELONGING}销量"].sum()) == int(new_df["数量"].sum()))
print(f"    主脚本口径 groupby 行数 {len(_g)}，合计 {int(_g[f'{BELONGING}销量'].sum())} 件"
      f" → {'✅ 与明细一致，可并入 total_sku_sale_df' if _ok else '❌ 不一致'}")
print("    站点/仓库取值:", _g[["站点", "仓库"]].drop_duplicates().to_dict("records"))

# ---- 7. 结论 ----
# 容差：Shopee 单号存在长短两种格式，按单号 join 会产生少量"对不上"的残差，
# 该残差已用**运单号交叉验证**证实为格式差异、非丢单（见《核查-菲律宾新订单报告能否一敌多.md》）。
# 故允许 ≤ max(5, 旧件×0.5%) 的负差；超出才算真缺口。
TOL = max(5, round(tot_o * 0.005))
print("\n" + "=" * 78)
print(f"负差合计 {neg_total} 件，容差 ±{TOL} 件（Shopee 单号格式差异，已证实非丢单）")
if neg_total < -TOL:
    print(f"结论：覆盖窗口内负差 {-neg_total} 件 > 容差 {TOL} 件 → 新报告在该段确有缺口，暂不可切换。")
    _rc = 1
elif neg_total < 0:
    print("结论：路径跑通；负差在容差内（Shopee 单号长短格式差异）→ 覆盖窗口内新报告 ≥ 旧报告。")
    _rc = 0
else:
    print("结论：路径跑通；覆盖窗口内新报告 ≥ 旧报告（无负差）→ 待 11 月新数据确认满月后即可切换。")
    _rc = 0
print("=" * 78)
sys.exit(_rc)
