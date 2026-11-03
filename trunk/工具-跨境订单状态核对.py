# -*- coding: utf-8 -*-
"""
跨境 Shopee「订单号 + 状态」件数核对（2026-11-02 新增；任务⑱，独立工具、只读源数据）
=====================================================================================

为什么有这个工具
----------------
2026-11-02 处理 Shopee 跨境台湾站时踩到两件事，都属于「数量算错但代码不报错」：

  ① **算多**：台湾 202610 导出 16 行、17 件，其中 **11 行状态是「不成立」**（买家取消/付款失败，
     导出「取消原因」列有值）。旧取消词表里没有「不成立」→ 会把 11 件取消单当成销量。
  ② **算少**：主脚本 `status_map` 里没登记的状态会被 `.map()` 变成 NaN，随后
     `order_df[订单状态.notna()]` 一句把它们**静默丢掉** —— 不报错、不告警、销量凭空变少。

所以每月（或任何一次怀疑数字不对时）跑本工具，按 **订单号** 聚合并按 **状态列** 分桶，
把「导出件数 / 计入件数 / 剔除件数 / 未识别件数 / 成品件数」五项摆在一起，差额非 0 就退非 0 码。

口径（与主脚本 `month_sale_file.py` 严格同源，改这边必须同步那边）
------------------------------------------------------------------
  · 台湾   计入：已完成 / 已送達 / 已送达        剔除：不成立
  · 马来西亚 计入：Completed / Order Received / Shipping / To ship / Delivered
            剔除：Cancelled / Unpaid / Cancellation requested
  · 越南   计入：Hoàn thành / Đã giao             剔除：Đã hủy
  · 其余状态 = **未识别** → 主脚本会静默丢弃 ⇒ 本工具单独列出来并让退出码=2，逼人工确认。

用法（在项目根目录 sale_file_gen/ 下）
--------------------------------------
    set PYTHONPATH=C:\\Users\\dell\\AppData\\Roaming\\Python\\Python314\\site-packages
    C:\\Python314\\python.exe trunk\\工具-跨境订单状态核对.py [归属月] [快照日] [站点]

    · 站点省略或写 all = 自动扫描快照目录下**所有**「Shopee跨境*站点{期}订单报告.xlsx」
    · 例：`python trunk\\工具-跨境订单状态核对.py 202610 20261102 台湾`
    · 只读：不改任何成品、不改源数据；只在 result/核对/ 下写一份核对表。

输出
----
    trunk/result/核对/跨境订单状态核对-{归属月}.xlsx
      sheet 汇总      每站点一行：行数/导出件/计入件/剔除件/未识别件/成品件/差异/判定
      sheet 对账      与「东南亚销量表-虾皮跨境」列、「月销总表-虾皮境内」列的三方对撞
      sheet 状态分布  站点 × 原始状态 × 行数 × 件数 × 归桶（一眼看出有没有没见过的状态）
      sheet 订单明细  站点/订单号/状态/取消原因/下单时间/行数/件数/是否计入/归桶理由

退出码
------
    0 = 全部一致；1 = 有站点「计入件数 ≠ 成品件数」；2 = 出现未识别状态（需人工确认口径）
"""
import os
import re
import sys
import warnings

warnings.filterwarnings("ignore")

import pandas as pd
import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter as gl

HERE = os.path.dirname(os.path.abspath(__file__))          # = trunk/
SNAP_ROOT = os.path.join(HERE, "data", "数据归属周期")
BIG_TABLE_TMPL = os.path.join(HERE, "result", "东南亚销量表", "{period}",
                              "2026年东南亚销量表Southeast Asia Sales table.xlsx")
MAIN_TABLE_TMPL = os.path.join(HERE, "result", "数据归属周期", "{period}", "月销售总表.xlsx")
OUT_DIR = os.path.join(HERE, "result", "核对")

# ── 站点规则（⚠️ 与 month_sale_file.py 的 status_map 一一对应，那边改了必须同步这里）──
SITE_RULES = {
    "台湾": dict(
        id_cols=("订单编号", "Order ID"),
        date_cols=("订单成立时间", "订单创建时间", "Order Creation Date"),
        status_cols=("订单状态", "Order Status"),
        reason_cols=("取消原因", "Cancel Reason", "Cancellation Reason"),
        sku_cols=("商品规格货号", "SKU Reference No.", "Seller SKU"),
        qty_cols=("Quantity per Item", "数量", "Quantity"),
        valid={"已完成", "已送達", "已送达"},
        cancel={"不成立"},
    ),
    "马来西亚": dict(
        id_cols=("Order ID", "订单编号"),
        date_cols=("Order Creation Date", "订单成立时间", "Created Time"),
        status_cols=("Order Status", "订单状态"),
        reason_cols=("Cancel Reason", "Cancellation Reason", "取消原因"),
        sku_cols=("SKU Reference No.", "Seller SKU", "商品规格货号"),
        qty_cols=("Quantity", "Quantity per Item", "数量"),
        valid={"Completed", "Order Received", "Shipping", "To ship", "Delivered"},
        cancel={"Cancelled", "Unpaid", "Cancellation requested"},
    ),
    "越南": dict(
        id_cols=("订单编号", "Order ID"),
        date_cols=("订单成立时间", "Order Creation Date"),
        status_cols=("订单状态", "Order Status"),
        reason_cols=("取消原因", "Cancel Reason"),
        sku_cols=("商品规格货号", "SKU Reference No.", "Seller SKU"),
        qty_cols=("数量", "Quantity", "Quantity per Item"),
        valid={"Hoàn thành", "Đã giao"},
        cancel={"Đã hủy"},
    ),
}

# 兜底取消词（规则表里没登记到的站点 / 校验用）：只要命中就当取消
CANCEL_PAT = re.compile(r"Cancel|取消|ยกเลิก|Đã hủy|Batal|不成立", re.I)

# ── 退货窗口期前缀（⚠️ 与 month_sale_file.py 的 return_window_prefix_map 同源）──
# 已签收、只是仍在退货窗口内的订单，源状态是「<完成态>, but buyer can raise return/refund until 2026-11-09」
# 这种带日期后缀的长文本；主脚本按**前缀**归一化成完成态、计入销量。
# ⚠️ 2026-11-02 实测：马来西亚 202610 有 **56 件**属于这种状态。本工具第一版没做前缀归一，
#    直接把这 56 件判成「未识别」并对不上成品（266 vs 322）—— 这个告警本身是对的（说明查得到），
#    但归桶必须补前缀，否则每月都会假红。**不是漏单，是口径要按前缀认。**
RETURN_WINDOW = {
    "马来西亚": ("Order Received, but buyer can raise return/refund", "Completed"),
    "新加坡": ("Order Received, but buyer can raise return/refund", "Completed"),
    "菲律宾": ("Order Received, but buyer can raise return/refund", "Completed"),
    "越南": ("Người mua xác nhận đã nhận được hàng", "Hoàn thành"),
}
STATUS_UNTIL = re.compile(r"\s+until\s+\d{4}-\d{2}-\d{2}", re.I)

# 虾皮跨境 sheet 的月块：块首 C=3、块宽 9
CROSS_SHEET_KEY = "虾皮跨境"
CROSS_BASE, CROSS_BW = 3, 9
CROSS_OFF = {"马来西亚": 0, "巴西": 1, "台湾": 2, "墨西哥": 3,
             "新加坡": 4, "越南": 5, "哥伦比亚": 6, "智利": 7}   # 8 = 月总


def pick_col(df, cands):
    """按候选列表找列名（先去空白再比，兼容表头带空格/换行）"""
    norm = {str(c).strip(): c for c in df.columns}
    for c in cands:
        if c in norm:
            return norm[c]
    for k, v in norm.items():
        for c in cands:
            if c and c.lower() in k.lower():
                return v
    return None


def read_any(path):
    low = path.lower()
    if low.endswith(".csv"):
        for enc in ("utf-8-sig", "gbk", "utf-8"):
            try:
                return pd.read_csv(path, dtype=str, encoding=enc)
            except Exception:
                continue
        raise RuntimeError("csv 编码识别失败")
    return pd.read_excel(path, dtype=str)


def audit_site(site, path, period):
    """返回 (汇总 dict, 明细 df, 状态分布 df)"""
    rule = SITE_RULES.get(site)
    df = read_any(path)
    if df.shape[1] == 0 or len(df) == 0:
        return dict(站点=site, 文件=os.path.basename(path), 行数=0, 导出件=0, 计入件=0,
                    剔除件=0, 未识别件=0, 有效单=0, 剔除单=0, 未识别单=0,
                    越界件=0, 退货窗口件=0, 备注="空表（0 列/0 行）→ 该站本期无销量"), pd.DataFrame(), pd.DataFrame()

    c_id = pick_col(df, rule["id_cols"]) if rule else None
    c_st = pick_col(df, rule["status_cols"]) if rule else None
    c_qty = pick_col(df, rule["qty_cols"]) if rule else None
    c_dt = pick_col(df, rule["date_cols"]) if rule else None
    c_rs = pick_col(df, rule["reason_cols"]) if rule else None

    miss = [n for n, c in (("订单号", c_id), ("状态", c_st), ("数量", c_qty)) if c is None]
    if miss:
        raise RuntimeError(f"缺少必需列 {miss}；实际列：{list(df.columns)[:12]}")

    out = pd.DataFrame({
        "站点": site,
        "订单号": df[c_id].astype(str).str.strip(),
        "状态": df[c_st].astype(str).str.strip(),
        "取消原因": df[c_rs].astype(str).str.strip() if c_rs else "",
        "下单时间": df[c_dt] if c_dt else "",
        "件数": pd.to_numeric(df[c_qty], errors="coerce").fillna(0),
    })

    # 状态归一：退货窗口期长文本 → 完成态（与主脚本 return_window_prefix_map 一致）
    rw = RETURN_WINDOW.get(site)
    if rw:
        pre, tgt = rw
        out["状态归一"] = out["状态"].where(~out["状态"].str.startswith(pre), tgt)
    else:
        out["状态归一"] = out["状态"]
    out["状态归并"] = out["状态归一"].str.replace(STATUS_UNTIL, " until …", regex=True)

    # 归桶：valid / cancel / 未识别（未识别 = 主脚本会静默丢弃的那些）
    valid, cancel = (rule["valid"], rule["cancel"]) if rule else (set(), set())

    def bucket(st):
        if st in valid:
            return "计入"
        if st in cancel:
            return "剔除"
        if CANCEL_PAT.search(st):
            return "剔除"
        return "未识别"

    out["归桶"] = out["状态归一"].map(bucket)
    # 月份越界（串月）：下单时间不在归属月内 —— 主脚本不按日期过滤，这种行会被算进当月
    if c_dt:
        t = pd.to_datetime(out["下单时间"], errors="coerce")
        ym = f"{period[:4]}-{period[4:]}"
        out["越界"] = t.notna() & (~t.dt.strftime("%Y-%m").eq(ym))
        out.loc[t.isna(), "越界"] = False
    else:
        out["越界"] = False

    # 订单级聚合
    g = out.groupby("订单号", sort=False)
    orders = pd.DataFrame({
        "站点": site,
        "订单号": [k for k, _ in g],
        "状态": g["状态归并"].first().values,
        "取消原因": g["取消原因"].first().values,
        "下单时间": g["下单时间"].first().values,
        "行数": g.size().values,
        "件数": g["件数"].sum().values,
        "归桶": g["归桶"].first().values,
        "越界": g["越界"].any().values,
    })
    # 同一订单出现多个不同状态 → 单独提示（极少见，但会导致归桶取错）
    multi = out.groupby("订单号")["状态"].nunique()
    if (multi > 1).any():
        orders.loc[orders["订单号"].isin(multi[multi > 1].index), "归桶"] = "多状态冲突"

    q = orders["件数"].astype(float)
    summ = dict(
        站点=site, 文件=os.path.basename(path), 行数=len(out), 导出件=int(round(q.sum())),
        计入件=int(round(q[orders["归桶"] == "计入"].sum())),
        剔除件=int(round(q[orders["归桶"] == "剔除"].sum())),
        未识别件=int(round(q[orders["归桶"].isin(["未识别", "多状态冲突"])].sum())),
        有效单=int((orders["归桶"] == "计入").sum()),
        剔除单=int((orders["归桶"] == "剔除").sum()),
        未识别单=int(orders["归桶"].isin(["未识别", "多状态冲突"]).sum()),
        越界件=int(round(q[orders["越界"]].sum())),
        退货窗口件=int(out.loc[out["状态归一"].isin(valid) & (out["状态"] != out["状态归一"]), "件数"].sum()),
        备注="",
    )
    stat = (out.groupby(["站点", "状态归并", "归桶"], sort=False)
               .agg(行数=("件数", "size"), 件数=("件数", "sum")).reset_index())
    stat["件数"] = stat["件数"].astype(int)
    return summ, orders, stat


def read_finished(period, month):
    """从成品读三站件数：东南亚表《虾皮跨境》月块 + 月销总表「虾皮境内」"""
    res = {}
    big = BIG_TABLE_TMPL.format(period=period)
    if os.path.isfile(big):
        wb = openpyxl.load_workbook(big, data_only=True)
        name = next((n for n in wb.sheetnames if CROSS_SHEET_KEY in n), None)
        if name:
            ws = wb[name]
            base = CROSS_BASE + (month - 1) * CROSS_BW
            for site, off in CROSS_OFF.items():
                col = base + off
                tot_row2 = ws.cell(2, col).value            # 行2 = 合计行
                res[site] = int(round(float(tot_row2 or 0)))
            # 月总列做恒等校验
            col_tot = base + 8
            res["_月总"] = int(round(float(ws.cell(2, col_tot).value or 0)))
            res["_八站和"] = sum(res[s] for s in CROSS_OFF)
            res["_工作表"] = name
            res["_块首"] = gl(base)
    main = MAIN_TABLE_TMPL.format(period=period)
    if os.path.isfile(main):
        wb = openpyxl.load_workbook(main, data_only=True)
        ws = wb["2026年"]
        col0 = next((c.column for c in ws[1] if str(c.value).strip() == "虾皮境内"), None)
        if col0:
            col = col0 + month - 1
            s = 0
            for r in range(4, ws.max_row + 1):
                v = ws.cell(r, col).value
                if isinstance(v, (int, float)):
                    s += v
                elif isinstance(v, str) and v.strip() not in ("", "00", "-"):
                    try:
                        s += float(v)
                    except ValueError:
                        pass
            res["_月销虾皮境内"] = int(round(s))
            res["_月销列"] = gl(col)
    return res


def write_xlsx(path, summ_df, stat_df, detail_df, recon_rows, period, month):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    wb = openpyxl.Workbook()
    head_fill = PatternFill("solid", fgColor="1F4E79")
    head_font = Font(color="FFFFFF", bold=True)
    warn_fill = PatternFill("solid", fgColor="FFF2CC")
    bad_fill = PatternFill("solid", fgColor="FFC7CE")
    ok_fill = PatternFill("solid", fgColor="C6EFCE")

    def dump(ws, df, widths=None, freeze="A2"):
        ws.append(list(df.columns))
        for c in ws[1]:
            c.fill, c.font = head_fill, head_font
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        for row in df.itertuples(index=False):
            ws.append([("" if (isinstance(v, float) and pd.isna(v)) else v) for v in row])
        ws.freeze_panes = freeze
        ws.auto_filter.ref = ws.dimensions
        for i, w in enumerate(widths or [], start=1):
            ws.column_dimensions[gl(i)].width = w

    ws = wb.active
    ws.title = "汇总"
    dump(ws, summ_df, [10, 42, 8, 10, 10, 10, 11, 9, 9, 10, 9, 46])
    for r in range(2, ws.max_row + 1):
        j = ws.cell(r, 12).value or ""
        if "≠" in str(j) or "未识别" in str(j) or "越界" in str(j):
            for c in range(1, ws.max_column + 1):
                ws.cell(r, c).fill = bad_fill
    ws.append([])
    ws.append([f"归属月 {period}｜月块 {month} 月｜成品对照取自 东南亚销量表《虾皮跨境》第 {month} 月块 与 月销总表「虾皮境内」列"])

    ws2 = wb.create_sheet("对账")
    dump(ws2, pd.DataFrame(recon_rows, columns=["核对项", "单据侧(件)", "成品侧(件)", "差异", "判定"]),
         [26, 12, 12, 10, 60])
    for r in range(2, ws2.max_row + 1):
        v = str(ws2.cell(r, 5).value or "")
        fill = ok_fill if v.startswith("✅") else bad_fill
        for c in range(1, 6):
            ws2.cell(r, c).fill = fill

    ws3 = wb.create_sheet("状态分布")
    dump(ws3, stat_df, [10, 26, 10, 8, 10])

    ws4 = wb.create_sheet("订单明细")
    dump(ws4, detail_df, [10, 26, 14, 30, 20, 7, 8, 10, 9])

    wb.save(path)


def main(period="202610", snap="20261102", site_arg="all"):
    month = int(period[4:])
    snap_dir = os.path.join(SNAP_ROOT, period, "快照日期", snap)
    cross_dir = os.path.join(snap_dir, "东南亚", "跨境")
    if not os.path.isdir(cross_dir):
        cross_dir = os.path.join(snap_dir, "跨境")            # 兼容老布局
    print("=" * 96)
    print(f"跨境 Shopee 订单状态核对 | 归属月={period} 快照={snap} 目录={cross_dir}")
    print("=" * 96)
    if not os.path.isdir(cross_dir):
        print(f"❌ 找不到跨境报告目录：{cross_dir}")
        return 1

    files = sorted(f for f in os.listdir(cross_dir)
                   if "Shopee" in f and f"站点{period}订单报告" in f
                   and f.lower().endswith((".xlsx", ".xls", ".csv")))
    if site_arg and site_arg.lower() != "all":
        files = [f for f in files if site_arg in f]
    if not files:
        print("❌ 该目录没有可核对的 Shopee 跨境订单报告")
        return 1

    summ, stats, details = [], [], []
    for f in files:
        site = next((s for s in SITE_RULES if s in f), None)
        if site is None:
            print(f"  [跳过] {f} —— 未登记规则的站点（如巴西/墨西哥目的地导出，不进月销口径）")
            continue
        try:
            s, od, st = audit_site(site, os.path.join(cross_dir, f), period)
        except Exception as e:
            print(f"  ❌ {site} 读取失败：{e}")
            summ.append(dict(站点=site, 文件=f, 行数=0, 导出件=0, 计入件=0, 剔除件=0,
                             未识别件=0, 有效单=0, 剔除单=0, 未识别单=0, 越界件=0,
                             退货窗口件=0, 备注=f"读取失败：{e}"))
            continue
        summ.append(s)
        if len(od):
            details.append(od)
        if len(st):
            stats.append(st)
        print(f"  {site:6s} 行={s['行数']:4d} 导出={s['导出件']:5d} 计入={s['计入件']:5d} "
              f"剔除={s['剔除件']:4d} 未识别={s['未识别件']:4d} 有效单={s['有效单']:3d}")

    fin = read_finished(period, month)
    if fin.get("_工作表"):
        print(f"\n成品：{fin['_工作表']} 第 {month} 月块（块首 {fin['_块首']}）→ "
              f"台湾 {fin.get('台湾', 0)}、马来 {fin.get('马来西亚', 0)}、越南 {fin.get('越南', 0)}、"
              f"月总 {fin.get('_月总', 0)}")

    # 逐站对撞
    for s in summ:
        got = fin.get(s["站点"])
        if got is None:
            s["备注"] = (s["备注"] + "；成品未找到该站列").strip("；")
            continue
        diff = s["计入件"] - got
        if diff:
            s["备注"] = (s["备注"] + f"；计入 {s['计入件']} ≠ 成品 {got}（差 {diff:+d}）").strip("；")
        if s["未识别件"]:
            s["备注"] = (s["备注"] + f"；⚠️ 未识别 {s['未识别件']} 件（可能被主脚本静默丢弃）").strip("；")
        if s["越界件"]:
            s["备注"] = (s["备注"] + f"；⚠️ 越界 {s['越界件']} 件（下单时间不在 {period}）").strip("；")
        if s.get("退货窗口件"):
            s["备注"] = (s["备注"] + f"；含退货窗口期 {s['退货窗口件']} 件（按口径计入）").strip("；")

    summ_df = pd.DataFrame(summ)[["站点", "文件", "行数", "导出件", "计入件", "剔除件", "未识别件",
                                  "有效单", "剔除单", "未识别单", "越界件", "退货窗口件", "备注"]]
    summ_df = summ_df.rename(columns={"文件": "源文件"})
    stat_df = pd.concat(stats, ignore_index=True) if stats else pd.DataFrame(
        columns=["站点", "状态归并", "归桶", "行数", "件数"])
    detail_df = pd.concat(details, ignore_index=True) if details else pd.DataFrame(
        columns=["站点", "订单号", "状态", "取消原因", "下单时间", "行数", "件数", "归桶", "越界"])

    # 对账表
    recon = []
    for s in summ:
        got = fin.get(s["站点"])
        if got is None:
            continue
        diff = s["计入件"] - got
        recon.append((f"{s['站点']}：订单级计入 ↔ 虾皮跨境列", s["计入件"], got, diff,
                      "✅ 一致" if diff == 0 else f"❌ 差 {diff:+d}"))
    # ⚠️ 只核一个站时**不能**做「三站合计 ↔ 虾皮境内」：另外两个站的件数没进来，差必定非 0（假红）。
    cross3 = sum(s["计入件"] for s in summ if s["站点"] in ("马来西亚", "越南", "台湾"))
    only_one = bool(site_arg) and site_arg.lower() != "all"
    if "_月销虾皮境内" in fin and not only_one:
        d = cross3 - fin["_月销虾皮境内"]
        recon.append(("三站跨境合计 ↔ 月销「虾皮境内」列（列 %s）" % fin.get("_月销列", "?"),
                      cross3, fin["_月销虾皮境内"], d, "✅ 一致" if d == 0 else f"❌ 差 {d:+d}"))
    if "_八站和" in fin:
        d = fin["_八站和"] - fin["_月总"]
        recon.append(("虾皮跨境 8 站之和 ↔ 月总列", fin["_八站和"], fin["_月总"], d,
                      "✅ 一致" if d == 0 else f"❌ 差 {d:+d}"))

    out = os.path.join(OUT_DIR, f"跨境订单状态核对-{period}.xlsx")
    write_xlsx(out, summ_df, stat_df, detail_df, recon, period, month)

    print("\n" + "-" * 96)
    print(summ_df.to_string(index=False))
    print("\n对账：")
    for a, b, c, d, v in recon:
        print(f"  {v}  {a}：单据 {b} vs 成品 {c}（差 {d:+d}）")
    print(f"\n核对表 → {out}")

    bad = [r for r in recon if r[3] != 0]
    unknown = sum(s["未识别件"] for s in summ)
    if bad:
        print("\n❌ 有站点与成品不一致，先别用本期数，回去查源报告/成品。")
        return 1
    if unknown:
        print(f"\n⚠️ 有 {unknown} 件属于「未识别状态」——请人工确认口径后补进 SITE_RULES 与主脚本 status_map。")
        return 2
    print("\n✅ 全部一致（计入件数 == 成品件数；无未识别状态）。")
    return 0


if __name__ == "__main__":
    os.chdir(HERE)
    a = sys.argv[1:]
    sys.exit(main(a[0] if len(a) > 0 else "202610",
                  a[1] if len(a) > 1 else "20261102",
                  a[2] if len(a) > 2 else "all"))
