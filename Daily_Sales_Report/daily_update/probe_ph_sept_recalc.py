# -*- coding: utf-8 -*-
"""菲律宾 9 月销量重算探针（只读，不动任何成品）

口径 100% 复刻 trunk/month_sale_file.py：
  · Shopee 菲律宾 L244-395：status_map + 退货窗口前缀归一化 + 剔 Cancelled/NaN，数量=Quantity
  · Lazada 菲律宾 L397-496：每条=1 件，白名单 8 态
  · Tiktok 菲律宾 L498-588：映射后保留 Shipped/Completed/To ship
  · 新订单报告 L589-660：已交运+已发货；SKU 剥「前导数字串-」前缀；≥ start_date(2026-09-13)

新口径 = 段A（DAILY MONITORING 监控表 9/1~9/12）+ 段B（新订单报告 9/13~9/30）
旧口径 = 旧三平台报告全月（trunk 现行）
验证：① 重叠区(9/13~9/14)两源对账 ② 监控表 9/1~9/12 vs 旧报告同段
"""
import os, re, sys
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter as gcl

MON   = r"C:\Users\dell\Desktop\DAILY MONITORING - SEPTEMBER 1-13(1).xlsx"
SNAP  = r"D:\Desktop\work-main\sale_file_gen\trunk\data\数据归属周期\202609\快照日期\20261006\东南亚"
NEWRPT= os.path.join(SNAP, "菲律宾202609订单报告.xlsx")
OUTDIR = os.path.dirname(os.path.abspath(__file__))

D0, D1, D2 = "2026-09-01", "2026-09-12", "2026-09-13"   # 段A=[D0,D1]  段B=[D2,月末]
SEP = slice("2026-09-01", "2026-09-30")

# ---------- 口径表（与 trunk 逐字一致） ----------
SHOPEE_MAP = {
    "Completed":"Completed", "Order Received":"Completed",
    "Shipping":"Shipping", "To ship":"Shipping", "Delivered":"Delivered",
    "Cancelled":"Cancelled", "Unpaid":"Cancelled", "Cancellation requested":"Cancelled",
}
RW_PREFIX = "Order Received, but buyer can raise return/refund"   # → Completed
LZD_KEEP = {"confirmed","shipped","delivered","packed","pending",
            "ready_to_ship","Lost by 3PL","Package scrapped"}
TK_MAP = {
    "Completed":"Completed","已完成":"Completed","Selesai":"Completed",
    "Shipped":"Shipped","已发货":"Shipped","Dikirim":"Shipped",
    "To ship":"To ship","待发货":"To ship","Perlu dikirim":"To ship",
    "Canceled":"Canceled","Cancelled":"Canceled","已取消":"Canceled",
    "Dibatalkan":"Canceled","Unpaid":"Canceled","Belum dibayar":"Canceled",
}
TK_KEEP = {"Shipped","Completed","To ship"}

def norm_sku(v):
    s = str(v or "").strip().upper()
    return "" if s in ("", "NAN", "NONE") else s

def q10(v):
    try: return float(v)
    except Exception: return 0.0

def day(v):
    """日期 → 'YYYY-MM-DD'；支持 datetime / '2026-09-14 00:00:00' / '2026-09-14'"""
    if v is None: return ""
    if hasattr(v, "strftime"): return v.strftime("%Y-%m-%d")
    s = str(v).strip()
    m = re.match(r"(\d{4}-\d{2}-\d{2})", s)
    return m.group(1) if m else ""

def agg(rows):
    """rows: [(sku, qty)] → dict"""
    out = {}
    for sku, q in rows:
        if sku: out[sku] = out.get(sku, 0.0) + q
    return out

unmapped = {"shopee": {}, "lzd": {}, "tk": {}}

# ---------- ① 监控表三平台（流式，Tiktok 104 万行只过筛 9 月） ----------
def read_monitor():
    wb = load_workbook(MON, read_only=True, data_only=True)
    # --- SHP Data: A=DATE B=OrderID C=OrderStatus E=SKU G=Quantity ---
    shp = []
    ws = wb["SHP Data"]
    for r in ws.iter_rows(min_row=2, values_only=True):
        d = day(r[0])
        if not (SEP.start <= d <= SEP.stop): continue
        st = str(r[2] or "").strip()
        if st.startswith(RW_PREFIX): st = "Completed"
        st2 = SHOPEE_MAP.get(st)
        if st is not None and st2 is None and st not in unmapped["shopee"]:
            unmapped["shopee"][st] = 0
        if st in unmapped["shopee"]: unmapped["shopee"][st] += 1
        if st2 in (None, "Cancelled"): continue
        shp.append((d, norm_sku(r[4]), q10(r[6])))
    # --- LZD Data: A=orderNumber B=sellerSku C=createTime F=status ---
    lzd = []
    ws = wb["LZD Data"]
    for r in ws.iter_rows(min_row=2, values_only=True):
        d = day(r[2])
        if not (SEP.start <= d <= SEP.stop): continue
        st = str(r[5] or "").strip()
        if st not in LZD_KEEP:
            unmapped["lzd"][st] = unmapped["lzd"].get(st, 0) + 1
            continue
        lzd.append((d, norm_sku(r[1]), 1.0))   # trunk：每条=1 件
    # --- Tiktok Data: A=SKU(单号) B=OrderStatus C=SellerSKU F=Quantity G=CreatedTime ---
    tk = []
    ws = wb["Tiktok Data"]
    for r in ws.iter_rows(min_row=2, values_only=True):
        d = day(r[6])
        if not (SEP.start <= d <= SEP.stop): continue
        st2 = TK_MAP.get(str(r[1] or "").strip())
        if st2 is None:
            unmapped["tk"][str(r[1] or "").strip()] = unmapped["tk"].get(str(r[1] or "").strip(), 0) + 1
            continue
        if st2 not in TK_KEEP: continue
        tk.append((d, norm_sku(r[2]), q10(r[5])))
    wb.close()
    return shp, lzd, tk

# ---------- ② 新订单报告（≥D2） ----------
def read_new_report():
    raw = pd.read_excel(NEWRPT, dtype=str)
    _line = ["订单ID","Sku","数量","多属性Id","平台数量","sku图片","sku中文名称"]
    _parent = [c for c in raw.columns if c not in _line]
    raw[_parent] = raw[_parent].ffill()
    v = raw[raw["订单状态"].isin(["已交运","已发货"])].copy()
    v["sku"] = v["Sku"].str.replace(r"^\d+-", "", regex=True).str.strip().str.upper()
    v["d"] = pd.to_datetime(v["创建时间"], errors="coerce").dt.strftime("%Y-%m-%d")
    v["q"] = pd.to_numeric(v["数量"], errors="coerce")
    v["plat"] = v["平台类型"].fillna("手工订单")
    segB = v[v["d"] >= D2]
    segO = v[(v["d"] >= D0) & (v["d"] <= "2026-09-14")]   # 与监控表重叠区
    return segB, segO, v

# ---------- ③ 旧三平台报告（trunk 现行口径，当月导出不过滤；此处自加日期窗对账） ----------
def read_old_reports():
    out = {"Shopee": [], "Lazada": [], "Tiktok": []}
    # Shopee
    d = pd.read_excel(os.path.join(SNAP, "Shopee菲律宾站点202609订单报告.xlsx"),
                      parse_dates=["Order Creation Date"])
    d["订单状态"] = d["Order Status"].astype(str).str.strip()
    m = d["订单状态"].str.startswith(RW_PREFIX)
    d.loc[m, "订单状态"] = "Completed"
    d["st"] = d["订单状态"].map(SHOPEE_MAP)
    d = d[(d["st"].notna()) & (d["st"] != "Cancelled")]
    for _, r in d.iterrows():
        out["Shopee"].append((r["Order Creation Date"].strftime("%Y-%m-%d"),
                              norm_sku(r["SKU Reference No."]), q10(r["Quantity"])))
    # Lazada
    d = pd.read_excel(os.path.join(SNAP, "Lazada菲律宾站点202609订单报告.xlsx"),
                      parse_dates=["createTime"])
    d = d[d["status"].astype(str).str.strip().isin(LZD_KEEP)]
    for _, r in d.iterrows():
        out["Lazada"].append((r["createTime"].strftime("%Y-%m-%d"),
                              norm_sku(r["sellerSku"]), 1.0))
    # Tiktok
    d = pd.read_csv(os.path.join(SNAP, "Tiktok菲律宾站点202609订单报告.csv"),
                    parse_dates=["Created Time"], low_memory=False)
    d["st"] = d["Order Status"].astype(str).str.strip().map(TK_MAP)
    d = d[d["st"].isin(TK_KEEP)]
    for _, r in d.iterrows():
        out["Tiktok"].append((r["Created Time"].strftime("%Y-%m-%d"),
                              norm_sku(r["Seller SKU"]), q10(r["Quantity"])))
    return out

def sub(rows, lo, hi):
    return [(d, s, q) for (d, s, q) in rows if lo <= d <= hi]

def show(tag, dic):
    tot = sum(dic.values())
    print(f"  {tag:<34} SKU数={len(dic):>4}  件数={tot:>8.0f}")
    return tot

def main():
    print("=" * 76)
    print("① 读监控表（DAILY MONITORING）三平台 9 月明细")
    print("=" * 76)
    mshp, mlzd, mtk = read_monitor()
    print(f"  监控表 9 月行数：Shopee={len(mshp)} Lazada={len(mlzd)} Tiktok={len(mtk)}")
    print(f"  未映射状态值：Shopee={unmapped['shopee'] or '无'} | Lazada={list(unmapped['lzd'])[:8] or '无'} | Tiktok={list(unmapped['tk'])[:8] or '无'}")

    print()
    print("=" * 76)
    print("② 段A = 监控表 9/1~9/12（新口径前半段）")
    print("=" * 76)
    A = {}
    for name, rows in [("Shopee", mshp), ("Lazada", mlzd), ("Tiktok", mtk)]:
        seg = sub(rows, D0, D1)
        d = agg([(s, q) for _, s, q in seg])
        show(f"{name} 9/1~9/12", d)
        for k, v in d.items(): A[k] = A.get(k, 0.0) + v
    show("段A 合计", A)

    print()
    print("=" * 76)
    print("③ 段B = 新订单报告 ≥9/13（新口径后半段）")
    print("=" * 76)
    segB, segO, vAll = read_new_report()
    B = {}
    for plat, g in segB.groupby("plat"):
        d = agg(list(zip(g["sku"], g["q"])))
        show(f"{plat} ≥9/13", d)
        for k, val in d.items(): B[k] = B.get(k, 0.0) + val
    show("段B 合计", B)

    print()
    print("=" * 76)
    print("④ 新口径全月 = 段A + 段B")
    print("=" * 76)
    NEW = {}
    for d in (A, B):
        for k, v in d.items(): NEW[k] = NEW.get(k, 0.0) + v
    show("菲律宾 9 月【新口径】", NEW)

    print()
    print("=" * 76)
    print("⑤ 旧口径（trunk 现行：旧三平台报告全月）")
    print("=" * 76)
    old = read_old_reports()
    OLD = {}
    for name, rows in old.items():
        d = agg([(s, q) for _, s, q in rows])
        show(f"{name} 全月", d)
        for k, v in d.items(): OLD[k] = OLD.get(k, 0.0) + v
    show("菲律宾 9 月【旧口径】", OLD)

    print()
    print("=" * 76)
    print("⑥ 交叉验证 1：重叠区 9/13~9/14 监控表 vs 新报告")
    print("=" * 76)
    OV_M = {}
    for name, rows in [("Shopee", mshp), ("Lazada", mlzd), ("Tiktok", mtk)]:
        for _, s, q in sub(rows, D2, "2026-09-14"):
            if s: OV_M[s] = OV_M.get(s, 0.0) + q
    OV_N = agg(list(zip(segO["sku"], segO["q"])))
    show("监控表 9/13~9/14", OV_M)
    show("新报告 9/13~9/14", OV_N)
    keys = set(OV_M) | set(OV_N)
    diffs = [(k, OV_M.get(k, 0), OV_N.get(k, 0)) for k in keys if abs(OV_M.get(k, 0) - OV_N.get(k, 0)) > 1e-9]
    print(f"  有差异的 SKU：{len(diffs)} 个")
    for k, a, b in sorted(diffs)[:25]:
        print(f"    {k:<12} 监控表={a:>6.0f}  新报告={b:>6.0f}  差={a-b:>+5.0f}")
    print(f"  合计差 = {sum(OV_M.values())-sum(OV_N.values()):+.0f}")

    print()
    print("=" * 76)
    print("⑦ 交叉验证 2：监控表 9/1~9/12 vs 旧平台报告 9/1~9/12")
    print("=" * 76)
    for name, rows in old.items():
        mo = {"Shopee": mshp, "Lazada": mlzd, "Tiktok": mtk}[name]
        dm = agg([(s, q) for _, s, q in sub(mo, D0, D1)])
        do = agg([(s, q) for _, s, q in sub(rows, D0, D1)])
        show(f"{name} 监控表 9/1~9/12", dm)
        show(f"{name} 旧报告 9/1~9/12", do)
        ks = set(dm) | set(do)
        dd = [(k, dm.get(k, 0), do.get(k, 0)) for k in ks if abs(dm.get(k, 0) - do.get(k, 0)) > 1e-9]
        print(f"    差异 SKU {len(dd)} 个，合计差 = {sum(dm.values())-sum(do.values()):+.0f}")
        for k, a, b in sorted(dd)[:10]:
            print(f"      {k:<12} 监控表={a:>5.0f} 旧报告={b:>5.0f}")

    print()
    print("=" * 76)
    print("⑧ 新旧口径差异分解（全月）")
    print("=" * 76)
    ks = set(NEW) | set(OLD)
    rows = []
    for k in ks:
        a, b = NEW.get(k, 0.0), OLD.get(k, 0.0)
        if abs(a - b) > 1e-9:
            rows.append((k, a, b, a - b))
    rows.sort(key=lambda x: -abs(x[3]))
    print(f"  新口径={sum(NEW.values()):.0f}  旧口径={sum(OLD.values()):.0f}  差={sum(NEW.values())-sum(OLD.values()):+.0f}")
    print(f"  有差异 SKU {len(rows)} 个（前 30）：")
    for k, a, b, d in rows[:30]:
        print(f"    {k:<12} 新={a:>6.0f}  旧={b:>6.0f}  差={d:>+6.0f}")

    # 输出明细 csv（供写表探针使用）
    allk = sorted(set(NEW) | set(OLD))
    pd.DataFrame([{
        "sku": k,
        "new_qty": NEW.get(k, 0.0),
        "old_qty": OLD.get(k, 0.0),
        "segA": A.get(k, 0.0), "segB": B.get(k, 0.0),
    } for k in allk]).to_csv(
        os.path.join(OUTDIR, "ph_sept_recalc.csv"), index=False, encoding="utf-8-sig")
    print(f"\n  明细已写出：{os.path.join(OUTDIR, 'ph_sept_recalc.csv')}（{len(allk)} SKU）")

if __name__ == "__main__":
    main()
