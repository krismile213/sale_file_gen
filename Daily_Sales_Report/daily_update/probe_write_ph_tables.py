# -*- coding: utf-8 -*-
"""菲律宾 9 月销量重算 · 写表探针（产出两份独立副本，不动原件）

新口径 = 段A（DAILY MONITORING 监控表 9/1~9/12）+ 段B（新订单报告 9/13~9/30，含手工单 145）
口径 100% 复刻 trunk/month_sale_file.py（Shopee L244-395 / Lazada L397-496 / Tiktok L498-588 /
新报告 L589-660）。跨境 Lazada（AY=168）与线下（BC=0）不在两源范围内，保持原件值。

改动范围（只碰菲律宾 9 月）：
  东南亚销量表副本
    · Philippines sheet：9月块渠道列 AZ(52)本土LZD / BA(53)本土SHP / BB(54)Tiktok / BD(56)合计
      + 补 G569 行（手工单新 SKU，原件无行）
    · SEA Total sheet：CA(79)菲律宾、CI(87)东南亚合计
    · 两张 sheet 行2 小计静态值
  月销总表副本
    · EM(143)菲律宾9月、HS(227)东南亚9月、KA(287)总计9月（后两者按差额联动）
  其余全部保持原件原样（脚本末段做逐格 diff 自证）
"""
import os, re, shutil, sys
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter as gcl
from copy import copy

BASE = r"D:\Desktop\work-main\sale_file_gen"
SEA_SRC = os.path.join(BASE, r"trunk\result\东南亚销量表\202609\2026年东南亚销量表Southeast Asia Sales table.xlsx")
MON_SRC = os.path.join(BASE, r"trunk\result\数据归属周期\202609\月销售总表.xlsx")
MON_FILE = r"C:\Users\dell\Desktop\DAILY MONITORING - SEPTEMBER 1-13(1).xlsx"
SNAP = os.path.join(BASE, r"trunk\data\数据归属周期\202609\快照日期\20261006\东南亚")
NEWRPT = os.path.join(SNAP, "菲律宾202609订单报告.xlsx")
OUTDIR = os.path.join(BASE, "菲律宾重算_202609")

D0, D1, D2 = "2026-09-01", "2026-09-12", "2026-09-13"
SEP = slice("2026-09-01", "2026-09-30")

SHOPEE_MAP = {"Completed":"Completed","Order Received":"Completed","Shipping":"Shipping",
              "To ship":"Shipping","Delivered":"Delivered","Cancelled":"Cancelled",
              "Unpaid":"Cancelled","Cancellation requested":"Cancelled"}
RW_PREFIX = "Order Received, but buyer can raise return/refund"
LZD_KEEP = {"confirmed","shipped","delivered","packed","pending",
            "ready_to_ship","Lost by 3PL","Package scrapped"}
TK_MAP = {"Completed":"Completed","已完成":"Completed","Selesai":"Completed",
          "Shipped":"Shipped","已发货":"Shipped","Dikirim":"Shipped",
          "To ship":"To ship","待发货":"To ship","Perlu dikirim":"To ship",
          "Canceled":"Canceled","Cancelled":"Canceled","已取消":"Canceled",
          "Dibatalkan":"Canceled","Unpaid":"Canceled","Belum dibayar":"Canceled"}
TK_KEEP = {"Shipped","Completed","To ship"}

def norm_sku(v):
    s = str(v or "").strip().upper()
    return "" if s in ("", "NAN", "NONE") else s

def q10(v):
    try: return float(v)
    except Exception: return 0.0

def day(v):
    if v is None: return ""
    if hasattr(v, "strftime"): return v.strftime("%Y-%m-%d")
    m = re.match(r"(\d{4}-\d{2}-\d{2})", str(v).strip())
    return m.group(1) if m else ""

# ---------------- 数据准备 ----------------
def build_new_quota():
    """返回 {plat: {sku: qty}}，plat ∈ Lazada/Shopee/Tiktok（新口径，9月全月）"""
    wb = load_workbook(MON_FILE, read_only=True, data_only=True)
    segA = {"Lazada": {}, "Shopee": {}, "Tiktok": {}}
    # SHP Data: A=DATE C=OrderStatus E=SKU G=Quantity
    for r in wb["SHP Data"].iter_rows(min_row=2, values_only=True):
        d = day(r[0])
        if not (D0 <= d <= D1): continue
        st = str(r[2] or "").strip()
        if st.startswith(RW_PREFIX): st = "Completed"
        st2 = SHOPEE_MAP.get(st)
        if st2 in (None, "Cancelled"): continue
        s = norm_sku(r[4])
        if s: segA["Shopee"][s] = segA["Shopee"].get(s, 0.0) + q10(r[6])
    # LZD Data: A=orderNumber B=sellerSku C=createTime F=status
    for r in wb["LZD Data"].iter_rows(min_row=2, values_only=True):
        d = day(r[2])
        if not (D0 <= d <= D1): continue
        if str(r[5] or "").strip() not in LZD_KEEP: continue
        s = norm_sku(r[1])
        if s: segA["Lazada"][s] = segA["Lazada"].get(s, 0.0) + 1.0
    # Tiktok Data: B=OrderStatus C=SellerSKU F=Quantity G=CreatedTime
    for r in wb["Tiktok Data"].iter_rows(min_row=2, values_only=True):
        d = day(r[6])
        if not (D0 <= d <= D1): continue
        st2 = TK_MAP.get(str(r[1] or "").strip())
        if st2 not in TK_KEEP: continue
        s = norm_sku(r[2])
        if s: segA["Tiktok"][s] = segA["Tiktok"].get(s, 0.0) + q10(r[5])
    wb.close()

    raw = pd.read_excel(NEWRPT, dtype=str)
    _line = ["订单ID","Sku","数量","多属性Id","平台数量","sku图片","sku中文名称"]
    _parent = [c for c in raw.columns if c not in _line]
    raw[_parent] = raw[_parent].ffill()
    v = raw[raw["订单状态"].isin(["已交运","已发货"])].copy()
    v["sku"] = v["Sku"].str.replace(r"^\d+-", "", regex=True).str.strip().str.upper()
    v["d"] = pd.to_datetime(v["创建时间"], errors="coerce").dt.strftime("%Y-%m-%d")
    v["q"] = pd.to_numeric(v["数量"], errors="coerce")
    segB = v[v["d"] >= D2]
    plat_map = {"Lazada": "Lazada", "Shopee": "Shopee", "TikTok": "Tiktok"}
    out = {"Lazada": {}, "Shopee": {}, "Tiktok": {}}
    for plat, g in segB.groupby("plat" if "plat" in segB else "平台类型"):
        key = plat_map.get(str(plat).strip())
        if key is None:  # 手工单经 ffill 继承平台类型（实测=Shopee），不会到这里
            print(f"  ⚠️ 未映射平台 {plat!r} {len(g)} 行 {g['q'].sum():.0f} 件，已忽略"); continue
        for s, q in zip(g["sku"], g["q"]):
            if s: out[key][s] = out[key].get(s, 0.0) + q
    # 合并
    NEW = {}
    for plat in out:
        NEW[plat] = dict(segA[plat])
        for s, q in out[plat].items():
            NEW[plat][s] = NEW[plat].get(s, 0.0) + q
    return NEW, segA, out

def n0(v):
    if v is None: return 0.0
    if isinstance(v, str):
        s = v.strip()
        if s in ("", "00"): return 0.0
        try: return float(s)
        except Exception: return 0.0
    return float(v or 0)

# ---------------- 主流程 ----------------
def main():
    print("① 计算新口径（监控表 9/1~9/12 + 新报告 ≥9/13）…")
    NEW, segA, segB = build_new_quota()
    tot = {p: sum(d.values()) for p, d in NEW.items()}
    print(f"   段A: LZD={sum(segA['Lazada'].values()):.0f} SHP={sum(segA['Shopee'].values()):.0f} TK={sum(segA['Tiktok'].values()):.0f}")
    print(f"   段B: LZD={sum(segB['Lazada'].values()):.0f} SHP={sum(segB['Shopee'].values()):.0f} TK={sum(segB['Tiktok'].values()):.0f}（SHP 含手工单 145）")
    print(f"   新口径合计: LZD={tot['Lazada']:.0f} SHP={tot['Shopee']:.0f} TK={tot['Tiktok']:.0f}")
    all_sku = sorted(set().union(*[set(d) for d in NEW.values()]))
    print(f"   涉及 SKU {len(all_sku)} 个")

    os.makedirs(OUTDIR, exist_ok=True)
    sea_dst = os.path.join(OUTDIR, "2026年东南亚销量表Southeast Asia Sales table.xlsx")
    mon_dst = os.path.join(OUTDIR, "月销售总表.xlsx")
    shutil.copy2(SEA_SRC, sea_dst)
    shutil.copy2(MON_SRC, mon_dst)
    print(f"\n② 已复制副本 → {OUTDIR}")

    # ============ 东南亚销量表 ============
    print("\n③ 写东南亚销量表副本…")
    wb = load_workbook(sea_dst, data_only=False)
    ph = wb["Philippines 菲律宾销量表"]
    sea = wb["SEA Total 东南亚总销量表"]

    ph_rows = {}
    for r in range(3, ph.max_row + 1):
        v = ph.cell(row=r, column=1).value
        if v: ph_rows[str(v).strip().upper()] = r
    sea_rows = {}
    for r in range(3, sea.max_row + 1):
        v = sea.cell(row=r, column=1).value
        if v: sea_rows[str(v).strip().upper()] = r

    # --- 补 G569 行（菲律宾 sheet 缺行；样式复制末行） ---
    need_add = [s for s in all_sku if s not in ph_rows]
    if need_add:
        ref = ph.max_row
        newr = ref + 1
        for col in range(1, 75):
            src = ph.cell(row=ref, column=col)
            dst = ph.cell(row=newr, column=col)
            dst._style = copy(src._style)
        ph.cell(row=newr, column=1).value = need_add[0]
        ph.cell(row=newr, column=2).value = "Google Pixel 11 Pro XL"  # 来自新报告 sku中文名称
        ph_rows[need_add[0]] = newr
        old_ref = ph.auto_filter.ref or f"A2:BV{ref}"
        ph.auto_filter.ref = re.sub(r"BV\d+$", f"BV{newr}", old_ref)
        print(f"   补行：{need_add[0]} → 菲律宾 sheet 行{newr}（样式复制行{ref}，auto_filter → {ph.auto_filter.ref}）")

    # --- 逐 SKU 写 9 月块（只写有变化的 SKU） ---
    PLAT_COL = {"Lazada": 52, "Shopee": 53, "Tiktok": 54}   # AZ / BA / BB
    changed_ph = 0
    for s in all_sku:
        r = ph_rows.get(s)
        if r is None: continue
        # 渠道列
        wrote = False
        for plat, col in PLAT_COL.items():
            old = n0(ph.cell(row=r, column=col).value)
            new = NEW[plat].get(s, 0.0)
            if abs(new - old) > 1e-9:
                ph.cell(row=r, column=col).value = int(new) if new else (0 if new == 0 else new)
                wrote = True
        # 线下 BC(55) 恒 0；跨境 AY(51) 不动
        # BD(56) = AY+AZ+BA+BB+BC（静态值，与原件一致）
        ay = n0(ph.cell(row=r, column=51).value)
        az = n0(ph.cell(row=r, column=52).value)
        ba = n0(ph.cell(row=r, column=53).value)
        bb = n0(ph.cell(row=r, column=54).value)
        bc = n0(ph.cell(row=r, column=55).value)
        bd_new = ay + az + ba + bb + bc
        bd_old = n0(ph.cell(row=r, column=56).value)
        if abs(bd_new - bd_old) > 1e-9:
            ph.cell(row=r, column=56).value = int(bd_new)
            wrote = True
        if wrote: changed_ph += 1
    print(f"   菲律宾 sheet：{changed_ph} 个 SKU 行的 9 月块已更新")

    # --- 新口径为 0 但旧值 >0 的 SKU（含不在 all_sku 的）：渠道列与 BD 归零 ---
    zeroed = 0
    for s, r in ph_rows.items():
        if s in all_sku: continue
        vals = [n0(ph.cell(row=r, column=c).value) for c in (52, 53, 54)]
        if any(v > 0 for v in vals):
            for col in (52, 53, 54):
                ph.cell(row=r, column=col).value = 0
            ay = n0(ph.cell(row=r, column=51).value)
            ph.cell(row=r, column=56).value = int(ay + 0)
            zeroed += 1
    if zeroed:
        print(f"   另有 {zeroed} 个 SKU 新口径无单 → 本土三渠道归零（跨境/线下列不变）")

    # --- SEA Total：CA(79) = 菲律宾新口径合计；CI(87) = 旧 + Δ ---
    changed_sea = 0
    for s, r in sea_rows.items():
        # 新口径合计
        new_tot = sum(NEW[p].get(s, 0.0) for p in PLAT_COL)
        # 跨境部分：菲律宾 sheet 的 AY（若该 SKU 在菲律宾 sheet 有行）
        rp = ph_rows.get(s)
        ay = n0(ph.cell(row=rp, column=51).value) if rp else 0.0
        new_ca = new_tot + ay
        old_ca = n0(sea.cell(row=r, column=79).value)
        if abs(new_ca - old_ca) > 1e-9:
            sea.cell(row=r, column=79).value = int(new_ca)
            old_ci = n0(sea.cell(row=r, column=87).value)
            sea.cell(row=r, column=87).value = int(old_ci + (new_ca - old_ca))
            changed_sea += 1
    print(f"   SEA Total：{changed_sea} 个 SKU 行的 CA/CI 已更新")

    # --- 行2 小计静态值 ---
    s_az = sum(NEW["Lazada"].values()); s_ba = sum(NEW["Shopee"].values()); s_bb = sum(NEW["Tiktok"].values())
    ay2 = n0(ph.cell(row=2, column=51).value); bc2 = n0(ph.cell(row=2, column=55).value)
    bd2 = ay2 + s_az + s_ba + s_bb + bc2
    ph.cell(row=2, column=52).value = int(s_az)
    ph.cell(row=2, column=53).value = int(s_ba)
    ph.cell(row=2, column=54).value = int(s_bb)
    ph.cell(row=2, column=56).value = int(bd2)
    sea.cell(row=2, column=79).value = int(bd2)
    old_ci2 = n0(sea.cell(row=2, column=87).value)
    old_bd2 = 3123.0
    sea.cell(row=2, column=87).value = int(old_ci2 + (bd2 - old_bd2))
    print(f"   行2小计：菲律宾 AZ={s_az:.0f} BA={s_ba:.0f} BB={s_bb:.0f} BD={bd2:.0f}；SEA Total CA={bd2:.0f} CI={old_ci2 + (bd2 - old_bd2):.0f}")
    wb.save(sea_dst)
    wb.close()
    print(f"   已保存 {sea_dst}")

    # ============ 月销总表 ============
    print("\n④ 写月销总表副本…")
    wb2 = load_workbook(mon_dst, data_only=False)
    ms = wb2["2026年"]
    m_rows = {}
    for r in range(4, ms.max_row + 1):
        v = ms.cell(row=r, column=1).value
        if v: m_rows[str(v).strip().upper()] = r
    changed_m = 0
    for s, r in m_rows.items():
        # 菲律宾新口径合计（含跨境 AY）
        new_tot = sum(NEW[p].get(s, 0.0) for p in PLAT_COL)
        rp = ph_rows.get(s)
        ay = n0(ph.cell(row=rp, column=51).value) if rp else 0.0
        new_em = new_tot + ay
        old_em = ms.cell(row=r, column=143).value
        old_em_n = n0(old_em)
        if abs(new_em - old_em_n) <= 1e-9:
            continue
        # EM(143)
        ms.cell(row=r, column=143).value = int(new_em) if new_em > 0 else (old_em if isinstance(old_em, str) and old_em.strip() == "00" else 0)
        # HS(227) 与 KA(287) 按差额联动（注意：结果为 0 也要写，不能因"与旧值相同"跳过——
        # 唯一不写的合法情形是 delta==0，而 delta==0 的 SKU 已在上面 continue 掉）
        delta = new_em - old_em_n
        for col in (227, 287):
            ov = ms.cell(row=r, column=col).value
            ovn = n0(ov)
            newv = ovn + delta
            if abs(newv - ovn) > 1e-9 or (ov is None and abs(newv) > 1e-9):
                ms.cell(row=r, column=col).value = int(newv)
        changed_m += 1
    print(f"   月销总表：{changed_m} 个 SKU 行的 EM/HS/KA 已更新")
    wb2.save(mon_dst)
    wb2.close()
    print(f"   已保存 {mon_dst}")

    print("\n⑤ 完成。副本位于：")
    print(f"   {sea_dst}")
    print(f"   {mon_dst}")

if __name__ == "__main__":
    main()
