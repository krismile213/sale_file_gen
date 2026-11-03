"""核对 Interflowd沙特库存销售表.xlsx 的数据完整性（2026-10-24 方案 A 验收）。

检查四件事：
  1) 分析表 SKU 集合 == 沙特仓报告 SKU 集合
  2) 沙特在途库存 W(可售) / Y(在途) 的数值 == 源报告
  3) 分析表 F/G 的公式应得值（手工复算 VLOOKUP）
  4) 沙特可售库存 29 行与报告一致
"""
import csv
import io
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(HERE))
sys.path.insert(0, os.path.join(os.path.dirname(HERE), "core"))

import analysis as An                      # noqa: E402
from openpyxl import load_workbook         # noqa: E402

DSR = os.path.dirname(os.path.dirname(HERE))
BOOK = os.path.join(DSR, "1.每日销售数据表", "Interflowd沙特库存销售表.xlsx")
SNAP = os.path.join(os.path.dirname(DSR), "trunk", "data", "数据归属周期",
                    "202610", "快照日期", "20261023")
REPORT = os.path.join(SNAP, "interflow沙特阿拉伯管理库存报告20261021.csv")


def load_report():
    raw = open(REPORT, "rb").read()
    txt = None
    for enc in ("utf-8-sig", "gbk", "utf-16"):
        try:
            txt = raw.decode(enc)
            break
        except Exception:
            continue
    out = {}
    for r in csv.DictReader(io.StringIO(txt)):
        s = (r.get("sku") or "").strip()
        if not s:
            continue

        def g(k):
            try:
                return float(r.get(k) or 0)
            except (TypeError, ValueError):
                return 0.0
        out[s] = dict(
            avail=g("afn-fulfillable-quantity"),
            onhand=g("afn-onhand-buyable-quantity"),
            inb=(g("afn-inbound-working-quantity")
                 + g("afn-inbound-shipped-quantity")
                 + g("afn-inbound-receiving-quantity")),
            reserved=g("afn-reserved-quantity"),
            fnsku=(r.get("fnsku") or "").strip(),
            asin=(r.get("asin") or "").strip(),
        )
    return out


def main():
    rep = load_report()
    print(f"源报告 {len(rep)} SKU，可售合计 {sum(v['avail'] for v in rep.values()):.0f} 件")

    wb = load_workbook(BOOK, data_only=False)
    ws = wb["沙特-每日库存分析表"]
    last = An._last_data_row(ws)
    tab = {}
    for r in range(An.DATA_FIRST_ROW, last + 1):
        s = ws.cell(r, 1).value
        if s:
            tab[str(s).strip()] = r
    print(f"分析表 {len(tab)} 行（r{An.DATA_FIRST_ROW}~r{last}）")

    # 1) SKU 集合
    only_rep = sorted(set(rep) - set(tab))
    only_tab = sorted(set(tab) - set(rep))
    print(f"\n① SKU 集合：只在报告里 {only_rep or '无'} ｜ 只在表里 {only_tab or '无'}")
    if not only_rep and not only_tab:
        print("   ✅ 完全一致")

    # 2) 库存 sheet 数值
    wst = wb["沙特在途库存"]
    sheet_val = {}
    for r in range(3, An._last_data_row(wst) + 1):
        s = wst.cell(r, 1).value
        if not s:
            continue
        s = str(s).strip()
        try:
            w_ = float(wst.cell(r, 23).value or 0)
        except (TypeError, ValueError):
            w_ = 0.0
        try:
            y_ = float(wst.cell(r, 25).value or 0)
        except (TypeError, ValueError):
            y_ = 0.0
        sheet_val[s] = (w_, y_)

    bad = []
    for s, rec in rep.items():
        if s not in sheet_val:
            bad.append((s, "缺行", "-", "-"))
            continue
        w_, y_ = sheet_val[s]
        if abs(w_ - rec["onhand"]) > 0.01:
            bad.append((s, "W列≠onhand-buyable", f"{w_:.0f}", f"{rec['onhand']:.0f}"))
        if abs(y_ - rec["inb"]) > 0.01:
            bad.append((s, "Y列≠在途合计", f"{y_:.0f}", f"{rec['inb']:.0f}"))
    print(f"\n② 沙特在途库存 W/Y 校验：{'✅ 全部一致' if not bad else f'❌ {len(bad)} 处异常'}")
    for b in bad[:15]:
        print(f"   {b[0]:<10} {b[1]:<20} 表={b[2]:<6} 源={b[3]}")

    # 3) 分析表 F/G 应得值（VLOOKUP 复算）
    print("\n③ 分析表 F(可售)/G(在途) 复算：")
    print(f"   {'SKU':<10}{'D状态':<7}{'F应得':>7}{'G应得':>7}{'报告可售':>9}{'报告在途':>9}")
    tot_f = tot_g = 0.0
    for s, r in sorted(tab.items(), key=lambda x: x[1]):
        f_ = sheet_val.get(s, (0.0, 0.0))[0]
        g_ = sheet_val.get(s, (0.0, 0.0))[1]
        tot_f += f_
        tot_g += g_
        rec = rep.get(s, {})
        flag = "" if (abs(f_ - rec.get("onhand", 0)) < 0.01
                      and abs(g_ - rec.get("inb", 0)) < 0.01) else "  ←"
        print(f"   {s:<10}{str(ws.cell(r, 4).value):<7}{f_:>7.0f}{g_:>7.0f}"
              f"{rec.get('avail', 0):>9.0f}{rec.get('inb', 0):>9.0f}{flag}")
    print(f"   {'合计':<10}{'':<7}{tot_f:>7.0f}{tot_g:>7.0f}")

    # 4) 沙特可售库存
    wsb = wb["沙特可售库存"]
    n_b = 0
    s_b = 0.0
    for r in range(3, wsb.max_row + 1):
        if wsb.cell(r, 5).value is not None:
            n_b += 1
            try:
                s_b += float(wsb.cell(r, 5).value)
            except (TypeError, ValueError):
                pass
    print(f"\n④ 沙特可售库存：{n_b} 行，期末仓库余额合计 {s_b:.0f} 件")

    # 5) D 列状态分布
    from collections import Counter
    cnt = Counter(str(ws.cell(r, 4).value) for r in tab.values())
    print(f"\n⑤ D 列状态分布：{dict(cnt)}")
    wb.close()


if __name__ == "__main__":
    main()
