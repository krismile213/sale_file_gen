# -*- coding: utf-8 -*-
"""V3（20261021）33 个 SKU 在新口径下的逐条归因。一次性核对用。"""
import os
import sys
import openpyxl

DESK = os.path.expanduser("~/Desktop")
if not os.path.isdir(DESK):
    DESK = r"C:\Users\dell\Desktop"
PROD = os.path.join(DESK, "东南亚商品资料及库存管理.xlsx")
V3 = os.path.join(r"C:\Users\dell\Desktop\work-main\sale_file_gen\Daily_Sales_Report",
                  "断货分析", "东南亚", "明细", "20261021",
                  "东南亚缺货60天明细_20261021_V3_商品资料口径.xlsx")


def num(v):
    if v is None:
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).replace(",", "").strip())
    except Exception:
        return 0.0


def main():
    wb = openpyxl.load_workbook(V3, read_only=True, data_only=True)
    v3 = []
    for r in wb["汇总"].iter_rows(min_row=5, max_col=14, values_only=True):
        if str(r[0]).strip() in ("印尼", "泰国", "菲律宾") and r[1]:
            v3.append((str(r[0]).strip(), str(r[1]).strip()))
    wb.close()

    wb = openpyxl.load_workbook(PROD, read_only=True, data_only=True)
    P = {}
    for site, sh, cs, ci, ct in [("印尼", "Indonesia stock", 2, 6, 8),
                                 ("泰国", "Thailand stock", 2, 6, 8),
                                 ("菲律宾", "Philippines stock", 3, 8, 10)]:
        ws = wb[sh]
        d = {}
        for r in ws.iter_rows(min_row=3, max_col=22, values_only=True):
            s = str(r[0]).strip() if r[0] else ""
            if not s or s == ".":
                continue
            d[s] = {"状态": str(r[cs]).strip() if r[cs] else "",
                    "评级": str(r[1]).strip() if r[1] else "",
                    "在库": num(r[ci]), "在途": num(r[ct])}
        P[site] = d
    wb.close()

    print(f"V3 共 {len(v3)} 个")
    print(f"{'国家':5s}|{'SKU':7s}|{'状态':8s}|{'在库':>6s}|{'在途':>6s}|{'新归类':12s}|说明")
    print("-" * 92)
    cnt = {}
    detail = []
    for site, sku in v3:
        p = P[site].get(sku)
        if not p:
            k, note, st, ins, tr = "剔除-无档案", "备货表查无此 SKU", "-", "-", "-"
        else:
            st, ins, tr = p["状态"], p["在库"], p["在途"]
            if st != "在售":
                k, note = "剔除-非在售", f"状态={st}"
            elif ins > 0:
                k, note = "剔除-有现货", f"In stock={ins:.0f}"
            elif tr > 0:
                k, note = "B 缺货在补", f"在途 {tr:.0f}（计划到货）"
            else:
                k, note = "A 真断货", "仍然断货"
        cnt[k] = cnt.get(k, 0) + 1
        detail.append((site, sku, st, ins, tr, k, note))
        ins_s = f"{ins:6.0f}" if isinstance(ins, float) else f"{ins:>6s}"
        tr_s = f"{tr:6.0f}" if isinstance(tr, float) else f"{tr:>6s}"
        print(f"{site:5s}|{sku:7s}|{st:8s}|{ins_s}|{tr_s}|{k:12s}|{note}")
    print()
    print("归因汇总：")
    for k, v in sorted(cnt.items(), key=lambda x: -x[1]):
        print(f"  {k}: {v}")
    return detail, cnt


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
