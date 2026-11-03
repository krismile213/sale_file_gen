# -*- coding: utf-8 -*-
"""诊断：东南亚缺货明细的「在售」判定，用商品资料 vs 日销表 差多少。

背景（2026-10-29）：
  用户指出「资产文件/商品资料.xlsx」的「状态」列（在售/停售/不销售）才是最准确的，
  之前用「日销表 印尼状态列 / 泰菲近60天有销量代理」会把已停售的 SKU 误收进来。

本脚本只做统计，不改任何产物。
"""
import os
import sys
import collections
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.abspath(os.path.join(HERE, "..", "..", ".."))
SEA = os.path.join(ROOT, "Daily_Sales_Report", "东南亚日销")
PROD = os.path.join(os.path.dirname(ROOT), "资产文件", "商品资料.xlsx")

SITES = {
    "印尼": dict(f="2026年印尼智慧物流销售表1-12月.xlsx", sheet="2026年10月",
                 c_sku=0, c_name=1, c_model=2, c_state=9, c_stock=11, c_avail=12, c_mtotal=13),
    "泰国": dict(f="2026年泰国易仓销售表1-12月.xlsx", sheet="10月",
                 c_sku=0, c_name=1, c_model=2, c_state=None, c_stock=10, c_avail=11, c_mtotal=12),
    "菲律宾": dict(f="2026年菲律宾仓销售表1-12月.xlsx", sheet="10月",
                  c_sku=0, c_name=1, c_model=2, c_state=None, c_stock=10, c_avail=11, c_mtotal=12),
}


def load_product():
    wb = openpyxl.load_workbook(PROD, read_only=True, data_only=True)
    ws = wb["商品资料"]
    out = {}
    for r in ws.iter_rows(min_row=3, values_only=True):
        s = str(r[0]).strip() if r[0] else ""
        if not s:
            continue
        out[s] = dict(
            state=str(r[12]).strip() if r[12] else "",
            score=str(r[5]).strip() if r[5] else "",
            typ=str(r[6]).strip() if r[6] else "",
        )
    wb.close()
    return out


def _n(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return 0.0


def main():
    prod = load_product()
    print(f"商品资料：{len(prod)} 个 SKU，状态 {dict(collections.Counter(v['state'] for v in prod.values()))}")
    print()

    for site, cfg in SITES.items():
        wb = openpyxl.load_workbook(os.path.join(SEA, cfg["f"]), read_only=True, data_only=True)
        ws = wb[cfg["sheet"]]
        rows = []
        for r in ws.iter_rows(min_row=3, values_only=True):
            sku = str(r[cfg["c_sku"]]).strip() if r[cfg["c_sku"]] else ""
            if not sku or sku == "合计":
                continue
            rows.append(dict(
                sku=sku,
                stock=_n(r[cfg["c_stock"]]),
                avail=_n(r[cfg["c_avail"]]),
                state=str(r[cfg["c_state"]]).strip() if cfg["c_state"] is not None and r[cfg["c_state"]] else "",
            ))
        wb.close()

        miss = [x["sku"] for x in rows if x["sku"] not in prod]
        zero = [x for x in rows if x["stock"] <= 0 and x["avail"] <= 0]
        zero_prod = [x for x in zero if prod.get(x["sku"], {}).get("state") == "在售"]
        zero_stop = [x for x in zero if prod.get(x["sku"], {}).get("state") == "停售"]
        zero_other = [x for x in zero if prod.get(x["sku"], {}).get("state") not in ("在售", "停售")]

        print(f"== {site}")
        print(f"   主表 SKU {len(rows)}｜商品资料匹配 {len(rows) - len(miss)}"
              f"（{100 * (len(rows) - len(miss)) / max(len(rows), 1):.1f}%）｜未匹配 {len(miss)} {miss[:5]}")
        print(f"   零库存零在途候选 {len(zero)} 个 → 商品资料在售 {len(zero_prod)}｜停售 {len(zero_stop)}"
              f"｜其它 {len(zero_other)}")
        if zero_stop:
            print(f"   ⚠ 停售却零库存（不该进缺货清单）：{[(x['sku'], prod[x['sku']]['score']) for x in zero_stop]}")
        if cfg["c_state"] is not None:
            ds = collections.Counter(x["state"] for x in zero)
            print(f"   日销表状态列在候选池里的分布：{dict(ds)}")
            diff = [x["sku"] for x in zero
                    if x["state"] == "在售" and prod.get(x["sku"], {}).get("state") != "在售"]
            print(f"   日销表说在售、商品资料说非在售：{len(diff)} 个 {diff}")
            diff2 = [x["sku"] for x in zero
                     if x["state"] != "在售" and prod.get(x["sku"], {}).get("state") == "在售"]
            print(f"   日销表说非在售、商品资料说在售：{len(diff2)} 个 {diff2}")
        print()


if __name__ == "__main__":
    main()
