# -*- coding: utf-8 -*-
"""探针：12 本簿「分析表 ASIN 列」的公式覆盖审计。

背景（2026-10-30）：用户反映 interflow 沙特表 ASIN 大量为空，并加了一份商品资料。
读 data_only=True 会得到空值，但那只是「无缓存值」；要判断真假空，必须看公式本体。

输出：每张分析表 → 数据行数 / 有公式 / 静态值 / 真空白 / 公式指向的源 sheet 是否命中。
"""
import os
import sys
from glob import glob

import openpyxl
from openpyxl.utils import column_index_from_string as CI, get_column_letter as L

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))
BOOK = os.path.join(ROOT, "1.每日销售数据表")


def audit(path, sheet, asin_col, src_sheet):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=False)
    sh = next((x for x in wb.sheetnames if x.strip() == sheet.strip()), None)
    if sh is None:
        print(f"  !! 找不到 sheet {sheet}")
        wb.close()
        return
    ws = wb[sh]
    ci = CI(asin_col)
    rows = hasf = static = empty = 0
    miss = []
    srcs = set()
    for r in ws.iter_rows(min_row=7, max_col=ci, values_only=True):
        sku = r[0] if r else None
        if not sku or str(sku).strip() in ("", "SKU", "合计"):
            continue
        rows += 1
        v = r[ci - 1] if len(r) >= ci else None
        if isinstance(v, str) and v.startswith("="):
            hasf += 1
            if "!A:A" in v:
                srcs.add(v.split("(")[1].split("!")[0] if "(" in v else "?")
        elif v is None or str(v).strip() == "":
            empty += 1
            miss.append(str(sku).strip())
        else:
            static += 1
    wb.close()
    flag = "✅" if empty == 0 else "⛔"
    print(f"  {flag} {sh.strip():22s} {asin_col:>3s} 数据行{rows:4d} 公式{hasf:4d} 静态{static:4d} 真空{empty:3d}"
          f"  源={sorted(srcs)[:2]}" + (f"  例:{miss[:5]}" if miss else ""))


if __name__ == "__main__":
    # (文件, 分析表, ASIN列, 期望源 sheet)
    JOBS = [
        ("北美加拿大库存销售表.xlsx", "美国-每日库存分析表", "CR", "美国可售+在途库存"),
        ("北美加拿大库存销售表.xlsx", "加拿大-每日库存分析表", "CR", "加拿大可售+在途库存"),
        ("北美加拿大库存销售表.xlsx", "墨西哥-每日库存分析表 ", "CR", "墨西哥可售+在途库存"),
        ("澳大利亚库存销售表.xlsx", "澳大利亚-每日库存分析表", "CH", "在途库存"),
        ("Interflowd澳大利亚库存销售表.xlsx", "澳大利亚-每日库存分析表", "CH", "在途库存"),
        ("Interflowd沙特库存销售表.xlsx", "沙特-每日库存分析表", "CH", "沙特在途库存"),
        ("沙特库存销售表.xlsx", "沙特阿拉伯-每日库存分析表", "CH", "沙特阿拉伯在途库存"),
        ("阿联酋库存销售表.xlsx", "阿联酋-每日库存分析表", "CH", "阿联酋在途库存"),
        ("Interflowd阿联酋库存销售表.xlsx", "阿联酋-每日库存分析表", "CH", "阿联酋在途库存"),
        ("印度库存销售表.xlsx", "印度-每日库存分析表", "CH", "在途库存"),
        ("日本库存销售表.xlsx", "日本-每日库存分析表", "CH", "可售+在途库存"),
        ("欧盟英国库存销售表.xlsx", "英国-每日库存分析表", "CR", "英国可售+在途库存"),
        ("欧盟英国库存销售表.xlsx", "欧洲-每日库存分析表", "DB", "欧洲可售+在途库存"),
    ]
    for f, s, cl, src in JOBS:
        p = os.path.join(BOOK, f)
        if not os.path.exists(p):
            print(f"!! 缺文件 {f}")
            continue
        print(f"{f} / {s.strip()}")
        audit(p, s, cl, src)
