# -*- coding: utf-8 -*-
"""探针：商品资料 SKU→ASIN 与 12 本簿报告源表 SKU→ASIN 是否一致（能否当兜底源）。

2026-10-30 用户新加 `资产文件/商品资料.xlsx`（SKU=A / ASIN=B / 状态=M，646 SKU ASIN 全有），
想用它补分析表 ASIN。先回答：两边会不会打架。
"""
import os
from glob import glob as gglob

import openpyxl

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))
WORK = os.path.dirname(ROOT)                     # sale_file_gen
MAIN = os.path.dirname(WORK)                     # work-main
BOOK = os.path.join(ROOT, "1.每日销售数据表")
PROD = os.path.join(MAIN, "资产文件", "商品资料.xlsx")


def load_prod():
    wb = openpyxl.load_workbook(PROD, read_only=True, data_only=True)
    ws = wb["商品资料"]
    m, blank = {}, 0
    for r in ws.iter_rows(min_row=3, max_col=2, values_only=True):
        s = str(r[0]).strip() if r[0] else ""
        if not s or s == ".":
            continue
        a = str(r[1]).strip() if r[1] else ""
        if a:
            m[s] = a
        else:
            blank += 1
    wb.close()
    return m, blank


def src_pairs(path):
    """遍历簿内所有「报告源 sheet」（表头行含 sku/asin 的报告），返回 {sku: asin} 合并。"""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    out = {}
    for s in wb.sheetnames:
        ws = wb[s]
        try:
            hdr = list(ws.iter_rows(min_row=1, max_row=4, max_col=12, values_only=True))
        except Exception:
            continue
        for hr in hdr:
            low = [(str(v).strip().lower() if v is not None else "") for v in (hr or [])]
            if "sku" in low and "asin" in low:
                si, ai = low.index("sku"), low.index("asin")
                ri = hdr.index(hr) + 2
                for r in ws.iter_rows(min_row=ri, max_col=max(si, ai) + 1, values_only=True):
                    k = r[si] if len(r) > si else None
                    if not k:
                        continue
                    a = r[ai] if len(r) > ai else None
                    k, a = str(k).strip(), (str(a).strip() if a else "")
                    if k and a and k not in out:
                        out[k] = a
                break
    wb.close()
    return out


if __name__ == "__main__":
    prod, pblank = load_prod()
    print(f"商品资料：{len(prod)} 个 SKU 有 ASIN（空 {pblank}）\n")
    tot_new = tot_conf = 0
    for p in sorted(gglob(os.path.join(BOOK, "*.xlsx"))):
        b = os.path.basename(p)
        if b.startswith("~$"):
            continue
        pairs = src_pairs(p)
        if not pairs:
            continue
        conf = [(k, pairs[k], prod[k]) for k in pairs if k in prod and prod[k] != pairs[k]]
        only_prod = [k for k in pairs if k not in prod]
        tot_new += len(only_prod)
        tot_conf += len(conf)
        print(f"{b[:26]:28s} 报告源 SKU {len(pairs):4d} | 商品资料缺 {len(only_prod):3d} | ASIN 冲突 {len(conf):3d}"
              + (f" | 冲突例 {conf[:2]}" if conf else ""))
    print(f"\n合计：报告有而商品资料无 {tot_new}，两边 ASIN 不一致 {tot_conf}")
