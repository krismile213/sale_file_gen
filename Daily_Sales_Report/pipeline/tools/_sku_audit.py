"""SKU 识别体检：13 本簿 × 每张分析表的库存源，看「源里有值但被正则挡掉」的候选。

输出三块：
  ① 每张表：源 sheet、识别到多少 SKU
  ② 被正则挡掉的「像 SKU」的值（含数字+大写字母、长度 4~12）→ 这是真实盲区
  ③ 三态分布

用法：
    python pipeline/tools/_sku_audit.py            # 全部
    python pipeline/tools/_sku_audit.py 阿联酋 日本  # 指定簿（按文件名模糊匹配）
"""
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(HERE))
sys.path.insert(0, os.path.join(os.path.dirname(HERE), "core"))

import config as C                      # noqa: E402
import analysis as An                   # noqa: E402
import runner as R                      # noqa: E402
from openpyxl import load_workbook      # noqa: E402

# 「像 SKU」的宽松判据：至少含 2 位数字、至少 1 个大写字母、总长 4~14
_LIKE = re.compile(r"^[A-Z0-9][A-Z0-9\-_ ]{3,13}$")


def looks_like_sku(s):
    if not _LIKE.match(s):
        return False
    if sum(c.isdigit() for c in s) < 2:
        return False
    if not any(c.isalpha() for c in s):
        return False
    return True


def audit_book(fname, filt):
    if filt and not any(k in fname for k in filt):
        return
    p = os.path.join(C.WORK_DIR, fname)
    if not os.path.exists(p):
        print(f"\n==== {fname}  ❌ 文件不存在")
        return

    # 欧盟簿必须走 strip + 普通模式，否则 openpyxl 直接 OOM
    tmp = None
    if os.path.getsize(p) > 8 * 1024 * 1024:
        tmp = os.path.join(C.DSR, ".workbuddy", "_audit_strip.xlsx")
        R.strip_pivot_records(p, tmp)
        load_p = tmp
    else:
        load_p = p

    wb = load_workbook(load_p, data_only=False)
    print(f"\n==== {fname}")
    for sn in wb.sheetnames:
        if "分析表" not in sn:
            continue
        ws = wb[sn]
        srcs = An.source_sheets_of(ws, wb)
        if not srcs:
            print(f"  {sn}: ❌ 找不到库存源")
            continue
        stock = An.stock_snapshot_merged(wb, srcs)

        # 收集源里所有原始值，挑出被正则挡掉的
        rejected = {}
        for src in srcs:
            if src not in wb.sheetnames:
                continue
            w = wb[src]
            hdr, hr, sku_c, sub = An._pick_header_row(w)
            if hdr is None:
                continue
            start = (sub[1] if sub else hr) + 1
            for r in range(start, min(w.max_row, 5000) + 1):
                v = w.cell(r, sku_c).value
                if v is None:
                    continue
                s = str(v).strip()
                if not s or An._SKU_RE.match(s):
                    continue
                if looks_like_sku(s):
                    rejected[s] = rejected.get(s, 0) + 1

        cnt = {}
        for rec in stock.values():
            st = An.tri_state(rec)
            cnt[st] = cnt.get(st, 0) + 1
        n_av = sum(1 for v in stock.values() if (v.get("avail") or 0) > 0)
        n_ib = sum(1 for v in stock.values() if (v.get("inb") or 0) > 0)

        print(f"  {sn}")
        print(f"    源: {' + '.join(srcs)}")
        print(f"    识别 {len(stock)} SKU ｜ 有可售 {n_av} ｜ 有在途 {n_ib} ｜ 三态 {cnt}")
        if rejected:
            lst = sorted(rejected.items(), key=lambda x: -x[1])[:12]
            print(f"    ⚠ 被正则挡掉的像-SKU 值 {len(rejected)} 个: "
                  + "、".join(f"{k}×{v}" for k, v in lst))
    wb.close()
    if tmp and os.path.exists(tmp):
        os.remove(tmp)


def main():
    filt = sys.argv[1:]
    for b in C.WORKBOOKS:
        if b.get("skip"):
            continue
        try:
            audit_book(b["file"], filt)
        except Exception as e:
            print(f"\n==== {b['file']}  !! 失败 {e!r}")


if __name__ == "__main__":
    main()
