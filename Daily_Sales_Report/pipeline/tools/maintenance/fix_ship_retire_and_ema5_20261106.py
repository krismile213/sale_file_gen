# -*- coding: utf-8 -*-
"""2026-11-06 货件台账纠错（两件事，幂等，--apply 落盘）：

① 登记「已到货退场」：用户今天在 13 本簿里手工删掉了 56 个货件列，但货件号没写进
   `202611/货件/到货.txt` —— 跑批的「到货即退场」只认这个登记表，不登记就会把删掉的
   列原样建回来（等于撤销手工清理）。
   本轮只登记**行5 标签里明确写了「到货」**的票（43 条），其余 13 条（行5 没标到货，
   可能仍在途）留给跑批重建 —— 用户 2026-11-06 拍板的口径。
   证据来源 = `1.每日销售数据表/11月04日/` 归档快照（删列前的状态）vs 当前簿。

② 清掉误落的 FBA15MRX4YLH：这票是 interflow 英国仓 EMA5（27/10 发货、185 件/3 SKU），
   以前被 interflow 兜底分支算成"美国"，写进 `Interflowd北美加拿大库存销售表 /
   美国-每日库存分析表` L 列。识别已修（shipments.int_target_of 加英国分支），
   这列必须清掉，否则 185 件美国在途永久挂账。

用法：
    python pipeline/tools/maintenance/fix_ship_retire_and_ema5_20261106.py            # 预演
    python pipeline/tools/maintenance/fix_ship_retire_and_ema5_20261106.py --apply
"""
import os
import sys
import glob
import shutil
import argparse
import datetime

HERE = os.path.dirname(os.path.abspath(__file__))
PIPE = os.path.dirname(os.path.dirname(HERE))
ROOT = os.path.dirname(PIPE)
sys.path.insert(0, PIPE)
sys.path.insert(0, os.path.join(PIPE, "core"))

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import shipments as Sh

DSR = ROOT
PROJ = os.path.dirname(DSR)
BOOKS = os.path.join(DSR, "1.每日销售数据表")
ARCHIVE = os.path.join(BOOKS, "11月04日")
ARRIVE_TXT = os.path.join(PROJ, "trunk", "data", "数据归属周期", "202611", "货件", "到货.txt")
BAK = os.path.join(DSR, ".workbuddy", "backup_fix_ship_20261106")

WRONG_ID = "FBA15MRX4YLH"
WRONG_BOOK = "Interflowd北美加拿大库存销售表.xlsx"
WRONG_SHEET = "美国-每日库存分析表"


def ship_cols(path):
    """{sheet: {货件号: (列号, 行5标签)}} —— 只读行1/行5，快。"""
    out = {}
    wb = load_workbook(path, read_only=True, data_only=True)
    for ws in wb.worksheets:
        if "分析表" not in ws.title:
            continue
        rows = list(ws.iter_rows(min_row=1, max_row=5, values_only=True))
        r1, r5 = rows[0], rows[4]
        m = {}
        for c in range(9, 40):
            v5 = r5[c - 1] if c - 1 < len(r5) else None
            if isinstance(v5, str) and "在途差额" in v5:
                break
            v1 = r1[c - 1] if c - 1 < len(r1) else None
            if v1 not in (None, ""):
                m[str(v1).strip()] = (c, v5)
        out[ws.title.strip()] = m
    wb.close()
    return out


def plan_retire():
    """返回 [dict(book, sheet, id, label)] —— 归档有、当前无、行5 含「到货」。"""
    cur = {os.path.basename(p): ship_cols(p) for p in glob.glob(os.path.join(BOOKS, "*.xlsx"))}
    old = {os.path.basename(p): ship_cols(p) for p in glob.glob(os.path.join(ARCHIVE, "*.xlsx"))}
    out = []
    for bk, sheets in old.items():
        if bk not in cur:                      # 簿已不存在（如新建簿）→ 不判
            continue
        for sn, m in sheets.items():
            have = cur[bk].get(sn, {})
            for sid, (col, lab) in m.items():
                if sid in have:
                    continue
                if isinstance(lab, str) and "到货" in lab:
                    out.append(dict(book=bk, sheet=sn, id=sid, label=lab, col=col))
    return out


def read_arrive():
    if not os.path.exists(ARRIVE_TXT):
        return set()
    with open(ARRIVE_TXT, encoding="utf-8-sig") as fh:
        return {l.strip() for l in fh if l.strip() and not l.strip().startswith("#")}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--apply", action="store_true")
    args = ap.parse_args()

    todo = plan_retire()
    have = read_arrive()
    add = [r for r in todo if r["id"] not in have]
    print(f"① 待登记退场：{len(todo)} 条，其中新登记 {len(add)} 条")
    for r in add:
        print(f"   + {r['id']:16s} {r['book'][:18]:18s} {r['sheet'][:12]:12s} 行5={r['label']!r}")

    # ② 误落列
    wpath = os.path.join(BOOKS, WRONG_BOOK)
    hit = None
    if os.path.exists(wpath):
        sc = ship_cols(wpath)
        for sn, m in sc.items():
            if WRONG_ID in m:
                hit = (sn, m[WRONG_ID][0], m[WRONG_ID][1])
    print(f"② 误落的 {WRONG_ID}: {('在 ' + str(hit)) if hit else '未找到（已清过）'}")

    if not args.apply:
        print("\n[预演] 未落盘（加 --apply 执行）")
        return 0

    os.makedirs(BAK, exist_ok=True)
    if add:
        shutil.copy2(ARRIVE_TXT, os.path.join(
            BAK, f"到货.txt.{datetime.datetime.now():%Y%m%d_%H%M%S}"))
        with open(ARRIVE_TXT, "a", encoding="utf-8") as fh:
            fh.write("\n# 2026-11-06 登记：用户手工删列（行5 已标「到货」）的 43 票，"
                     "避免跑批重建\n")
            for r in add:
                fh.write(r["id"] + "\n")
        print(f"  ✓ 已追加 {len(add)} 行到 {os.path.basename(ARRIVE_TXT)}")

    if hit:
        sn, col, lab = hit
        shutil.copy2(wpath, os.path.join(BAK, WRONG_BOOK))
        wb = load_workbook(wpath)
        ws = wb[sn]
        n = Sh.clear_shipment_col(ws, col)
        wb.calculation.fullCalcOnLoad = True
        wb.save(wpath)
        wb.close()
        print(f"  ✓ {WRONG_BOOK} / {sn}: 清空 {get_column_letter(col)} 列"
              f"（{lab!r}）共 {n} 格")
    return 0


if __name__ == "__main__":
    sys.exit(main())
