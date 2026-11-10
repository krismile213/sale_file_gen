# -*- coding: utf-8 -*-
"""把美客多「空运装箱单」写进「墨西哥-每日库存分析表」的货件列。

数据源：桌面装箱单 xlsx（r4 有「货件编号：」，r1 有「发货日期：」，r7+ 明细 C=SKU / H=数量）
目标：L / M 两列（K 之后、R 之前的空位；H 列公式已覆盖 I..R）

默认 dry-run，--apply 才写；写前自动备份 .bak_mlship_<时间>。
"""
import argparse
import os
import re
import shutil
from collections import Counter
from copy import copy
from datetime import datetime

from openpyxl import load_workbook
from openpyxl.formula.translate import Translator
from openpyxl.utils import get_column_letter as gl

WB = r"D:\Desktop\work-main\sale_file_gen\Daily_Sales_Report\1.每日销售数据表\美客多库存销售表.xlsx"
SHEET = "墨西哥-每日库存分析表"
MASTER = r"D:\Desktop\work-main\sale_file_gen\商品资料.xlsx"

PACKINGS = [
    r"C:\Users\dell\Desktop\5-2美客多墨西哥空运装箱单20261014.xlsx",
    r"C:\Users\dell\Desktop\2美客多墨西哥空运装箱单20261021.xlsx",
]
TARGET_COLS = ["L", "M"]


def parse_packing(path):
    wb = load_workbook(path, data_only=True)
    ws = wb["墨西哥"] if "墨西哥" in wb.sheetnames else wb.worksheets[0]
    ship = date = None
    for r in range(1, 8):
        for c in range(1, ws.max_column + 1):
            v = ws.cell(r, c).value
            if not isinstance(v, str):
                continue
            m = re.search(r"货件编号[：:]\s*([0-9A-Za-z]+)", v)
            if m:
                ship = m.group(1)
            m = re.search(r"发货日期[：:]\s*(\d{4})/(\d{1,2})/(\d{1,2})", v)
            if m:
                date = (int(m.group(1)), int(m.group(2)), int(m.group(3)))
    qty = Counter()
    for r in range(7, ws.max_row + 1):
        sku = ws.cell(r, 3).value
        q = ws.cell(r, 8).value
        if sku and isinstance(q, (int, float)):
            qty[str(sku).strip()] += int(q)
    wb.close()
    return dict(path=path, ship=ship, date=date, qty=qty)


def master_info(skus):
    """SKU → (B列型号名, C列评级)。型号名取英文品名 For 之后那段。"""
    wb = load_workbook(MASTER, read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]
    out = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        sku = row[0]
        if not sku or str(sku).strip() not in skus:
            continue
        en = str(row[2] or "")
        m = re.search(r"\bFor\s+(.+?)(?:-\d+\s*Pcs|$)", en)
        name = m.group(1).strip() if m else en
        out[str(sku).strip()] = (name, row[5] or "")
    wb.close()
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--apply", action="store_true")
    args = ap.parse_args()

    packs = [parse_packing(p) for p in PACKINGS]
    packs.sort(key=lambda x: x["date"])
    for p in packs:
        print(f"装箱单 {os.path.basename(p['path'])}: 货件 {p['ship']} 发货 "
              f"{p['date'][0]}/{p['date'][1]}/{p['date'][2]}  {len(p['qty'])} SKU / "
              f"{sum(p['qty'].values())} 件")

    wb = load_workbook(WB)
    ws = wb[SHEET]

    # 最后一次数据行
    last = max(r for r in range(7, ws.max_row + 1) if ws.cell(r, 1).value)
    row_of = {str(ws.cell(r, 1).value).strip(): r
              for r in range(7, last + 1) if ws.cell(r, 1).value}

    # 缺失 SKU（新品还没进分析表）
    all_sku = set().union(*[set(p["qty"]) for p in packs])
    missing = sorted(s for s in all_sku if s not in row_of)
    info = master_info(set(missing)) if missing else {}

    print(f"\n分析表现有数据行 7..{last}（{len(row_of)} 个 SKU）")
    if missing:
        print(f"需补行的新品 {len(missing)} 个：")
        for s in missing:
            nm, sc = info.get(s, ("?", "?"))
            print(f"    + {s}  {nm}  评级 {sc}  件 {sum(p['qty'].get(s,0) for p in packs)}")

    if not args.apply:
        for col, p in zip(TARGET_COLS, packs):
            print(f"\n[{col}] {p['ship']} 标签 "
                  f"{p['date'][1]}/{p['date'][2]}空运 → 写 {len(p['qty'])} SKU / "
                  f"{sum(p['qty'].values())} 件")
        print("\n[dry] 未写入。加 --apply 执行。")
        return

    bak = f"{WB}.bak_mlship_{datetime.now():%m%d_%H%M%S}"
    shutil.copy2(WB, bak)
    print(f"\n[备份] {os.path.basename(bak)}")

    # ---- 1) 补行（复制最后一行公式/样式，整体下移） ----
    added = {}
    cur_last = last
    for s in missing:
        cur_last += 1
        for c in range(1, ws.max_column + 1):
            src = ws.cell(last, c)
            dst = ws.cell(cur_last, c)
            v = src.value
            if isinstance(v, str) and v.startswith("="):
                dst.value = Translator(v, origin=f"{gl(c)}{last}").translate_formula(f"{gl(c)}{cur_last}")
            else:
                dst.value = v
            dst.font = copy(src.font)
            dst.fill = copy(src.fill)
            dst.border = copy(src.border)
            dst.alignment = copy(src.alignment)
            dst.number_format = src.number_format
        nm, sc = info.get(s, (s, ""))
        ws.cell(cur_last, 1).value = s
        ws.cell(cur_last, 2).value = nm
        ws.cell(cur_last, 3).value = sc
        row_of[s] = cur_last
        added[s] = cur_last
    if added:
        print(f"[补行] {cur_last - last} 行："
              + ", ".join(f"{k}@{v}" for k, v in added.items()))

    # ---- 2) 清空 L/M 1..cur_last（保留样式） ----
    for col in TARGET_COLS:
        for r in range(1, cur_last + 1):
            ws[f"{col}{r}"] = None

    # ---- 3) 写列头 + 数量 ----
    for col, p in zip(TARGET_COLS, packs):
        tag = f"{p['date'][1]}/{p['date'][2]}空运"
        c1 = ws[f"{col}1"]
        c1.value = p["ship"]
        c1.number_format = "@"
        ws[f"{col}5"] = tag
        for r in (2, 3, 4):
            src = ws[f"I{r}"].value
            if isinstance(src, str):
                ws[f"{col}{r}"] = re.sub(r"\bI(\d+)\b", lambda m: f"{col}{m.group(1)}", src)
        ws[f"{col}6"] = f"=SUBTOTAL(9,{col}7:{col}1515)"
        for r in range(1, cur_last + 1):
            src, dst = ws[f"I{r}"], ws[f"{col}{r}"]
            dst.font = copy(src.font)
            dst.fill = copy(src.fill)
            dst.border = copy(src.border)
            dst.alignment = copy(src.alignment)
            if r <= 6:
                dst.number_format = src.number_format
        written = 0
        for sku, q in p["qty"].items():
            r = row_of.get(sku)
            if not r:
                continue
            ws[f"{col}{r}"] = q
            written += q
        print(f"[{col}] {p['ship']} 标签 {tag} → {len(p['qty'])} SKU / {written} 件")
        ws.column_dimensions[col].width = 13

    wb.calculation.fullCalcOnLoad = True
    wb.save(WB)
    print("[完成] 已保存")


if __name__ == "__main__":
    main()
