# -*- coding: utf-8 -*-
"""修 2026-10-30 YXU1 货件错簿（FBA19VFTFF83）

背景：FBA19VFTFF83 配送地址 YXU1（加拿大伦敦仓），但 shipments._CA_PREFIX 不含 YXU，
      na_sheet_of() 兜底成「美国-每日库存分析表」，545 件/72 个 SKU 写进美国表 I 列；
      美国 H(在途合计) 把 I 列算了进去、加拿大 H 漏了这 545 件（AC 在途差额跟着错）。

本脚本（单次修复，幂等）：
  1. 读美国表 I 列（行1 货件号、行5 标签、行7+ 各 SKU 数量）
  2. 清空美国表 I 列
  3. 按 SKU 写入加拿大表第一个空货件列（应为 I）
  4. 给美国表 J 列补行6 SUBTOTAL（历史扩列漏了公式）

保存走 runner 的保真路径（strip_pivot_records → openpyxl → restore_parts）。
"""
import os
import re
import sys

import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
PIPE = os.path.dirname(os.path.dirname(HERE))          # pipeline/
sys.path.insert(0, PIPE)
sys.path.insert(0, os.path.join(PIPE, "core"))

import runner  # noqa: E402

BOOK = os.path.join(
    os.path.dirname(PIPE),
    "1.每日销售数据表", "北美加拿大库存销售表.xlsx",
)
US, CA = "美国-每日库存分析表", "加拿大-每日库存分析表"
DATA_FIRST_ROW = 7
FIRST_SHIP_COL = 9      # I
LAST_SHIP_COL = 40


def main():
    if not os.path.exists(BOOK):
        sys.exit(f"找不到工作簿：{BOOK}")
    lock = os.path.join(os.path.dirname(BOOK), "~$" + os.path.basename(BOOK)[2:])
    if os.path.exists(lock):
        sys.exit(f"工作簿被占用（{lock}），请先关闭 WPS/Excel")

    stripped = BOOK + ".pivotstrip.xlsx"
    runner.strip_pivot_records(BOOK, stripped)
    wb = openpyxl.load_workbook(stripped)
    ws_us, ws_ca = wb[US], wb[CA]

    # 1. 读美国表 I 列
    c = FIRST_SHIP_COL
    sid = ws_us.cell(1, c).value
    label = ws_us.cell(5, c).value
    if not sid:
        sys.exit("美国表 I 列没有货件号，脚本可能已跑过，退出")
    print(f"美国表 I 列 = {sid} / 标签 {label!r}")
    items = {}
    for r in range(DATA_FIRST_ROW, ws_us.max_row + 1):
        sku = ws_us.cell(r, 1).value
        v = ws_us.cell(r, c).value
        if sku and isinstance(v, (int, float)) and v:
            items[str(sku).strip()] = v
    print(f"  数量行 {len(items)} 个 SKU，合计 {sum(items.values()):.0f} 件")

    # 2. 清空美国表 I 列
    cleared = 0
    for r in range(1, ws_us.max_row + 1):
        cell = ws_us.cell(r, c)
        if cell.value is not None:
            cell.value = None
            cleared += 1
    print(f"  清空美国表 I 列 {cleared} 格")

    # 3. 找加拿大表目标列（第一个行1为空的货件列）
    tgt = None
    for cc in range(FIRST_SHIP_COL, LAST_SHIP_COL + 1):
        if not ws_ca.cell(1, cc).value:
            tgt = cc
            break
    if tgt is None:
        sys.exit("加拿大表没有空货件列")
    print(f"加拿大表目标列 = {openpyxl.utils.get_column_letter(tgt)}（行1 原为空）")
    if ws_ca.cell(5, tgt).value not in (None, ""):
        print(f"  注意：该列行5 已有标签 {ws_ca.cell(5, tgt).value!r}，将被覆盖为 {label!r}")

    ws_ca.cell(1, tgt).value = sid
    ws_ca.cell(5, tgt).value = label

    ca_rows = {}
    for r in range(DATA_FIRST_ROW, ws_ca.max_row + 1):
        sku = ws_ca.cell(r, 1).value
        if sku:
            ca_rows[str(sku).strip()] = r

    hit = miss = 0
    miss_skus = []
    for sku, q in items.items():
        r = ca_rows.get(sku)
        if r is None:
            miss += 1
            miss_skus.append(sku)
            continue
        ws_ca.cell(r, tgt).value = q
        hit += 1
    print(f"  写入加拿大表 {hit} 个 SKU；未命中 {miss} 个 {miss_skus}")

    # 加拿大目标列行6 合计
    col = openpyxl.utils.get_column_letter(tgt)
    if ws_ca.cell(6, tgt).value in (None, ""):
        ws_ca.cell(6, tgt).value = f"=SUBTOTAL(9,{col}7:{col}2838)"
        print(f"  补 {col}6 合计公式")

    # 4. 美国表 J 列补合计（历史扩列漏公式）
    fixed = []
    for cc in range(FIRST_SHIP_COL, LAST_SHIP_COL + 1):
        if ws_us.cell(1, cc).value and ws_us.cell(6, cc).value in (None, ""):
            cl = openpyxl.utils.get_column_letter(cc)
            ws_us.cell(6, cc).value = f"=SUBTOTAL(9,{cl}7:{cl}2838)"
            fixed.append(f"{cl}6")
    if fixed:
        print(f"  美国表补合计公式：{fixed}")

    tmp = BOOK + ".fixtmp.xlsx"
    wb.save(tmp)
    wb.close()
    runner.restore_parts(BOOK, tmp)
    os.replace(tmp, BOOK)
    try:
        os.remove(stripped)
    except OSError:
        pass
    print("已保存：", BOOK)


if __name__ == "__main__":
    main()
