# -*- coding: utf-8 -*-
"""新建 Interflowd英国库存销售表.xlsx（2026-11-06，用户拍板）。

背景（2026-11-06 用户反馈）：
  FBA15MRX4YLH 是 interflow 发往英国仓 EMA5 的货件（185 件 / 3 SKU），
  以前 EMA5 不在 UK_WAREHOUSES 精确表里 → interflow 分支兜底成「美国」，
  185 件写进 Interflowd北美「美国-每日库存分析表」L 列 —— 美国库存报告里
  这些 SKU 的在途为 0，这 185 件永远不会"到货"，在途永久挂账
  （和 2026-10-20 沙特 RUH8 的 FBA15MHHH99G 完全同型）。

  实证：G599/G600/G601 在「欧盟英国簿-英国表」里根本没有行（467 个 SKU 全不命中），
  而 interflow 美国/澳洲簿里都有 → 它们属于 interflow 账号，不能塞主账号英国表。

做法：以 Interflowd阿联酋簿为模板派生（结构/公式/CF 完全一致）：
  1) 全表文本 阿联酋 → 英国（单元格 / 公式 / 数组公式 / 数据有效性 / 条件格式）
  2) 五个 sheet 改名
  3) 分析表只保留 EMA5 这票的 SKU（G599/G600/G601），属性带品名/ASIN
  4) 库存 sheet：可售（结余式）预填同 3 个 SKU=0；在途（inv 式）清空表头以下
     —— inv 式跑批会自动追加新 SKU，所以空的没关系
  5) 订单 sheet 清空数据，保留 35/36 列公式（append_sales 的模板）
  6) 货件列清掉模板簿残留的阿联酋货件号

⚠️ 报告还没下载：config 里这两份报告标 optional=True（缺了只告警、不算体检硬伤），
   等 interflow 英国站有库存/订单导出后放进快照目录即可自动接上。

用法：
    python pipeline/tools/build_interflow_uk.py --dry-run
    python pipeline/tools/build_interflow_uk.py
"""
import os
import re
import sys
import csv
import shutil
import argparse

HERE = os.path.dirname(os.path.abspath(__file__))
PIPE = os.path.dirname(HERE)
ROOT = os.path.dirname(PIPE)
sys.path.insert(0, PIPE)
sys.path.insert(0, os.path.join(PIPE, "core"))

from openpyxl import load_workbook
from openpyxl.worksheet.formula import ArrayFormula
import analysis as An
import shipments as Sh

DSR = ROOT
PROJ = os.path.dirname(DSR)
BOOKS = os.path.join(DSR, "1.每日销售数据表")
SRC = os.path.join(BOOKS, "Interflowd阿联酋库存销售表.xlsx")
DST = os.path.join(BOOKS, "Interflowd英国库存销售表.xlsx")

# 货件 TSV（唯一已知的英国仓货件）：FBA15MRX4YLH / EMA5 / 27-10-2026
SHIP_TSV = os.path.join(PROJ, "trunk", "data", "数据归属周期", "202611", "货件",
                        "interflow", "FBA15MRX4YLH.tsv")

OLD, NEW = "阿联酋", "英国"
RENAMES = [
    (f"数据透视图{OLD}", f"数据透视图{NEW}"),
    (f"每日销售数据{OLD}", f"每日销售数据{NEW}"),
    (f"{OLD}可售库存", f"{NEW}可售库存"),
    (f"{OLD}在途库存", f"{NEW}在途库存"),
    (f"{OLD}-每日库存分析表", f"{NEW}-每日库存分析表"),
]
AN = f"{NEW}-每日库存分析表"
SL = f"每日销售数据{NEW}"
BL = f"{NEW}可售库存"
ST = f"{NEW}在途库存"

# 属性列（不是历史数据，按 SKU 带回）
ATTR_COLS = {64: "膜类型", 78: "品牌", 80: "初次销售日期", 86: "ASIN"}


def relocate_formula(v, row):
    """把属性列里的公式本地化到新行（XLOOKUP 的 A 行引用改成当前行）。"""
    if not isinstance(v, str) or not v.startswith("="):
        return v
    return re.sub(r"(XLOOKUP\(\s*\$?A)\$?\d+", r"\g<1>%d" % row, v, count=1)


def read_ship_skus(path):
    """货件 TSV → [{sku, name, asin, qty}]（保持票内顺序）。"""
    rows = []
    for enc in ("utf-8-sig", "gb18030", "latin-1"):
        try:
            with open(path, encoding=enc, newline="") as fh:
                rows = list(csv.reader(fh, delimiter="\t"))
            break
        except Exception:
            continue
    out, seen, in_detail = [], set(), False
    for r in rows:
        if not r:
            continue
        if r[0] == "卖家 SKU":
            in_detail = True
            continue
        if not in_detail:
            continue
        sku = r[0].strip()
        if not An._SKU_RE.match(sku) or sku in seen:
            continue
        seen.add(sku)
        try:
            q = float(r[9])
        except (IndexError, ValueError):
            q = 0
        out.append(dict(sku=sku, name=(r[1].strip() if len(r) > 1 else ""),
                        asin=(r[2].strip() if len(r) > 2 else ""), qty=q))
    return out


def replace_text(wb):
    """全表文本 阿联酋 → 英国（单元格/公式/数组公式 + 数据有效性 + 条件格式）。"""
    n = 0
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if isinstance(v, ArrayFormula):
                    if OLD in v.text:
                        c.value = ArrayFormula(ref=c.coordinate,
                                               text=v.text.replace(OLD, NEW))
                        n += 1
                elif isinstance(v, str) and OLD in v:
                    c.value = v.replace(OLD, NEW)
                    n += 1
        try:
            for dv in ws.data_validations.dataValidation:
                for attr in ("formula1", "formula2"):
                    f = getattr(dv, attr, None)
                    if isinstance(f, str) and OLD in f:
                        setattr(dv, attr, f.replace(OLD, NEW))
                        n += 1
        except Exception as e:
            print(f"  !! {ws.title} 数据有效性替换失败：{e}")
        try:
            for rng in ws.conditional_formatting:
                for rule in rng.rules:
                    f = getattr(rule, "formula", None)
                    if not f:
                        continue
                    newf = [str(x).replace(OLD, NEW) for x in f]
                    if newf != [str(x) for x in f]:
                        rule.formula = newf
                        n += 1
        except Exception as e:
            print(f"  !! {ws.title} 条件格式替换失败：{e}")
    return n


def clear_values(ws, from_row, cols=None):
    n = 0
    maxc = ws.max_column
    for r in range(from_row, (ws.max_row or from_row) + 1):
        for c in (cols if cols else range(1, maxc + 1)):
            if ws.cell(r, c).value is not None:
                ws.cell(r, c).value = None
                n += 1
    return n


def rebuild_analysis(ws, recs):
    """分析表重建为给定 SKU 集合，保留模板行的公式与样式。"""
    last = An._last_data_row(ws)
    maxc = min(ws.max_column, 300)
    tpl = An.pick_template_row(ws, last)
    fc, lc = Sh.detect_sheet_cols(ws)
    log = [f"  原数据区 r{An.DATA_FIRST_ROW}~r{last}，模板行 r{tpl}，货件列 {fc}~{lc}"]

    if tpl != An.DATA_FIRST_ROW:
        An.clone_row(ws, tpl, An.DATA_FIRST_ROW, clear_range=(fc, lc), max_col=maxc)
        log.append(f"  模板 r{tpl} → r{An.DATA_FIRST_ROW}")
    for r in range(An.DATA_FIRST_ROW + 1, last + 1):
        for c in range(1, maxc + 1):
            ws.cell(r, c).value = None

    def _fill(r, rec):
        ws.cell(r, 1).value = rec["sku"]
        ws.cell(r, 2).value = rec["name"] or None
        ws.cell(r, 4).value = An.STATE_IN_TRANSIT   # 还没到仓：在途
        # ASIN 列是公式（XLOOKUP 到在途库存），先按报告里的 ASIN 落静态值兜底，
        # 等库存报告到位后跑批会按公式刷新
        ws.cell(r, 86).value = rec["asin"] or None

    _fill(An.DATA_FIRST_ROW, recs[0])
    rest = recs[1:]
    names = {r["sku"]: r["name"] for r in rest}
    added = An.append_sku_rows_auto(ws, [r["sku"] for r in rest], names,
                                    clear_range=(fc, lc),
                                    state_map={r["sku"]: An.STATE_IN_TRANSIT for r in rest})
    by = {r["sku"]: r for r in recs}
    for r, s in added:
        ws.cell(r, 2).value = names.get(s) or None
        ws.cell(r, 86).value = by[s]["asin"] or None
    log.append(f"  重建 {1 + len(added)} 行（r{An.DATA_FIRST_ROW}~r{An.DATA_FIRST_ROW + len(added)}）")
    return log, fc, lc


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--dry-run", action="store_true")
    args = ap.parse_args()

    if not os.path.exists(SRC):
        print(f"✗ 模板不存在：{SRC}")
        return 1
    if not os.path.exists(SHIP_TSV):
        print(f"✗ 货件 TSV 不存在：{SHIP_TSV}")
        return 1

    recs = read_ship_skus(SHIP_TSV)
    print(f"货件 {os.path.basename(SHIP_TSV)}: {len(recs)} 个 SKU / "
          f"{int(sum(r['qty'] for r in recs))} 件 → {[r['sku'] for r in recs]}")
    if not recs:
        print("✗ 没读到 SKU，中止")
        return 1

    if not args.dry_run:
        if os.path.exists(DST):
            bak = DST + ".bak"
            shutil.copy2(DST, bak)
            print(f"  已有同名文件，先备份到 {os.path.basename(bak)}")
        shutil.copy2(SRC, DST)
        print(f"  复制模板 → {os.path.basename(DST)}")

    wb = load_workbook(SRC if args.dry_run else DST)
    print(f"  sheets: {wb.sheetnames}")

    n = replace_text(wb)
    print(f"  文本 {OLD}→{NEW}：{n} 格")

    for old, new in RENAMES:
        if old in wb.sheetnames:
            wb[old].title = new
    print(f"  改名后: {wb.sheetnames}")

    ws = wb[AN]
    for line in rebuild_analysis(ws, recs)[0]:
        print(line)
    fc, lc = Sh.detect_sheet_cols(ws)

    ws.cell(5, 3).value = NEW
    print(f"  C5={NEW}  D5: {ws.cell(5, 4).value!r} → 'Amazon.co.uk'")
    ws.cell(5, 4).value = "Amazon.co.uk"

    # 库存 sheet：可售（结余式）预填 3 行；在途（inv 式）清空表头以下
    if BL in wb.sheetnames:
        wsb = wb[BL]
        clear_values(wsb, 3)
        hdr = {str(wsb.cell(2, c).value).strip().lower(): c
               for c in range(1, wsb.max_column + 1) if wsb.cell(2, c).value}
        for i, rec in enumerate(recs):
            r = 3 + i
            if hdr.get("asin"):
                wsb.cell(r, hdr["asin"]).value = rec["asin"] or None
            if hdr.get("msku") or hdr.get("sku"):
                wsb.cell(r, hdr.get("msku") or hdr.get("sku")).value = rec["sku"]
            if hdr.get("disposition"):
                wsb.cell(r, hdr["disposition"]).value = "SELLABLE"
            wsb.cell(r, hdr.get("ending warehouse balance") or 5).value = 0
        print(f"  {BL}: 预填 {len(recs)} 行（可售 0，等报告）")
    if ST in wb.sheetnames:
        k = clear_values(wb[ST], 3)
        print(f"  {ST}: 清空 {k} 格（inv 式，跑批自动追加 SKU）")

    # 货件列：清掉模板残留的阿联酋货件号 **和静态标签**
    # ⚠️ 2026-11-06 踩坑：只清第一列的 行2~5，别的列留下的静态行5 标签
    # （'10/28空运'）会让 plan_columns 认为该列已被占用（行5 非空且不是占位公式），
    # 结果新票排不进这个空列位、表头还多出一个没有货件号的假标签。
    n3 = 0
    for c in range(fc, lc + 1):
        for r in (1, 2, 3, 4, 5):
            v = ws.cell(r, c).value
            if v is None or (isinstance(v, str) and v.strip() == ""):
                continue
            if r > 1 and isinstance(v, str) and v.startswith("="):
                continue                      # 占位公式保留（会自动显示 M/D）
            ws.cell(r, c).value = None
            n3 += 1
    print(f"  货件列 {fc}~{lc}: 清表头 {n3} 格")

    # 订单 sheet
    ws2 = wb[SL]
    keep = {35, 36}
    n1 = 0
    for c in range(1, ws2.max_column + 1):
        if c in keep:
            continue
        if ws2.cell(3, c).value is not None:
            ws2.cell(3, c).value = None
            n1 += 1
    n2 = clear_values(ws2, 4)
    print(f"  {SL}: r3 清 {n1} 格（保留 35/36 公式），r4+ 清 {n2} 格")

    if args.dry_run:
        print("\n[dry-run] 不落盘")
        wb.close()
        return 0

    wb.calculation.fullCalcOnLoad = True
    wb.save(DST)
    wb.close()
    print(f"\n✓ 已生成 {DST}")
    print(f"  大小 {os.path.getsize(DST) / 1024:.0f} KB")
    return 0


if __name__ == "__main__":
    sys.exit(main())
