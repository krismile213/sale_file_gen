# -*- coding: utf-8 -*-
"""新建 Interflowd沙特库存销售表.xlsx（2026-10-24 方案 A）。

背景（查证见 memory/2026-10-24.md §2）：
  interflow 账号有两个物理仓，但系统只读了阿联酋仓：
    · 阿联酋 FC  → interflow阿拉伯联合酋长国管理库存报告  → Interflowd阿联酋簿 ✅
    · 沙特 RUH8  → interflow沙特阿拉伯管理库存报告        → 无人接收 ❌（116 件可售全丢）
  订单报告两份哈希相同（同一份下载存两名），含 Amazon.ae 58 件 + Amazon.sa 5 件；
  Interflowd阿联酋 的 expect_channel=['Amazon.ae'] 把 5 件 sa 挡掉了。

做法：以 Interflowd阿联酋 为模板（结构/公式/CF 完全一致，且无真实透视表部件，
      可安全用 openpyxl 改写），派生一本沙特簿：
  1) 全表文本 阿联酋 → 沙特（含公式里的 sheet 引用、C5 站点标签、透视图锚点）
  2) 五个 sheet 改名
  3) 分析表重建为「沙特仓 29 个 SKU」——原来的 51 行含 22 个只有阿联酋仓库存的 SKU，
     留着会在沙特簿里假报断货，所以只保留沙特仓报告里的 SKU，属性列（分类/膜类型/
     品牌/ASIN/初次销售日期）按 SKU 从原行带回
  4) 库存 sheet 清空数据区（值），保留表头与样式 → 跑批时按沙特报告重新追加 29 行
  5) 订单 sheet 清空（含行3 订单，但**保留 35/36 列公式**给 append_sales 当模板）

用法：
    python pipeline/tools/build_interflow_sa.py            # 建到 1.每日销售数据表/
    python pipeline/tools/build_interflow_sa.py --dry-run  # 只报告，不落盘
"""
import os
import re
import sys
import csv
import io
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

DSR = ROOT                                            # Daily_Sales_Report
PROJ = os.path.dirname(DSR)                           # sale_file_gen
BOOKS = os.path.join(DSR, "1.每日销售数据表")
SRC = os.path.join(BOOKS, "Interflowd阿联酋库存销售表.xlsx")
DST = os.path.join(BOOKS, "Interflowd沙特库存销售表.xlsx")

SNAP = os.path.join(PROJ, "trunk", "data", "数据归属周期", "202610", "快照日期", "20261021")
SA_REPORT = os.path.join(SNAP, "interflow沙特阿拉伯管理库存报告20261021.csv")
SA_ROWS = []          # main() 里由 read_sa_rows() 填充，供「沙特可售库存」预填用

RENAMES = [
    ("数据透视图阿联酋", "数据透视图沙特"),
    ("每日销售数据阿联酋", "每日销售数据沙特"),
    ("阿联酋可售库存", "沙特可售库存"),
    ("阿联酋在途库存", "沙特在途库存"),
    ("阿联酋-每日库存分析表", "沙特-每日库存分析表"),
]

AN_OLD = "阿联酋-每日库存分析表"
AN_NEW = "沙特-每日库存分析表"
SL_OLD = "每日销售数据阿联酋"
SL_NEW = "每日销售数据沙特"
ST_OLD = "阿联酋在途库存"
ST_NEW = "沙特在途库存"
BL_OLD = "阿联酋可售库存"
BL_NEW = "沙特可售库存"

# 属性列（不是历史数据，按 SKU 带回）
ATTR_COLS = {
    64: "膜类型",
    78: "品牌",
    80: "初次销售日期",
    86: "ASIN",      # ⚠️ 这一列是**公式**，带回时必须重写行引用，见 relocate_formula()
}


def relocate_formula(v, row):
    """把属性列里的公式「本地化」到新行。

    ⚠️ 2026-10-30 事故：ATTR_COLS 里的 86(ASIN) 是 `=XLOOKUP(A42,在途!A:A,在途!C:C)`。
    派生时原样抄过来，行引用还是**源簿的行号**，于是新簿里
      · 引用到空行 → 查出 0/空（表现为「ASIN 大量为空」）
      · 引用到别人 → 显示成**别的 SKU 的 ASIN**（G583 显示成 G587 的 B0HHMH12QL）
    这类错误不报错、数字也不异常，只靠肉眼根本发现不了，所以在这里一次性挡住。
    """
    if not isinstance(v, str) or not v.startswith("="):
        return v
    return re.sub(r"(XLOOKUP\(\s*\$?A)\$?\d+", r"\g<1>%d" % row, v, count=1)


def read_sa_rows():
    """沙特仓报告 → [{sku, fnsku, asin, disposition, fulfillable}, ...]（按报告顺序去重）。

    比 read_sa_skus 多带属性字段，用来预填「沙特可售库存」那张结余式 sheet。
    """
    raw = open(SA_REPORT, "rb").read()
    txt = None
    for enc in ("utf-8-sig", "gbk", "utf-16"):
        try:
            txt = raw.decode(enc)
            break
        except Exception:
            continue
    out, seen = [], set()
    for r in csv.DictReader(io.StringIO(txt)):
        s = (r.get("sku") or "").strip()
        if not s or s in seen or not An._SKU_RE.match(s):
            continue
        seen.add(s)
        try:
            q = float(r.get("afn-fulfillable-quantity") or 0)
        except (TypeError, ValueError):
            q = 0.0
        out.append(dict(sku=s,
                        fnsku=(r.get("fnsku") or "").strip(),
                        asin=(r.get("asin") or "").strip(),
                        disposition=(r.get("disposition") or "SELLABLE").strip(),
                        fulfillable=int(q)))
    return out


def read_sa_skus():
    """沙特仓报告里的 SKU（按报告顺序去重）。"""
    raw = open(SA_REPORT, "rb").read()
    txt = None
    for enc in ("utf-8-sig", "gbk", "utf-16"):
        try:
            txt = raw.decode(enc)
            break
        except Exception:
            continue
    out = []
    for r in csv.DictReader(io.StringIO(txt)):
        s = (r.get("sku") or "").strip()
        if s and s not in out and An._SKU_RE.match(s):
            out.append(s)
    return out


def replace_text(wb):
    """全表文本 阿联酋 → 沙特。返回改动处数。

    三个地方都要改，漏一个就是悬空引用（Excel 打开会报 #REF! 或弹修复框）：
      ① 单元格值 / 公式（含数组公式）—— ws.iter_rows 能覆盖
      ② 数据有效性下拉框的 formula1 —— **不在单元格里**，在 sheet 级
         dataValidations 节点（如 B1 的 `=数据透视图阿联酋!$E$1` 日期选择器）
      ③ 条件格式规则 formula —— 也在 sheet 级 conditionalFormatting 节点
         （如 `I4<数据透视图阿联酋!$E$1`、`SUM(阿联酋在途库存!$W:$W)`）

    2026-10-24 踩坑：第一版只做了 ①，结果 CF 和下拉框里 12 处仍指向
    `数据透视图阿联酋` / `阿联酋在途库存` —— 沙特簿里这两个 sheet 不存在，
    断货标记和日期锚点全部失效，且 Excel 打开会提示修复。
    """
    n = 0
    for ws in wb.worksheets:
        # ① 单元格
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if isinstance(v, ArrayFormula):
                    if "阿联酋" in v.text:
                        c.value = ArrayFormula(ref=c.coordinate,
                                               text=v.text.replace("阿联酋", "沙特"))
                        n += 1
                elif isinstance(v, str) and "阿联酋" in v:
                    c.value = v.replace("阿联酋", "沙特")
                    n += 1

        # ② 数据有效性（下拉框 / 日期选择器）
        try:
            for dv in ws.data_validations.dataValidation:
                for attr in ("formula1", "formula2"):
                    f = getattr(dv, attr, None)
                    if isinstance(f, str) and "阿联酋" in f:
                        setattr(dv, attr, f.replace("阿联酋", "沙特"))
                        n += 1
        except Exception as e:
            print(f"  !! {ws.title} 数据有效性替换失败：{e}")

        # ③ 条件格式规则
        try:
            for rng in ws.conditional_formatting:
                for rule in rng.rules:
                    f = getattr(rule, "formula", None)
                    if not f:
                        continue
                    new = [str(x).replace("阿联酋", "沙特") for x in f]
                    if new != [str(x) for x in f]:
                        rule.formula = new
                        n += 1
        except Exception as e:
            print(f"  !! {ws.title} 条件格式替换失败：{e}")
    return n


def capture_attrs(ws, last):
    """{sku: {列号: 值}}，只抓属性列 + B/C。"""
    out = {}
    for r in range(An.DATA_FIRST_ROW, last + 1):
        sku = ws.cell(r, 1).value
        if not sku:
            continue
        sku = str(sku).strip()
        d = {"B": ws.cell(r, 2).value, "C": ws.cell(r, 3).value}
        for col in ATTR_COLS:
            d[col] = ws.cell(r, col).value
        out[sku] = d
    return out


def rebuild_analysis(ws, skus, attrs, dry=False):
    """分析表重建为指定 SKU 集合，保留模板行的公式与样式。"""
    last = An._last_data_row(ws)
    maxc = min(ws.max_column, 300)
    tpl = An.pick_template_row(ws, last)
    fc, lc = Sh.detect_sheet_cols(ws)
    log = [f"  原数据区 r{An.DATA_FIRST_ROW}~r{last}（{last - An.DATA_FIRST_ROW + 1} 行），"
           f"模板行 r{tpl}，货件列 {fc}~{lc}"]

    # 1) 模板行克隆到第 7 行（货件列清空，避免抄到别的 SKU 的件数）
    if tpl != An.DATA_FIRST_ROW:
        An.clone_row(ws, tpl, An.DATA_FIRST_ROW, clear_range=(fc, lc), max_col=maxc)
        log.append(f"  模板 r{tpl} → r{An.DATA_FIRST_ROW}")

    # 2) 清空 8..last
    for r in range(An.DATA_FIRST_ROW + 1, last + 1):
        for c in range(1, maxc + 1):
            ws.cell(r, c).value = None

    # 3) 第 7 行 = 第一个 SKU
    def _fill(r, sku):
        a = attrs.get(sku) or {}
        ws.cell(r, 1).value = sku
        ws.cell(r, 2).value = a.get("B")
        ws.cell(r, 3).value = a.get("C")
        if not ws.cell(r, 4).value:
            ws.cell(r, 4).value = An.STATE_ON_SALE
        for col in ATTR_COLS:
            if a.get(col) is not None:
                ws.cell(r, col).value = relocate_formula(a[col], r)

    _fill(An.DATA_FIRST_ROW, skus[0])

    # 4) 其余 SKU 用补行机制追加（公式平移 / 样式照抄 / CF 跟随）
    rest = skus[1:]
    names = {s: (attrs.get(s) or {}).get("B") for s in rest}
    added = An.append_sku_rows_auto(ws, rest, names, clear_range=(fc, lc),
                                    state_map={s: An.STATE_ON_SALE for s in rest})
    for r, s in added:
        a = attrs.get(s) or {}
        for col in ATTR_COLS:
            if a.get(col) is not None:
                ws.cell(r, col).value = relocate_formula(a[col], r)
    log.append(f"  重建 {1 + len(added)} 行（r{An.DATA_FIRST_ROW}~r{An.DATA_FIRST_ROW + len(added)}）")
    # 缺属性的 SKU 提示
    miss = [s for s in skus if not (attrs.get(s) or {}).get("B")]
    if miss:
        log.append(f"  ⚠ 无型号属性的 SKU {len(miss)} 个：{miss[:8]}")
    return log


def clear_values(ws, from_row, cols=None):
    """清空数据区的**值**（保留样式与公式列）。"""
    n = 0
    maxc = ws.max_column
    rng = range(from_row, (ws.max_row or from_row) + 1)
    for r in rng:
        for c in (cols if cols else range(1, maxc + 1)):
            if ws.cell(r, c).value is not None:
                ws.cell(r, c).value = None
                n += 1
    return n


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--dry-run", action="store_true")
    args = ap.parse_args()

    if not os.path.exists(SRC):
        print(f"✗ 模板不存在：{SRC}")
        return 1
    if not os.path.exists(SA_REPORT):
        print(f"✗ 沙特库存报告不存在：{SA_REPORT}")
        return 1

    global SA_ROWS
    SA_ROWS = read_sa_rows()
    skus = [r["sku"] for r in SA_ROWS]
    print(f"沙特仓报告 SKU {len(skus)} 个，可售合计 {sum(r['fulfillable'] for r in SA_ROWS)} 件")
    print(f"  {skus}")
    if not skus:
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

    # 1) 文本替换
    n = replace_text(wb)
    print(f"  文本 阿联酋→沙特：{n} 格")

    # 2) 改 sheet 名
    for old, new in RENAMES:
        if old in wb.sheetnames:
            wb[old].title = new
    print(f"  改名后: {wb.sheetnames}")

    # 3) 分析表重建
    ws = wb[AN_NEW]
    attrs = capture_attrs(ws, An._last_data_row(ws))
    print(f"  抓到属性 {len(attrs)} 个 SKU")
    for line in rebuild_analysis(ws, skus, attrs):
        print(line)

    # 4) 站点标签 / 渠道
    ws.cell(5, 3).value = "沙特"
    old_ch = ws.cell(5, 4).value
    ws.cell(5, 4).value = "Amazon.sa"
    print(f"  C5=沙特  D5: {old_ch!r} → 'Amazon.sa'")

    # 5) 库存 sheet 清空值
    for sn in (ST_NEW, BL_NEW):
        if sn in wb.sheetnames:
            k = clear_values(wb[sn], 3)
            print(f"  {sn}: 清空 {k} 格")

    # 5b) 沙特可售库存：按沙特仓报告**预填 29 行**
    #     这张 sheet 是 kind="balance"（结余式），runner 只**更新已有行、不追加新行**。
    #     清空后不预填，它就永远是空的 —— 而「数据透视图沙特」有 1084 处 SUMIFS 引用它，
    #     跑批日志会报「更新0行 未命中0行」，透视区全变 0。
    #     模板（阿联酋簿）这张表也只有 10 行，说明它本来就是「少量 SKU 的期末余额快照」，
    #     所以这里按报告原样填，之后每天由 runner 覆盖更新。
    if BL_NEW in wb.sheetnames and SA_ROWS:
        wsb = wb[BL_NEW]
        hdr_b = {}
        for c in range(1, wsb.max_column + 1):
            v = wsb.cell(2, c).value
            if v is not None:
                hdr_b[str(v).strip().lower()] = c
        c_fnsku = hdr_b.get("fnsku")
        c_asin = hdr_b.get("asin")
        c_msku = hdr_b.get("msku") or hdr_b.get("sku")
        c_disp = hdr_b.get("disposition")
        c_bal = None
        for k, cc in hdr_b.items():
            if "ending warehouse balance" in k or "期末仓库余额" in k:
                c_bal = cc
                break
        if c_bal is None:
            c_bal = 5
        n_b = 0
        for i, rec in enumerate(SA_ROWS):
            r = 3 + i
            if c_fnsku:
                wsb.cell(r, c_fnsku).value = rec.get("fnsku") or None
            if c_asin:
                wsb.cell(r, c_asin).value = rec.get("asin") or None
            if c_msku:
                wsb.cell(r, c_msku).value = rec.get("sku") or None
            if c_disp:
                wsb.cell(r, c_disp).value = rec.get("disposition") or "SELLABLE"
            wsb.cell(r, c_bal).value = rec.get("fulfillable")
            n_b += 1
        print(f"  {BL_NEW}: 预填 {n_b} 行（可售合计 {sum(int(rec.get('fulfillable') or 0) for rec in SA_ROWS)} 件）")

    # 6) 货件列（I..R）：清掉模板簿残留的阿联酋货件号与标签
    #    行1=货件号、行2=发货日期、行3=时效、行4=预计到货、行5=列标题。
    #    行1 必须清空，否则 plan_columns 认为该列已被占用，新货件排不进来。
    #    行2~5 的公式（=IFERROR(LEFT(J$1...))、=TEXT(数据透视图沙特!$F$1...)）是通用模板，
    #    引用行1 为空时自然求值为空，保留即可；只清第 9 列那组静态值。
    fc, lc = Sh.detect_sheet_cols(ws)
    n3 = 0
    for c in range(fc, lc + 1):
        if ws.cell(1, c).value is not None:
            ws.cell(1, c).value = None
            n3 += 1
    for r in (2, 3, 4, 5):
        if ws.cell(r, fc).value is not None:
            ws.cell(r, fc).value = None
            n3 += 1
    print(f"  货件列 {fc}~{lc}: 清表头 {n3} 格")

    # 7) 订单 sheet：清 3 行数据列 + 4 行以后全部，保留 35/36 列公式
    ws2 = wb[SL_NEW]
    keep = {35, 36}
    n1 = 0
    for c in range(1, ws2.max_column + 1):
        if c in keep:
            continue
        if ws2.cell(3, c).value is not None:
            ws2.cell(3, c).value = None
            n1 += 1
    n2 = clear_values(ws2, 4)
    print(f"  {SL_NEW}: r3 清 {n1} 格（保留 35/36 公式），r4+ 清 {n2} 格")

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
