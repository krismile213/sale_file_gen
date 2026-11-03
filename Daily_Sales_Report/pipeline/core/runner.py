# -*- coding: utf-8 -*-
"""日销表程序化更新 - 主程序（第一版）

流程：备份 → 逐工作簿：库存sheet替换值 → 每日销售数据append
      → openpyxl 保存 → 库存记录按天加列+只写当天固化值(XML手术)
      → zip级保真回填(customXml/缩略图/printerSettings/pivot refreshOnLoad)

用法：
  python runner.py --dry-run                 # 只打印计划
  python runner.py --only 日本,澳大利亚       # 只跑指定文件（名字含匹配）
  python runner.py                           # 全量
"""
import os, sys, csv, shutil, zipfile, re, argparse
from collections import Counter
from copy import copy
from datetime import datetime, timedelta, date
from openpyxl import load_workbook
from openpyxl.formula.translate import Translator
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.formula import ArrayFormula

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import config as C
import record_extend

WEEK_TXT = ["周一", "周二", "周三", "周四", "周五", "周六", "周日"]

# 结余式"可售库存"sheet 的表头别名（strip+lower 匹配）→ 报告字段
BALANCE_HDR = {
    "卖家sku": "sku", "msku": "sku", "sku": "sku", "卖家 sku": "sku",
    "fnsku": "fnsku", "asin": "asin",
    "期末仓库余额": "afn-fulfillable-quantity",
    "ending warehouse balance": "afn-fulfillable-quantity",
}

# 数量/价格列白名单（标识列 sku/fnsku/asin/product-name/condition 不覆盖）
INV_VALUE_FIELDS = {
    "mfn-fulfillable-quantity", "afn-warehouse-quantity", "afn-fulfillable-quantity",
    "afn-unsellable-quantity", "afn-reserved-quantity", "afn-total-quantity",
    "per-unit-volume", "afn-inbound-working-quantity", "afn-inbound-shipped-quantity",
    "afn-inbound-receiving-quantity", "afn-researching-quantity",
    "afn-reserved-future-supply", "afn-future-supply-buyable",
    "afn-fulfillable-quantity-local", "afn-fulfillable-quantity-remote",
    "afn-fc-transfer-quantity", "afn-onhand-buyable-quantity",
    "store", "your-price", "sales-price",
}


# ---------------- 报告读取 ----------------
def _num(v):
    if v is None:
        return None
    s = str(v).strip().replace(",", "")
    if s == "":
        return None
    try:
        f = float(s)
        return int(f) if f == int(f) else f
    except ValueError:
        return v


def _enc_order(path):
    """按报告文件名给首选编码（日文界面报告是 Shift-JIS，必须排在 gb18030 之前，
    否则文本列会被 gb18030 '成功'读成乱码而数值列正常 —— 静默出错）。"""
    name = os.path.basename(str(path))
    for pat, enc in getattr(C, "REPORT_ENC", {}).items():
        if pat in name:
            return [enc] + [e for e in C.ENC_CANDIDATES if e != enc]
    return list(C.ENC_CANDIDATES)


def read_report(path, kind, sep=",", colmap=None):
    """返回 {sku: {field: value}}；统一字段名为 inv 名"""
    data, order = {}, []
    if kind == "age":
        with open(path, encoding="utf-8-sig", newline="") as f:
            for row in csv.DictReader(f):
                d = {}
                for k, v in row.items():
                    if k is None:
                        continue
                    kk = k.strip()
                    if kk in C.AGE_FIELD_MAP:
                        tgt = C.AGE_FIELD_MAP[kk]
                        d[tgt] = _num(v) if not isinstance(tgt, int) else tgt
                # age 缺失的派生字段
                ful = d.get("afn-fulfillable-quantity") or 0
                res = d.get("afn-reserved-quantity") or 0
                uns = d.get("afn-unsellable-quantity") or 0
                w, sh, rc = (d.get("afn-inbound-working-quantity") or 0,
                             d.get("afn-inbound-shipped-quantity") or 0,
                             d.get("afn-inbound-receiving-quantity") or 0)
                fct = d.get("afn-fc-transfer-quantity") or 0
                d.setdefault("afn-warehouse-quantity", ful + res + uns)
                d.setdefault("afn-total-quantity", ful + res + uns + w + sh + rc)
                d.setdefault("afn-onhand-buyable-quantity", ful + fct)
                sku = (row.get("sku") or "").strip()
                if sku:
                    data[sku] = d
                    order.append(sku)
    else:  # inv
        rows = None
        for enc in _enc_order(path):
            try:
                with open(path, encoding=enc, newline="") as f:
                    all_rows = list(csv.reader(f, delimiter=sep))
                hdr_rep = [h.strip() for h in all_rows[0]]
                if colmap or "sku" in hdr_rep:
                    rows = all_rows   # 位置映射模式：编码可解即用
                    break
            except UnicodeDecodeError:
                continue
        if rows is None:
            raise RuntimeError(f"无法解码或找不到 sku 列: {path}")
        # 报告类型闸（2026-10-12 新增，自会话区回传）：inv / balance 两类 sheet 都靠报告里的
        # afn-* 列取值。若把「库龄报告」（age：列头是 available / fc-transfer / inbound-*）
        # 误当成「管理库存报告」导出，并命名成 *管理库存报告*.csv，解析出的 dict 里就没有
        # afn-fulfillable-quantity 等字段 → 值区一个都匹配不上，而"报告里没有的 SKU 清空"
        # 的语义会把整张 sheet 的值清掉（2026-10-12 印度站实测：在途库存被清 1529 格）。
        # 位置映射（colmap=JP24，日文乱码表头）不受此闸影响。
        if not colmap:
            hdr_rep = [h.strip() for h in rows[0]]
            if not any(h.startswith("afn-") for h in hdr_rep):
                print(f"    [!] 报告类型不符: {os.path.basename(str(path))} 无 afn-* 列"
                      f"（疑似库龄报告），按缺报告处理。表头前几列={hdr_rep[:6]}")
                return {}, []
        body = rows[1:]
        for row in body:
            d = {}
            if colmap:
                names = getattr(C, colmap)
                for i, fld in enumerate(names):
                    if i < len(row):
                        d[fld] = _num(row[i])
            else:
                for i, h in enumerate(rows[0]):
                    if i < len(row):
                        d[h.strip()] = _num(row[i])
            sku = (d.get("sku") or "").strip()
            if sku:
                data[sku] = d
                order.append(sku)
    return data, order


# ---------------- 库存 sheet 更新 ----------------
def _header_row(ws, fallback=2):
    af = ws.auto_filter.ref
    if af:
        m = re.match(r"[A-Z]+(\d+)", af)
        if m:
            return int(m.group(1))
    return fallback


def _last_data_row(ws, key_col, start, limit=None):
    # max_row 不越界：普通模式 iter_rows 超出会实例化空 cell（文件膨胀+dimension 虚胖）
    top = ws.max_row if limit is None else min(limit, ws.max_row)
    last = start - 1
    for r, row in enumerate(ws.iter_rows(min_row=start, max_row=top, min_col=key_col, max_col=key_col), start):
        if row[0].value is not None:
            last = r
    return last


def update_stock_sheet(ws, data, order, kind="inv", colmap=None):
    if not data:
        print(f"    [!] {ws.title}: 报告数据为空，跳过（防止清空整表）")
        return None
    hdr = _header_row(ws)
    headers = {}
    for c in ws[hdr]:
        if c.value is not None:
            headers[str(c.value).strip()] = c.column

    if kind == "balance":
        # 结余式可售库存 sheet：FNSKU/ASIN/MSKU/Disposition/期末仓库余额
        alias = {h: col for h, col in headers.items() if h.strip().lower() in BALANCE_HDR}
        key_col = alias.get("sku") or next((c for h, c in alias.items() if BALANCE_HDR[h.strip().lower()] == "sku"), None)
        bal_col = next((c for h, c in alias.items() if BALANCE_HDR[h.strip().lower()] == "afn-fulfillable-quantity"), None)
        if not key_col or not bal_col:
            print(f"    [!] {ws.title}: 结余式表头识别失败 {list(headers)[:6]}，跳过")
            return None
        start = hdr + 1
        last = _last_data_row(ws, key_col, start)
        updated, missing = 0, 0
        for r in range(start, last + 1):
            sku = ws.cell(row=r, column=key_col).value
            if sku is None:
                continue
            d = data.get(str(sku).strip())
            if d and d.get("afn-fulfillable-quantity") is not None:
                ws.cell(row=r, column=bal_col).value = d["afn-fulfillable-quantity"]
                updated += 1
            else:
                missing += 1
        print(f"    [库存] {ws.title}(结余式): 更新{updated}行 未命中{missing}行 (数据区 r{start}~r{last})")
        return dict(sheet=ws.title, updated=updated, cleared=0, appended=0, last=last)

    # inv 模式
    if colmap:
        # 位置模式（日文乱码表头）：第 i 列 = JP24[i-1]，只覆盖前 len(colmap) 列
        names = getattr(C, colmap)
        headers = {}
        for i, nm in enumerate(names, start=1):
            headers[nm] = i
        # 加工列：colmap 之外的列若数据首行是公式，记为 formula_cols
    key_col = headers.get("sku")
    if not key_col:
        print(f"    [!] {ws.title}: 找不到 sku 列，跳过")
        return None
    start = hdr + 1
    last = _last_data_row(ws, key_col, start)
    skus_in_sheet = {}

    # 表头名 → 报告字段（限白名单）
    col_field = {col: h for h, col in headers.items() if h in INV_VALUE_FIELDS}
    # 本地公式列（表头不在报告字段、数据首行是公式）→ append 时要补公式
    first_row = start
    formula_cols = {}
    for col in range(1, ws.max_column + 1):
        if col in col_field.values() or col == key_col:
            continue
        v = ws.cell(row=first_row, column=col).value
        if isinstance(v, str) and v.startswith("="):
            formula_cols[col] = v

    updated, cleared, appended = 0, 0, 0
    # 1) 既有行更新
    for r in range(start, last + 1):
        sku = ws.cell(row=r, column=key_col).value
        if sku is None:
            continue
        sku = str(sku).strip()
        skus_in_sheet[sku] = r
        d = data.get(sku)
        if d is None:
            # 报告没有该 SKU：值区清空（等效整表替换），公式列不动
            for col, fld in col_field.items():
                if ws.cell(row=r, column=col).value is not None:
                    ws.cell(row=r, column=col).value = None
                    cleared += 1
            continue
        for col, fld in col_field.items():
            if fld in d and d[fld] is not None:
                ws.cell(row=r, column=col).value = d[fld]
                updated += 1
    # 2) 报告新增 SKU → append
    ref_row = last if last >= start else start
    for sku in order:
        if sku in skus_in_sheet:
            continue
        r = (max((max(skus_in_sheet.values()) if skus_in_sheet else last), last)) + 1
        skus_in_sheet[sku] = r
        d = data[sku]
        for col in range(1, ws.max_column + 1):
            src = ws.cell(row=ref_row, column=col)
            dst = ws.cell(row=r, column=col)
            if src.has_style:
                dst._style = copy(src._style)
        for h, col in headers.items():
            if h in data[sku]:
                ws.cell(row=r, column=col).value = data[sku][h]
        for col, f0 in formula_cols.items():
            L = get_column_letter(col)
            ws.cell(row=r, column=col).value = Translator(f0, origin=f"{L}{first_row}").translate_formula(f"{L}{r}")
        appended += 1
    print(f"    [库存] {ws.title}: 更新{updated}格 清空{cleared}格 新增{appended}行 (数据区 r{start}~r{last})")
    return dict(sheet=ws.title, updated=updated, cleared=cleared, appended=appended, last=last)


# ---------------- 多国库存 / verbatim 整块重贴（multi） ----------------
# 工作簿「多国库存」sheet = 亚马逊「多国库存报告」原样粘贴（列名同名），
# A1=下载日期、B1==数据透视表!B11 锚点、行2=报告表头、行3起=数据，另有 autoFilter。
# 报告行数每天变（本地履约 SKU 增减），故按【整块重贴 + 多退少补】处理：
# 逐行按列名写入、超出的旧行清空、autoFilter 跟随数据末行。
def read_multi(path):
    """返回 (报告表头列表, 数据行列表)；整块重贴用（UTF-8-BOM CSV）。"""
    with open(path, encoding="utf-8-sig", newline="") as f:
        allrows = list(csv.reader(f))
    hdr = [h.strip() for h in allrows[0]]
    body = [r for r in allrows[1:] if any(str(x).strip() for x in r)]
    return hdr, body


def update_multi_sheet(ws, rep_hdr, rep_rows):
    """把报告整块重贴到 sheet（按列名对齐；保留 A1 锚点行 / 表头 / 样式）。"""
    hdr = _header_row(ws)
    sheet_cols = {}
    for c in ws[hdr]:
        if c.value is not None:
            sheet_cols[str(c.value).strip()] = c.column
    pairs = [(i, sheet_cols[h]) for i, h in enumerate(rep_hdr) if h in sheet_cols]
    if not pairs:
        print(f"    [!] {ws.title}: 报告列名与表头不匹配 {rep_hdr[:6]}，跳过")
        return None
    cols = sorted({col for _i, col in pairs})
    first_col, last_col = min(cols), max(cols)
    start = hdr + 1
    last = _last_data_row(ws, first_col, start)
    n = len(rep_rows)
    for k, rv in enumerate(rep_rows):
        r = start + k
        if r > last:      # 新行：继承最末数据行样式，避免丢边框/格式
            for c in range(1, ws.max_column + 1):
                src = ws.cell(row=last, column=c)
                if src.has_style:
                    ws.cell(row=r, column=c)._style = copy(src._style)
        for i, col in pairs:
            ws.cell(row=r, column=col).value = _num(rv[i]) if i < len(rv) else None
    cleared = 0
    for r in range(start + n, last + 1):
        if any(ws.cell(row=r, column=c).value is not None for c in cols):
            for c in cols:
                ws.cell(row=r, column=c).value = None
            cleared += 1
    new_last = max(last, start + n - 1)
    try:
        ws.auto_filter.ref = "%s%d:%s%d" % (
            get_column_letter(first_col), hdr, get_column_letter(last_col), new_last)
    except Exception:
        pass
    print(f"    [库存] {ws.title}(整块重贴): 写入{n}行 清空{cleared}行 "
          f"(原 r{start}~r{last} → r{start}~r{start + n - 1})")
    return dict(sheet=ws.title, rows=n, cleared=cleared)


# ---------------- 美客多 Full 仓库存（ml_full） ----------------
def read_ml_full(path):
    """ML后台"库存总览(Resumen)"报告：r10主表头/r11r12子表头，数据r13起，SKU在c4。
    列 c2..c34 ↔ 工作簿 c1..c33（表头逐列核对零差异）。返回按报告顺序的行值列表。"""
    wb = load_workbook(path, read_only=True, data_only=True)
    sn = "Resumen" if "Resumen" in wb.sheetnames else wb.sheetnames[0]
    ws = wb[sn]
    rows = []
    for row in ws.iter_rows(min_row=13, max_col=34, values_only=True):
        sku = row[3]
        if sku is None or not str(sku).strip():
            continue
        rows.append(list(row[1:34]))
    wb.close()
    return rows


def update_ml_stock_detail(ws, rows, snap_date):
    """美客多「库存明细」整块替换（等效手工粘贴报告块）：
    数据区 r5..（c3=SKU）；新块顺序写入，多余旧行 delete_rows（已核实无跨表公式引用尾部行）；
    A1 下载日锚点 → 快照日（全簿3122处公式引用它联动透视图/分析表/订单数据）。"""
    first = last = None
    for r, row in enumerate(ws.iter_rows(min_row=5, min_col=3, max_col=3), 5):
        v = row[0].value
        if v is not None and str(v).strip():
            if first is None:
                first = r
            last = r
    if first is None:
        first, last = 5, 4
    old_n = last - first + 1
    n = len(rows)
    n_clr = 0
    for i, vals in enumerate(rows):
        for c, v in enumerate(vals, 1):
            # 整块替换 = 报告空值也必须落成空格。只写非空值会留下上一块的残格：
            # 2026-10-21 上午把订单报告误命名成 Full 库存报告，整张「库存明细」被写成
            # 1170 行订单数据；下午换回正确报告后重跑，B/E/H 三列（报告恒空/大面积为空）
            # 若不清空就会残留订单文案，透视图/分析表照旧读到脏数据。
            if v is not None:
                ws.cell(row=first + i, column=c).value = v
            else:
                cell = ws.cell(row=first + i, column=c)
                if cell.value is not None:
                    cell.value = None
                    n_clr += 1
    n_del = 0
    if old_n > n:
        n_del = old_n - n
        ws.delete_rows(first + n, n_del)
    new_last = first + n - 1
    # autoFilter 收缩/扩展
    af = ws.auto_filter.ref
    if af:
        m = re.match(r"([A-Z]+)(\d+):([A-Z]+)(\d+)", af)
        if m:
            ws.auto_filter.ref = f"{m.group(1)}{m.group(2)}:{m.group(3)}{new_last}"
    old_anchor = ws["A1"].value
    ws["A1"].value = datetime.strptime(snap_date, "%Y-%m-%d")
    print(f"    [库存] {ws.title}: 整块替换 {old_n}→{n} 行 (r{first}~r{new_last}，删尾{n_del}，"
          f"清空格{n_clr})，锚点A1: {old_anchor} → {snap_date}")
    return dict(sheet=ws.title, old=old_n, new=n, deleted=n_del)


# ---------------- eMAG罗马尼亚库存（emag_stock：优惠更新文件第3个sheet「库存」） ----------------
# 数据源=eMAG后台预填的优惠更新文件（含库存那份，sheet: 说明/数据定义/库存/库存_errors），
# 「库存」sheet r3=字段名 r4=必填性 r5=中文表头，数据 r6 起：
#   c6=PNK  c7=产品代码(G***)  c10=proposed_quantity(数量)  c11=stock_type(ready_to_sell等)
# 工作簿「库存明细」手工口径=每天把4列按行重贴：c4←PNK c5←产品代码 c27←数量(int) c28←stock_type，
# 其余列不动 + A1锚点滚到快照日。
# ⚠️ 该文件 XML dimension 记录失真（库存sheet标成1×1），read_only 模式会漏读全部数据——
#     必须普通模式 load_workbook 打开（2026-10-13 实测踩坑）。
# 安全性：分析表/透视图对库存明细的 1851 处引用全部为开放列区间
#     （AA5:AA29898 / E5:E29898 / AB5:AB29898 + A1锚点），行数增减/重排不影响（probe_emag_stock.py 实测）。
def read_emag_stock(path):
    wb = load_workbook(path, data_only=True)   # 普通模式，勿改 read_only（dimension 失真）
    sn = "库存" if "库存" in wb.sheetnames else wb.sheetnames[min(2, len(wb.sheetnames) - 1)]
    ws = wb[sn]
    rows = []
    for r in range(6, ws.max_row + 1):
        pnk = ws.cell(row=r, column=6).value
        code = ws.cell(row=r, column=7).value
        qty = ws.cell(row=r, column=10).value
        st = ws.cell(row=r, column=11).value
        if pnk is None and code is None:
            continue
        q = _num(qty)   # 报告里数量是文本 '4'
        rows.append(dict(
            pnk=str(pnk).strip() if pnk is not None else "",
            code=str(code).strip() if code is not None else "",
            qty=int(q) if isinstance(q, float) and q == int(q) else q,
            st=str(st).strip() if st is not None else "",
        ))
    wb.close()
    return rows


def update_emag_stock_detail(ws, rows, snap_date):
    """罗马尼亚「库存明细」整块重贴（等效手工每天贴4列）：
    c4←PNK c5←产品代码 c27←当前库存(int) c28←stock_type；数据 r5 起；其余列不动；
    A1 锚点 → 快照日（透视图 E1/F1 直接引用它）。"""
    first = last = None
    for r, row in enumerate(ws.iter_rows(min_row=5, min_col=5, max_col=5), 5):
        v = row[0].value
        if v is not None and str(v).strip():
            if first is None:
                first = r
            last = r
    if first is None:
        first, last = 5, 4
    old_n = last - first + 1
    n = len(rows)
    for i, x in enumerate(rows):
        r = first + i
        ws.cell(row=r, column=4).value = x["pnk"] or None
        ws.cell(row=r, column=5).value = x["code"] or None
        ws.cell(row=r, column=27).value = x["qty"]
        ws.cell(row=r, column=28).value = x["st"] or None
    n_del = 0
    if old_n > n:
        n_del = old_n - n
        ws.delete_rows(first + n, n_del)
    new_last = first + n - 1
    af = ws.auto_filter.ref
    if af:
        m = re.match(r"([A-Z]+)(\d+):([A-Z]+)(\d+)", af)
        if m:
            ws.auto_filter.ref = f"{m.group(1)}{m.group(2)}:{m.group(3)}{new_last}"
    old_anchor = ws["A1"].value
    ws["A1"].value = datetime.strptime(snap_date, "%Y-%m-%d")
    print(f"    [库存] {ws.title}: 整块重贴 {old_n}→{n} 行 (r{first}~r{new_last}，删尾{n_del})，"
          f"锚点A1: {old_anchor} → {snap_date}")
    return dict(sheet=ws.title, old=old_n, new=n, deleted=n_del)


# ---------------- 美客多订单（ml_sales：xlsx月报 + 中文→西语正典规范化） ----------------
# 背景：ML后台语言漂移（202608/202610 中文、202609 西语），而工作簿「订单数据」正典=西语，
# c64 purchase-date / c65 有效性两个公式在文本上查「墨西哥透视图」的西语月名表(O1:P12)与
# 西语关键词表(K2:L38)——中文原文进来全 #N/A，分析表 SUMIFS(BL,BM=1) 整列漏计。
# 故 append 前把语言相关列规范化回西语正典（字符串与规则见 config.ML_*）。
_EN_MONTHS = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto",
              "septiembre", "octubre", "noviembre", "diciembre"]
_ML_MONTH_NUM = {nm: i + 1 for i, nm in enumerate(_EN_MONTHS)}
_ML_ZH_MON = re.compile(r"(十月|十一月|十二月|一月|二月|三月|四月|五月|六月|七月|八月|九月)")
# 数值列：单位/金额/单价（含退货、投诉区的单位列）
_ML_NUMERIC_COLS = {7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 23, 47, 60}


def _es_snap(snap_date):
    d = datetime.strptime(snap_date, "%Y-%m-%d")
    return f"{d.day} de {_EN_MONTHS[d.month - 1]}"


def ml_norm_date(s):
    """'9 de 十月 de 2026 23:29 hs.' → ('9 de octubre de 2026 23:29 hs.', datetime(2026,10,9))。
    西语原文原样返回；解析不出返回 (原文, None)。"""
    if s is None:
        return "", None
    s = str(s).strip()
    m = re.search(r"(\d{1,2}) de (\S+) de (\d{4})(.*)", s)
    if not m:
        return s, None
    mes = C.ML_MONTHS.get(m.group(2), m.group(2))
    num = _ML_MONTH_NUM.get(mes)
    if not num:
        return s, None
    try:
        d = datetime(int(m.group(3)), num, int(m.group(1)))
    except ValueError:
        return s, None
    return f"{int(m.group(1))} de {mes} de {m.group(3)}{m.group(4)}", d


def ml_norm_status(s, snap_date):
    """状态(c3) 中文→西语正典；西语原样；未识别返回 None（调用方保留原文并告警）。"""
    s = str(s or "").strip()
    if s in C.ML_STATUS_MAP:
        t = C.ML_STATUS_MAP[s]
        return t.format(snap=_es_snap(snap_date)) if "{" in t else t
    for pat, t in C.ML_STATUS_PATTERNS:
        m = re.match(pat, s)
        if m:
            gd = {k: v for k, v in (m.groupdict() or {}).items()}
            gd.setdefault("d", "")
            gd.setdefault("snap", _es_snap(snap_date))
            gd.setdefault("s", "" if gd.get("d") == "1" else "s")   # 西语单复数: día/días
            return t.format(**gd)
    return None


def ml_norm_desc(s, snap_date):
    """状态描述(c4) 正则→西语；未命中原样返回（c65 会回退查 c3）。"""
    s = str(s or "").strip()
    for pat, t in C.ML_DESC_PATTERNS:
        m = re.match(pat, s)
        if m:
            gd = {k: (v if v is not None else "") for k, v in (m.groupdict() or {}).items()}
            if gd.get("wd"):
                gd["wd"] = C.ML_WEEKDAYS.get(gd["wd"], gd["wd"])
            # 月份：优先用模式里捕获的中文月名；否则在"星期X月"段剔除后搜索
            # （'星期一月12日'里的'一月'是星期+乱码de，不是一月份）
            if gd.get("zhm"):
                gd["m"] = C.ML_MONTHS[gd["zhm"]]
            else:
                gd["m"] = _zh_or_es_month(re.sub(r"星期.月", "", s), snap_date)
            gd.setdefault("d", "")
            gd.setdefault("d2", "")
            gd.setdefault("snap", _es_snap(snap_date))
            try:
                return t.format(**gd)
            except (KeyError, IndexError):
                return s
    return s


def _zh_or_es_month(text, snap_date):
    m = _ML_ZH_MON.search(text)
    if m:
        return C.ML_MONTHS[m.group(1)]
    for en in _EN_MONTHS:
        if en in text:
            return en
    return _EN_MONTHS[datetime.strptime(snap_date, "%Y-%m-%d").month - 1]


def ml_normalize_row(vals, snap_date):
    """报告数据行(c1..c63) → 西语正典行。返回 (row63, 解析出的销售日期)。"""
    vals = list(vals[:63]) + [None] * max(0, 63 - len(vals))
    row = [_cell_str(v) for v in vals]                   # datetime/float 统一转文本
    row = [x if x != "" else " " for x in row]          # ML导出空值=正典' '
    row[1], d = ml_norm_date(row[1])                     # c2 Fecha de venta
    st = ml_norm_status(row[2], snap_date)               # c3 Estado
    st_known = st is not None
    if st is not None:
        row[2] = st
    row[3] = ml_norm_desc(row[3], snap_date)             # c4 Descripción
    for ci in (4, 5, 17, 24, 40):                        # c5/c6/c18/c25/c41 值映射
        v = row[ci].strip()
        if v in C.ML_VALUE_MAP:
            row[ci] = C.ML_VALUE_MAP[v]
    for ci in _ML_NUMERIC_COLS:                          # 数值列
        row[ci - 1] = _num(row[ci - 1]) if isinstance(row[ci - 1], str) and row[ci - 1].strip() else None
    return row, d, st_known


def _read_k_table(wb):
    """透视图 K2:L38 有效性关键词表（作者口径，语义已用缓存值1183行复刻验证0差异）。"""
    kw = []
    if "墨西哥透视图" in wb.sheetnames:
        wsp = wb["墨西哥透视图"]
        for r in range(2, 60):
            k = wsp.cell(row=r, column=11).value
            if k is None:
                break
            kw.append((str(k), wsp.cell(row=r, column=12).value))
    return kw


def _ml_validity(kw, D, Cstat, S):
    """复刻 订单数据!BM（c65 有效性）: SWITCH(S2=' ',0, XLOOKUP(D2)→FIND(D2)→FIND(C2))；无命中=None(#N/A)。"""
    if S == " ":
        return 0
    for k, l in kw:
        if D == k:
            return l
    hits = [l for k, l in kw if k in D]
    if hits:
        return hits[-1]
    hits = [l for k, l in kw if k in Cstat]
    if hits:
        return hits[-1]
    return None


def append_ml_sales(ws, report_path, window_lo=None, snap_date=None):
    hdr = _header_row(ws)
    headers = {str(c.value).strip(): c.column for c in ws[hdr] if c.value is not None}
    key_col = headers.get("# de venta", 1)
    first_row = hdr + 1
    last = _last_data_row(ws, key_col, first_row)
    have = set()
    for r in range(first_row, last + 1):
        v = ws.cell(row=r, column=key_col).value
        if v is not None and str(v).strip():
            have.add(str(v).strip())
    # 公式列源：c64 purchase date / c65 有效性（数组公式，从首个数据行平移）
    f_src = {}
    for col, L in ((64, "BL"), (65, "BM")):
        v = ws.cell(row=first_row, column=col).value
        if isinstance(v, ArrayFormula):
            f_src[col] = (L, v.text)
        elif isinstance(v, str) and v.startswith("="):
            f_src[col] = (L, v)
    if 64 not in f_src or 65 not in f_src:
        print(f"    [!] {ws.title}: c64/c65 公式列缺失，ml_sales 跳过（防写坏统计链）")
        return None
    # 读报告（数据行=第1列为纯数字）
    wb_r = load_workbook(report_path, read_only=True, data_only=True)
    ws_r = wb_r[wb_r.sheetnames[0]]
    raw = list(ws_r.iter_rows(min_row=1, values_only=True))
    wb_r.close()
    body = [r for r in raw if r and r[0] is not None and str(r[0]).strip().isdigit()]
    kw = _read_k_table(ws.parent)
    snap_es = _es_snap(snap_date)
    new_rows, n_win, n_dup = [], 0, 0
    v_cnt, na_samples = Counter(), {}
    for r in body:
        row, d, st_known = ml_normalize_row(r, snap_date)
        if window_lo and d is not None and d.strftime("%Y-%m-%d") < window_lo:
            n_win += 1
            continue
        k = str(row[0]).strip()
        if k in have:
            n_dup += 1
            continue
        have.add(k)
        v = _ml_validity(kw, row[3], row[2], row[18])
        v_cnt[v] += 1
        if v is None and len(na_samples) < 8:
            na_samples[f"{row[2][:30]} | {row[3][:40]}"] = 1
        new_rows.append(row)
    # 写入
    r = last
    ref_row = last if last >= first_row else first_row
    for row in new_rows:
        r += 1
        for c in range(1, 67):
            src = ws.cell(row=ref_row, column=c)
            dst = ws.cell(row=r, column=c)
            if src.has_style:
                dst._style = copy(src._style)
        for c in range(1, 64):
            v = row[c - 1]
            if v is not None:
                ws.cell(row=r, column=c).value = v
        for col, (L, f0) in f_src.items():
            ws.cell(row=r, column=col).value = Translator(f0, origin=f"{L}{first_row}").translate_formula(f"{L}{r}")
    af = ws.auto_filter.ref
    if af:
        m = re.match(r"([A-Z]+)(\d+):([A-Z]+)(\d+)", af)
        if m:
            ws.auto_filter.ref = f"{m.group(1)}{m.group(2)}:{m.group(3)}{r}"
    win = f" 窗口外过滤{n_win}行" if n_win else ""
    dup = f" 去重跳过{n_dup}行" if n_dup else ""
    na = ""
    if v_cnt.get(None):
        na = f" ⚠️#N/A {v_cnt[None]}行(不计销): " + "; ".join(list(na_samples)[:5])
    print(f"    [销售] {ws.title} ← {os.path.basename(report_path)}: 报告{len(body)}行{win}{dup} "
          f"追加{len(new_rows)}行 (r{last}→r{r})")
    print(f"        有效性复刻: 计销(1)={v_cnt.get(1, 0)} 不计(0)={v_cnt.get(0, 0)}{na}")
    return dict(sheet=ws.title, total=len(body), appended=len(new_rows), blocked=n_win,
                valid=v_cnt.get(1, 0), invalid=v_cnt.get(0, 0), na=v_cnt.get(None, 0))


# ---------------- 库存记录加列（按天插列 + 只写当天固化值，保存后 XML 手术） ----------------
def build_record_jobs(wb_cfg, stock_plans):
    """按 records 配置收集每个记录表的【跑批当天】填充数据。

    源 = 同名基座 stock sheet 的报告数据（美国库存记录 ← 美国可售+在途库存）；
    源报告缺失（没进 stock_plans，如缺报告/0 行被安全闸跳过）→ 不建任务：
    没数据就不写列，避免老逻辑"复制上一列"造成的假历史/复制作旧值。
    在售字段默认 afn-onhand-buyable-quantity，墨西哥记录表用 fulfillable（config 覆盖）。
    """
    jobs = []
    for rsn in wb_cfg.get("records", []):
        base = rsn.replace("库存记录", "").strip()
        src_sheet = base + "可售+在途库存"
        plan = next((p for p in stock_plans if p[0].get("sheet") == src_sheet), None)
        if plan is None:
            # 预留库存记录等无日期列的表本来就没有 src，手术时会自动跳过
            continue
        data = plan[1] or {}
        if not data:
            # 空报告（0 行）不能当"全 0"写进记录列——等同伪造历史，直接不收
            print(f"    [记录] {rsn}: 源报告 {src_sheet} 解析出 0 行，不加列")
            continue
        insale_field = getattr(C, "RECORD_INSALE_FIELD", {}).get(
            rsn.strip(), "afn-onhand-buyable-quantity")
        insale, intra = record_extend.fill_from_report(data, insale_field)
        jobs.append(dict(sheet=rsn, insale=insale, intra=intra))
    return jobs


def extend_stock_record(wb, sheet_name, until_date_str):
    """已废弃：老逻辑（新列=复制上一列）会把活公式历史列静默刷新成最新快照、
    或把固化值列的前一天数值抄进新列。按天加列 + 当天固化值必须在 openpyxl
    保存之后做 XML 手术：record_extend.extend_record_file(path, sheet, date, ...)。"""
    raise RuntimeError(
        f"extend_stock_record 已废弃（{sheet_name} {until_date_str}）："
        "请改走保存后的 record_extend.extend_record_file()")


# ---------------- 每日销售数据 append ----------------
# ID 类列保持文本（人工口径=Excel导入向导按文本处理；写成 int 会被显示成 1.7E+14 科学计数法）
TEXT_COLS = {"order-item-id", "order-id", "asin1", "isbn", "fnsku", "订单编号", "# de venta"}


def fix_id_text(ws):
    """存量修复：ID 类长数字列的 int → 文本。幂等，每天跑无害。"""
    hdr = _header_row(ws)
    headers = {str(c.value).strip(): c.column for c in ws[hdr] if c.value is not None}
    cols = [c for h, c in headers.items() if h.lower() in TEXT_COLS]
    n = 0
    for col in cols:
        for row in ws.iter_rows(min_row=hdr + 1, min_col=col, max_col=col):
            c = row[0]
            if isinstance(c.value, int):
                c.value = str(c.value)
                n += 1
    return n


def _cell_str(v):
    """xlsx 单元格 → 报告口径字符串：datetime 转标准文本、float 整数化、None 空串"""
    if v is None:
        return ""
    if hasattr(v, "strftime"):
        return v.strftime("%Y-%m-%d %H:%M:%S")
    if isinstance(v, float) and v == int(v):
        return str(int(v))
    return str(v).strip()


def _read_report_rows(report_path):
    """txt/csv → csv.reader；xlsx → 取数据行最多的 sheet 全表。返回 (header, body)。"""
    if str(report_path).lower().endswith((".xlsx", ".xlsm")):
        wb_r = load_workbook(report_path, read_only=True, data_only=True)
        sn_r = max(wb_r.sheetnames, key=lambda s: wb_r[s].max_row or 0)
        ws_r = wb_r[sn_r]
        rows = [[_cell_str(c.value) for c in row] for row in ws_r.iter_rows()]
        wb_r.close()
        if not rows:
            return [], []
        return rows[0], rows[1:]
    for enc in _enc_order(report_path):
        try:
            with open(report_path, encoding=enc, newline="") as f:
                rows = list(csv.reader(f, delimiter=sep_default(report_path)))
            return rows[0], rows[1:]
        except UnicodeDecodeError:
            continue
    raise RuntimeError(f"报告解码失败: {report_path}")


def sep_default(p):
    return "\t" if str(p).lower().endswith(".txt") else ","


_CHANNEL_OWNERS = None


def _channel_owners():
    """{sales-channel: {簿文件名}} —— 全 12 簿 `expect_channel` 的并集。

    用途：`append_sales` 判定「这条渠道属于**别的**簿」——即源报告错位（把沙特数据
    下载成阿联酋报告名）时的兜底。只拒「已知属于其它簿」的渠道；完全没见过的渠道
    （如新开的欧盟站点）**保留并告警**，避免误杀合法行。见 config.sales 的 expect_channel。
    """
    owners = {}
    for bk in C.WORKBOOKS:
        specs = bk.get("sales")
        specs = specs if isinstance(specs, list) else ([specs] if specs else [])
        for sc in specs:
            for ch in (sc.get("expect_channel") or ()):
                owners.setdefault(ch, set()).add(bk.get("file", "?"))
    return owners


def _resolve_report(snap_dir, tmpls, fmt):
    """报告模板 → 第一个存在的路径（2026-10-24 加通配符回退）。

    模板可以是单个字符串，也可以是候选列表；含 `*` / `?` 时按 glob 匹配并取
    **字典序最新**的一个（文件名带日期，字典序=时间序），并忽略 `xxx (1).csv`
    这类重复下载副本。

    为什么需要：interflow 沙特仓的库存报告下载频率低于其它站点，
    20261023 快照目录里最新一份仍叫 `…20261021.csv`（内容确实是新的，只是文件名没变），
    死认 {d} 会天天判「缺报告」→ 库存永远不更新。
    """
    import glob as _glob
    for t in (tmpls if isinstance(tmpls, list) else [tmpls]):
        try:
            c = t.format(**fmt)
        except (KeyError, IndexError):
            c = t
        p = c if os.path.isabs(c) else os.path.join(snap_dir, c)
        if "*" in p or "?" in p:
            hits = [h for h in _glob.glob(p) if "(" not in os.path.basename(h)]
            hits = hits or _glob.glob(p)
            if hits:
                return sorted(hits)[-1]
            continue
        if os.path.exists(p):
            return p
    return None


def screen_channels(body, ch_idx, expect_channel, book, owners=None,
                    allow_mixed=False):
    """渠道错位筛查 → (kept, stat)。纯函数，便于单测。

    2026-10-19 事故后加严：整份报告错位（外来渠道占比 ≥50%，或报告里一条本簿渠道
    都没有）时 **整份拒收**（kept=[]）。原来只逐行拒、剩下多少导多少 —— 沙特那次
    错位报告混进 37 行，看着像"跑过了"，其实当天真实销售根本没拿到，属于静默漏数。

    stat: dict(total, foreign, unknown, own, foreign_cnt, unknown_cnt, mislabel)
    mislabel=True ⇒ 调用方必须把整份报告当"缺报告"处理（不入库、记异常）。
    """
    stat = dict(total=0, foreign=0, unknown=0, own=0, foreign_cnt={},
                unknown_cnt={}, mislabel=False)
    if not expect_channel or ch_idx is None or not body:
        return body, stat
    global _CHANNEL_OWNERS
    if _CHANNEL_OWNERS is None:
        _CHANNEL_OWNERS = _channel_owners()
    own = set(expect_channel)
    foreign_cnt, unknown_cnt = {}, {}
    kept = []
    for row in body:
        ch = row[ch_idx].strip() if len(row) > ch_idx else ""
        # 空渠道 / 本簿渠道 / Non-Amazon* → 一律放行
        if not ch or ch in own or ch.startswith("Non-Amazon"):
            if ch in own:
                stat["own"] += 1
            kept.append(row)
            continue
        other = _CHANNEL_OWNERS.get(ch)
        if other and book not in other:
            stat["foreign"] += 1
            foreign_cnt[ch] = foreign_cnt.get(ch, 0) + 1
            continue
        stat["unknown"] += 1
        unknown_cnt[ch] = unknown_cnt.get(ch, 0) + 1
        kept.append(row)
    stat["total"] = len(body)
    stat["foreign_cnt"], stat["unknown_cnt"] = foreign_cnt, unknown_cnt
    if allow_mixed:
        # 2026-10-24：interflow 沙特/阿联酋两份订单报告是**同一份下载**（sha256 相同），
        # 里面天然混着 Amazon.ae 58 件 + Amazon.sa 5 件。两个簿各取自己渠道是设计意图，
        # 不是"导错账号"。此时只做逐行过滤，不升级成整份拒收。
        stat["mixed_allowed"] = True
        return kept, stat
    # 整份错位：外来占比 ≥50%，或整份报告里没有一条本簿渠道（却有别家渠道的行）
    if stat["foreign"] >= max(1, stat["total"]) * 0.5 or \
            (stat["own"] == 0 and stat["foreign"] > 0):
        stat["mislabel"] = True
        return [], stat
    return kept, stat


def screen_sku_overlap(body, sku_idx, stock_skus, book, report_path, min_rate=0.6):
    """SKU 归属闸（2026-10-22 新增）→ (kept, stat|None)。

    **为什么渠道闸挡不住**：`Amazon.com` / `Amazon.ca` 同时属于「北美加拿大簿」和
    「Interflowd北美加拿大簿」，`screen_channels` 只要发现本簿在 owner 集合里就放行。
    2026-10 实测：interflow 北美订单报告被换成了主账号的（5325 行，含 .mx 1047），
    渠道全部合法，runner 照单全收 —— `Interflowd北美加拿大库存销售表 / 每日销售数据`
    被灌进 4288 行主账号订单（占 10 月 90%），两本分析表销量虚高。

    判据只能靠 SKU 归属：订单报告的 SKU 集与本簿**库存 sheet** 的 SKU 集重合率。
    实测合法组合全部 ≥96.5%（欧盟英国 96.5%、其余 100%），错位报告只有 15%。
    阈值 60% 两边都留足余量。
    """
    if not stock_skus or sku_idx is None or not body:
        return body, None
    rep = {r[sku_idx].strip() for r in body
           if len(r) > sku_idx and r[sku_idx].strip()}
    if not rep:
        return body, None
    hit = len(rep & stock_skus)
    rate = hit / len(rep)
    if rate >= min_rate:
        return body, None
    print(f"    [销售] ✗✗ {os.path.basename(str(report_path))}: **SKU 归属不符**"
          f"（报告 {len(rep)} 个 SKU 仅 {hit} 个在本簿库存里，{rate:.0%} < {min_rate:.0%}）")
    print(f"    [销售] ✗✗ 疑似导成了**别的账号**的同渠道报告（同名 sales-channel 不同卖家），"
          f"已整份拒收、本簿今日销售不更新 —— 请切到本账号重新导出")
    return [], dict(total=len(body), sku_total=len(rep), sku_hit=hit, rate=rate)


def append_sales(ws, report_path, sep="\t", key=None, expect_channel=None, book=None,
                 stock_skus=None, allow_mixed=False):
    hdr = _header_row(ws)
    headers = {str(c.value).strip(): c.column for c in ws[hdr] if c.value is not None}
    oid_col = (headers.get(key) if key else None) \
        or headers.get("order-item-id") or headers.get("amazon-order-id")
    if not oid_col:
        print(f"    [销售] {ws.title}: 找不到 order-item-id 列，跳过")
        return None
    # 已有订单键：组合键 oid+amazon-order-id+sku+purchase-date
    # （纯 oid 去重会把 MCF 多渠道配送单误杀——亚马逊给它们的 order-item-id 全是 '0'）
    aoi_col = headers.get("amazon-order-id")
    sku_col_h = headers.get("sku")
    pd_col_h = headers.get("purchase-date")
    use_comp = all((aoi_col, sku_col_h, pd_col_h))

    def _comp(oid, aoi, sku, pdv):
        p4 = str(pdv).strip()[:19] if pdv is not None else ""
        return "\x01".join(["" if oid is None else str(oid).strip(),
                            "" if aoi is None else str(aoi).strip(),
                            "" if sku is None else str(sku).strip(), p4])

    _cols = sorted({c for c in (oid_col, aoi_col, sku_col_h, pd_col_h) if c})
    have = set()
    for r, row in enumerate(ws.iter_rows(min_row=hdr + 1, min_col=_cols[0], max_col=_cols[-1]), hdr + 1):
        vals = {c: row[c - _cols[0]].value for c in _cols}
        oidv = vals.get(oid_col)
        if oidv is None:
            continue
        if use_comp:
            have.add(_comp(oidv, vals.get(aoi_col), vals.get(sku_col_h), vals.get(pd_col_h)))
        else:
            have.add(str(oidv).strip())
    # 读报告（xlsx 或 txt/csv）
    hdr_rep, body = _read_report_rows(report_path)
    rep_idx = {h.strip(): i for i, h in enumerate(hdr_rep)}
    # 列映射
    col_map = {c: rep_idx[h] for h, c in headers.items() if h in rep_idx}
    col_names = {c: h for h, c in headers.items()}
    first_row = hdr + 1
    # 公式列（辅助列）：数据首行是公式 → 新行平移（含数组公式 ArrayFormula）
    formula_cols = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(row=first_row, column=c).value
        if isinstance(v, ArrayFormula):
            formula_cols[c] = v.text
        elif isinstance(v, str) and v.startswith("="):
            formula_cols[c] = v
    # 追加
    last = _last_data_row(ws, oid_col, first_row)
    r = last
    n = 0
    h_oid = key if (key and key in headers) else ("order-item-id" if "order-item-id" in headers else "amazon-order-id")
    oi = rep_idx.get(h_oid)
    if oi is None:
        print(f"    [销售] {ws.title}: 报告缺去重键列 {h_oid!r}，跳过")
        return None
    # 窗口下界过滤：月度报告按"最后更新日期"导出，会混入 purchase-date 更老的行
    # （2024/2025 退款散单、上月窗口外的单）——挡在窗口外，保持当月表【上月7天+当月】口径
    pd_idx = rep_idx.get("purchase-date")
    win_lo = getattr(C, "WINDOW_LO", None)
    n_blocked = 0
    if pd_idx is not None and win_lo:
        kept = []
        for row in body:
            d = row[pd_idx].strip()[:10] if len(row) > pd_idx else ""
            if d and d < win_lo:
                n_blocked += 1
                continue
            kept.append(row)
        body = kept
    # 渠道校验（2026-10-16 拒收别家行；2026-10-19 加严：整份错位整份拒收）
    ch_idx = rep_idx.get("sales-channel")
    body, cstat = screen_channels(body, ch_idx, expect_channel, book,
                                  allow_mixed=allow_mixed)
    if cstat["mislabel"]:
        br = "、".join(f"{k}×{v}" for k, v in sorted(cstat["foreign_cnt"].items(), key=lambda x: -x[1]))
        print(f"    [销售] ✗✗ {ws.title}: 报告 {os.path.basename(report_path)} **整份错位**"
              f"（{cstat['foreign']}/{cstat['total']} 行渠道属于其它簿：{br}）")
        print(f"    [销售] ✗✗ {ws.title}: 已整份拒收、本簿今日销售不更新 —— 请重下该站点报告后重跑；"
              f"错位报告里的本店订单不会丢（订单号去重，修正报告里都有）")
        return dict(sheet=ws.title, total=cstat["total"], appended=0,
                    blocked=n_blocked, foreign=cstat["foreign"],
                    unknown=cstat["unknown"], mislabel=True,
                    report=os.path.basename(report_path))
    # SKU 归属闸：渠道合法但账号不对（同名渠道不同卖家）—— 见 screen_sku_overlap
    body, sstat = screen_sku_overlap(body, rep_idx.get("sku"), stock_skus, book, report_path)
    if sstat:
        return dict(sheet=ws.title, total=sstat["total"], appended=0,
                    blocked=n_blocked, foreign=0, unknown=0, mislabel=True,
                    sku_rate=sstat["rate"], report=os.path.basename(report_path))
    ao_ir = rep_idx.get("amazon-order-id")
    sku_ir = rep_idx.get("sku")
    use_comp_rep = use_comp and ao_ir is not None and sku_ir is not None and pd_idx is not None
    new_rows = []
    for row in body:
        if not (len(row) > oi and row[oi].strip()):
            continue
        if use_comp_rep:
            ck = _comp(row[oi],
                       row[ao_ir] if ao_ir < len(row) else None,
                       row[sku_ir] if sku_ir < len(row) else None,
                       row[pd_idx] if pd_idx < len(row) else None)
            if ck in have:
                continue
        elif row[oi].strip() in have:
            continue
        new_rows.append(row)
    for row in new_rows:
        r += 1
        for col, ci in col_map.items():
            v = row[ci].strip() if ci < len(row) else ""
            if v and (col_names.get(col) or "").strip().lower() in TEXT_COLS:
                ws.cell(row=r, column=col).value = v        # ID 列保持文本
            else:
                ws.cell(row=r, column=col).value = _num(v) if v else None
        for col, f0 in formula_cols.items():
            L = get_column_letter(col)
            ws.cell(row=r, column=col).value = Translator(f0, origin=f"{L}{first_row}").translate_formula(f"{L}{r}")
        n += 1
    # autoFilter 扩展
    af = ws.auto_filter.ref
    if af:
        m = re.match(r"([A-Z]+)(\d+):([A-Z]+)(\d+)", af)
        if m:
            ws.auto_filter.ref = f"{m.group(1)}{m.group(2)}:{m.group(3)}{r}"
    blk = f" 窗口外过滤{n_blocked}行" if n_blocked else ""
    frn = f" 渠道外拒收{cstat['foreign']}行" if cstat["foreign"] else ""
    print(f"    [销售] {ws.title}: 报告{len(body)}行{blk}{frn} 追加{n}行 (r{last}→r{r})")
    if cstat["foreign"]:
        br = "、".join(f"{k}×{v}" for k, v in sorted(cstat["foreign_cnt"].items(), key=lambda x: -x[1]))
        if cstat.get("mixed_allowed"):
            # 混渠道是设计意图（interflow 沙特/阿联酋共用一份下载），不是导错站点，
            # 用普通提示而非 !! 告警，避免每天刷一条假红灯。
            print(f"    [销售] {ws.title}: 跳过 {cstat['foreign']} 行其它渠道({br})"
                  f" —— 混渠道报告，按 expect_channel 逐行切分")
        else:
            print(f"    [销售] !! {ws.title}: 拒收 {cstat['foreign']} 行「渠道属于其它簿」({br})"
                  f" —— 疑似源报告错位，请核对下载的站点")
    if cstat["unknown"]:
        br = "、".join(f"{k}×{v}" for k, v in sorted(cstat["unknown_cnt"].items(), key=lambda x: -x[1]))
        print(f"    [销售] ?? {ws.title}: {cstat['unknown']} 行渠道未登记（已保留）({br})"
              f" —— 若是新站点请补进 config 的 expect_channel")
    return dict(sheet=ws.title, total=len(body), appended=n, blocked=n_blocked,
                foreign=cstat["foreign"], unknown=cstat["unknown"], mislabel=False,
                report=os.path.basename(report_path))


# ---------------- 保真回填 ----------------
# xl/metadata.xml + xl/richData/*：富数据（Excel/WPS 的"数据类型"，如股票/错误型富值）。
# openpyxl 不认识这些部件，存盘时会连同单元格上的 vm=/cm= 索引一起整个丢掉 → 必须按原件补回
# （见 restore_richdata）。无富数据的簿（日本/欧盟英国/罗马尼亚/美客多）原件就没有，不受影响。
_RESTORE_PREFIX = ("customXml/", "docProps/custom.xml", "docProps/thumbnail", "xl/printerSettings/",
                   "xl/media/", "xl/metadata.xml", "xl/richData/")
# 透视表/透视缓存：openpyxl 会把它整个建模再回写（欧盟簿 pivotCacheRecords 11MB/20303 条，
# 实测 10.2s），而且重新序列化不如原字节保真 → 保存后一律用原字节覆盖同名部件。
_PIVOT_PREFIX = ("xl/pivotCache/", "xl/pivotTables/")
_PIVOT_RECORDS_RE = re.compile(r"^xl/pivotCache/pivotCacheRecords\d+\.xml$")
_EMPTY_RECORDS = (b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
                  b'<pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0"/>')


def strip_pivot_records(path, dst):
    """把 pivotCacheRecords*.xml 换成空壳再交给 openpyxl 加载。
    部件名 / rels / [Content_Types] 全部不动，所以 openpyxl 仍会写出同名部件；
    保存后由 restore_parts 用原字节覆盖回去。无透视件的簿返回 0，不走这条分支。"""
    with zipfile.ZipFile(path) as zo:
        infos = zo.infolist()
        blobs = {i.filename: zo.read(i.filename) for i in infos}
    hits = [i.filename for i in infos if _PIVOT_RECORDS_RE.match(i.filename)]
    if not hits or os.environ.get("DSR_NO_PIVOT_STRIP"):
        return 0
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zn:
        for i in infos:
            zn.writestr(i, _EMPTY_RECORDS if i.filename in hits else blobs[i.filename])
    return len(hits)


def _ct_patch(ct, parts):
    b = ct.decode("utf-8")
    names = set(parts)
    # 透视部件：openpyxl 若整个丢了某个 pivot 部件，补回时要保证 [Content_Types] 有对应 Override
    for prefix, ctype in (
        ("xl/pivotCache/pivotCacheDefinition", "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml"),
        ("xl/pivotCache/pivotCacheRecords", "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml"),
        ("xl/pivotTables/pivotTable", "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml"),
    ):
        for n in names:
            if n.startswith(prefix) and f'PartName="/{n}"' not in b:
                b = b.replace("</Types>",
                              f'<Override PartName="/{n}" ContentType="{ctype}"/></Types>')
    # 富数据部件（openpyxl 存盘会整个丢，见 restore_richdata）→ 补回时必须补 CT，否则 Excel 报"文件损坏"
    for part, ctype in (
        ("xl/metadata.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"),
        ("xl/richData/rdrichvalue.xml", "application/vnd.ms-excel.rdrichvalue+xml"),
        ("xl/richData/rdrichvaluestructure.xml", "application/vnd.ms-excel.rdrichvaluestructure+xml"),
    ):
        if part in names and f'PartName="/{part}"' not in b:
            b = b.replace("</Types>", f'<Override PartName="/{part}" ContentType="{ctype}"/></Types>')
    if any(n.endswith(".vml") for n in names):
        if 'Extension="vml"' not in b:
            b = b.replace("<Types ", "<Types ", 1).replace(
                "</Types>",
                '<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/></Types>')
    if any(n.endswith(".bin") for n in names):
        if 'Extension="bin"' not in b:
            b = b.replace(
                "</Types>",
                '<Default Extension="bin" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings"/></Types>')
    if "docProps/custom.xml" in names and "/docProps/custom.xml" not in b:
        b = b.replace(
            "</Types>",
            '<Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/></Types>')
    for i in range(1, 12):
        if f"customXml/itemProps{i}.xml" in names and f"/customXml/itemProps{i}.xml" not in b:
            b = b.replace(
                "</Types>",
                f'<Override PartName="/customXml/itemProps{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.customXmlProperties+xml"/></Types>')
    return b.encode("utf-8")


def _rels_patch(rl, parts):
    b = rl.decode("utf-8")
    add = []
    if "docProps/custom.xml" in parts and "docProps/custom.xml" not in b:
        add.append('<Relationship Id="rIdCustom999" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties" Target="docProps/custom.xml"/>')
    if any(n.startswith("docProps/thumbnail") for n in parts) and "thumbnail" not in b:
        ext = "jpeg" if any(n.endswith(".jpeg") for n in parts) else "wmf"
        add.append(f'<Relationship Id="rIdThumb999" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail" Target="docProps/thumbnail.{ext}"/>')
    if add:
        b = b.replace("</Relationships>", "".join(add) + "</Relationships>")
    return b.encode("utf-8")


def _wb_rels_patch(rl, parts):
    """xl/_rels/workbook.xml.rels 里补回富数据的三条关系（sheetMetadata / rdRichValue /
    rdRichValueStructure）。openpyxl 存盘会重写这个 rels 且把这三条丢掉。"""
    b = rl.decode("utf-8")
    add = []
    if "xl/metadata.xml" in parts and "relationships/sheetMetadata" not in b:
        add.append('<Relationship Id="rIdMeta999" Type="http://schemas.openxmlformats.org/'
                   'officeDocument/2006/relationships/sheetMetadata" Target="metadata.xml"/>')
    if "xl/richData/rdrichvalue.xml" in parts and "relationships/rdRichValue" not in b:
        add.append('<Relationship Id="rIdRv999" Type="http://schemas.microsoft.com/office/2017/06/'
                   'relationships/rdRichValue" Target="richData/rdrichvalue.xml"/>')
    if "xl/richData/rdrichvaluestructure.xml" in parts and "relationships/rdRichValueStructure" not in b:
        add.append('<Relationship Id="rIdRvs999" Type="http://schemas.microsoft.com/office/2017/06/'
                   'relationships/rdRichValueStructure" Target="richData/rdrichvaluestructure.xml"/>')
    if add:
        b = b.replace("</Relationships>", "".join(add) + "</Relationships>")
    return b.encode("utf-8")


# ---- 富数据：单元格上的 vm=/cm= 索引（openpyxl 保存时丢失，需按原件原位补回）----
_XML_ATTR_RE = re.compile(r'(\w[\w:.-]*)="([^"]*)"')


def _sheet_name_map(wbx_b, rels_b):
    """由 workbook.xml + 其 rels 得到 {sheet 名: 部件路径(如 xl/worksheets/sheet5.xml)}。"""
    rels = {}
    for tag in re.findall(r"<Relationship\b[^>]*/?>", rels_b.decode("utf-8", "ignore")):
        a = dict(_XML_ATTR_RE.findall(tag))
        if a.get("Id"):
            rels[a["Id"]] = a.get("Target", "")
    out = {}
    for m in re.finditer(r"<sheet\b[^>]*/?>", wbx_b.decode("utf-8", "ignore")):
        a = dict(_XML_ATTR_RE.findall(m.group(0)))
        nm, rid = a.get("name"), a.get("r:id")
        if nm is None or rid is None:
            continue
        t = rels.get(rid, "").lstrip("/")
        if not t:
            continue
        out[nm] = t if t.startswith("xl/") else "xl/" + t
    return out


def _collect_vmcm(orig_path):
    """{sheet 名: {单元格: 'vm="1" cm="2"'}} —— 原件里带富数据索引的格（通常只有极少数几个）。"""
    out = {}
    with zipfile.ZipFile(orig_path) as zo:
        try:
            sm = _sheet_name_map(zo.read("xl/workbook.xml"), zo.read("xl/_rels/workbook.xml.rels"))
        except KeyError:
            return out
        for nm, path in sm.items():
            try:
                raw = zo.read(path)
            except KeyError:
                continue
            if b'vm="' not in raw and b'cm="' not in raw:   # 便宜预筛，避免无谓解码
                continue
            d = {}
            for m in re.finditer(r"<c\b[^>]*/?>", raw.decode("utf-8", "ignore")):
                a = dict(_XML_ATTR_RE.findall(m.group(0)))
                if "r" in a:
                    got = [f'{k}="{a[k]}"' for k in ("cm", "vm") if k in a]
                    if got:
                        d[a["r"]] = " ".join(got)
            if d:
                out[nm] = d
    return out


def _inject_vmcm(xml, cellmap):
    """把 vm=/cm= 补回对应格。安全闸：①目标格尚未带 vm/cm（幂等）；②目标格仍是公式格
    （紧跟 `<f`）——防止万一发生删行导致引用错位时，把普通值格误标成富数据。"""
    for ref, attrs in cellmap.items():
        pat = re.compile(r'<c r="%s"((?![^>]*\b(?:vm|cm)=)[^>]*)>(<f\b)' % re.escape(ref))
        xml = pat.sub(lambda m: '<c r="%s" %s%s>%s' % (ref, attrs, m.group(1), m.group(2)), xml, count=1)
    return xml


def restore_richdata(orig_path, new_path):
    """把原件里的 xl/metadata.xml + xl/richData/* 与单元格 vm=/cm= 索引按原字节补回 new_path。
    幂等：new_path 已有同名部件时不重复补；无富数据的簿为 no-op。返回补回的部件数。
    运行于 restore_parts 内部；也可单独调用（如手工回填历史簿）。"""
    try:
        vmmap = _collect_vmcm(orig_path)
    except Exception as e:
        print(f"    [保真] 富数据解析跳过: {e!r}")
        return 0
    with zipfile.ZipFile(orig_path) as zo:
        orig_names = set(zo.namelist())
        parts = [n for n in orig_names
                 if n == "xl/metadata.xml" or n.startswith("xl/richData/")]
        blobs = {n: zo.read(n) for n in parts}
    if not parts and not vmmap:
        return 0
    with zipfile.ZipFile(new_path) as zn:
        new_names = set(zn.namelist())
        data = {i.filename: zn.read(i.filename) for i in zn.infolist()}
        order = [i.filename for i in zn.infolist()]
    added = [n for n in parts if n not in new_names]
    for n in added:
        data[n] = blobs[n]
        order.append(n)
    # vm/cm 原位回填（按 sheet 名匹配新文件的部件路径）
    reinjected = 0
    if vmmap and "xl/workbook.xml" in data and "xl/_rels/workbook.xml.rels" in data:
        nmap = _sheet_name_map(data["xl/workbook.xml"], data["xl/_rels/workbook.xml.rels"])
        for nm, cellmap in vmmap.items():
            part = nmap.get(nm)
            if part and part in data:
                newx = _inject_vmcm(data[part].decode("utf-8", "ignore"), cellmap)
                if newx != data[part].decode("utf-8", "ignore"):
                    data[part] = newx.encode("utf-8")
                    reinjected += 1
    if added:
        present = set(added) | set(parts)
        data["[Content_Types].xml"] = _ct_patch(data["[Content_Types].xml"], present)
        if "xl/_rels/workbook.xml.rels" in data:
            data["xl/_rels/workbook.xml.rels"] = _wb_rels_patch(data["xl/_rels/workbook.xml.rels"], present)
    if not added and not reinjected:
        return 0
    tmp = new_path + ".rich.tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zo2:
        zo2.writestr("[Content_Types].xml", data["[Content_Types].xml"])
        for n in order:
            if n != "[Content_Types].xml":
                zo2.writestr(n, data[n])
    shutil.move(tmp, new_path)
    if added or reinjected:
        print(f"    [保真] 富数据补回部件{len(added)}个 vm/cm索引{reinjected}个sheet")
    return len(added)


def restore_parts(orig_path, new_path):
    with zipfile.ZipFile(orig_path) as zo:
        orig_names = set(zo.namelist())
        lost = [n for n in orig_names if n.startswith(_RESTORE_PREFIX)]
        blobs = {n: zo.read(n) for n in lost}
        # 透视部件也要读进来，供下面同名覆盖
        for n in orig_names:
            if n.startswith(_PIVOT_PREFIX) and n not in blobs:
                blobs[n] = zo.read(n)
    # 富数据：原件里带 vm=/cm= 索引的格（openpyxl 存盘必丢，需原位补回）
    vmmap = _collect_vmcm(orig_path)
    with zipfile.ZipFile(new_path) as zn:
        new_names = set(zn.namelist())
        data = {i.filename: zn.read(i.filename) for i in zn.infolist()}
        order = [i.filename for i in zn.infolist()]
    restored = [n for n in lost if n not in new_names]
    for n in restored:
        data[n] = blobs[n]
        order.append(n)
    # vm/cm 原位回填（按 sheet 名匹配新文件的部件路径；只在尚未带该属性的格上补）
    rich_sheets = 0
    if vmmap and "xl/workbook.xml" in data and "xl/_rels/workbook.xml.rels" in data:
        nmap = _sheet_name_map(data["xl/workbook.xml"], data["xl/_rels/workbook.xml.rels"])
        for nm, cellmap in vmmap.items():
            part = nmap.get(nm)
            if not part or part not in data:
                continue
            old = data[part].decode("utf-8", "ignore")
            newx = _inject_vmcm(old, cellmap)
            if newx != old:
                data[part] = newx.encode("utf-8")
                rich_sheets += 1
    if rich_sheets:
        print(f"    [保真] 富数据 vm/cm 索引原位回填 {rich_sheets} 个 sheet")
    # 透视表/透视缓存：同名覆盖为原字节（openpyxl 刚写出的是空壳记录或重新序列化结果）。
    # 只处理新文件里已存在的同名部件；若 openpyxl 整个丢了某个部件，仍按原字节补回并补 CT。
    forced = []
    for n in orig_names:
        if not n.startswith(_PIVOT_PREFIX):
            continue
        if data.get(n) == blobs[n]:
            continue
        data[n] = blobs[n]
        if n not in order:
            order.append(n)
            forced.append(n)
    if restored or forced or rich_sheets:
        present = set(restored) | set(forced)
        data["[Content_Types].xml"] = _ct_patch(data["[Content_Types].xml"], present)
        data["_rels/.rels"] = _rels_patch(data["_rels/.rels"], present)
        # 富数据补回时 workbook.xml.rels 也要补 sheetMetadata/rdRichValue 三条关系
        if "xl/_rels/workbook.xml.rels" in data:
            data["xl/_rels/workbook.xml.rels"] = _wb_rels_patch(data["xl/_rels/workbook.xml.rels"], present)
    # pivot refreshOnLoad
    piv = 0
    for n in list(data):
        if "pivotCacheDefinition" in n and n.endswith(".xml"):
            x = data[n].decode("utf-8")
            if 'refreshOnLoad="1"' not in x:
                if "refreshOnLoad=" in x:
                    x = x.replace('refreshOnLoad="0"', 'refreshOnLoad="1"')
                else:
                    x = re.sub(r"(<pivotCacheDefinition\b)", r'\1 refreshOnLoad="1"', x, count=1)
                data[n] = x.encode("utf-8")
                piv += 1
    tmp = new_path + ".rs.tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zo2:
        zo2.writestr("[Content_Types].xml", data["[Content_Types].xml"])
        for n in order:
            if n != "[Content_Types].xml":
                zo2.writestr(n, data[n])
    shutil.move(tmp, new_path)
    return restored, piv


# ---------------- 主流程 ----------------
def backup(files, snap_day="20261006"):
    bdir = os.path.join(C.DSR, ".workbuddy", f"backup_daily_{snap_day}")
    os.makedirs(bdir, exist_ok=True)
    for f in files:
        src = os.path.join(C.WORK_DIR, f)
        dst = os.path.join(bdir, f)
        if not os.path.exists(dst):
            shutil.copy2(src, dst)
    print(f"备份 → {bdir}")


def _prev_month_daily(fname, pm_y, pm_m, src="每日销售数据"):
    """上月归档簿的销量源表（read_only，公式句柄——重算用原始列，不依赖公式缓存）。
    封存上月必须用归档簿：当前簿在换月时会裁掉上月早期订单行（如 10月版删到 9-24 起），
    直接用当前簿会把上月封存成「月尾几天」的残值。
    找不到归档/源表时返回 (None, None)，调用方回落当前簿并打警告。
    """
    arch = os.path.join(C.WORK_DIR, f"{pm_y}年每日数据", f"{pm_y}年{pm_m}月",
                        f"{pm_y}年{pm_m}月-{fname}")
    if not fname or not os.path.exists(arch):
        return None, None
    wb_a = load_workbook(arch, read_only=True, data_only=False)
    if src not in wb_a.sheetnames:
        wb_a.close()
        return None, None
    return wb_a, wb_a[src]


def _site_append_rows(wb, snap_day):
    """站点表补行（2026-10-27 用户拍板：「销量是最核心的东西，要保证准确」）。

    背景：G598（Amazon.it 10月 1 件）/ K001（Amazon.se 10月 1 件）在站点表里没有行，
          站点表合计比分析表少 2 件 —— 补行闸只作用于分析表，站点表没人管。

    证据闸：本站本月销量 > 0（month_sales 按站点渠道过滤，amzn.gr 已归并），
    不看库存 / 首现日期 —— 有真实成交就证明该进表。行追加复用
    append_sku_rows_auto（克隆列最完整行、公式平移、CF 范围跟随，站点表
    少「库存」列的左移布局也兼容，2026-10-27 手工补行已验证）。

    必须在 apply_month_summary **之前**调用：新行落位后封存当期就把月度值写上。
    开关 config.SITE_APPEND_ENABLED=False 完整回退旧行为。
    返回补行总数。
    """
    if not getattr(C, "SITE_APPEND_ENABLED", False):
        return 0
    if "每日销售数据" not in wb.sheetnames:
        return 0
    import analysis as An
    ch_map = getattr(C, "SITE_SHEET_CHANNELS", {}) or {}
    ym = f"{snap_day[:4]}-{snap_day[4:6]}"
    n_total = 0
    for sn in list(wb.sheetnames):
        ch = ch_map.get(str(sn).strip())
        if not ch:
            continue
        try:
            sold = An.month_sales(wb["每日销售数据"], ym, channel=ch) or {}
        except Exception as e:
            print(f"    [补行·站点] {sn} 月销汇总失败：{e}")
            continue
        sold = {k: v for k, v in sold.items() if v and v > 0}
        if not sold:
            continue
        ws = wb[sn]
        have = An.sku_set(ws)
        need = [k for k in sold if k not in have and An._SKU_RE.match(k)]
        if not need:
            continue
        # 品名：优先「欧洲-每日库存分析表」B 列，回退任意分析表
        names = {}
        order = ([s for s in wb.sheetnames if str(s).strip().startswith("欧洲")]
                 + [s for s in wb.sheetnames if "分析表" in s])
        for asn in order:
            aws = wb[asn]
            for r in range(7, min((aws.max_row or 7) + 1, 5000)):
                k = aws.cell(r, 1).value
                if k is None:
                    break
                k = str(k).strip()
                if k in need and k not in names:
                    names[k] = aws.cell(r, 2).value
            if all(k in names for k in need):
                break
        added = An.append_sku_rows_auto(ws, need,
                                        name_map={k: names.get(k, "") for k in need},
                                        state_map={k: "在售" for k in need})
        if added:
            print(f"    [补行·站点] {sn}: {len(added)} 行 ← 本站({ch})本月有销量但表里没有 → "
                  + "、".join(f"行{r}·{k}" for r, k in added[:6])
                  + (f" 等 {len(added)} 个" if len(added) > 6 else ""))
            n_total += len(added)
    return n_total


def _month_summary_sheets(wb, An):
    """月度封存要处理的 sheet 列表。

    ① 名字含「分析」的（英国/欧洲-每日库存分析表 …）—— 原有行为，永远处理。
    ② config.SITE_MONTHLY_FREEZE 为 True 时，追加**站点表**（意大利/法国/…）：
       判据是 scan_analysis 能解出 SUMIFS 口径、源表存在、且有满 12 个月标签列。
       站点表行5 的日列头是公式链（=AB5+1），2026-10-27 起 scan_analysis 走
       「行7 最右 SUMIFS」兜底，因此能解出与主表一致的 spec。
       不满足判据的 sheet 自动跳过，不需要逐簿配置。
    """
    out = [s for s in wb.sheetnames if "分析" in s]
    if not getattr(C, "SITE_MONTHLY_FREEZE", True):
        return out
    for sn in wb.sheetnames:
        if sn in out:
            continue
        try:
            sc = An.scan_analysis(wb[sn])
        except Exception:
            continue
        if not sc:
            continue
        if sc["spec"].get("src") not in wb.sheetnames:
            continue
        if len(sc["month_cols"]) < 12:
            continue
        out.append(sn)
    return out


_MSX_CACHE = {}   # (period, site) -> {sku: qty} | None（None=查过但没有，避免重复打开）


def _month_summary_from_msx(pm_y, pm_m, site):
    """月销总表「{Y}年」sheet 的 site 站点 {m}月 逐 SKU 值（任务㉒，2026-11-03）。

    月销总表 = 公司上报口径（站点**本地月**归月， Europeans 组 London 时区），
    日销明细 = UTC 月。换月封存后两者差跨月单（单月几十件）+33 件漏记补录这类
    事后修正——resolve 重算永远对不齐上报数。用户拍板（11-03「都修」）：
    分析表上月封存列**优先取月销总表**，对不齐的问题在源头（月销表）修。

    路径：{PROJECT}/trunk/result/数据归属周期/{YYYYMM}/月销售总表.xlsx
    结构：行1 块标签（C=欧洲/O=英国/AA=美国/AM=加拿大/AY=日本/BK=澳洲/BW=阿联酋/
          CI=沙特/CU=墨西哥/DG=印度），行2 块内 1~12月，行4+ A=SKU。
    找不到簿/块/月列返回 None，调用方回落 resolve 重算（老行为）。
    结果带模块级缓存（13 簿 × 2 分析表每天会读 26 次）。
    """
    key = (pm_y, pm_m, site)
    if key in _MSX_CACHE:
        return _MSX_CACHE[key]
    out = None
    period = "%04d%02d" % (pm_y, pm_m)
    path = os.path.join(C.PROJECT, "trunk", "result", "数据归属周期", period,
                        "月销售总表.xlsx")
    wb = None
    try:
        if os.path.exists(path):
            import analysis as _An
            wb = load_workbook(path, read_only=True, data_only=True)
            sh = "%s年" % pm_y
            if sh in wb.sheetnames:
                ws = wb[sh]
                it = ws.iter_rows(min_row=1, max_row=2, max_col=200, values_only=True)
                r1 = next(it, None)
                r2 = next(it, None)
                if r1 and r2:
                    start = next((i for i, v in enumerate(r1)
                                  if v is not None and str(v).strip() == site), None)
                    if start is not None:
                        want = "%d月" % pm_m
                        off = next((j for j in range(start, min(start + 12, len(r2)))
                                    if r2[j] is not None and str(r2[j]).strip() == want),
                                   None)
                        if off is not None:
                            out = {}
                            for row in ws.iter_rows(min_row=4, max_col=off + 1,
                                                    values_only=True):
                                sku = row[0]
                                if sku is None or not str(sku).strip():
                                    continue
                                v = row[off] if off < len(row) else None
                                if isinstance(v, (int, float)) and v:
                                    k = _An.normalize_sku(sku)
                                    out[k] = out.get(k, 0) + v
    except Exception as e:
        print(f"    [月销总表] {period}/{site} 读取失败，回落重算：%r" % (e,))
        out = None
    finally:
        if wb is not None:
            wb.close()
    _MSX_CACHE[key] = out
    return out


# 任务㉒：能直取月销总表的站点白名单 —— 只限「一个站点只有一本簿」的站点。
# 北美（美国/加拿大/墨西哥）、澳洲、中东（沙特/阿联酋）各有 Interflowd 孪生簿
# （10-30 备忘：主簿+孪生簿跨簿相加=月销表口径，实测澳洲 1656+276=1932 分毫不差）。
# 这些站点若单独写月销总表值会与孪生簿重复计数，必须保持归档簿 resolve。
MSX_DIRECT_SITES = {"欧洲", "英国", "日本", "印度"}


def apply_month_summary(wb, snap_day, fname=""):
    """全簿通用：月汇总区「上月封存 + 当月列滚动」。
    口径自动提取自各分析表日销列 SUMIFS 公式（analysis.scan_analysis / resolve_month_sales），
    与表内口径永远一致，无需逐簿配置。幂等：freeze_month 覆盖写 + 清残留。
    上月源（任务㉒，2026-11-03）：**分析表**优先=月销总表（上报口径逐 SKU 值），
    月销总表缺该站点时回落归档簿重算（老行为）；站点表仍=归档簿重算。
    当月源=当前簿。
    """
    import analysis as An

    y, m = int(snap_day[:4]), int(snap_day[4:6])
    pm_y, pm_m = (y - 1, 12) if m == 1 else (y, m - 1)
    prev_ym, cur_ym = f"{pm_y:04d}-{pm_m:02d}", f"{y:04d}-{m:02d}"

    for sn in _month_summary_sheets(wb, An):
        ws = wb[sn]
        try:
            sc = An.scan_analysis(ws)
        except Exception as e:
            print(f"    [月汇总] {sn}: !! 扫描失败 {e!r}")
            continue
        if not sc:
            print(f"    [月汇总] {sn}: !! 无法解析日销 SUMIFS 口径，跳过")
            continue
        spec, d4, d5 = sc["spec"], sc["d4"], sc["d5"]
        mc = sc["month_cols"]
        src = spec["src"]
        if src not in wb.sheetnames:
            print(f"    [月汇总] {sn}: !! 源表 {src} 不存在，跳过")
            continue
        daily = wb[src]
        awb, adaily = _prev_month_daily(fname, pm_y, pm_m, src=src)
        try:
            # 任务㉒：分析表上月源优先月销总表（上报口径）。站点名=sheet 名去「-每日库存分析表」；
            # 月销总表不存在（期初未生成/站点无块）时 _month_summary_from_msx 返回 None 自动回落重算
            _site = sn.split("-")[0].strip() if "分析" in sn else None
            if _site:
                _site = {"澳大利亚": "澳洲", "沙特阿拉伯": "沙特"}.get(_site, _site)
            _ms_msx = (_month_summary_from_msx(pm_y, pm_m, _site)
                       if _site in getattr(C, "MSX_DIRECT_SITES", MSX_DIRECT_SITES) else None)
            for ym, mm in ((prev_ym, pm_m), (cur_ym, m)):
                col = mc.get(mm)
                if col is None:
                    print(f"    [月汇总] {sn}: !! 行5 无「{mm}月」标签列，跳过")
                    continue
                ms = tag = None
                if ym == prev_ym and _ms_msx is not None:
                    ms, tag = _ms_msx, "月销总表"          # ===== 任务㉒：上报口径直取 =====
                if ms is None:
                    wsrc = daily if ym == cur_ym else (adaily or daily)
                    try:
                        ms = An.resolve_month_sales(wsrc, ym, spec, d4, d5)
                    except Exception as e:
                        print(f"    [月汇总] {sn}: {ym} !! 重算失败 {e!r}")
                        continue
                    tag = ("归档簿" if adaily is not None
                           else "当前簿(缺归档,上月可能少算!)") if ym == prev_ym else "当前簿"
                nf = An.freeze_month(ws, col, ms)
                if ym == prev_ym:
                    # 上月日均同步刷新为活公式：=IF(封存列="","",ROUND(封存列/天数,2))，
                    # 与封存列永远联动（2026-10-15 起，原静态值 round 1 会把月销1件抹成0）
                    # 站点表（非「分析」表）受 config.SITE_MONTHLY_FILL_AVG 控制：
                    # 它们的 G 列就是「上月日均」且是染色基准，默认一并刷新。
                    _is_ana = "分析" in sn
                    _do_avg = _is_ana or getattr(C, "SITE_MONTHLY_FILL_AVG", True)
                    import calendar as _cal
                    days = _cal.monthrange(pm_y, pm_m)[1]
                    if _do_avg:
                        avg = {k: (v / days) for k, v in ms.items()}
                        acol, na = An.fill_prev_avg(ws, avg, month_col=col, days=days)
                        if acol:
                            from openpyxl.utils import get_column_letter as _gl
                            print(f"    [上月日均] {sn}: {_gl(acol)} 列 {na} 格"
                                  f"（{prev_ym} ÷ {days} 天，公式联动封存列）")
                else:
                    tag = "当前簿"
                print(f"    [月汇总] {sn}: {ym}（{mm}月列）{nf} 格 / {sum(ms.values()):.0f} 件（源={tag}）")
        finally:
            if awb is not None:
                awb.close()


# 分析表站点 → (sales-channel, region)。region 只给「欧洲」用（month_sales 的 region 模式）。
# 不在表里的（罗马尼亚-EMAG / 美客多墨西哥）不走亚马逊销量口径，销量补行时跳过。
_SITE_CHANNEL = {
    "英国": ("Amazon.co.uk", None),
    "欧洲": (None, "欧洲"),
    "美国": ("Amazon.com", None),
    "加拿大": ("Amazon.ca", None),
    "墨西哥": ("Amazon.com.mx", None),
    "日本": ("Amazon.co.jp", None),
    "印度": ("Amazon.in", None),
    "沙特": ("Amazon.sa", None),
    "沙特阿拉伯": ("Amazon.sa", None),
    "澳大利亚": ("Amazon.com.au", None),
    "阿联酋": ("Amazon.ae", None),
}


def _channel_of_sheet(sheet_name):
    """分析表名 → (channel, region)。认不出返回 (None, None)。"""
    site = str(sheet_name or "").strip().split("-")[0].strip()
    return _SITE_CHANNEL.get(site, (None, None))


def _first_seen_map():
    """读「断货分析/型号首库存登记.xlsx」→ {sku: 首现日期 'YYYYMMDD'}。

    文件不存在返回 None，调用方据此跳过新品补行（不会静默按老逻辑补全量）。
    表由 tools/first_stock_register.py --apply 生成，跑批前刷新一次即可。
    """
    p = os.path.join(C.DSR, "断货分析", "型号首库存登记.xlsx")
    if not os.path.exists(p):
        return None
    try:
        wb = load_workbook(p, data_only=True, read_only=True)
        ws = wb["首库存登记"]
        out = {}
        for r in ws.iter_rows(min_row=2, values_only=True):
            if r and r[0] and r[2]:
                out[str(r[0]).strip()] = str(r[2]).strip()
        wb.close()
        return out
    except Exception as e:
        print(f"    [补行] 读型号首库存登记失败：{e}")
        return None


def _sales_cache_per_sheet(wb, fname, snap_day):
    """{分析表名: {sku: 本月销量}}。一次算好，补行证据闸和销量补行段共用。

    口径与分析表所属站点严格对应（英国=Amazon.co.uk / 欧洲=region / 其余=各自 channel），
    避免法国站的销量被当成英国站的证据。美客多 / EMAG 的「订单数据」不是亚马逊口径，
    没有 expect_channel → 不参与，返回空 dict（这两本靠库存驱动补行）。
    """
    import analysis as An
    cfg = next((b for b in C.WORKBOOKS if b.get("file") == fname), None)
    specs = (cfg or {}).get("sales")
    specs = specs if isinstance(specs, list) else ([specs] if specs else [])
    dsn = next((s.get("sheet") for s in specs
                if s.get("sheet") in wb.sheetnames and s.get("expect_channel")), None)
    if not dsn:
        return {}
    ym = f"{snap_day[:4]}-{snap_day[4:6]}"
    out = {}
    for sn in wb.sheetnames:
        if "分析表" not in sn:
            continue
        ch, reg = _channel_of_sheet(sn)
        if ch is None and reg is None:
            continue
        try:
            out[sn.strip()] = An.month_sales(wb[dsn], ym, channel=ch, region=reg) or {}
        except Exception as e:
            print(f"    [销量] {sn} 月销汇总失败：{e}")
    return out


def _stock_map_per_sheet(wb):
    """{分析表名: {sku: {name, avail, inb, listing}}}。

    每个分析表只读自己的库存源（An.source_sheets_of 严格 1:1 映射），
    所以英国表的「在途」不会串到欧洲表。三态判定与补行都用这一份快照，
    避免同一次跑批里同一 SKU 在不同段落到不同结论。
    """
    import analysis as An
    out = {}
    for sn in wb.sheetnames:
        if "分析表" not in sn:
            continue
        try:
            srcs = An.source_sheets_of(wb[sn], wb)
            if not srcs:
                continue
            out[sn.strip()] = An.stock_snapshot_merged(wb, srcs)
        except Exception as e:
            print(f"    [状态] {sn} 读库存源失败：{e}")
    return out


def _sync_states_and_start_dates(wb, stock_map):
    """把表内已有数据行的 D 列按库存刷成三态，并登记起售日期。

    · D 列：只有 在售/在途/未上架 三个合法值；**其他值（用户手工写的备注/颜色标记）不覆盖**。
      但为了不把用户手填的东西冲掉，只处理「当前值是这三种之一或为空」的行。
    · 起售日期：不写进 12 本表（表里没有空位，插列会打断公式/CF/透视表），
      写进「断货分析/型号首库存登记.xlsx」新增的「起售日期」列 —— 一次性登记，
      已有值不覆盖（用户口径：起售日期是**首次**有库存那天，不是最近一次）。
    返回 (改状态行数, 新登记起售日期数)。
    """
    import analysis as An

    _legal = (An.STATE_ON_SALE, An.STATE_IN_TRANSIT, An.STATE_NOT_LISTED)
    n_st = 0
    for sn in wb.sheetnames:
        if "分析表" not in sn:
            continue
        stock = stock_map.get(sn.strip()) or {}
        if not stock:
            continue
        ws = wb[sn]
        last = An._last_data_row(ws)
        for r in range(An.DATA_FIRST_ROW, last + 1):
            sku = ws.cell(r, 1).value
            if sku is None:
                continue
            sku = str(sku).strip()
            if sku not in stock:
                continue
            cur = ws.cell(r, 4).value
            cur_s = str(cur).strip() if cur is not None else ""
            if cur_s and cur_s not in _legal:
                continue                      # 用户手填的备注，不动
            new = An.tri_state(stock[sku])
            if new != cur_s:
                ws.cell(r, 4).value = new
                n_st += 1

    n_start = _record_start_dates()

    # 让粉/浅粉断货标记排除「在途」「未上架」——在途不该显示断货
    try:
        import analysis as An2
        for sn in wb.sheetnames:
            if "分析表" in sn:
                An2.sync_state_cf(wb[sn])
    except Exception:
        pass
    return n_st, n_start


def _record_start_dates():
    """把「首可售日」登记成「起售日期」（型号首库存登记.xlsx 新增 L 列）。

    用户口径：「一旦有库存那天就可以当作是这个SKU的起售日期」。
    只在 L 列**为空**时写入 —— 起售日期是历史事实，不能被后续跑批改写。
    返回新登记条数。
    """
    p = os.path.join(C.DSR, "断货分析", "型号首库存登记.xlsx")
    if not os.path.exists(p):
        return 0
    try:
        wb = load_workbook(p)
        ws = wb["首库存登记"]
        hdr = {str(ws.cell(1, c).value or "").strip(): c for c in range(1, ws.max_column + 1)}
        col_start = hdr.get("起售日期")
        if not col_start:
            col_start = ws.max_column + 1
            ws.cell(1, col_start).value = "起售日期"
            from copy import copy as _cp
            for attr in ("font", "alignment", "fill", "border"):
                try:
                    setattr(ws.cell(1, col_start), attr,
                            _cp(getattr(ws.cell(1, 1), attr)))
                except Exception:
                    pass
            try:
                ws.column_dimensions[
                    ws.cell(1, col_start).column_letter].width = 12
            except Exception:
                pass
        col_av = hdr.get("首可售日") or 4
        col_inb = hdr.get("首在途日") or 5
        n = 0
        for r in range(2, ws.max_row + 1):
            sku = ws.cell(r, 1).value
            if not sku:
                continue
            _cur = ws.cell(r, col_start).value
            _av = ws.cell(r, col_av).value
            _ib = ws.cell(r, col_inb).value
            # 校正（2026-10-24）：起售日期必须来自「首可售日」。
            # 早期版本在没有首可售日时退而用「首在途日」，但用户口径明确
            # 「有库存那天才算起售」，在途不算 → 这类值要清掉，等真有可售再写。
            if _cur and not _av and str(_cur).strip() == str(_ib or "").strip():
                ws.cell(r, col_start).value = None
                _cur = None
            if _cur:
                continue                     # 已登记，不覆盖
            if _av:
                ws.cell(r, col_start).value = str(_av).strip()
                n += 1
        if n:
            wb.save(p)
        wb.close()
        return n
    except Exception as e:
        print(f"    [起售] 登记起售日期失败：{e}")
        return 0


def apply_shipments(wb, snap_day, fname=""):
    """货件目录（含地区子文件夹）→ 本簿分析表「货件列」匹配填写。
    幂等：行1 已有货件号复用列；空列位=行1 空且 行5 空或占位公式（=TEXT(...)）；
    新列填 行1 货件号 + 行5 静态标签「M/D运输」+ 数据行；行2-4 公式自动从行5 解析。
    到货标记（行5 补「到货」）仍由人工维护。
    缺失 SKU 自动补行（2026-10-16 起**全 12 簿通用**）：货件里有、A 列没有的 SKU
    在数据区末尾按票内顺序补行（克隆末行 → 公式平移 / 文本属性照抄 / 数值留空 /
    货件列清空），再由 apply_sheet 回填该票数量。幂等，重复跑不会重复补。
    """
    import shipments as Sh
    import analysis as An

    # 「当月总销量」表头动态化 + 日列「组合」按快照日同步（全 12 簿通用，与货件有无无关）
    snap_date = f"{snap_day[:4]}-{snap_day[4:6]}-{snap_day[6:]}"
    for sn in list(wb.sheetnames):
        if "分析表" not in sn:
            continue
        ch = An.dynamic_current_headers(wb[sn])
        for c, old, new in ch:
            print(f"    [表头] {sn}: {Sh.col_letter(c)}5 {old!r} → 动态跟随锚点")
        hid, btn = An.sync_day_outline(wb[sn], snap_date)
        if hid:
            print(f"    [组合] {sn}: 折叠 {len(hid)} 个日列 "
                  f"{Sh.col_letter(min(hid))}..{Sh.col_letter(max(hid))}"
                  f"（{snap_date} 起），按钮落 {Sh.col_letter(btn)}")
        elif hid == []:
            print(f"    [组合] {sn}: 无可折叠日列")

    # ── 补行（2026-10-23 上架三态改造）─────────────────────────────────────
    # 用户拍板原话：
    #   「listing和库存的判定可以用来额外辅助，一旦有库存那天就可以当作是这个SKU的起售日期」
    #   「在途已经有了，可以标记在途，不然怎么和在途列匹配？等有库存之后就直接更新成在售」
    #
    # 旧口径（C.TRISTATE_ENABLED=False 时完整保留）：
    #   listing 活着 且 首现日期 >= 当月 1 号 → 补行，状态只有 在售/未上架。
    #   G599 这类「listing 列是空」的 SKU 就永远进不了表。
    #
    # 新口径（TRISTATE_ENABLED=True）：
    #   ① **库存说话**：可售>0 → 在售；可售=0 且在途>0 → 在途；两者都 0 且 listing 不通 → 未上架
    #   ② 补行不再要求 listing 活着 —— 只要 可售>0 或 在途>0 就补（G597/G598 立刻能进表）
    #   ③ 月份闸保留（C.TRISTATE_KEEP_MONTH_GATE）：只补本月新出现的，避免首跑灌 150+ 老行
    #   ④ 起售日期 = 首次有可售库存的那天，写进 D 列右侧的「起售日期」列（见 _ensure_start_col）
    stock_map = _stock_map_per_sheet(wb) if C.TRISTATE_ENABLED else None
    first_seen = _first_seen_map()
    month_start = snap_day[:6] + "01"
    # 销量缓存（2026-10-24）：补行证据闸要「本月有销量」这条硬证据，
    # 所以订单数据的定位与月销汇总必须**提前到补行之前**算一次，
    # 后面的「销量驱动补行」段直接复用，不重复解析。
    sold_cache = _sales_cache_per_sheet(wb, fname, snap_day)
    n_list = 0
    if first_seen is None:
        print("    [补行] 型号首库存登记.xlsx 不存在 → 跳过新品补行（先跑 first_stock_register.py --apply）")
    else:
        for _sn in list(wb.sheetnames):
            if "分析表" not in _sn:
                continue
            _ws = wb[_sn]
            _srcs = An.source_sheets_of(_ws, wb)
            if not _srcs:
                print(f"    [补行] {_sn}: 找不到库存源 sheet，跳过")
                continue
            _have = An.sku_set(_ws)
            _stock = (stock_map or {}).get(_sn.strip(), {})
            _need, _names, _states, _skip_old, _n_inb, _n_zero = [], {}, {}, 0, 0, 0

            # 候选集：旧口径看 listing 名单；新口径看库存源 + listing 名单的并集
            _cands = []
            for _sku, _nm, _alive in An.listing_skus_merged(wb, _srcs):
                _cands.append((_sku, _nm, _alive))
            if C.TRISTATE_ENABLED:
                _seen = {s for s, _, _ in _cands}
                for _sku, _rec in _stock.items():
                    if _sku not in _seen:
                        _cands.append((_sku, _rec.get("name", ""), None))

            for _sku, _nm, _alive in _cands:
                if _sku in _have or _sku in _need:
                    continue
                # ── 补行闸（2026-10-24 用户拍板）────────────────────────────
                # 原话：「正则可以打开，前提是销量对得上」
                # 旧闸：首现日期 >= 当月 1 号 → 会把「9 月首现、本月刚到货」挡在表外。
                # 新闸：本站有货 或 有在途 或 本月有销量 —— 三条硬证据任一条即可。
                #
                # ⚠️ 顺序很重要（2026-10-24 踩坑）：证据闸必须**先于**三态闸。
                # 原写法先用 tri_state 判「未上架」就 skip，导致「有销量但
                # afn-listing-exists 显示 No/空」的 SKU 永远进不来 —— 而能卖出去
                # 本身就证明 listing 是通的，那个 False 只是导出滞后。用户也明确说
                # 「listing 和库存的判定可以用来额外辅助」，不是一票否决。
                if C.TRISTATE_ENABLED:
                    _rec = _stock.get(_sku) or {}
                    _av = _rec.get("avail") or 0
                    _ib = _rec.get("inb") or 0
                    _q = (sold_cache or {}).get(_sn.strip(), {}).get(_sku, 0)
                    _st = An.tri_state_of_sku(_stock, _sku)
                    if C.TRISTATE_REQUIRE_EVIDENCE and not C.TRISTATE_KEEP_MONTH_GATE:
                        if not (_av > 0 or _ib > 0 or _q > 0):
                            # ── 零证据白名单（2026-10-26 用户拍板）──────────────────
                            # 原话：「只要识别到 G599 或者 G602 的数据的时候就可以列进来，
                            #        也方便后面货件列的比对」
                            # G599 这类 SKU：listing 已建（afn-listing-exists=Yes）、
                            # 但 可售/在途/本月销量 全 0，被证据闸挡在表外 → 货件列没法比对。
                            # 放行条件三条全满足：① 命中真编号正则（挡掉 amzn.gr.* 临时
                            # SKU）② listing 不是明确 False ③ 本站库存源里确实有这条记录。
                            if (C.ZERO_EVIDENCE_SKU_RE
                                    and re.match(C.ZERO_EVIDENCE_SKU_RE, _sku)
                                    and _st != An.STATE_NOT_LISTED
                                    and _rec):
                                _n_zero += 1      # 放行，按 在售/可售=0 补行
                            else:
                                _skip_old += 1    # 没库存没销量 → 不补（不管首现日期）
                                continue
                        if _st == An.STATE_NOT_LISTED and _q > 0:
                            # 有销量但 listing 显示不通 → 按在售（能卖就是通）
                            _st = An.STATE_ON_SALE
                    elif _st == An.STATE_NOT_LISTED:
                        _skip_old += 1
                        continue
                else:
                    if not _alive:
                        continue                      # 旧口径：listing 没活，不登记
                    _st = "在售"
                    _fs = first_seen.get(_sku, "")
                    if not _fs or _fs < month_start:
                        _skip_old += 1                # 老 SKU / 未登记 → 不追溯
                        continue
                _need.append(_sku)
                _names[_sku] = _nm or (_stock.get(_sku, {}) or {}).get("name", "")
                _states[_sku] = _st
                if _st == An.STATE_IN_TRANSIT:
                    _n_inb += 1
            if not _need:
                continue
            _fc, _lc = Sh.detect_sheet_cols(_ws)
            _added = An.append_sku_rows_auto(_ws, _need, _names,
                                             clear_range=(_fc, _lc), state_map=_states)
            if _added:
                _txt = "、".join(f"行{r}·{k}" for r, k in _added[:8])
                _more = "" if len(_added) <= 8 else f" 等 {len(_added)} 个"
                _tag = "（含在途 %d）" % _n_inb if _n_inb else ""
                if _n_zero:
                    _tag += "（含零库存 %d：listing 已建、货未到）" % _n_zero
                _why = "无库存无销量" if (C.TRISTATE_REQUIRE_EVIDENCE
                                          and not C.TRISTATE_KEEP_MONTH_GATE) else "未上架/老 SKU"
                print(f"    [补行] {_sn}: 源 {'+'.join(_srcs)} 新增 {len(_added)} 行 → {_txt}{_more}{_tag}"
                      + (f"（另有 {_skip_old} 个{_why}不补）" if _skip_old else ""))
                n_list += len(_added)
    if n_list:
        _mode = "库存三态" if C.TRISTATE_ENABLED else "旧口径 listing+月份"
        print(f"    [补行] 本簿合计新增 {n_list} 行（{_mode}）")

    # ── 状态回刷 + 起售日期（2026-10-23 新增，只在三态模式下生效）────────────
    # 表内已有行也要跟库存走：G597/G598 之前是「未上架」，货发出去后要变「在途」，
    # 到仓可售后变「在售」。起售日期只写一次（首次出现可售库存那天），之后不再覆盖。
    if C.TRISTATE_ENABLED and stock_map:
        n_st, n_start = _sync_states_and_start_dates(wb, stock_map)
        if n_st:
            print(f"    [状态] 按库存刷新 {n_st} 行状态（在售/在途/未上架）")
        if n_start:
            print(f"    [起售] 新登记 {n_start} 个 SKU 的起售日期（= 首次有可售库存那天）")

    # ── 销量驱动补行（2026-10-23 用户拍板第二条）─────────────────────────────
    # 用户原话：「或者哪天识别到有销量了再去更新到表格里」。
    # 老 SKU 平时不追溯（首现日期早于当月），但**本月一旦有销量就必须进表** ——
    # 否则日销列没它的位置，月销汇总会少算，缺货表也看不到它。
    # 按分析表所属站点取销量（英国=Amazon.co.uk / 欧洲=region / 其余=各自 channel），
    # 避免法国站的销量被补进英国表。美客多 / EMAG 的「订单数据」不是亚马逊口径，跳过。
    n_sold = 0
    if sold_cache:
        for _sn in list(wb.sheetnames):
            if "分析表" not in _sn:
                continue
            _sold = sold_cache.get(_sn.strip()) or {}
            if not _sold:
                continue
            _ws = wb[_sn]
            _have = An.sku_set(_ws)
            _need = [k for k in _sold if k not in _have and An._SKU_RE.match(k)]
            if not _need:
                continue
            _srcs = An.source_sheets_of(_ws, wb)
            _lmap = {s: (n, a) for s, n, a in An.listing_skus_merged(wb, _srcs)}
            _names, _states = {}, {}
            for k in _need:
                _nm, _alive = _lmap.get(k, ("", True))
                _names[k] = _nm
                _states[k] = "在售" if _alive else "未上架"
            _fc, _lc = Sh.detect_sheet_cols(_ws)
            _added = An.append_sku_rows_auto(_ws, _need, _names,
                                             clear_range=(_fc, _lc), state_map=_states)
            if _added:
                print(f"    [补行·销量] {_sn}: {len(_added)} 行 ← 本月有销量但表里没有 → "
                      + "、".join(f"行{r}·{k}" for r, k in _added[:6])
                      + (f" 等 {len(_added)} 个" if len(_added) > 6 else ""))
                n_sold += len(_added)
    if n_sold:
        print(f"    [补行·销量] 本簿合计 {n_sold} 行（有销量才补，不看首现日期）")

    ships = Sh.load_shipments(C.SHIP_DIR)
    # 先清货件列数据区的文本残留（否则 H=调拨在途合计 整列 #VALUE!，见 Sh.sanitize_shipment_cols）
    nclr = Sh.sanitize_shipment_cols(wb)
    if nclr:
        print(f"    [修复] 货件列文本残留清除 {nclr} 格（H=调拨在途合计 不再 #VALUE!）")
    mine = [s for s in ships if s["book"] == fname]
    # 「到货即退场」（2026-10-20 用户维护方式）：已登记到货（到货.txt）且表内行1 已无
    # 该货件号（用户已手工删列）的货件，不再重新落列——否则每次跑批都把删掉的列建回来。
    # 行1 仍存在的到货列不动（apply_sheet 对已存在列只补空格，幂等不变）。
    row1_ids = {}
    for _sn in wb.sheetnames:
        if "分析表" not in _sn:
            continue
        _ws = wb[_sn]
        _ids = set()
        for _c in range(1, min(_ws.max_column or 1, 500) + 1):
            _v = _ws.cell(1, _c).value
            if _v is not None and str(_v).strip():
                _ids.add(str(_v).strip())
        row1_ids[_sn.strip()] = _ids
    _kept, _retired = [], []
    for s in mine:
        if s.get("arrived") and str(s["id"]).strip() not in row1_ids.get(s["sheet"].strip(), set()):
            _retired.append(s)
        else:
            _kept.append(s)
    for s in _retired:
        print(f"    [货件] {s['id']}（{s['src']}）已登记到货且表内无此列 → 不重新落列")
    if _retired:
        print(f"    [货件] 到货退场 {len(_retired)} 票（如需重新落列：从"
              f" trunk/data/数据归属周期/*/货件/到货.txt 删除对应行）")
    mine = _kept
    orphans = [s for s in ships if s.get("book") == Sh.SKIP]
    if orphans and fname == Sh.BOOK_INT_NA:
        # 无对应簿的货件（interflow 发往沙特）：只在国际簿提示一次，避免 12 本各打一遍。
        # 2026-10-20 前这类票被兜底算成"美国"，77 件沙特货件永久挂在途。
        for s in orphans:
            print(f"    [货件] !! {Sh.sid_label(s)}：Interflowd 无沙特簿 → "
                  f"跳过不落列（{int(s['total'])} 件；订单/库存同样无统计出口）")
    if not mine:
        print(f"    [货件] 本簿无货件（目录共 {len(ships) - len(orphans)} 票）")
        return
    by_sheet = {}
    for s in mine:
        s["label"] = Sh.label_of(s, s["transport"], arrived=s.get("arrived"))
        if s["ship_md"] is None:
            # PDF 货件（美客多/eMAG）文字层无发货日期 ⇒ 解析器 ship_md=None。
            # 但该票可能早已落列、且**人工补过 M/D**（apply_sheet 对已存在列不动行5）。
            # 先看表内现有标签，避免每次跑批都误报「缺发货日期」（2026-10-16）。
            ws0 = next((wb[x] for x in wb.sheetnames if x.strip() == s["sheet"].strip()), None)
            cur = None
            if ws0 is not None:
                for _c in range(1, ws0.max_column + 1):
                    if str(ws0.cell(1, _c).value or "").strip() == s["id"]:
                        v = ws0.cell(5, _c).value
                        cur = str(v) if v is not None else ""
                        break
            if cur and re.match(r"^\d{1,2}/\d{1,2}", cur):
                print(f"    [货件] {s['id']}（{s['src']}）PDF 无发货日期，沿用表内已填标签「{cur}」")
            else:
                print(f"    [货件] !! {s['id']}（{s['src']}）缺发货日期 → 行5 只写运输方式「{s['label']}」，"
                      f"请人工补 M/D（行2/3/4 公式依赖行5 前 5 字）")
        by_sheet.setdefault(s["sheet"].strip(), []).append(s)

    for sheet_key, subs in by_sheet.items():
        real = next((x for x in wb.sheetnames if x.strip() == sheet_key), None)
        if real is None:
            print(f"    [货件] !! 找不到分析表 [{sheet_key}]")
            continue
        ws = wb[real]
        fc, lc = Sh.detect_sheet_cols(ws)
        # 缺失 SKU 自动补行（全 12 簿通用，2026-10-16 用户拍板）：
        # 补在数据区末尾、按票内出现顺序，货件列留空 → 紧接的 apply_sheet 回填。
        have = An.sku_set(ws)
        need, names = [], {}
        for s in subs:
            for sku in s["items"]:
                if sku not in have and sku not in need:
                    need.append(sku)
                names.setdefault(sku, s["names"].get(sku))
        if need:
            added = An.append_sku_rows_auto(ws, need, names, clear_range=(fc, lc))
            if added:
                rows = "、".join(f"行{r}·{k}" for r, k in added)
                print(f"    [货件] {real}: 补缺失 SKU 行 {len(added)} 个 → {rows}")
            skipped = [k for k in need if k not in {x[1] for x in added}]
            if skipped:
                print(f"    [货件] {real}: 补行跳过（A 列已存在）{skipped}")
        assign, existing, new, overflow, recycle = Sh.plan_columns(ws, subs, fc, lc)
        res = Sh.apply_sheet(ws, subs, assign, new, recycle)
        for sid, c, lab, filled, miss, refill in res:
            # 日志显示**表内实际行5 标签**（已存在列 apply_sheet 不重写行5，
            # 人工补过 M/D 的列若照抄解析器标签会误导，2026-10-16）。
            _cur = ws.cell(5, c).value
            lab = str(_cur) if _cur not in (None, "") else lab
            note = f"，未命中 {miss}" if miss else ""
            rf = f"，回填空格 {refill}" if refill else ""
            rc = "，回收已到货列" if c in recycle else ""
            print(f"    [货件] {real}: {sid} → {Sh.col_letter(c)} [{lab}] {filled} 格{rf}{note}{rc}")
        if existing:
            print(f"    [货件] {real}: 复用已有列 {len(existing)} 个")
        if recycle:
            print(f"    [货件] {real}: 回收「已到货」列 {len(recycle)} 个 {sorted(Sh.col_letter(c) for c in recycle)}")
        if overflow:
            print(f"    [货件] {real}: !! 列位不足 {overflow}")


def _anchor_skip_why(skipped, sheet_missing):
    """锚点不滚动的原因串（dry/实跑共用）。"""
    miss = list(skipped) + [f"{x}（sheet不存在）" for x in sheet_missing]
    return f"库存报告缺/空: {miss}" if miss else "无库存报告"


def process(wb_cfg, dry=False, snap_dir=None, snap_day="20261006", snap_date="2026-10-06"):
    f = wb_cfg["file"]
    print(f"\n==== {f} ====")
    if wb_cfg.get("skip"):
        print(f"    跳过：{wb_cfg['skip']}")
        return dict(file=f, hard=[], soft=[], skipped=True)
    path = os.path.join(C.WORK_DIR, f)
    fmt = dict(d=snap_day, d4=snap_day[4:], m=C.ORDER_MONTH)

    # 读报告（缺文件 → 跳过该项，不中断整批）
    stock_plans, sales_plans = [], []
    stock_skipped = []          # 记录"缺/空报告"的库存 sheet → 用于锚点滚动判定
    stock_sheet_missing = []    # 记录"工作簿里没有这个 sheet"的库存项（config 与簿不符）
    sales_missing = []          # 记录"缺报告"的销售 sheet（当日销售不会更新）
    sales_stats = []            # append_sales 的返回（含 mislabel 整份错位标记）
    for s in wb_cfg["stocks"]:
        # report 可以是单个模板，也可以是候选列表（按顺序取第一个存在的）：
        # 亚马逊有的报告下载名不带日期（如 欧洲多国库存.csv），有的带（需用户改名）
        tmpls = s["report"] if isinstance(s["report"], list) else [s["report"]]
        rp = _resolve_report(snap_dir, tmpls, fmt)
        if rp is None:
            print(f"    [缺报告] {s['sheet']} ← {' / '.join(os.path.basename(t.format(**fmt)) for t in tmpls)} 不存在，跳过该项")
            stock_skipped.append(s["sheet"])
            continue
        if s["kind"] == "ml_full":
            data = read_ml_full(rp)
            if not data:
                # 安全闸（2026-10-12 自会话区回传）：ml_full 是"整块替换"，n=0 时
                # update_ml_stock_detail 会把整张「库存明细」的数据行全部 delete_rows 删光。
                # 触发场景：ML 后台「库存总览」与「刊登/商品模板」下载后同名，商品模板
                # （sheet = 帮助/hidden/商品，SKU 在 E 列而非 c4）被误命名成 Full 库存报告时，
                # read_ml_full 从 r13/c4 起读必然得到 0 行 —— 此时必须当"缺报告"处理。
                print(f"    [缺报告] {s['sheet']} ← {os.path.basename(rp)} 解析出 0 行"
                      f"（疑似商品模板或错误报告），跳过该项以免清空库存明细")
                stock_skipped.append(s["sheet"])
                continue
            stock_plans.append((s, data, None))
            print(f"    报告 {os.path.basename(rp)}: {len(data)} SKU (ml_full)")
            continue
        if s["kind"] == "emag_stock":
            data = read_emag_stock(rp)
            if not data:
                # 安全闸（2026-10-13 补）：eMAG 空白优惠模板「库存」sheet 无数据 → 0 行时
                # 整块替换会把「库存明细」全删——当"缺报告"处理（与 ml_full 安全闸同理）。
                print(f"    [缺报告] {s['sheet']} ← {os.path.basename(rp)} 解析出 0 行"
                      f"（疑似空白优惠模板），跳过该项以免清空库存明细")
                stock_skipped.append(s["sheet"])
                continue
            stock_plans.append((s, data, None))
            print(f"    报告 {os.path.basename(rp)}: {len(data)} 行 (emag_stock)")
            continue
        if s["kind"] == "multi":
            data = read_multi(rp)
            if not data[1]:
                # 安全闸（2026-10-20 补）：报告只有表头 → 整块重贴会把「多国库存」数据行
                # 全部清空，与 ml_full/emag_stock 的安全闸同理，按缺报告处理。
                print(f"    [缺报告] {s['sheet']} ← {os.path.basename(rp)} 解析出 0 行"
                      f"（只有表头），跳过该项以免清空库存")
                stock_skipped.append(s["sheet"])
                continue
            stock_plans.append((s, data, None))
            print(f"    报告 {os.path.basename(rp)}: {len(data[1])} 行 (multi)")
            continue
        data, order = read_report(rp, s["kind"], colmap=s.get("colmap"))
        if not data:
            # 安全闸（2026-10-20 补）：报告只有表头/无数据行时 update_stock_sheet 会整份跳过。
            # 若不拦在这里，空数据照样进 stock_plans →「全部库存报告都入库」判定通过 →
            # 锚点照样滚动 = 虚报库存新鲜度（10/20 实测：沙特库存 CSV 只有表头，表里还是
            # 10-19 的数据，锚点 B1/C1 却被写成 10-20）。
            print(f"    [缺报告] {s['sheet']} ← {os.path.basename(rp)} 解析出 0 行"
                  f"（只有表头或无数据行），跳过该项以免锚点虚报")
            stock_skipped.append(s["sheet"])
            continue
        stock_plans.append((s, data, order))
        print(f"    报告 {os.path.basename(rp)}: {len(data)} SKU ({s['kind']}{'/' + s['colmap'] if s.get('colmap') else ''})")
    # 库存记录：收集跑批当天填充数据（保存后做 XML 手术，见 record_extend）
    rec_jobs = build_record_jobs(wb_cfg, stock_plans)
    scfg = wb_cfg.get("sales") or []
    if isinstance(scfg, dict):
        scfg = [scfg]
    for sc in scfg:
        # report 支持候选列表（与 stocks 一致）：亚马逊后台同一报告的下改名会漂移
        # （如欧盟站点订单报告 10-13「亚马逊欧洲站点」→ 10-14「亚马逊欧盟站点」），
        # 按顺序取第一个存在的，避免改名导致销售 append 被静默跳过。
        tmpls = sc["report"] if isinstance(sc["report"], list) else [sc["report"]]
        rp = _resolve_report(snap_dir, tmpls, fmt)
        if rp is not None:
            sales_plans.append((sc, rp))
        else:
            print(f"    [缺报告] 销售 ← {' / '.join(os.path.basename(t.format(**fmt)) for t in tmpls)} 不存在，跳过append")
            sales_missing.append(sc["sheet"])

    if dry:
        for s, data, order in stock_plans:
            if s["kind"] == "multi":
                print(f"    [dry] 将整块重贴: {s['sheet']} ← {len(data[1])} 行")
            else:
                print(f"    [dry] 将更新库存sheet: {s['sheet']} ← {len(data)} SKU")
        for rn in wb_cfg.get("records", []):
            job = next((j for j in rec_jobs if j["sheet"].strip() == rn.strip()), None)
            if job:
                print(f"    [dry] 库存记录按天加列（只写 {snap_date} 固化值）: {rn}")
            else:
                print(f"    [dry] 库存记录不加列（无源报告数据/无日期列）: {rn}")
        for sc, rp in sales_plans:
            print(f"    [dry] 销售append: {sc['sheet']} ← {os.path.basename(rp)}")
        anchors_ok = bool(stock_plans) and not stock_skipped and not stock_sheet_missing
        if f in C.ANCHORS:
            if anchors_ok:
                print(f"    [dry] 锚点滚动: {C.ANCHORS[f]} → {snap_date}")
            else:
                why = _anchor_skip_why(stock_skipped, stock_sheet_missing)
                print(f"    [dry] 锚点不滚动（{why}，避免下载日期虚报）")
        return dict(file=f, hard=[], soft=[], dry=True)

    if not stock_plans and not sales_plans:
        print("    [i] 无可用报告，整份跳过（不写文件）")
        return dict(file=f, hard=["无可用报告，整份跳过"], soft=[])
    backup([f], snap_day)
    orig = path
    tmp = path + ".new.xlsx"
    # 性能：pivotCacheRecords 换成空壳再加载（省 openpyxl 建模+回写，欧盟簿实测约 10s），
    # 原字节在 restore_parts 里按原样覆盖回去。无透视件的簿 nstripped=0，直接用原文件。
    stripped = path + ".pivotstrip.xlsx"
    nstripped = strip_pivot_records(path, stripped)
    load_path = stripped if nstripped else path
    try:
        wb = load_workbook(load_path, data_only=False)
    finally:
        if nstripped and os.path.exists(stripped):
            try:
                os.remove(stripped)
            except OSError:
                pass
    for s, data, order in stock_plans:
        if s["sheet"] not in wb.sheetnames:
            print(f"    [!] sheet 不存在: {s['sheet']}")
            stock_sheet_missing.append(s["sheet"])
            continue
        # update_* 返回 None = 该 sheet 没写成（空报告/表头识别失败/列名不匹配）——
        # 均计入 stock_skipped 以阻止锚点滚动：下载日期=库存数据快照日，
        # sheet 没更新就把共享日期滚到新快照日 = 虚报库存新鲜度（10/20 事故）。
        ok = True
        if s["kind"] == "ml_full":
            update_ml_stock_detail(wb[s["sheet"]], data, snap_date)
        elif s["kind"] == "emag_stock":
            update_emag_stock_detail(wb[s["sheet"]], data, snap_date)
        elif s["kind"] == "multi":
            ok = update_multi_sheet(wb[s["sheet"]], data[0], data[1]) is not None
        else:
            kind = "balance" if s["kind"] == "balance" else "inv"
            ok = update_stock_sheet(wb[s["sheet"]], data, order, kind=kind,
                                    colmap=s.get("colmap")) is not None
        if not ok:
            stock_skipped.append(s["sheet"])
    for rn in wb_cfg.get("records", []):
        # 表名可能带尾随空格（'英国库存记录 '）——加列已移到保存后做 XML 手术，
        # 这里只校验 sheet 存在性，实际写入见 restore_parts 之后的 rec_jobs 循环
        target = rn if rn in wb.sheetnames else next((n for n in wb.sheetnames if n.strip() == rn.strip()), None)
        if not target:
            print(f"    [!] 记录sheet不存在: {rn}")
    # 存量修复：订单号 int→文本（不依赖订单报告是否到位，每天幂等）
    for sc, _rp in sales_plans:
        if sc["sheet"] in wb.sheetnames:
            nfix = fix_id_text(wb[sc["sheet"]])
            if nfix:
                print(f"    [修复] {sc['sheet']}: {nfix} 个订单号 int→文本")
    # 本簿库存 SKU 集 —— SKU 归属闸的参照（订单报告的 SKU 必须大部分落在本簿库存里）
    stock_sku_set = set()
    for s in (wb_cfg.get("stocks") or []):
        sn = s.get("sheet")
        if sn not in wb.sheetnames:
            continue
        _ws = wb[sn]
        for row in _ws.iter_rows(min_col=1, max_col=1, max_row=min(_ws.max_row, 3000),
                                 values_only=True):
            v = row[0] if row else None
            if isinstance(v, str) and v.strip():
                stock_sku_set.add(v.strip())
    for sc, rp in sales_plans:
        if sc["sheet"] in wb.sheetnames:
            if sc.get("kind") == "ml_sales":
                r = append_ml_sales(wb[sc["sheet"]], rp, window_lo=sc.get("window_lo"), snap_date=snap_date)
            else:
                r = append_sales(wb[sc["sheet"]], rp, sc.get("sep") or "\t", key=sc.get("key"),
                                 expect_channel=sc.get("expect_channel"), book=f,
                                 stock_skus=stock_sku_set,
                                 allow_mixed=bool(sc.get("allow_mixed_channel")))
            if r:
                sales_stats.append(r)
        else:
            print(f"    [!] sheet 不存在: {sc['sheet']}")
    # 锚点滚动（"组合跟着更新"）：模拟透视件的下载日期格 → 快照日
    # 仅在该簿**全部**库存报告都入库时滚动；有缺失就不动（锚点=下载日期，
    # 部分 sheet 只有旧数据时把共享日期滚到新快照日 = 虚报库存新鲜度）。
    if f in C.ANCHORS:
        if stock_plans and not stock_skipped and not stock_sheet_missing:
            for sn, coord in C.ANCHORS[f]:
                if sn in wb.sheetnames:
                    old = wb[sn][coord].value
                    wb[sn][coord].value = datetime.strptime(snap_day, "%Y%m%d")
                    print(f"    [锚点] {sn}!{coord}: {old} → {snap_date}")
                else:
                    print(f"    [!] 锚点sheet不存在: {sn!r}")
        else:
            why = _anchor_skip_why(stock_skipped, stock_sheet_missing)
            print(f"    [锚点] 不滚动（{why}，避免下载日期虚报）")
    # 货件列匹配填写 + 缺失 SKU 追加行（仅欧盟簿；幂等）
    apply_shipments(wb, snap_day, fname=f)
    # 站点表补行（2026-10-27）：本站本月销量>0 而站点表无行 → 补。
    # 必须在 apply_month_summary 之前，新行当期就能拿到月度封存值。
    _n_site = _site_append_rows(wb, snap_day)
    if _n_site:
        print(f"    [补行·站点] 本簿合计 {_n_site} 行（本站有销量才补）")
    # 月汇总：上月封存 + 当月列滚动（全部簿；口径=各分析表日销 SUMIFS 自动提取）
    apply_month_summary(wb, snap_day, fname=f)
    # 打开即全量重算：openpyxl 不写公式缓存值，日销列 SUMIFS / F / G / H 等公式格在
    # 关闭了自动重算的 Excel/WPS 或部分在线阅读器里会不刷新（条件格式因此不显色）。
    # fullCalcOnLoad=1 让任何表格软件一打开就重算全簿（2026-10-16）。
    try:
        wb.calculation.fullCalcOnLoad = True
        wb.calculation.calcMode = "auto"
    except Exception:
        pass
    wb.save(tmp)
    wb.close()
    restored, piv = restore_parts(orig, tmp)
    # 库存记录按天加列 + 只写当天固化值（XML 手术，必须在 openpyxl 保存后做：
    # openpyxl 的 insert_cols 不平移 mergeCells/CF/cols/公式，会静默损坏表结构）
    for job in rec_jobs:
        try:
            record_extend.extend_record_file(tmp, job["sheet"], snap_date,
                                             job["insale"], job["intra"])
        except Exception as e:
            print(f"    [记录] {job['sheet']}: 加列失败 {e!r}（库存/销售数据已保存，记录表现状不变）")
    # 本机 os.replace 偶发 PermissionError(WinError 5)（WPS/资源管理器占用）→ cp -f 兜底
    try:
        os.replace(tmp, orig)
    except OSError:
        try:
            shutil.copyfile(tmp, orig)
            os.remove(tmp)
        except OSError as e2:
            raise RuntimeError(
                f"无法覆盖原文件：{os.path.basename(orig)} 正被其他程序占用"
                f"（多半是 WPS/Excel 开着它）。请关闭后重跑；本次该簿全部更新已丢弃: {e2!r}"
            ) from e2
    print(f"    [保真] 回填{len(restored)}部件 pivot刷新{piv}处 → 已覆盖原文件")

    # ---- 本簿异常汇总（供 main() 汇总与退出码判定）----
    hard, soft = [], []
    for st in sales_stats:
        if st.get("mislabel"):
            hard.append(f"销售报告整份错位已拒收: {st['report']} → {st['sheet']}"
                        f"（外来渠道 {st['foreign']}/{st['total']} 行）")
        elif st.get("unknown"):
            soft.append(f"{st['sheet']}: {st['unknown']} 行渠道未登记")
        elif st.get("total") and not st.get("appended"):
            soft.append(f"{st['sheet']}: 报告 {st['total']} 行全部去重（无新增）")
    for sn in sales_missing:
        hard.append(f"销售报告缺失，当日不更新: {sn}")
    for sn in stock_skipped:
        hard.append(f"库存报告缺/空，当日不更新: {sn}")
    for sn in stock_sheet_missing:
        hard.append(f"库存 sheet 不存在（config 与工作簿不符），当日不更新: {sn}")
    if hard or soft:
        for h in hard:
            print(f"    [汇总] ✗ {h}")
        for s in soft:
            print(f"    [汇总] · {s}")
    else:
        print("    [汇总] ✅ 无异常")
    return dict(file=f, hard=hard, soft=soft)


def _run_one(wb_cfg, dry, snap_dir, snap_day, snap_date, work_dir=None):
    """并行调度单元：把子进程 stdout 原样带回主进程打印；异常不炸整批（与串行容错一致）。
    work_dir 仅供测试沙箱覆盖（bench_run.py 用），生产不传。
    返回 (簿名, 输出文本, 异常dict) —— 异常dict 供 main() 做「今日异常汇总」与退出码。"""
    import io, contextlib, traceback
    if work_dir:
        C.WORK_DIR = work_dir
    buf = io.StringIO()
    iss = dict(file=wb_cfg["file"], hard=[], soft=[])
    try:
        with contextlib.redirect_stdout(buf):
            iss = process(wb_cfg, dry=dry, snap_dir=snap_dir,
                          snap_day=snap_day, snap_date=snap_date) or iss
    except Exception as e:
        print(f"    [ERROR] {wb_cfg['file']}: {e!r}", file=buf)
        traceback.print_exc(file=buf)
        iss["hard"] = [f"跑批异常 {type(e).__name__}: {e}"] + list(iss.get("hard") or [])
    return wb_cfg["file"], buf.getvalue(), iss


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--only", default="")
    ap.add_argument("--snap", default="", help="手工指定快照目录（默认自动取最新 YYYYMMDD 目录）")
    ap.add_argument("--period", default="", help="数据归属周期目录名（默认 config.CURRENT_PERIOD）")
    ap.add_argument("--jobs", type=int, default=3,
                    help="并行处理的工作簿数。openpyxl 解析是 CPU 密集型，用多进程；默认 3。"
                         "内存紧张或只想看顺序日志可设 1")
    args = ap.parse_args()
    if args.snap:
        snap_dir = os.path.normpath(args.snap)
        snap_day = os.path.basename(snap_dir)
        if not re.match(r"^\d{8}$", snap_day):
            raise SystemExit(f"--snap 目录名必须是 YYYYMMDD: {snap_dir}")
    else:
        snap_dir, snap_day = C.latest_snapshot(args.period or None)
    snap_date = f"{snap_day[:4]}-{snap_day[4:6]}-{snap_day[6:]}"
    print(f"===== 快照目录: {snap_dir}")
    print(f"===== 数据日: {snap_date} | 订单报告月: {C.ORDER_MONTH} =====")
    cfgs = [c for c in C.WORKBOOKS
            if not args.only or any(k in c["file"] for k in args.only.split(","))]
    jobs = max(1, args.jobs)
    work_dir = os.environ.get("DSR_WORK_DIR") or None
    results = []
    if jobs == 1 or len(cfgs) <= 1:
        for wb_cfg in cfgs:
            _f, out, iss = _run_one(wb_cfg, args.dry_run, snap_dir, snap_day, snap_date, work_dir)
            print(out, end="")
            results.append(iss)
    else:
        from concurrent.futures import ProcessPoolExecutor
        from functools import partial
        fn = partial(_run_one, dry=args.dry_run, snap_dir=snap_dir,
                     snap_day=snap_day, snap_date=snap_date, work_dir=work_dir)
        print(f"===== 并行度: {jobs}（{len(cfgs)} 本簿，多进程） =====")
        with ProcessPoolExecutor(max_workers=jobs) as ex:
            for _f, out, iss in ex.map(fn, cfgs):
                print(out, end="")
                results.append(iss)

    # ---- 今日异常汇总（2026-10-19：沙特报告整份错位只打在两行日志里，不看全文就漏）----
    hard_all = [(i["file"], h) for i in results for h in (i.get("hard") or [])]
    soft_all = [(i["file"], s) for i in results for s in (i.get("soft") or [])]
    if hard_all or soft_all:
        print("\n" + "=" * 60)
        print(f"今日异常汇总（{snap_date}，{len(results)} 本簿）")
        print("=" * 60)
        for f_, h in hard_all:
            print(f"  ✗ {f_}: {h}")
        for f_, s in soft_all:
            print(f"  · {f_}: {s}")
        print("=" * 60)
    if hard_all:
        print(f">>> 有 {len(hard_all)} 项硬伤（报告错位/缺失/异常），相关簿当日数据不完整，"
              f"请处理后重跑；退出码 3")
        return 3
    if soft_all:
        print(f">>> 有 {len(soft_all)} 项软提醒（已正常入库）；退出码 0")
    return 0


if __name__ == "__main__":
    sys.exit(main())
