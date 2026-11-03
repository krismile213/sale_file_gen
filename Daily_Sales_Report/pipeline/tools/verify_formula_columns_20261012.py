# -*- coding: utf-8 -*-
"""
20261012 快照 → 全链复算验证 v2（只读）。
v2 修正：AA/分析表用「值化副本 or 公式文本模式」判读；逐行比较统一规范化；
每日销售数据日期改用 C 列 purchase-date 原始值（AP 是公式无缓存）；
美客多/罗马尼亚分析表日销复算改用「日销表程序版 output」值化簿。
"""
# --- pipeline 路径引导（2026-10-14 目录收敛自动注入）---
import os as _os, sys as _sys
_bp = _os.path.dirname(_os.path.abspath(__file__))
for _i in range(6):
    if _os.path.isdir(_os.path.join(_bp, "1.每日销售数据表")):
        break
    _bp = _os.path.dirname(_bp)
DSR = _bp
_bcore = _os.path.join(DSR, "pipeline", "core")
if _bcore not in _sys.path:
    _sys.path.insert(0, _bcore)
# --- 引导结束 ---
import csv
import json
import os
import sys
from datetime import datetime
from collections import defaultdict

sys.stdout.reconfigure(encoding="utf-8")
import openpyxl

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import config as C
from runner import read_report, read_multi, read_emag_stock, read_ml_full, _num

SNAP = r"C:\Users\dell\Desktop\work-main\sale_file_gen\trunk\data\数据归属周期\202610\快照日期\20261010\20261012"
# [normalize] DSR 由上方引导块赋值
BOOK = DSR + r"\1.每日销售数据表"
OUT_VER = DSR + r"\日销表程序版\output"          # 值化簿（分析表已写值）
SESSION_OUT = r"C:\Users\dell\Desktop\sale_file_gen\sale_file_gen\Daily_Sales_Report\output"


def parse_dt(v):
    if isinstance(v, datetime):
        return v.date()
    s = str(v).strip()
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d/%m/%Y %H:%M", "%d/%m/%Y"):
        try:
            return datetime.strptime(s, fmt).date()
        except Exception:
            continue
    try:
        return datetime.fromisoformat(s.replace("Z", "+00:00")).date()
    except Exception:
        return None

FIELDS = ["afn-fulfillable-quantity", "afn-onhand-buyable-quantity",
          "afn-inbound-working-quantity", "afn-inbound-shipped-quantity",
          "afn-inbound-receiving-quantity"]
R = {}


def norm(v):
    """统一规范化：数字化、日期化或去空白字符串。"""
    if v is None:
        return None
    if isinstance(v, datetime):
        return v.strftime('%Y-%m-%d %H:%M') if (v.hour or v.minute) else v.strftime('%Y-%m-%d')
    if isinstance(v, (int, float)):
        return int(v) if float(v).is_integer() else round(float(v), 6)
    s = str(v).strip()
    if s == '':
        return None
    n = _num(s)
    return n if n is not None and not isinstance(n, str) else s


# ============ A/B. 英国 / 欧洲库存 sheet（值 + AA 公式模式） ============
def check_stock_sheet(tag, sheet_name, report_path, kind, colmap=None):
    wb = openpyxl.load_workbook(BOOK + r"\欧盟英国库存销售表.xlsx", read_only=True, data_only=False)
    ws = wb[sheet_name]
    header, hdr_row = {}, None
    for r in (1, 2):
        vals = [(i, v) for i, v in enumerate(next(ws.iter_rows(min_row=r, max_row=r, max_col=30, values_only=True)), 1) if v]
        if any(str(v).startswith("afn-") for _, v in vals):
            header = {str(v).strip(): i for i, v in vals}
            hdr_row = r
            break
    data, order = read_report(report_path, kind, colmap=colmap)
    ok = bad = missing = aa_pat = aa_other = 0
    for row in ws.iter_rows(min_row=hdr_row + 1, max_row=ws.max_row, max_col=27, values_only=True):
        sku = row[header["sku"] - 1]
        if sku is None or not str(sku).strip():
            continue
        aa = row[26]
        if isinstance(aa, str) and aa.replace(" ", "") == f"=P{aa_row}+Q{aa_row}+R{aa_row}" if False else False:
            pass
        d = data.get(str(sku).strip())
        if d is None:
            missing += 1
        else:
            for fld in FIELDS:
                col = header.get(fld)
                if col and row[col - 1] is not None:
                    if norm(row[col - 1]) == norm(d.get(fld)):
                        ok += 1
                    else:
                        bad += 1
                        if bad <= 3:
                            print(f"    ✗ {sku} {fld}: 表={row[col-1]} 报告={d.get(fld)}")
        # AA 公式模式（值化前只能看公式文本）
        if isinstance(aa, str) and aa.startswith("="):
            aa_pat += 1
        elif aa is not None:
            aa_other += 1
    R[tag] = {"report_skus": len(data), "cells_ok": ok, "cells_bad": bad,
              "sheet_skus_not_in_report": missing,
              "AA_col": {"formula_cells": aa_pat, "static_cells": aa_other}}
    print(f"[{tag}] 报告{len(data)}SKU | 值格 一致{ok}/不一致{bad} | sheet有报告无 {missing} | AA列 公式{aa_pat}格/静态{aa_other}格")
    wb.close()
    # AA 语义：在值化簿（程序版 output 欧盟英国）读值复算 P+Q+R
    p_out = OUT_VER + r"\欧盟英国库存销售表.xlsx"
    if os.path.isfile(p_out):
        wbv = openpyxl.load_workbook(p_out, read_only=True, data_only=True)
        wsv = wbv[sheet_name]
        hdr2 = None
        for r in (1, 2):
            vals = [(i, v) for i, v in enumerate(next(wsv.iter_rows(min_row=r, max_row=r, max_col=30, values_only=True)), 1) if v]
            if any(str(v).startswith("afn-") for _, v in vals):
                hdr2 = r
                break
        aa_ok = aa_bad = 0
        for row in wsv.iter_rows(min_row=hdr2 + 1, max_row=hdr2 + 800, max_col=27, values_only=True):
            aa, p, q, rr = row[26], row[15], row[16], row[17]
            if aa is None:
                continue
            if norm(aa) == norm((p or 0) + (q or 0) + (rr or 0)):
                aa_ok += 1
            else:
                aa_bad += 1
                if aa_bad <= 2:
                    print(f"    ✗ AA 语义(output值化簿): aa={aa} P+Q+R={(p or 0)+(q or 0)+(rr or 0)}")
        R[tag]["AA_semantics_valuebook"] = {"ok_800": aa_ok, "bad": aa_bad}
        print(f"    AA=P+Q+R（值化簿抽查800行）: {aa_ok} 对 / {aa_bad} 错")
        wbv.close()


# ============ C. 多国库存 ============
def check_multi():
    wb = openpyxl.load_workbook(BOOK + r"\欧盟英国库存销售表.xlsx", read_only=True, data_only=False)
    ws = wb["多国库存"]
    hdr, body = read_multi(SNAP + r"\欧洲多国库存.csv")
    sh_hdr = [str(v).strip() if v is not None else "" for v in next(ws.iter_rows(min_row=2, max_row=2, max_col=10, values_only=True))]
    hdr_ok = sh_hdr[:len(hdr)] == hdr
    sheet_rows = []
    for row in ws.iter_rows(min_row=3, max_row=ws.max_row, max_col=len(hdr), values_only=True):
        if any(v is not None for v in row):
            sheet_rows.append([norm(v) for v in row])
    n = min(len(sheet_rows), len(body))
    diff = 0
    for a, b in zip(sheet_rows[:n], body[:n]):
        for x, y in zip(a, b):
            if x != norm(y):
                diff += 1
    R["多国库存"] = {"report_rows": len(body), "sheet_rows": len(sheet_rows), "header_ok": hdr_ok,
                    "rows_compared": n, "cell_diff": diff}
    print(f"[多国库存] 报告{len(body)}行 / sheet{len(sheet_rows)}行 / 表头一致={hdr_ok} / 逐格 diff={diff} (比{n}行)")
    wb.close()


# ============ D. 每日销售数据 vs 欧洲订单报告（日期取 C 列原始值） ============
def check_sales():
    wb = openpyxl.load_workbook(BOOK + r"\欧盟英国库存销售表.xlsx", read_only=True, data_only=True)
    ws = wb["每日销售数据"]
    agg = defaultdict(int)
    n_rows = 0
    day = datetime(2026, 10, 12).date()
    for row in ws.iter_rows(min_row=3, max_row=ws.max_row, max_col=19, values_only=True):
        sku, qty, pd, ch = row[11], row[14], row[2], row[6]   # L, O, C, G
        if sku is None or qty is None or pd is None:
            continue
        d = None
        if isinstance(pd, datetime):
            d = pd.date()
        else:
            s = str(pd)
            try:
                d = datetime.fromisoformat(s.replace("Z", "+00:00")).date() if "T" in s else datetime.strptime(s[:10], "%Y-%m-%d").date()
            except Exception:
                continue
        if d == day:
            agg[(str(sku).strip(), str(ch).strip() if ch else "")] += int(_num(qty))
        n_rows += 1
    rep = defaultdict(int)
    n_rep = 0
    with open(SNAP + r"\亚马逊欧洲站点202610订单报告.txt", encoding="utf-8-sig", newline="") as f:
        for row in csv.DictReader(f, delimiter="\t"):
            sku = (row.get("sku") or "").strip()
            qty = _num(row.get("quantity"))
            pd = (row.get("purchase-date") or "").strip()
            if not sku or qty is None or not pd:
                continue
            try:
                d = datetime.fromisoformat(pd.replace("Z", "+00:00")).date()
            except Exception:
                continue
            if d == day:
                rep[(sku, (row.get("sales-channel") or "").strip())] += int(qty)
            n_rep += 1
    keys = set(agg) | set(rep)
    ok = bad = 0
    for k in sorted(keys):
        a, b = agg.get(k, 0), rep.get(k, 0)
        if a == b:
            ok += 1
        else:
            bad += 1
            if bad <= 5:
                print(f"    ✗ {k}: sheet={a} 报告={b}")
    R["每日销售数据"] = {"sheet_rows": n_rows, "report_rows_202610txt": n_rep,
                        "day_1012_groups_ok": ok, "day_1012_groups_bad": bad}
    print(f"[每日销售数据] sheet {n_rows} 行 / 报告 {n_rep} 行；10-12 当天 SKU×渠道: 一致 {ok} / 不一致 {bad}")
    wb.close()


# ============ E. 罗马尼亚库存明细 + 分析表（值化簿） ============
def check_ro():
    wb = openpyxl.load_workbook(BOOK + r"\罗马尼亚库存销售表.xlsx", read_only=True, data_only=True)
    sd = wb["库存明细"]
    rows = read_emag_stock(SNAP + r"\EMAG罗马尼亚库存报告20261012.xlsx")
    rep = [(norm(d["pnk"]), norm(d["code"]), norm(d["qty"]), norm(d["st"])) for d in rows]
    got, a1 = [], None
    for row in sd.iter_rows(min_row=5, max_row=sd.max_row, max_col=28, values_only=True):
        code = row[4]
        if code is None or not str(code).strip():
            continue
        got.append((norm(row[3]), norm(code), norm(row[26]), norm(row[27])))
    a1 = sd["A1"].value
    wb.close()
    same = sum(1 for a, b in zip(got, rep) if a == b)
    R["罗马尼亚库存明细"] = {"report_rows": len(rep), "sheet_rows": len(got), "row_equal": same, "anchor": str(a1)}
    print(f"[罗马尼亚库存明细] 报告{len(rep)}行 / sheet{len(got)}行 / 逐行相等 {same} / A1={a1}")
    # 分析表 X/日销（程序版 output 值化簿）
    p_out = OUT_VER + r"\罗马尼亚库存销售表.xlsx"
    wbv = openpyxl.load_workbook(p_out, read_only=True, data_only=True)
    od = wbv["订单数据"]
    agg = defaultdict(int)
    for row in od.iter_rows(min_row=2, max_row=od.max_row, max_col=34, values_only=True):
        sku = row[4]        # E 列 c5
        qty = row[7]        # H 列 c8
        ag = parse_dt(row[32])   # AG 列 c33 purchase-date（文本日期兼容）
        valid = row[33]     # AH 列 c34 有效性
        if sku is None or qty is None or ag is None or norm(valid) != 1:
            continue
        agg[(str(sku).strip(), ag)] += int(qty)
    an = wbv["罗马尼亚-每日库存分析表"]
    dates = {}
    for row in an.iter_rows(min_row=5, max_row=5, min_col=25, max_col=63, values_only=True):
        for i, v in enumerate(row, 25):
            if isinstance(v, datetime):
                dates[i] = v.date()
    x_ok = x_bad = d_ok = d_bad = n = 0
    for row in an.iter_rows(min_row=7, max_row=an.max_row, min_col=1, max_col=63, values_only=True):
        sku = row[0]
        if sku is None or not str(sku).strip():
            continue
        n += 1
        xc = sum(agg.get((str(sku).strip(), d), 0) for d in dates.values() if d.month == 9)
        if _num(row[23]) == xc:
            x_ok += 1
        else:
            x_bad += 1
            if x_bad <= 3:
                print(f"    ✗ X {sku}: 表={row[23]} 复算={xc}")
        for i, d in dates.items():
            if _num(row[i - 1]) == agg.get((str(sku).strip(), d), 0):
                d_ok += 1
            else:
                d_bad += 1
                if d_bad <= 3:
                    print(f"    ✗ 日销 {sku} {d}: 表={row[i-1]} 复算={agg.get((str(sku).strip(), d), 0)}")
    R["罗马尼亚分析表"] = {"sku_rows": n, "X_9月总": {"ok": x_ok, "bad": x_bad},
                          "日销窗": {"ok": d_ok, "bad": d_bad, "days": len(dates)}, "value_book": p_out}
    print(f"[罗马尼亚分析表(值化簿)] SKU {n} 行；X9月总 {x_ok}对/{x_bad}错；日销窗 {len(dates)}天 {d_ok}对/{d_bad}错")
    wbv.close()


# ============ F. 美客多库存明细 + 分析表（会话区 output 值化簿） ============
def check_ml():
    wb = openpyxl.load_workbook(BOOK + r"\美客多库存销售表.xlsx", read_only=True, data_only=True)
    sd = wb["库存明细"]
    rows = read_ml_full(SNAP + r"\美客多墨西哥Full库存报告20261012.xlsx")
    rep = [[norm(v) for v in r] for r in rows]
    got = []
    for row in sd.iter_rows(min_row=5, max_row=sd.max_row, max_col=34, values_only=True):
        if row[3] is None or not str(row[3]).strip():
            continue
        got.append([norm(v) for v in row[0:33]])   # read_ml_full: 报告c2..c34 ↔ 工作簿c1..c33
    a1 = sd["A1"].value
    wb.close()
    same = sum(1 for a, b in zip(got, rep) if a == b)
    R["美客多库存明细"] = {"report_rows": len(rep), "sheet_rows": len(got), "row_equal": same, "anchor": str(a1)}
    print(f"[美客多库存明细] 报告{len(rep)}行 / sheet{len(got)}行 / 逐行相等 {same} / A1={a1}")
    p_out = SESSION_OUT + r"\美客多库存销售表.xlsx"
    wbv = openpyxl.load_workbook(p_out, read_only=True, data_only=True)
    od = wbv["订单数据"]
    agg = defaultdict(int)
    for row in od.iter_rows(min_row=2, max_row=od.max_row, max_col=66, values_only=True):
        sku = row[18]
        qty = row[6]
        bl = parse_dt(row[63])
        valid = row[64]     # BM c65 有效性
        if sku is None or qty is None or bl is None or norm(valid) != 1:
            continue
        agg[(str(sku).strip(), bl)] += int(qty)
    an = wbv["墨西哥-每日库存分析表"]
    dates = {}
    for row in an.iter_rows(min_row=5, max_row=5, min_col=25, max_col=63, values_only=True):
        for i, v in enumerate(row, 25):
            if isinstance(v, datetime):
                dates[i] = v.date()
    x_ok = x_bad = d_ok = d_bad = n = 0
    for row in an.iter_rows(min_row=7, max_row=an.max_row, min_col=1, max_col=63, values_only=True):
        sku = row[0]
        if sku is None or not str(sku).strip():
            continue
        n += 1
        xc = sum(agg.get((str(sku).strip(), d), 0) for d in dates.values() if d.month == 9)
        if _num(row[23]) == xc:
            x_ok += 1
        else:
            x_bad += 1
            if x_bad <= 3:
                print(f"    ✗ X {sku}: 表={row[23]} 复算={xc}")
        for i, d in dates.items():
            if _num(row[i - 1]) == agg.get((str(sku).strip(), d), 0):
                d_ok += 1
            else:
                d_bad += 1
                if d_bad <= 3:
                    print(f"    ✗ 日销 {sku} {d}: 表={row[i-1]} 复算={agg.get((str(sku).strip(), d), 0)}")
    R["美客多分析表"] = {"sku_rows": n, "X_9月总": {"ok": x_ok, "bad": x_bad},
                        "日销窗": {"ok": d_ok, "bad": d_bad, "days": len(dates)}, "value_book": p_out}
    print(f"[美客多分析表(值化簿)] SKU {n} 行；X9月总 {x_ok}对/{x_bad}错；日销窗 {len(dates)}天 {d_ok}对/{d_bad}错")
    wbv.close()


if __name__ == "__main__":
    print("=== 20261012 全链复算 v2 ===")
    check_stock_sheet("英国库存", "英国可售+在途库存", SNAP + r"\亚马逊英国管理库存报告20261012.csv", "inv")
    check_stock_sheet("欧洲库存", "欧洲可售+在途库存", DSR + r"\欧洲库存1012.csv", "age")
    check_multi()
    check_sales()
    check_ro()
    check_ml()
    with open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "verify_formula_20261012.json"), "w", encoding="utf-8") as f:
        json.dump(R, f, ensure_ascii=False, indent=1)
    print("=== 完成 ===")
