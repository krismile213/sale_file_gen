# -*- coding: utf-8 -*-
"""修复两个非亚马逊渠道的历史数据缺口（2026-11-04）。

① eMAG（罗马尼亚）：明细里 18 行 `订单日期` 为空（当年报告列名漂移，日期列没映射上），
   日销列 SUMIFS 按日期取不到 → 10 月少 13 件。用订单报告按单号回填日期。
② 美客多：明细里 63 件的中文状态描述未被中文→西语规则覆盖 → c65 有效性 #N/A
   → 判不计销。用同一条归一管线（runner.ml_norm_status / ml_norm_desc）重跑一遍，
   幂等：已经是西语的行不会被改动。

两个缺口都同时影响「当前簿」和「10 月封存簿」——上月月汇总块用 resolve_month_sales
读封存簿重算，只修当前簿的话汇总还是旧值。故两处都修。

用法：
    python pipeline/tools/maintenance/fix_sales_gaps_20261104.py           # 预演
    python pipeline/tools/maintenance/fix_sales_gaps_20261104.py --apply   # 落盘
"""
from __future__ import annotations

import argparse
import glob
import os
import shutil
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
CORE = os.path.join(os.path.dirname(os.path.dirname(HERE)), "core")
DSR = os.path.dirname(os.path.dirname(CORE))
REPO = os.path.dirname(DSR)
for p in (CORE, os.path.join(DSR, "pipeline")):
    if p not in sys.path:
        sys.path.insert(0, p)

import openpyxl  # noqa: E402

import config as C          # noqa: E402
import runner as R          # noqa: E402

SNAP_DATE = "2026-11-04"
BACKUP = os.path.join(DSR, ".workbuddy", "backup_fix_gaps_20261104")
EMAG_FILE = "罗马尼亚库存销售表.xlsx"
ML_FILE = "美客多库存销售表.xlsx"


def _targets(fname, period="202610"):
    """当前簿 + 指定月份封存簿（存在才返回）。"""
    cur = os.path.join(DSR, "1.每日销售数据表", fname)
    y, m = period[:4], int(period[4:])
    arch = os.path.join(DSR, "1.每日销售数据表", f"{y}年每日数据", f"{y}年{m}月",
                        f"{y}年{m}月-{fname}")
    return [p for p in (cur, arch) if os.path.exists(p)]


def _backup(path):
    os.makedirs(BACKUP, exist_ok=True)
    dst = os.path.join(BACKUP, os.path.basename(path))
    if not os.path.exists(dst):
        shutil.copy2(path, dst)
    return dst


def _hdr(ws, need):
    """找表头行 → (行号, {列名: 列号})。
    ⚠️ **重名列取第一个**：美客多簿里 c3「Estado」=状态、c38「Estado」=墨西哥州，
    若用 last-wins 会把状态写到地址列上去（写坏数据）。"""
    for r in range(1, 6):
        names = [(str(c.value).strip() if c.value is not None else "") for c in ws[r]]
        if all(k in names for k in need):
            h = {}
            for i, n in enumerate(names):
                if n:
                    h.setdefault(n, i + 1)
            return r, h
    return None, None


# ---------------------------------------------------------------- ① eMAG

def emag_report_dates():
    """扫全部快照里 eMAG 订单报告 → ({订单编号: 日期串}, {订单编号: 状态}, 报告份数)。"""
    alias = C.EMAG_COL_ALIAS
    out, st = {}, {}
    pats = os.path.join(REPO, "trunk", "data", "数据归属周期", "*", "快照日期", "*",
                        "EMAG全站点*订单报告.xlsx")
    files = sorted(glob.glob(pats))
    for p in files:
        try:
            wb = openpyxl.load_workbook(p, read_only=True, data_only=True)
            ws = wb[wb.sheetnames[0]]
            rows = ws.iter_rows(values_only=True)
            raw = next(rows, None)
            if not raw:
                wb.close()
                continue
            hdr = [alias.get(str(x).strip(), str(x).strip()) if x is not None else ""
                   for x in raw]
            if "订单编号" not in hdr or "订单日期" not in hdr:
                wb.close()
                continue
            i_oid, i_dt = hdr.index("订单编号"), hdr.index("订单日期")
            i_st = hdr.index("订单状态") if "订单状态" in hdr else None
            for r in rows:
                if not r or len(r) <= max(i_oid, i_dt) or r[i_oid] is None:
                    continue
                oid = str(r[i_oid]).strip()
                d = r[i_dt]
                if d is not None and str(d).strip():
                    out[oid] = str(d).strip()
                    if i_st is not None and i_st < len(r):
                        st[oid] = str(r[i_st] or "").strip()
            wb.close()
        except Exception as e:      # 单份报告坏了不影响整体
            print(f"  [eMAG] 跳过 {os.path.basename(p)}: {e!r}")
    return out, st, len(files)


def fix_emag(paths, dates, statuses, apply=False):
    import status_sync as SS
    for p in paths:
        wb = openpyxl.load_workbook(p)
        ws = wb["订单数据"]
        hr, h = _hdr(ws, ("订单编号", "订单日期"))
        if not hr:
            print(f"  [eMAG] {os.path.basename(p)}: 表头识别失败，跳过")
            wb.close()
            continue
        c_oid, c_dt = h["订单编号"], h["订单日期"]
        filled = notfound = 0
        samples = []
        for r in range(hr + 1, ws.max_row + 1):
            oid = ws.cell(r, c_oid).value
            if oid is None or not str(oid).strip():
                continue
            cur = ws.cell(r, c_dt).value
            if cur is not None and str(cur).strip():
                continue
            d = dates.get(str(oid).strip())
            if not d:
                notfound += 1
                continue
            if apply:
                ws.cell(r, c_dt).value = d
            filled += 1
            if len(samples) < 3:
                samples.append((str(oid).strip(), d, ws.cell(r, h["产品代码"]).value if "产品代码" in h else None))
        tag = "" if apply else "[预演] "
        print(f"  {tag}[eMAG] {os.path.basename(p)}: 回填 {filled} 行日期"
              f"，报告里查不到 {notfound} 行  例={samples}")
        # ③ 滞后取消：报告已 Storno/Canceled、明细还是 Finished → 改状态
        stt = SS.fix_emag_sheet(ws, statuses, dry=not apply)
        if stt.get("ok"):
            print(f"  {tag}[eMAG] {os.path.basename(p)}: 滞后取消 {stt['fixed']} 行"
                  f"（{stt['qty']:.0f} 件停止计销）例={stt['samples'][:3]}")
        if apply and (filled or stt.get("fixed")):
            _backup(p)
            wb.save(p)
        wb.close()


# ---------------------------------------------------------------- ② 美客多

def fix_ml(paths, apply=False):
    for p in paths:
        wb = openpyxl.load_workbook(p)
        ws = wb["订单数据"]
        hr, h = _hdr(ws, ("# de venta", "Estado", "Descripción del estado"))
        if not hr:
            print(f"  [ML] {os.path.basename(p)}: 表头识别失败，跳过")
            wb.close()
            continue
        c_st, c_de = h["Estado"], h["Descripción del estado"]
        n_st = n_de = 0
        samples = []
        for r in range(hr + 1, ws.max_row + 1):
            if ws.cell(r, h["# de venta"]).value is None:
                continue
            st = ws.cell(r, c_st).value
            de = ws.cell(r, c_de).value
            if st is not None and str(st).strip():
                new = R.ml_norm_status(st, SNAP_DATE)
                if new and new != str(st).strip():
                    if apply:
                        ws.cell(r, c_st).value = new
                    n_st += 1
            if de is not None and str(de).strip():
                new = R.ml_norm_desc(de, SNAP_DATE)
                if new and new != str(de).strip():
                    if apply:
                        ws.cell(r, c_de).value = new
                    n_de += 1
                    if len(samples) < 3:
                        samples.append((str(de)[:26], new[:40]))
        tag = "" if apply else "[预演] "
        print(f"  {tag}[ML] {os.path.basename(p)}: 状态归一 {n_st} 行 / 描述归一 {n_de} 行"
              f"  例={samples}")
        if apply and (n_st or n_de):
            _backup(p)
            wb.save(p)
        wb.close()


# ---------------------------------------------------------------- ② 美客多状态回写

def ml_report_map():
    """{ '# de venta'(订单号): (Estado, Descripción) }，取各期最新一份月度报告（后读覆盖）。
    ⚠️ ML 报告的**表头行位置随版本变，且表头语言随界面变**（202609 西语 / 202610+ 中文），
    所以这里按内容找表头行，列名中英双认。"""
    out = {}
    files = []
    for per in ("202610", "202611"):
        cand = sorted(glob.glob(os.path.join(REPO, "trunk", "data", "数据归属周期", per,
                                             "快照日期", "*", "美客多*站点*订单报告.xlsx")))
        if cand:
            files.append(cand[-1])
    print(f"  [ML] 候选报告 {len(files)} 份")
    for p in files:
        try:
            wb = openpyxl.load_workbook(p, read_only=True, data_only=True)
            ws = wb[wb.sheetnames[0]]
            rows = list(ws.iter_rows(values_only=True))
            wb.close()
            hdr_i = None
            for i, r in enumerate(rows[:15]):
                cells = [str(x).strip() if x is not None else "" for x in (r or [])]
                if "订单号" in cells or "# de venta" in cells:
                    hdr_i = i
                    hdr = cells
                    break
            if hdr_i is None:
                print(f"  [ML] 跳过 {os.path.basename(p)}: 找不到表头行")
                continue
            i_oid = hdr.index("订单号") if "订单号" in hdr else hdr.index("# de venta")
            i_st = next((hdr.index(k) for k in ("状态", "Estado") if k in hdr), None)
            i_de = next((hdr.index(k) for k in ("状态描述", "Descripción del estado") if k in hdr), None)
            n = 0
            for r in rows[hdr_i + 1:]:
                if not r or len(r) <= i_oid or r[i_oid] is None:
                    continue
                out[str(r[i_oid]).strip()] = (
                    str(r[i_st] or "") if i_st is not None and i_st < len(r) else "",
                    str(r[i_de] or "") if i_de is not None and i_de < len(r) else "")
                n += 1
            print(f"  [ML] 读报告 {os.path.basename(p)}: {n} 单")
        except Exception as e:
            print(f"  [ML] 跳过 {os.path.basename(p)}: {e!r}")
    return out


def fix_ml_stale_status(paths, repmap, apply=False):
    """报告判定「不计销」、明细却还是「计销」→ 用报告的状态/描述覆盖明细（只改文本，不动数量）。"""
    _ML_UNCANCEL_KW = ("cancelad", "canceled", "cancelled", "devoluci", "reembols",
                       "取消", "退货", "退款")
    import openpyxl.worksheet.formula  # noqa: F401  (确保数组公式类可用)
    for p in paths:
        wb = openpyxl.load_workbook(p)
        ws = wb["订单数据"]
        hr, h = _hdr(ws, ("# de venta", "Estado", "Descripción del estado"))
        if not hr:
            print(f"  [ML] {os.path.basename(p)}: 表头识别失败，跳过")
            wb.close()
            continue
        c_st, c_de = h["Estado"], h["Descripción del estado"]
        # 关键词表取本簿「墨西哥透视图」
        kw = []
        if "墨西哥透视图" in wb.sheetnames:
            wsp = wb["墨西哥透视图"]
            for rr in range(2, 60):
                k = wsp.cell(row=rr, column=11).value
                if k is None:
                    break
                kw.append((str(k), wsp.cell(row=rr, column=12).value))
        n_fix = 0
        q_fix = 0.0
        samples = []
        for r in range(hr + 1, ws.max_row + 1):
            oid = ws.cell(r, h["# de venta"]).value
            if oid is None:
                continue
            got = repmap.get(str(oid).strip())
            if not got:
                continue
            st_r, de_r = got
            st_r_n = R.ml_norm_status(st_r, SNAP_DATE) or st_r
            de_r_n = R.ml_norm_desc(de_r, SNAP_DATE)
            # 只同步「取消/退货/退款」这类**两套口径都判不计销**的终态。
            # `已延迟/Está demorado` 只有日销口径判 0（K36），月销表算有效 —— 那是口径差，
            # 不属于滞后取消，硬改会把日销推向另一个方向，故不在这里处理。
            if not any(k in f"{st_r} {st_r_n}".lower() for k in _ML_UNCANCEL_KW):
                continue
            if R._ml_validity(kw, de_r_n, st_r_n, "x") not in (0, "0"):
                continue                     # 报告侧不是「不计销」，不动
            st_d = str(ws.cell(r, c_st).value or "")
            de_d = str(ws.cell(r, c_de).value or "")
            if R._ml_validity(kw, de_d, st_d, "x") in (0, "0"):
                continue                     # 明细侧已判不计销，无需改
            try:
                qv = float(ws.cell(r, 7).value or 0)
            except Exception:
                qv = 0.0
            n_fix += 1
            q_fix += qv
            if len(samples) < 4:
                samples.append((str(oid).strip(), st_d[:22], st_r_n[:26], qv))
            if apply:
                ws.cell(r, c_st).value = st_r_n
                ws.cell(r, c_de).value = de_r_n
        tag = "" if apply else "[预演] "
        print(f"  {tag}[ML] {os.path.basename(p)}: 状态回写 {n_fix} 行"
              f"（{q_fix:.0f} 件停止计销）例={samples}")
        if apply and n_fix:
            _backup(p)
            wb.save(p)
        wb.close()


# ---------------------------------------------------------------- 校验

def report_month(paths, kind):
    """统计各簿 10 月「有效件数」，用于修前修后对照。"""
    for p in paths:
        wb = openpyxl.load_workbook(p, read_only=True, data_only=True)
        ws = wb["订单数据"]
        if kind == "emag":
            tot = n = 0
            for r in ws.iter_rows(min_row=2, values_only=True):
                if not r or r[0] is None:
                    continue
                if not str(r[1] or "")[:7].startswith("2026-10"):
                    continue
                st = str(r[12] or "")
                if "Canceled" in st or "Storno" in st:
                    continue
                try:
                    tot += float(r[7] or 0)
                except Exception:
                    pass
                n += 1
            print(f"    {os.path.basename(p)}: 10月有效 {tot:.0f} 件 / {n} 行")
        else:
            kw = ["退款", "退货", "取消", "reembols", "devoluci", "cancelad", "anulaci"]
            tot = v1 = 0
            kwtab = []
            wb2 = openpyxl.load_workbook(p, read_only=False, data_only=True)
            if "墨西哥透视图" in wb2.sheetnames:
                wsp = wb2["墨西哥透视图"]
                for rr in range(2, 60):
                    k = wsp.cell(row=rr, column=11).value
                    if k is None:
                        break
                    kwtab.append((str(k), wsp.cell(row=rr, column=12).value))
            wb2.close()
            for r in list(ws.iter_rows(min_row=2, values_only=True)):
                if not r or r[0] is None:
                    continue
                if "de octubre de 2026" not in str(r[1] or ""):
                    continue
                try:
                    q = float(r[6] or 0)
                except Exception:
                    q = 0
                if any(k in str(r[2] or "").lower() for k in kw):
                    continue
                tot += q
                if R._ml_validity(kwtab, str(r[3] or ""), str(r[2] or ""), r[18]) in (1, "1"):
                    v1 += q
            print(f"    {os.path.basename(p)}: 10月月销口径 {tot:.0f} 件 / 日销有效性=1 {v1:.0f} 件")
        wb.close()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--apply", action="store_true")
    ap.add_argument("--only", default="all", choices=("all", "emag", "ml"))
    a = ap.parse_args()

    if a.only in ("all", "emag"):
        dates, statuses, nf = emag_report_dates()
        print(f"① eMAG：扫到 {nf} 份订单报告，覆盖 {len(dates)} 个订单号")
        tg = _targets(EMAG_FILE)
        print("  [修前]")
        report_month(tg, "emag")
        fix_emag(tg, dates, statuses, apply=a.apply)
        if a.apply:
            print("  [修后]")
            report_month(tg, "emag")

    if a.only in ("all", "ml"):
        print("② 美客多：中文状态/描述重归一 + 滞后取消状态回写")
        tg = _targets(ML_FILE)
        print("  [修前]")
        report_month(tg, "ml")
        fix_ml(tg, apply=a.apply)
        fix_ml_stale_status(tg, ml_report_map(), apply=a.apply)
        if a.apply:
            print("  [修后]")
            report_month(tg, "ml")

    if a.apply:
        print(f"\n备份目录：{BACKUP}")


if __name__ == "__main__":
    main()
