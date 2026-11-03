# -*- coding: utf-8 -*-
"""封存前状态校正（方案 B，2026-11-02 用户拍板）

问题
----
`runner.append_sales()` 只按组合键 (order-item-id, amazon-order-id, sku, purchase-date)
判断「该行是否已存在」，存在就 `continue`，**从不回写状态**。订单第一次被抓到时
几乎都是 `Pending`，之后：

  · 变 Shipped   → 件数不变，无影响
  · 变 Cancelled → 亚马逊把报告里的 quantity 改成 0，但明细里那条仍是
                   `Pending / qty=1` → 日销列 SUMIFS 只按 SKU+日期求和、
                   **不筛状态**，于是照样算成销量，且永远不回退

实测（202610）：507 单 / 517 件，日销表比月销表高 0.95%。9 月同构（115 件）。
→ 这是稳定机制，不是偶发，必须校正。

方案 B（用户选定）：只在换月封存前校正
--------------------------------------
不碰日常跑批路径（日内日销表仍略高，可接受）。`month_rollover.build_archive()`
归档前调用本模块，用当期**最新订单报告的最终状态**把明细校正一遍：

  报告最终 Cancelled、明细行却不是 Cancelled
      → 明细 `order-status` 写 'Cancelled'，`quantity` 写 0

两列都要改：
  · 日销列 SUMIFS 只求和 quantity → 必须置 0 才算不到
  · 月汇总块 resolve_month_sales 筛 E<>"Cancelled" → 状态也要改
改完再重算并冻结该月列（`freeze_month_cols`），归档簿的月汇总值才跟着降。

只改这两列，不改行数、不动公式、不增删行。

用法
----
    from status_sync import load_final_status, fix_sales_sheet, freeze_month_cols

    rep, snap, nfile = load_final_status("202610")
    st = fix_sales_sheet(wb["每日销售数据"], rep)      # 就地改
    fr = freeze_month_cols(wb, "2026-10")              # 重算月汇总列

命令行（修某一期已经封存的归档簿）：
    python pipeline/core/status_sync.py --period 202610 --apply
    python pipeline/core/status_sync.py --period 202610            # 预演
"""
import os
import sys
import glob
import argparse

import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))        # pipeline/core
DSR = os.path.dirname(os.path.dirname(HERE))             # Daily_Sales_Report
REPO = os.path.dirname(DSR)
if HERE not in sys.path:
    sys.path.insert(0, HERE)

# 明细 sheet 必备列（亚马逊系报表口径一致）
NEED = ("amazon-order-id", "purchase-date", "order-status", "sku", "quantity")


# ---------------------------------------------------------------- 报告侧

def _snap_root(period):
    return os.path.join(REPO, "trunk", "data", "数据归属周期", period, "快照日期")


def list_snaps(period):
    base = _snap_root(period)
    if not os.path.isdir(base):
        return []
    return sorted(d for d in os.listdir(base)
                  if os.path.isdir(os.path.join(base, d)))


def _read_report(path):
    """一份订单报告 → {(amazon-order-id, sku, purchase-date[:19]): status}
    只收 sales-channel 以 Amazon. 开头的行（eMAG/美客多不走这份报告）。"""
    try:
        f = open(path, "r", encoding="utf-8", errors="replace")
    except OSError:
        return None
    with f:
        hdr = f.readline().rstrip("\n").split("\t")
        idx = {h.strip(): i for i, h in enumerate(hdr)}
        if any(h not in idx for h in NEED):
            return None
        i_aoi, i_pd, i_st, i_ch, i_sku = (idx["amazon-order-id"], idx["purchase-date"],
                                          idx["order-status"], idx["sales-channel"],
                                          idx["sku"])
        mx = max(idx[h] for h in NEED)
        out = {}
        for line in f:
            p = line.rstrip("\n").split("\t")
            if len(p) <= mx:
                continue
            ch = p[i_ch].strip()
            if not ch.startswith("Amazon."):
                continue
            sku = p[i_sku].strip()
            pdv = p[i_pd].strip()
            if not sku or not pdv:
                continue
            out[(p[i_aoi].strip(), sku, pdv[:19])] = p[i_st].strip()
        return out


def load_final_status(period, snap=None):
    """当期最新（或指定）快照的全部订单最终状态。
    返回 (dict, snap_name, 报告份数)。dict 键与 append_sales 的组合键一致。"""
    snaps = list_snaps(period)
    if not snaps:
        raise SystemExit("找不到快照目录：%s" % _snap_root(period))
    snap = snap or snaps[-1]
    d = os.path.join(_snap_root(period), snap)
    if not os.path.isdir(d):
        raise SystemExit("快照不存在：%s" % d)
    rep, nfile = {}, 0
    for fn in sorted(os.listdir(d)):
        if not fn.endswith("订单报告.txt"):
            continue
        got = _read_report(os.path.join(d, fn))
        if got is None:
            continue
        nfile += 1
        rep.update(got)          # 同键以最后读到为准（英国重复报告不会翻倍）
    return rep, snap, nfile


# ---------------------------------------------------------------- 明细侧

def _det_sheet(wb):
    for sn in wb.sheetnames:
        if sn.startswith("每日销售数据"):
            return wb[sn]
    return None


def _cell_day(v):
    """明细里的 purchase-date → 取前 19 位（保留 T 分隔，与报告原文对齐）"""
    if v is None:
        return ""
    if hasattr(v, "strftime"):
        return v.strftime("%Y-%m-%dT%H:%M:%S")
    return str(v).strip()[:19]


def fix_sales_sheet(ws, rep, dry=False):
    """就地校正一张明细 sheet。返回统计 dict（scanned / fixed / qty / 样例）。"""
    hdr = head = None
    for r in range(1, 8):
        names = [(str(c.value).strip() if c.value is not None else "")
                 for c in ws[r]]
        if all(k in names for k in ("amazon-order-id", "purchase-date", "order-status")):
            hdr = r
            head = {n: i + 1 for i, n in enumerate(names) if n}
            break
    if hdr is None:
        return dict(sheet=ws.title, ok=False, err="表头缺关键列")
    c_aoi, c_pd, c_st = (head["amazon-order-id"], head["purchase-date"],
                         head["order-status"])
    c_sku = head.get("sku", 12)
    c_q = head.get("quantity", 15)
    lo, hi = min(c_aoi, c_pd, c_st, c_sku, c_q), max(c_aoi, c_pd, c_st, c_sku, c_q)
    # 索引换算（iter_rows 从 lo 列起，元组下标要减 lo）
    o_aoi, o_pd, o_st, o_sku, o_q = (c_aoi - lo, c_pd - lo, c_st - lo,
                                     c_sku - lo, c_q - lo)
    n_scan = n_fix = 0
    q_fix = 0.0
    n_rep_cancel = 0          # 报告说取消、明细也已取消（无需动）
    n_missing = 0             # 报告里查不到该行（老单/已删单，不动）
    samples = []
    for row in ws.iter_rows(min_row=hdr + 1, max_row=ws.max_row,
                            min_col=lo, max_col=hi):
        vals = [c.value for c in row]
        aoi = vals[o_aoi]
        if aoi is None:
            continue
        n_scan += 1
        key = (str(aoi).strip(), str(vals[o_sku] or "").strip(),
               _cell_day(vals[o_pd]))
        st_rep = rep.get(key)
        if st_rep is None:
            n_missing += 1
            continue
        if st_rep != "Cancelled":
            continue
        n_rep_cancel += 1
        st_now = str(vals[o_st] or "").strip()
        if st_now == "Cancelled":
            continue
        qv = vals[o_q]
        try:
            qv = float(qv or 0)
        except (TypeError, ValueError):
            qv = 0.0
        n_fix += 1
        q_fix += qv
        if len(samples) < 5:
            samples.append((str(aoi).strip(), key[1], st_now, qv))
        if not dry:
            row[o_st].value = "Cancelled"
            row[o_q].value = 0
    return dict(sheet=ws.title, ok=True, scanned=n_scan, fixed=n_fix, qty=q_fix,
                rep_cancel=n_rep_cancel, missing=n_missing, samples=samples)


# ---------------------------------------------------------------- 月汇总

def _col_static_sum(ws, col, first_row=7, max_row=3200):
    """读某月汇总列的现有静态值合计（改前的对照）。"""
    s = 0.0
    for r in range(first_row, first_row + max_row):
        if ws.cell(r, 1).value is None:
            break
        v = ws.cell(r, col).value
        if isinstance(v, (int, float)):
            s += v
    return s


def freeze_month_cols(wb, ym, dry=False):
    """重算并冻结 wb 内所有分析表的 ym 月列（源=本簿明细，已校正）。
    返回 [dict(sheet, col, before, after, n)]"""
    import analysis as An
    m = int(ym[5:7])          # ym = "YYYY-MM"
    out = []
    for sn in list(wb.sheetnames):
        if "每日库存分析表" not in sn:
            continue
        ws = wb[sn]
        try:
            sc = An.scan_analysis(ws)
        except Exception as e:
            out.append(dict(sheet=sn, err="扫描失败 %r" % (e,)))
            continue
        if not sc:
            out.append(dict(sheet=sn, err="无法解析日销口径"))
            continue
        col = sc["month_cols"].get(m)
        if col is None:
            out.append(dict(sheet=sn, err="行5 无「%d月」标签列" % m))
            continue
        src = sc["spec"].get("src")
        if src not in wb.sheetnames:
            out.append(dict(sheet=sn, err="源表 %r 不存在" % (src,)))
            continue
        ms = An.resolve_month_sales(wb[src], ym, sc["spec"], sc["d4"], sc["d5"])
        before = _col_static_sum(ws, col)
        if not dry:
            n = An.freeze_month(ws, col, ms)
        else:
            n = len([1 for v in ms.values() if v])
        out.append(dict(sheet=sn, col=col, before=before,
                        after=sum(ms.values()), n=n))
    return out


# ---------------------------------------------------------------- 整簿 / CLI

def fix_workbook(path, period, rep=None, snap=None, dry=False, freeze=True):
    """校正一本簿（就地）。返回 (wb, 明细统计, 月汇总统计)。dry=True 不落盘。"""
    if rep is None:
        rep, snap, _ = load_final_status(period, snap)
    wb = openpyxl.load_workbook(path, data_only=False)
    ws = _det_sheet(wb)
    if ws is None:
        wb.close()
        return None, dict(ok=False, err="无「每日销售数据」sheet"), []
    st = fix_sales_sheet(ws, rep, dry=dry)
    fr = freeze_month_cols(wb, "%s-%s" % (period[:4], period[4:6]),
                           dry=dry) if freeze else []
    if not dry:
        tmp = path + ".new.xlsx"
        wb.save(tmp)
        try:
            from runner import restore_parts
            restore_parts(path, tmp)
        except Exception as e:
            print("  [!] restore_parts 跳过：%r" % (e,))
        os.replace(tmp, path)
    wb.close()
    return None, st, fr


def _archive_files(period):
    y, m = period[:4], int(period[4:6])
    d = os.path.join(DSR, "1.每日销售数据表", "%s年每日数据" % y, "%s年%d月" % (y, m))
    return sorted(glob.glob(os.path.join(d, "%s年%d月-*.xlsx" % (y, m))))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--period", required=True, help="YYYYMM，如 202610")
    ap.add_argument("--apply", action="store_true", help="实跑（默认只预演）")
    ap.add_argument("--snap", default="", help="指定快照日期，默认最新")
    ap.add_argument("--no-freeze", action="store_true", help="只校正明细，不重算月汇总")
    a = ap.parse_args()
    dry = not a.apply

    rep, snap, nfile = load_final_status(a.period, a.snap or None)
    print("订单报告：快照 %s / %d 份 / 合并 %d 单" % (snap, nfile, len(rep)))
    files = _archive_files(a.period)
    if not files:
        print("找不到归档簿（{Y}年{M}月-*.xlsx）：%s" % a.period)
        return 2
    print("归档簿 %d 本" % len(files))
    tot_fix = tot_q = 0.0
    for p in files:
        _, st, fr = fix_workbook(p, a.period, rep=rep, dry=dry,
                                 freeze=not a.no_freeze)
        name = os.path.basename(p)
        if not st.get("ok"):
            print("  %-42s 跳过：%s" % (name, st.get("err")))
            continue
        tot_fix += st["fixed"]
        tot_q += st["qty"]
        print("  %-42s 扫%6d行 校正%4d行 / %5.0f件"
              % (name, st["scanned"], st["fixed"], st["qty"]))
        for x in fr:
            if "err" in x:
                continue
            d = x["after"] - x["before"]
            flag = "" if abs(d) < 1 else "  Δ%+.0f" % d
            print("      [月汇总] %-22s 列%s  %8.0f → %8.0f%s"
                  % (x["sheet"], openpyxl.utils.get_column_letter(x["col"]),
                     x["before"], x["after"], flag))
    print("\n合计校正 %d 行 / %.0f 件%s" % (tot_fix, tot_q, "（预演，未落盘）" if dry else ""))
    return 0


if __name__ == "__main__":
    sys.exit(main())
