# -*- coding: utf-8 -*-
"""给站点表补行（2026-10-27 用户确认：销量是最核心的东西，要保证准确）。

背景：G598（Amazon.it 10月 1 件）和 K001（Amazon.se 10月 1 件）在 10 张站点表里
      一行都没有，导致站点表合计比分析表少 2 件 —— 补行闸只作用于分析表。

做法：复用跑批自己的 `analysis.append_sku_rows_auto()`（克隆列最完整的行、公式平移、
      CF 范围自动跟随），补完立刻调 `runner.apply_month_summary()` 让新行拿到
      上月封存 + 当月销量 + 上月日均。保存流程与跑批一致：
      strip_pivot_records → openpyxl 保存 → restore_parts 补回透视件。

用法：
    python append_site_rows_20261027.py            # 临时副本上跑，只看不落盘
    python append_site_rows_20261027.py --apply    # 落盘到正式文件
"""
import argparse, os, shutil, sys, tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
CORE = os.path.abspath(os.path.join(HERE, "..", "..", "core"))
sys.path.insert(0, CORE)
sys.path.insert(0, HERE)

import analysis as An
import config as C
import runner as R
from openpyxl import load_workbook

SNAP = "20261027"
FNAME = "欧盟英国库存销售表.xlsx"
SRC = os.path.join(C.WORK_DIR, FNAME)
BACKUP = os.path.join(C.DSR, ".workbuddy", "backup_site_append_20261027")
# 站点表 → 10 月有销量但缺行的 SKU（按日销数据核对得出）
PLAN = {"意大利": ["G598"], "瑞典": ["K001"]}
# 10 月列 / 上月日均列（瑞典等少「库存」列的表整体左移一列）
COLS = {"意大利": ("AX", "G", "AY"), "瑞典": ("AW", "F", "AX")}


def name_map_of(wb, skus):
    eu = wb["欧洲-每日库存分析表"]
    out = {}
    for r in range(7, (eu.max_row or 7) + 1):
        v = eu.cell(r, 1).value
        if v in skus and v not in out:
            out[v] = eu.cell(r, 2).value
    return out


def run(wb):
    all_skus = {s for v in PLAN.values() for s in v}
    nm = name_map_of(wb, all_skus)
    print("品名:", {k: (v or "")[:24] for k, v in nm.items()})
    for sn, skus in PLAN.items():
        ws = wb[sn]
        last = An._last_data_row(ws)
        added = An.append_sku_rows_auto(ws, skus, name_map=nm,
                                        state_map={s: "在售" for s in skus})
        print("  %-5s last=%d 追加 %s" % (sn, last, added))
    print("--- 月度封存（让新行拿到 10 月销量 / 上月日均）---")
    R.apply_month_summary(wb, SNAP, fname=FNAME)


def verify(path):
    wb = load_workbook(path, data_only=False)
    print("\n--- 复验 ---")
    for sn, skus in PLAN.items():
        ws = wb[sn]
        d10, gavg, d9 = COLS[sn]
        for sku in skus:
            r = next((i for i in range(7, (ws.max_row or 7) + 1)
                      if ws.cell(i, 1).value == sku), None)
            if not r:
                print("  ✗ %s 没有 %s" % (sn, sku)); continue
            print("  %-5s %-6s 行%-4d 10月(%s)=%-4s 9月(%s)=%-4s 上月日均(%s)=%s" % (
                sn, sku, r, d10, ws["%s%d" % (d10, r)].value,
                d9, ws["%s%d" % (d9, r)].value,
                gavg, str(ws["%s%d" % (gavg, r)].value)[:46]))
    wb.close()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--apply", action="store_true")
    a = ap.parse_args()

    if not a.apply:
        tmpd = tempfile.mkdtemp(prefix="siteappend_")
        cp = os.path.join(tmpd, FNAME)
        shutil.copy2(SRC, cp)
        st = cp + ".strip.xlsx"
        n = R.strip_pivot_records(cp, st)
        wb = load_workbook(st if n else cp, data_only=False)
        run(wb)
        out = os.path.join(tmpd, "out.xlsx")
        wb.save(out); wb.close()
        print("\n（预览：副本未回写正式文件）")
        verify(out)
        print("副本:", out)
        return

    os.makedirs(BACKUP, exist_ok=True)
    bak = os.path.join(BACKUP, FNAME)
    if not os.path.exists(bak):
        shutil.copy2(SRC, bak)
    print("备份:", bak)

    stripped = SRC + ".pivotstrip.xlsx"
    n = R.strip_pivot_records(SRC, stripped)
    wb = load_workbook(stripped if n else SRC, data_only=False)
    run(wb)
    try:
        wb.calculation.fullCalcOnLoad = True
        wb.calculation.calcMode = "auto"
    except Exception:
        pass
    tmp = SRC + ".tmp.xlsx"
    wb.save(tmp); wb.close()
    restored, piv = R.restore_parts(SRC, tmp)
    print("restore_parts: 补回 %d 项（pivot %d）" % (restored, piv))
    if n and os.path.exists(stripped):
        try: os.remove(stripped)
        except OSError: pass
    try:
        os.replace(tmp, SRC)
    except PermissionError:
        with open(SRC, "wb") as f:
            f.write(open(tmp, "rb").read())
        os.remove(tmp)
    print("已落盘:", SRC)
    verify(SRC)


if __name__ == "__main__":
    main()
