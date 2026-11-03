# -*- coding: utf-8 -*-
"""修复欧盟簿两张分析表 1~7 月的整行错复制（2026-10-27）。

问题（`tools/_scan_dup_month_rows_20261027.py` 扫出）：
  · 英国分析表 7 个 SKU 的 1~7 月值完全相同 = [280,250,288,242,222,251,276]
    （与第 7 行 G415 的 1~7 月一模一样 → 明显是从首行整行复制下来）
  · 欧洲分析表 G598/K001 与 G600/G601 共用 [0,8,493,363,372,347,420]

做法：**只改这 11 个 SKU 的 1~7 月**，值取各自月份归档簿，
口径直接调跑批自己的 `analysis.resolve_month_sales()`（含 amzn.gr 归并、派生地区映射），
第 6 行合计按 (新值-旧值) 的差额修正，绝不整格重写。

用法：python fix_dup_month_rows_20261027.py [--apply]
"""
import argparse, os, re, shutil, sys, zipfile
from xml.etree import ElementTree as ET

HERE = os.path.dirname(os.path.abspath(__file__))
CORE = os.path.abspath(os.path.join(HERE, "..", "..", "core"))
sys.path.insert(0, CORE)
sys.path.insert(0, HERE)

import analysis as An
import config as C
from rebuild_monthly_from_daily_20261027 import (
    XLSX, smap, sst_of, month_cols, sku_rows, get_cells, patch_sheet, set_total)

TARGETS = {
    "英国-每日库存分析表": ["G598", "G212", "G272", "G302", "K001", "K004", "K005"],
    "欧洲-每日库存分析表": ["G598", "K001", "G600", "G601"],
}
MONTHS = list(range(1, 8))
BACKUP = os.path.join(C.ROOT if hasattr(C, "ROOT") else C.DSR, ".workbuddy",
                      "backup_dup_rows_20261027")


def arch_path(m):
    return os.path.join(C.WORK_DIR, "2026年每日数据", "2026年%d月" % m,
                        "2026年%d月-欧盟英国库存销售表.xlsx" % m)


def month_sales_from_archive(m, sheet):
    """从第 m 月归档簿重算该表口径的 {sku: qty}。"""
    from openpyxl import load_workbook
    p = arch_path(m)
    if not os.path.exists(p):
        return None, "归档缺失"
    wb = load_workbook(p, read_only=True, data_only=False)
    try:
        if sheet not in wb.sheetnames or "每日销售数据" not in wb.sheetnames:
            return None, "归档缺 sheet"
        sc = An.scan_analysis(wb[sheet])
        if not sc:
            return None, "归档口径解析失败"
        return An.resolve_month_sales(wb["每日销售数据"], "2026-%02d" % m,
                                      sc["spec"], sc["d4"], sc["d5"]), None
    finally:
        wb.close()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--apply", action="store_true")
    a = ap.parse_args()

    z = zipfile.ZipFile(XLSX)
    sm = smap(z); sst = sst_of(z)
    parts = {n: z.read(n) for n in z.namelist()}
    z.close()

    for sheet, skus in TARGETS.items():
        sp = sm[sheet]
        x = parts[sp].decode("utf-8")
        cols = month_cols(x, sst)
        rows = sku_rows(x, sst)
        row_of = {v: k for k, v in rows.items()}
        print("\n=== %s ===" % sheet)

        # 逐月取新值
        newv = {}
        for m in MONTHS:
            ms, err = month_sales_from_archive(m, sheet)
            if ms is None:
                print("  ! %d月 %s —— 跳过该月" % (m, err))
                continue
            newv[m] = ms

        plan, delta = {}, {m: 0 for m in newv}
        for sku in skus:
            r = row_of.get(sku)
            if not r:
                print("  ! %s 无此行" % sku)
                continue
            old = get_cells(x, r, [cols[m] for m in newv])
            cell = {}
            parts_s = []
            for m in newv:
                o = old.get(cols[m]) or 0
                n = int(newv[m].get(sku, 0))
                cell[cols[m]] = n
                delta[m] += n - int(o)
                parts_s.append("%d月 %d→%d" % (m, int(o), n))
            plan[r] = cell
            print("  %-8s %s" % (sku, "  ".join(parts_s)))

        print("  合计修正: %s" % "  ".join("%d月%+d" % (m, d) for m, d in delta.items()))
        if not a.apply:
            continue
        x = patch_sheet(x, plan)
        for m, d in delta.items():
            if not d:
                continue
            cur = None
            mm = re.search(r'<c r="%s6"(?![0-9])[^>]*?(?:/>|>(.*?)</c>)' % cols[m], x, re.S)
            v = re.search(r"<v>(-?[\d.]+)</v>", mm.group(1) or "") if mm else None
            cur = int(float(v.group(1))) if v else 0
            x, ok = set_total(x, cols[m], cur + d)
            if not ok:
                print("  ! %d月 合计公式未找到" % m)
        ET.fromstring(x)
        parts[sp] = x.encode("utf-8")

    if not a.apply:
        print("\n（预览，未落盘）")
        return
    os.makedirs(BACKUP, exist_ok=True)
    shutil.copy2(XLSX, os.path.join(BACKUP, os.path.basename(XLSX)))
    tmp = XLSX + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zo:
        for n, d in parts.items():
            zo.writestr(n, d)
    try:
        os.replace(tmp, XLSX)
    except PermissionError:
        with open(XLSX, "wb") as f:
            f.write(open(tmp, "rb").read())
        os.remove(tmp)
    print("\n已落盘，备份:", BACKUP)


if __name__ == "__main__":
    main()
