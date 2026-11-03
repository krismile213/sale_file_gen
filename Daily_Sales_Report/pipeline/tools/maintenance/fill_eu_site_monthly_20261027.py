# -*- coding: utf-8 -*-
"""
补齐欧盟英国簿 10 张站点表的 2026 年 8/9/10 月销量（2026-10-27）。

问题：意大利/法国/西班牙/德国/土耳其/荷兰/瑞典/波兰/比利时/爱尔兰
      这 10 张表的「12月..1月」月度块（2026 年）只有 1~7 月有值，
      8/9/10 月整列为 0。而英国/欧洲两张分析表的月度块是齐的。

数据源（已逐项验证）：
  · 8 月、9 月 ← trunk/result/数据归属周期/202609/月销售总表.xlsx
                  sheet「2026年-欧洲各国每月销售情况」，每 11 列一个月份块
                  （10 国 + 1 空列），块序 block_k = (k+1) 月。
                  验证：block6(7月) 10 个国别合计与站点表 7 月列**逐个精确相等**。
  · 10 月     ← 本簿「每日销售数据」sheet，有效性(AQ)=1 且 purchase-date 在 2026-10，
                  按 sales-channel 分站点、按 sku 汇总 quantity。
                  验证：欧洲 9 国合计 = 21007，与「欧洲-每日库存分析表」10月列完全相等；
                        德国 = 5942，与该表「10月总销量」完全相等。

为什么用 XML 手术：本簿 10.6MB，openpyxl 非 read_only 加载会 SIGTERM。

用法：
    python fill_eu_site_monthly_20261027.py            # 预览（只打印，不写）
    python fill_eu_site_monthly_20261027.py --apply    # 落盘（先备份）
"""
import argparse
import os
import re
import shutil
import zipfile
from collections import defaultdict
from xml.etree import ElementTree as ET

import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.abspath(os.path.join(HERE, "..", "..", "..", ".."))
XLSX = os.path.join(ROOT, "Daily_Sales_Report", "1.每日销售数据表",
                    "欧盟英国库存销售表.xlsx")
MONTH_FILE = os.path.join(ROOT, "trunk", "result", "数据归属周期", "202609",
                          "月销售总表.xlsx")
BACKUP = os.path.join(ROOT, "Daily_Sales_Report", ".workbuddy",
                      "backup_site_monthly_20261027")

DATA_FIRST_ROW = 7
# 站点表名 → (月销售总表国别列序号 0 基, 每日销售数据 sales-channel)
SITES = [
    ("意大利", 5, "Amazon.it"),
    ("法国", 4, "Amazon.fr"),
    ("西班牙", 3, "Amazon.es"),
    ("德国", 2, "Amazon.de"),
    ("土耳其", 1, "Amazon.com.tr"),
    ("荷兰", 6, "Amazon.nl"),
    ("瑞典", 8, "Amazon.se"),
    ("波兰", 7, "Amazon.pl"),
    ("比利时", 0, "Amazon.com.be"),
    ("爱尔兰", 9, "Amazon.ie"),
]
TARGET_MONTHS = (8, 9, 10)


# ---------------- 通用 ----------------
def gl(n):
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def ci(s):
    n = 0
    for ch in s:
        n = n * 26 + (ord(ch) - 64)
    return n


def sheet_map(z):
    rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    rid2t = dict((m[0], m[1]) for m in
                 re.findall(r'Id="(rId\d+)"[^>]*Target="([^"]+)"', rels))
    wb = z.read("xl/workbook.xml").decode("utf-8")
    sh = re.findall(r'<sheet [^>]*name="([^"]+)"[^>]*r:id="(rId\d+)"', wb)
    return {n: ("xl/" + rid2t[r].lstrip("/")) for n, r in sh}


def sst_of(z):
    if "xl/sharedStrings.xml" not in z.namelist():
        return []
    x = z.read("xl/sharedStrings.xml").decode("utf-8")
    return ["".join(re.findall(r"<t[^>]*>(.*?)</t>", si, re.S))
            for si in re.findall(r"<si>(.*?)</si>", x, re.S)]


def cell_val(x, ref, sst):
    m = re.search(r'<c r="%s"([^>]*?)(?:/>|>(.*?)</c>)' % ref, x, re.S)
    if not m:
        return None
    attrs, body = m.group(1), m.group(2) or ""
    if 't="s"' in attrs:
        v = re.search(r"<v>(\d+)</v>", body)
        return sst[int(v.group(1))] if v else None
    if 't="inlineStr"' in attrs:
        return "".join(re.findall(r"<t[^>]*>(.*?)</t>", body, re.S))
    v = re.search(r"<v>(.*?)</v>", body, re.S)
    return v.group(1) if v else None


# ---------------- 数据源 ----------------
def load_monthly():
    """{(国别序号, sku): {8: n, 9: n}} —— 从月销售总表 202609 取 8/9 月"""
    wb = openpyxl.load_workbook(MONTH_FILE, read_only=True, data_only=True)
    ws = wb["2026年-欧洲各国每月销售情况"]
    rows = list(ws.iter_rows(min_row=2, max_row=ws.max_row, values_only=True))
    r2 = rows[0]
    starts = [i for i, v in enumerate(r2)
              if v and ("（" in str(v) or "(" in str(v))
              and (i == 0 or not r2[i - 1])]
    if len(starts) < 8:
        raise RuntimeError("月销售总表月份块不足 8 个：%d" % len(starts))
    blk = {8: starts[6], 9: starts[7]}      # block_k = (k+1) 月
    out = defaultdict(dict)
    for r in rows[1:]:
        sku = r[0]
        if not sku or not isinstance(sku, str):
            continue
        sku = sku.strip()
        for m, s in blk.items():
            for j in range(10):
                v = r[s + j]
                if isinstance(v, (int, float)) and v:
                    out[(j, sku)][m] = int(v)
    wb.close()
    return out


def load_october():
    """{(sales-channel, sku): 10月销量}"""
    wb = openpyxl.load_workbook(XLSX, read_only=True, data_only=True)
    ws = wb["每日销售数据"]
    out = defaultdict(int)
    for r in ws.iter_rows(min_row=3, max_col=44, values_only=True):
        if not r[0] or r[42] != 1:
            continue
        if str(r[2])[:7] != "2026-10":
            continue
        sku, ch, q = r[11], r[6], r[14]
        if not sku or not isinstance(q, (int, float)):
            continue
        out[(ch, str(sku).strip())] += int(q)
    wb.close()
    return out


# ---------------- 站点表定位 ----------------
def month_block(x, sst):
    """返回 (起始列号, 结束列号)，即 2026 年 12月..1月 那一段（1 基）"""
    # 找第 5 行全部 "N月" 表头
    m = re.search(r'<row r="5"[^>]*>(.*?)</row>', x, re.S)
    if not m:
        raise RuntimeError("找不到第 5 行")
    heads = {}
    for c in re.finditer(r'<c r="([A-Z]+)5"([^>]*?)(?:/>|>(.*?)</c>)', m.group(1), re.S):
        col, attrs, body = c.group(1), c.group(2), c.group(3) or ""
        if 't="s"' in attrs:
            v = re.search(r"<v>(\d+)</v>", body)
            txt = sst[int(v.group(1))] if v else None
        elif 't="inlineStr"' in attrs:
            txt = "".join(re.findall(r"<t[^>]*>(.*?)</t>", body, re.S))
        else:
            txt = None
        if txt:
            heads[ci(col)] = txt.strip()
    # 找连续 12 列 12月..1月
    cands = []
    for c, t in heads.items():
        if t != "12月":
            continue
        if all(heads.get(c + k) == "%d月" % (12 - k) for k in range(12)):
            cands.append(c)
    if not cands:
        raise RuntimeError("找不到 12月..1月 月份块")
    if len(cands) == 1:
        return cands[0], cands[0] + 11
    # 两个块（2026 / 2025）：2026 块 12月 合计为 0
    for c in cands:
        v = cell_val(x, "%s6" % gl(c), sst)
        if v in (None, "", "0"):
            return c, c + 11
    return cands[0], cands[0] + 11


def sku_rows(x, sst):
    """{行号: SKU}"""
    out = {}
    for m in re.finditer(r'<row r="(\d+)"', x):
        r = int(m.group(1))
        if r < DATA_FIRST_ROW:
            continue
        v = cell_val(x, "A%d" % r, sst)
        if v and isinstance(v, str):
            out[r] = v.strip()
    return out


# ---------------- 单元格写入（单遍按行，避免 O(n^2) 全局搜索） ----------------
CELL_RE = r'<c r="([A-Z]+)%d"(?:[^>]*?)(?:/>|>.*?</c>)'


def patch_row(inner, row, targets, style_src, style_col_idx=None):
    """targets = {列字母: 整数值}。就地改这一行的 XML，返回新 inner。"""
    segs, gaps, pos = [], [], 0
    for m in re.finditer(CELL_RE % row, inner, re.S):
        gaps.append(inner[pos:m.start()])
        segs.append([ci(m.group(1)), m.group(0)])
        pos = m.end()
    gaps.append(inner[pos:])

    if style_col_idx is not None:
        for c, txt in segs:
            if c == style_col_idx:
                sm = re.search(r'\ss="(\d+)"', txt)
                if sm:
                    style_src = ' s="%s"' % sm.group(1)
                break

    for col, val in targets.items():
        ref = "%s%d" % (col, row)
        idx = ci(col)
        hit = None
        for k, (c, _) in enumerate(segs):
            if c == idx:
                hit = k
                break
        if hit is not None:
            attrs = re.match(r'<c r="[A-Z]+\d+"([^>]*?)(?:/>|>)', segs[hit][1]).group(1)
            attrs = re.sub(r'\s+t="[^"]*"', "", attrs)
            segs[hit][1] = '<c r="%s"%s><v>%d</v></c>' % (ref, attrs, val)
        elif val != 0:
            new = '<c r="%s"%s><v>%d</v></c>' % (ref, style_src, val)
            ins = len(segs)
            for k, (c, _) in enumerate(segs):
                if c > idx:
                    ins = k
                    break
            segs.insert(ins, [idx, new])
            gaps.insert(ins, "")
    return "".join(gaps[i] + segs[i][1] for i in range(len(segs))) + gaps[-1]


def patch_sheet(x, plan, style_src, style_col_idx=None):
    """plan = {行号: {列字母: 值}}。整表单遍重写。"""
    out, pos = [], 0
    for m in re.finditer(r'<row r="(\d+)"[^>]*>(.*?)</row>', x, re.S):
        r = int(m.group(1))
        out.append(x[pos:m.start()])
        if r in plan:
            out.append(x[m.start():m.start(2)])
            out.append(patch_row(m.group(2), r, plan[r], style_src, style_col_idx))
            out.append("</row>")
        else:
            out.append(m.group(0))
        pos = m.end()
    out.append(x[pos:])
    return "".join(out)


def style_of(x, col, row):
    m = re.search(r'<c r="%s%d"([^>]*?)(?:/>|>)' % (col, row), x, re.S)
    if not m:
        return ""
    s = re.search(r'\ss="(\d+)"', m.group(1))
    return ' s="%s"' % s.group(1) if s else ""


# ---------------- 主流程 ----------------
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--apply", action="store_true")
    a = ap.parse_args()

    monthly = load_monthly()
    october = load_october()
    print("月销售总表 8/9 月记录 %d 条；每日销售数据 10 月记录 %d 条"
          % (len(monthly), len(october)))

    z = zipfile.ZipFile(XLSX)
    sm = sheet_map(z)
    sst = sst_of(z)
    parts = {}
    report = []

    for site, cidx, channel in SITES:
        sp = sm[site]
        x = z.read(sp).decode("utf-8")
        start, end = month_block(x, sst)
        col = {m: gl(start + (12 - m)) for m in range(1, 13)}
        rows = sku_rows(x, sst)
        hits = {m: 0 for m in TARGET_MONTHS}
        tot = {m: 0 for m in TARGET_MONTHS}
        plan = {}
        for r, sku in rows.items():
            cell = {}
            for m in TARGET_MONTHS:
                v = (october.get((channel, sku), 0) if m == 10
                     else monthly.get((cidx, sku), {}).get(m, 0))
                v = int(v)
                cell[col[m]] = v
                tot[m] += v
                if v:
                    hits[m] += 1
            plan[r] = cell
        x = patch_sheet(x, plan, "", ci(col[7]))
        # 第 6 行合计：缓存值同步
        for m in TARGET_MONTHS:
            ref = "%s6" % col[m]
            mm = re.search(r'<c r="%s"([^>]*?)(?:/>|>(.*?)</c>)' % ref, x, re.S)
            if mm:
                x = (x[:mm.start()] + '<c r="%s"%s><v>%d</v></c>'
                     % (ref, mm.group(1), tot[m]) + x[mm.end():])
        ET.fromstring(x)
        parts[sp] = x
        report.append((site, "%s..%s" % (gl(start), gl(end)),
                       len(rows), tot, hits))
        print("  %-5s 块 %-8s SKU %-4d 8月=%-6d 9月=%-6d 10月=%-6d 命中 %d/%d/%d"
              % (site, "%s..%s" % (gl(start), gl(end)), len(rows),
                 tot[8], tot[9], tot[10], hits[8], hits[9], hits[10]),
              flush=True)

    for n in z.namelist():
        if n not in parts:
            parts[n] = z.read(n)
    z.close()

    if a.apply:
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
    else:
        print("\n（预览，未落盘）")


if __name__ == "__main__":
    main()
