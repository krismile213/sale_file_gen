# -*- coding: utf-8 -*-
"""把「日销量红/蓝染色带」从最左 8 列扩到全部日销列（2026-10-14 用户确认需要）。

背景：B1 图例规定「红色=当日销量>上月日均×2.8 且 上月日均>1；蓝色=销量<上月日均÷2.5」，
但这条条件格式的 sqref 只覆盖最左 8 个日销列（英国 AI7:AP、欧洲 AS7:AZ、其余 Y7:AF），
而日列实际有 39 列（AI..BU / AS..CE / Y..BK）→ 只有窗口最左 8 天（9/24~10/1）会出现红/蓝标。

做法：**只改 sheet XML 里对应 <conditionalFormatting> 的 sqref 列范围**（首列不动、末列扩到最后一个
日销列；行范围不变）。不动规则公式、不动 dxf、不动样式、不动数据 —— 逐字节最小改动。
CF 公式里的相对引用（如 AI7）以 sqref 左上角为锚，扩展区域后 Excel 自动按列平移，$AE7 绝对列不变。

用法：
  python extend_daycf_band_20261014.py            # 干跑：列出每表将扩的 sqref
  python extend_daycf_band_20261014.py --apply    # 落盘（先备份）
"""
import os
import re
import shutil
import zipfile
import datetime

import openpyxl

DSR = r"C:\Users\dell\Desktop\work-main\sale_file_gen\Daily_Sales_Report"
WORK = os.path.join(DSR, "1.每日销售数据表")
BAK = os.path.join(DSR, ".workbuddy", "backup_daycf_20261014")
APPLY = "--apply" in os.sys.argv


def col_idx(letter):
    n = 0
    for ch in letter:
        n = n * 26 + (ord(ch) - 64)
    return n


def col_letter(n):
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def attrs(t):
    return dict(re.findall(r'([\w:]+)="([^"]*)"', t))


def sheet_map(z):
    wb = z.read("xl/workbook.xml").decode("utf-8", "replace")
    rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8", "replace")
    rid2t = {}
    for m in re.finditer(r"<Relationship\b[^>]*/>", rels):
        a = attrs(m.group(0))
        rid2t[a.get("Id", "")] = a.get("Target", "")
    out = {}
    for m in re.finditer(r"<sheet\b[^>]*/>", wb):
        a = attrs(m.group(0))
        t = rid2t.get(a.get("r:id", "") or a.get("id", ""), "")
        if t.startswith("/"):
            t = t[1:]
        elif not t.startswith("xl/"):
            t = "xl/" + t
        out[a.get("name", "")] = t
    return out


def day_range(path, sheet):
    """该分析表日销列范围 (first, last)：行5 为 datetime 的列。read_only 流式，只读行5。"""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=False)
    try:
        ws = wb[sheet]
        vals = next(ws.iter_rows(min_row=5, max_row=5, values_only=True), ())
        cols = [i + 1 for i, v in enumerate(vals) if isinstance(v, datetime.datetime)]
        return (min(cols), max(cols)) if cols else (None, None)
    finally:
        wb.close()


def extend_sqref(sqref, fc, lc):
    """把 sqref 中「首列 == fc」的区间扩到 fc..lc（行不变）。返回 (新 sqref, 是否改动)。"""
    parts, changed = [], False
    for p in sqref.split():
        m = re.match(r"^\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)$", p)
        if m and col_idx(m.group(1)) == fc:
            parts.append(f"{col_letter(fc)}{m.group(2)}:{col_letter(lc)}{m.group(4)}")
            changed = True
        else:
            parts.append(p)
    return " ".join(parts), changed


def process_book(f):
    path = os.path.join(WORK, f)
    with zipfile.ZipFile(path) as z:
        smap = sheet_map(z)
        parts = {i.filename: z.read(i.filename) for i in z.infolist()}
        order = [i.filename for i in z.infolist()]
    print(f"\n==== {f}")
    hits = []
    for sn, part in smap.items():
        if "分析" not in sn or part not in parts:
            continue
        fc, lc = day_range(path, sn)
        if not fc:
            print(f"    [{sn}] 未识别日列，跳过")
            continue
        xml = parts[part].decode("utf-8", "replace")
        changed_here = []

        def repl(m):
            old = m.group(2)
            new, ch = extend_sqref(old, fc, lc)
            if ch:
                changed_here.append((old, new))
            return m.group(1) + new + m.group(3)

        new_xml = re.sub(r'(<conditionalFormatting sqref=")([^"]*)(")', repl, xml)
        if changed_here:
            print(f"    [{sn}] 日列 {col_letter(fc)}..{col_letter(lc)}：")
            for old, new in changed_here:
                print(f"        sqref {old}  →  {new}")
            parts[part] = new_xml.encode("utf-8")
            hits.append((sn, changed_here))
        else:
            print(f"    [{sn}] 日列 {col_letter(fc)}..{col_letter(lc)}：无「从首列起」的 CF 块，跳过")
    if APPLY and hits:
        try:
            os.makedirs(BAK, exist_ok=True)
            dst_bak = os.path.join(BAK, f)
            if not os.path.exists(dst_bak):
                shutil.copy2(path, dst_bak)
            tmp = path + ".cf.tmp.xlsx"
            with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zo:
                for n in order:
                    zo.writestr(n, parts[n])
            assert zipfile.ZipFile(tmp).testzip() is None, "zip 校验失败"
            try:
                os.replace(tmp, path)
            except OSError:
                shutil.copyfile(tmp, path)
                os.remove(tmp)
            print(f"    → 已落盘（{len(hits)} 张表）")
        except Exception as e:
            print(f"    !! 落盘失败（多半文件被 WPS/Excel/预览占用）：{e!r}")
            return 0
    return len(hits)


total = 0
for f in sorted(x for x in os.listdir(WORK) if x.endswith(".xlsx")):
    total += process_book(f)
print(f"\n合计处理 {total} 张分析表。")
if not APPLY:
    print("（干跑，未改盘；加 --apply 落盘）")
