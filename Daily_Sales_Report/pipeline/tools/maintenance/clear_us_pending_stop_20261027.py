# -*- coding: utf-8 -*-
"""
清掉美国表 9 个被误判为「待停售」的新品标记（2026-10-27）。

背景：`北美加拿大库存销售表.xlsx` / `美国-每日库存分析表` C 列是
「商品分类 + 待停售」混合列，尾部 362~450 行整块被人工标了 待停售，
其中 378~386 这 9 个是新品（D=在售、可售 0、在途 0、初次销售时间空），
用户确认只清这 9 个，其余 80 行保留。

C 列清空为空格（保留 s= 样式，不动其它列）。

用法：
    python clear_us_pending_stop_20261027.py            # 预览
    python clear_us_pending_stop_20261027.py --apply    # 落盘（先备份）
"""
import argparse
import os
import re
import shutil
import zipfile
from xml.etree import ElementTree as ET

HERE = os.path.dirname(os.path.abspath(__file__))
XLSX = os.path.abspath(os.path.join(
    HERE, "..", "..", "..", "1.每日销售数据表", "北美加拿大库存销售表.xlsx"))
BACKUP = os.path.abspath(os.path.join(
    HERE, "..", "..", "..", ".workbuddy", "backup_clear_pending_20261027"))
SHEET = "美国-每日库存分析表"

# 待清行的 SKU 白名单（行号由 SKU 反查，不写死行号，避免插入行后错位）
TARGET_SKUS = ["G566", "G592", "G593", "G594", "G595", "G596", "G597",
               "G598", "K013-016"]


def sheet_path(z, want):
    rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    rid2t = dict((m[0], m[1]) for m in
                 re.findall(r'Id="(rId\d+)"[^>]*Target="([^"]+)"', rels))
    wb = z.read("xl/workbook.xml").decode("utf-8")
    for n, r in re.findall(r'<sheet [^>]*name="([^"]+)"[^>]*r:id="(rId\d+)"', wb):
        if n.strip() == want:
            return "xl/" + rid2t[r].lstrip("/")
    raise KeyError(want)


def sst_of(z):
    if "xl/sharedStrings.xml" not in z.namelist():
        return []
    x = z.read("xl/sharedStrings.xml").decode("utf-8")
    out = []
    for si in re.findall(r"<si>(.*?)</si>", x, re.S):
        out.append("".join(re.findall(r"<t[^>]*>(.*?)</t>", si, re.S)))
    return out


def cell_text(x, row, col, sst):
    """取某单元格文本（支持 t="s" 与 inlineStr）"""
    m = re.search(r'<c r="%s%d"([^>]*?)(?:/>|>(.*?)</c>)' % (col, row), x, re.S)
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


def find_rows(x, sst, skus):
    rows = {}
    for m in re.finditer(r'<row r="(\d+)"', x):
        rows.setdefault(int(m.group(1)), None)
    hit = {}
    for r in sorted(rows):
        a = cell_text(x, r, "A", sst)
        if a and a.strip() in skus:
            hit[a.strip()] = r
    return hit


def clear_cell(x, row, col):
    """把 C{row} 变成空单元格，保留 s= 样式"""
    pat = re.compile(r'<c r="%s%d"((?:\s+(?!t=)[a-zA-Z]+="[^"]*")*)(?:\s+t="[^"]*")?\s*(?:/>|>.*?</c>)'
                     % (col, row), re.S)
    m = pat.search(x)
    if not m:
        return x, False
    return x[:m.start()] + '<c r="%s%d"%s/>' % (col, row, m.group(1)) + x[m.end():], True


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--apply", action="store_true")
    a = ap.parse_args()

    z = zipfile.ZipFile(XLSX)
    sp = sheet_path(z, SHEET)
    sst = sst_of(z)
    x = z.read(sp).decode("utf-8")

    hit = find_rows(x, sst, set(TARGET_SKUS))
    missing = [s for s in TARGET_SKUS if s not in hit]
    print("命中 %d/%d 个 SKU" % (len(hit), len(TARGET_SKUS)))
    if missing:
        print("未找到:", missing)

    x2 = x
    for sku, r in sorted(hit.items(), key=lambda kv: kv[1]):
        before = cell_text(x2, r, "C", sst)
        x2, ok = clear_cell(x2, r, "C")
        print("  行%-4d %-10s C: %r -> %s" % (r, sku, before, "清空" if ok else "失败"))
    ET.fromstring(x2)   # 良构校验

    if a.apply:
        os.makedirs(BACKUP, exist_ok=True)
        shutil.copy2(XLSX, os.path.join(BACKUP, os.path.basename(XLSX)))
        tmp = XLSX + ".tmp"
        parts = {}
        for n in z.namelist():
            parts[n] = x2.encode("utf-8") if n == sp else z.read(n)
        z.close()
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zo:
            for n, d in parts.items():
                zo.writestr(n, d)
        try:
            os.replace(tmp, XLSX)
        except PermissionError:
            with open(XLSX, "wb") as f:
                f.write(open(tmp, "rb").read())
            os.remove(tmp)
        print("已落盘，备份:", BACKUP)
    else:
        z.close()
        print("（预览，未落盘）")


if __name__ == "__main__":
    main()
