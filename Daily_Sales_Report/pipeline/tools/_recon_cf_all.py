# -*- coding: utf-8 -*-
"""17 张分析表染色现状侦察：B1 图例、红/蓝规则公式、行 1 内容边界、已有合并。

输出每表一段，供改造前确认「合并到哪一列」「文案里的数字从哪来」。
"""
import os
import re
import sys
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.abspath(os.path.join(HERE, "..", "core")))
import config as C

BOOKS = os.path.join(C.DSR, "1.每日销售数据表")


def sheet_map(z):
    wb = z.read("xl/workbook.xml").decode("utf-8", "replace")
    rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8", "replace")
    rid2t = {}
    for tag in re.findall(r"<Relationship\b[^>]*/?>", rels):
        rid = re.search(r'\bId="([^"]+)"', tag)
        tgt = re.search(r'\bTarget="([^"]+)"', tag)
        if rid and tgt:
            rid2t[rid.group(1)] = tgt.group(1)
    out = []
    for tag in re.findall(r"<sheet\b[^>]*/?>", wb):
        nm = re.search(r'\bname="([^"]+)"', tag)
        rid = re.search(r'\br:id="([^"]+)"', tag)
        if not (nm and rid):
            continue
        t = rid2t.get(rid.group(1), "")
        sp = t if t.startswith("xl/") else "xl/" + t.lstrip("/")
        out.append((nm.group(1), sp.replace("xl/xl/", "xl/")))
    return out


def sst_list(z):
    try:
        s = z.read("xl/sharedStrings.xml").decode("utf-8")
    except KeyError:
        return []
    return [re.sub(r"<[^>]+>", "", m.group(1))
            for m in re.finditer(r"<si>(.*?)</si>", s, re.S)]


def cell_text(body, attrs, ss):
    if 't="s"' in attrs:
        m = re.search(r"<v>(\d+)</v>", body)
        if m and int(m.group(1)) < len(ss):
            return ss[int(m.group(1))]
        return ""
    if 't="inlineStr"' in attrs or "<is>" in body:
        return "".join(re.findall(r"<t[^>]*>(.*?)</t>", body, re.S))
    m = re.search(r"<v>(.*?)</v>", body, re.S)
    return m.group(1) if m else ""


def row_cells(xml, r):
    m = re.search(r'<row r="%d"[^>]*>(.*?)</row>' % r, xml, re.S)
    if not m:
        return {}
    out = {}
    for c in re.finditer(r'<c r="([A-Z]+)%d"([^>]*?)(?:/>|>(.*?)</c>)' % r,
                         m.group(1), re.S):
        out[c.group(1)] = (c.group(2) or "", c.group(3) or "")
    return out


def colnum(letters):
    n = 0
    for ch in letters:
        n = n * 26 + ord(ch) - 64
    return n


def main():
    files = sorted(f for f in os.listdir(BOOKS) if f.endswith(".xlsx"))
    for fname in files:
        p = os.path.join(BOOKS, fname)
        z = zipfile.ZipFile(p)
        ss = sst_list(z)
        for sn, sp in sheet_map(z):
            if "分析表" not in sn:
                continue
            xml = z.read(sp).decode("utf-8")
            print("=" * 78)
            print("%s  ::  %s" % (fname, sn))
            print("  sheet part: %s  (%d KB)" % (sp, len(xml) // 1024))

            # B1
            rc = row_cells(xml, 1)
            b1 = rc.get("B")
            if b1:
                txt = cell_text(b1[1], b1[0], ss)
                print("  B1 长度 %d 字, 行数 %d" % (len(txt), txt.count("\n") + 1))
                print("    前80: %r" % txt[:80])
            # 行 1 有内容的列（B 之后第一个非空）
            others = sorted((colnum(k), k) for k, v in rc.items()
                            if k != "B" and cell_text(v[1], v[0], ss).strip())
            print("  行1 B 之后有内容的列: %s" % [k for _, k in others[:6]])
            if others:
                first = others[0][0]
                print("  → 可合并到第 %d 列 (%s 前一列 %s)"
                      % (first - 1,
                         chr(64 + first) if first <= 26 else str(first),
                         _col(first - 1)))
            # 已有合并
            mm = re.search(r"<mergeCells[^>]*>(.*?)</mergeCells>", xml, re.S)
            if mm:
                refs = re.findall(r'<mergeCell ref="([^"]+)"', mm.group(1))
                print("  已有 mergeCells %d 个: %s" % (len(refs), refs[:6]))
            else:
                print("  已有 mergeCells: 无")

            # 红规则
            reds = []
            for cf in re.finditer(r'<conditionalFormatting sqref="([^"]+)"[^>]*>(.*?)</conditionalFormatting>',
                                  xml, re.S):
                sq = cf.group(1)
                for r_ in re.finditer(r'<cfRule([^>]*)>(.*?)</cfRule>', cf.group(2), re.S):
                    attrs, inner = r_.group(1), r_.group(2)
                    f = re.search(r"<formula>(.*?)</formula>", inner, re.S)
                    if not f:
                        continue
                    fo = f.group(1)
                    dx = re.search(r'dxfId="(\d+)"', attrs)
                    pri = re.search(r'priority="(\d+)"', attrs)
                    reds.append((sq, dx.group(1) if dx else "-",
                                 pri.group(1) if pri else "-", fo))
            print("  CF 规则 %d 条，其中带公式 %d 条" % (xml.count("<cfRule"), len(reds)))
            # 只打印日列/实时速度列相关的红蓝
            for sq, dx, pri, fo in reds:
                if "*" in fo or "/" in fo:
                    print("    sq=%-14s dxf=%-3s pri=%-5s %s" % (sq, dx, pri, fo[:110]))
        z.close()


def _col(n):
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


if __name__ == "__main__":
    main()
