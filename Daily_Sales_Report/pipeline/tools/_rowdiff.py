# -*- coding: utf-8 -*-
"""临时诊断：对比分析表里「正常行」与「补出来的行」的列差异

找出补行缺了哪些列（尤其 E/F/G/H 之后的 AD/AE/AF/AG/AS..CE/CF）。
"""
import os, sys, zipfile, re

BOOK = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))),
                    "1.每日销售数据表", "欧盟英国库存销售表.xlsx")


def sheet_map(path):
    z = zipfile.ZipFile(path)
    wbxml = z.read("xl/workbook.xml").decode()
    rels = z.read("xl/_rels/workbook.xml.rels").decode()
    rmap = {}
    for tag in re.findall(r"<Relationship\b[^>]*/?>", rels):
        rid = re.search(r'\bId="([^"]+)"', tag)
        tgt = re.search(r'\bTarget="([^"]+)"', tag)
        if rid and tgt:
            rmap[rid.group(1)] = tgt.group(1)
    out = {}
    for m in re.finditer(r'<sheet[^>]*name="([^"]+)"[^>]*r:id="([^"]+)"', wbxml):
        t = rmap.get(m.group(2), "").lstrip("/")
        if not t.startswith("xl/"):
            t = "xl/" + t
        out[m.group(1)] = t
    return z, out


def rows_of(x):
    rows = {}
    for m in re.finditer(r'<row r="(\d+)"[^>]*>(.*?)</row>', x, re.S):
        rn = int(m.group(1))
        cs = {}
        for mc in re.finditer(r'<c r="([A-Z]+)(\d+)"([^>/]*)(?:/>|>(.*?)</c>)', m.group(2), re.S):
            col, _r, attr, body = mc.group(1), mc.group(2), mc.group(3), mc.group(4) or ""
            cs[col] = dict(attr=attr.strip(), hasF="<f" in body, hasV="<v" in body, body=body[:80])
        rows[rn] = dict(n=len(m.group(2)), cs=cs)
    return rows


def colidx(c):
    n = 0
    for ch in c:
        n = n * 26 + (ord(ch) - 64)
    return n


z, sm = sheet_map(BOOK)
for sn in ("英国-每日库存分析表", "欧洲-每日库存分析表"):
    x = z.read(sm[sn]).decode()
    rows = rows_of(x)
    # 统计每个列号有多少行出现
    from collections import Counter
    cnt = Counter()
    for rn, d in rows.items():
        if rn < 7:
            continue
        for c in d["cs"]:
            cnt[colidx(c)] += 1
    # 正常行 = 出现次数最多的那一档
    data_rows = [rn for rn in rows if rn >= 7 and rows[rn]["cs"]]
    print("=" * 60)
    print(sn, "数据行(非空)=", len(data_rows))
    # 用最后一档高频判断基准：取每行 col 数众数
    widths = Counter(len(rows[rn]["cs"]) for rn in data_rows)
    base = widths.most_common(1)[0][0]
    good = next(rn for rn in data_rows if len(rows[rn]["cs"]) == base)
    print("基准行 r%d 列数=%d" % (good, base))
    bad = [rn for rn in data_rows if len(rows[rn]["cs"]) < base * 0.75]
    print("残缺行:", [(rn, len(rows[rn]["cs"]), rows[rn]["cs"].get("A", {}).get("body", "")[:20]) for rn in bad])
    if bad:
        rn = bad[0]
        miss = [c for c in rows[good]["cs"] if c not in rows[rn]["cs"]]
        print("  残缺行 r%d 缺少 %d 列: %s" % (rn, len(miss), ",".join(sorted(miss, key=colidx)[:40])))
        print("  缺列顶格后缀:", ",".join(sorted(miss, key=colidx)[-10:]))
    print("基准行最后5列:", sorted(rows[good]["cs"], key=colidx)[-5:])
