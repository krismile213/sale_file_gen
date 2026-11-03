# -*- coding: utf-8 -*-
"""临时诊断：分析表每一行的「列覆盖完整度」

补行时 clone_row 是照抄**末数据行**的，如果末行本身列残缺（只有 49 格），
新行也会残缺 —— 缺 AD/AE（在途差额/实时速度）、AF/AG（可售天数）、
AS..CE（日销列）、CF（种类）等。条件格式范围也依赖列，缺列等于规则不生效。

本脚本统计每张分析表：所有数据行的单元格数分布 + 残缺行号清单。
"""
import os, sys, zipfile, re
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

BOOKDIR = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))),
                       "1.每日销售数据表")


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


def main(book_filter=None):
    for fn in sorted(os.listdir(BOOKDIR)):
        if not fn.endswith(".xlsx"):
            continue
        if book_filter and book_filter not in fn:
            continue
        p = os.path.join(BOOKDIR, fn)
        try:
            z, sm = sheet_map(p)
        except Exception as e:
            print(f"!! {fn}: {e}")
            continue
        for sn, tgt in sm.items():
            if "分析表" not in sn:
                continue
            x = z.read(tgt).decode()
            # 每行单元格数
            info = []
            for m in re.finditer(r'<row r="(\d+)"[^>]*>(.*?)</row>', x, re.S):
                rn = int(m.group(1))
                n = len(re.findall(r"<c\s+r=", m.group(2)))
                info.append((rn, n))
            data = [(r, n) for r, n in info if r >= 7]
            if not data:
                continue
            from collections import Counter
            dist = Counter(n for r, n in data)
            mode_n, mode_cnt = dist.most_common(1)[0]
            bad = [(r, n) for r, n in data if n < mode_n * 0.75]
            print(f"## {fn} | {sn!r}")
            print(f"   行数={len(data)}  常见列数={mode_n}({mode_cnt}行)  分布={dict(sorted(dist.items()))}")
            if bad:
                print(f"   ⚠ 残缺行 {len(bad)} 个: " + ", ".join(f"r{r}({n}格)" for r, n in bad[:20]))
        z.close()


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else None)
