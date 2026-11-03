# -*- coding: utf-8 -*-
"""对比几种红色口径的效果。日列范围直接取自红色 CF 规则的 sqref（最可靠）。"""
import os
import re
import sys
import zipfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _cf_inventory as inv  # noqa: E402

TARGETS = [
    ("欧盟英国库存销售表.xlsx", "欧洲-每日库存分析表"),
    ("欧盟英国库存销售表.xlsx", "英国-每日库存分析表"),
    ("北美加拿大库存销售表.xlsx", "美国-每日库存分析表"),
    ("澳大利亚库存销售表.xlsx", "澳大利亚-每日库存分析表"),
    ("沙特库存销售表.xlsx", "沙特阿拉伯-每日库存分析表"),
    ("美客多库存销售表.xlsx", "墨西哥-每日库存分析表"),
]


def col2num(s):
    n = 0
    for ch in s:
        n = n * 26 + (ord(ch) - 64)
    return n


def shared_strings(z):
    if "xl/sharedStrings.xml" not in z.namelist():
        return []
    sx = z.read("xl/sharedStrings.xml").decode("utf-8", "replace")
    return ["".join(re.findall(r"<t[^>]*>(.*?)</t>", si, re.S)) for si in re.findall(r"<si>(.*?)</si>", sx, re.S)]


def cell_val(body, attr, ss):
    if 't="s"' in attr:
        i = re.search(r"<v>(\d+)</v>", body)
        return ss[int(i.group(1))] if i and int(i.group(1)) < len(ss) else ""
    if 't="inlineStr"' in attr:
        return "".join(re.findall(r"<t[^>]*>(.*?)</t>", body, re.S))
    m = re.search(r"<v>(.*?)</v>", body, re.S)
    return m.group(1) if m else ""


def load(book, sheet):
    z = zipfile.ZipFile(os.path.join(inv.WORK, book))
    ss = shared_strings(z)
    sp = next(p for nm, p in inv.sheet_files(z) if nm == sheet)
    x = z.read(sp).decode("utf-8", "replace")
    z.close()

    rows = {}
    for rm in re.finditer(r'<row r="(\d+)"[^>]*>(.*?)</row>', x, re.S):
        r = int(rm.group(1))
        cells = {}
        for cm in re.finditer(r'<c r="([A-Z]{1,3})\d+"([^>]*?)(?:/>|>(.*?)</c>)', rm.group(2), re.S):
            cells[cm.group(1)] = (cm.group(3) or "", cm.group(2))
        rows[r] = cells

    # 日列 + 日均列：取「红色倍率档」规则（公式含 *N 且跨多列）
    daycols = []
    avgcol = None
    for sq, body in re.findall(
        r'<conditionalFormatting[^>]*sqref="([^"]*)"[^>]*>(.*?)</conditionalFormatting>', x, re.S
    ):
        for rule in re.findall(r"<cfRule[^>]*?/>|<cfRule[^>]*?>.*?</cfRule>", body, re.S):
            fm = re.findall(r"<formula>(.*?)</formula>", rule, re.S)
            if not fm:
                continue
            ff = fm[0].replace("&lt;", "<").replace("&gt;", ">")
            m = re.match(r"([A-Z]{1,3})7:([A-Z]{1,3})\d+$", sq)
            if not m:
                continue
            if re.search(r"\*[23]", ff) and col2num(m.group(2)) - col2num(m.group(1)) > 10:
                a, b = col2num(m.group(1)), col2num(m.group(2))

                def num2col(n):
                    s = ""
                    while n:
                        n, r = divmod(n - 1, 26)
                        s = chr(65 + r) + s
                    return s

                daycols = [num2col(i) for i in range(a, b + 1)]
                am = re.search(r"\$([A-Z]{1,3})7\*", ff)
                avgcol = am.group(1) if am else None
                break
        if daycols:
            break

    def num(r, c):
        if c is None or c not in rows.get(r, {}):
            return None
        body, attr = rows[r][c]
        try:
            return float(cell_val(body, attr, ss))
        except Exception:
            return None

    data = []
    for r in sorted(rows):
        if r < 7:
            continue
        ao = num(r, avgcol)
        for c in daycols:
            v = num(r, c)
            if v is not None:
                data.append((r, c, ao, v))
    return data, daycols, avgcol


def count(data, N, M, mode, K=20):
    n = 0
    vals = []
    for r, c, ao, v in data:
        hit = False
        if ao is None:
            hit = (mode == "abs" and v >= K)
            if hit:
                n += 1
                vals.append(v)
            continue
        if mode == "cur":
            hit = (ao > M and v > ao * N) or (0 < ao <= M and v >= 2 and v > ao * 4)
        elif mode == "nolo":
            hit = ao > M and v > ao * N
        elif mode == "abs":
            hit = v >= K
        elif mode == "abs_rel":
            hit = v >= K and v > ao * 1.5
        if hit:
            n += 1
            vals.append(v)
    return n, vals


def main():
    for book, sheet in TARGETS:
        data, daycols, avgcol = load(book, sheet)
        total = len(data)
        if not total:
            print(f"### {book[:22]} / {sheet}  —— 未取到数据，跳过")
            continue
        # 从表里反推 N/M（按表分组）
        N, M = {
            "欧洲-每日库存分析表": (3.0, 1.5),
            "英国-每日库存分析表": (2.5, 1.5),
            "美国-每日库存分析表": (2.8, 1.0),
            "澳大利亚-每日库存分析表": (2.5, 1.0),
            "沙特阿拉伯-每日库存分析表": (2.0, 0.5),
            "墨西哥-每日库存分析表": (2.0, 0.5),
        }[sheet]
        mx = max(v for _, _, _, v in data)
        print("=" * 100)
        print(f"### {book[:22]} / {sheet}  N={N} M={M}  日列 {daycols[0]}..{daycols[-1]} 日均列={avgcol}  非空格={total}  最高单日={mx:.0f}")
        print(f"{'口径':<26}{'红格数':>7}{'占比':>9}{'销量范围':>14}")
        print("-" * 100)
        for label, mode, K in (
            ("现状（倍率+低基数）", "cur", 0),
            ("只留倍率档", "nolo", 0),
            ("纯绝对 ≥10", "abs", 10),
            ("纯绝对 ≥15", "abs", 15),
            ("纯绝对 ≥20", "abs", 20),
            ("纯绝对 ≥30", "abs", 30),
            ("绝对≥10 且 >日均×1.5", "abs_rel", 10),
            ("绝对≥15 且 >日均×1.5", "abs_rel", 15),
            ("绝对≥20 且 >日均×1.5", "abs_rel", 20),
        ):
            n, vals = count(data, N, M, mode, K)
            rng = f"{min(vals):.0f}~{max(vals):.0f}" if vals else "-"
            print(f"{label:<26}{n:>7}{n/total*100:>8.2f}%{rng:>14}")


if __name__ == "__main__":
    main()
