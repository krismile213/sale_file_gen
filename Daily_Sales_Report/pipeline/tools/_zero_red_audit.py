# -*- coding: utf-8 -*-
"""欧洲表：逐格复算最终颜色，专门统计「销量=0 却被染色」的格子属于哪条规则。"""
import os
import re
import sys
import zipfile
from collections import Counter

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _cf_inventory as inv  # noqa: E402

BOOK = "欧盟英国库存销售表.xlsx"
SHEET = "欧洲-每日库存分析表"


def col2num(s):
    n = 0
    for ch in s:
        n = n * 26 + (ord(ch) - 64)
    return n


def num2col(n):
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def shared_strings(z):
    if "xl/sharedStrings.xml" not in z.namelist():
        return []
    sx = z.read("xl/sharedStrings.xml").decode("utf-8", "replace")
    return ["".join(re.findall(r"<t[^>]*>(.*?)</t>", si, re.S)) for si in re.findall(r"<si>(.*?)</si>", sx, re.S)]


def main():
    z = zipfile.ZipFile(os.path.join(inv.WORK, BOOK))
    ss = shared_strings(z)
    sp = next(p for nm, p in inv.sheet_files(z) if nm == SHEET)
    x = z.read(sp).decode("utf-8", "replace")
    z.close()

    rows = {}
    for rm in re.finditer(r'<row r="(\d+)"[^>]*>(.*?)</row>', x, re.S):
        r = int(rm.group(1))
        cells = {}
        for cm in re.finditer(r'<c r="([A-Z]{1,3})\d+"([^>]*?)(?:/>|>(.*?)</c>)', rm.group(2), re.S):
            cells[cm.group(1)] = (cm.group(3) or "", cm.group(2))
        rows[r] = cells

    def num(r, c):
        if c not in rows.get(r, {}):
            return None
        body, attr = rows[r][c]
        if 't="s"' in attr or 't="inlineStr"' in attr:
            return None
        m = re.search(r"<v>(.*?)</v>", body, re.S)
        if not m or not m.group(1).strip():
            return None
        try:
            return float(m.group(1))
        except Exception:
            return None

    def txt(r, c):
        if c not in rows.get(r, {}):
            return ""
        body, attr = rows[r][c]
        if 't="s"' in attr:
            i = re.search(r"<v>(\d+)</v>", body)
            return ss[int(i.group(1))].strip() if i and int(i.group(1)) < len(ss) else ""
        return "".join(re.findall(r"<t[^>]*>(.*?)</t>", body, re.S)).strip()

    daycols = [num2col(i) for i in range(col2num("AS"), col2num("CE") + 1)]
    dcols = [num2col(i) for i in range(col2num("AS"), col2num("CE") + 1)]

    # 按优先级复算：数值越小优先级越高
    stats = Counter()
    samples = []
    zero_red = []
    for r in sorted(rows):
        if r < 7:
            continue
        a = txt(r, "A")
        if not a:
            continue
        ao = num(r, "AO")
        f_ = num(r, "F")
        g_ = num(r, "G")
        h_ = num(r, "H")
        vals = {c: num(r, c) for c in daycols}
        nz = [(c, v) for c, v in vals.items() if v is not None and v > 0]
        total_sum = sum(v for v in vals.values() if v is not None)
        lastc = max(nz, key=lambda t: col2num(t[0]))[0] if nz else None

        for c in daycols:
            v = vals.get(c)
            if v is None:
                continue
            color = None
            rule = None
            # 优先级（从高到低）：红·高销量 > 红·低基数 > 粉(深/浅) > 蓝
            if ao is not None and ao > 1.5 and v > ao * 3:
                color, rule = "纯红 FFFF0000", "红·高销量档"
            elif ao is not None and 0 < ao <= 1.5 and v >= 2 and v > ao * 4:
                color, rule = "纯红 FFFF0000", "红·低基数档"
            elif (f_ is not None and f_ <= 0 and total_sum > 0 and lastc
                  and col2num(c) > col2num(lastc)):
                if (g_ or 0) == 0 and (h_ or 0) == 0:
                    color, rule = "深粉 FFFF8FA8", "粉·断货无补给"
                else:
                    color, rule = "浅粉 FFFFDDE7", "粉·断货有补给"
            elif ao is not None and ao > 0 and 0 < v < ao / 3:
                color, rule = "蓝 5BD7FF", "蓝·低销量"
            if rule:
                stats[rule] += 1
                if v == 0:
                    zero_red.append((r, a, c, ao, f_, g_, h_, color, rule))
                elif len(samples) < 8 and rule.startswith("粉"):
                    samples.append((r, a, c, v, ao, color, rule))

    print("=" * 96)
    print("① 欧洲表各规则命中格数")
    print("=" * 96)
    for k, n in stats.most_common():
        print(f"   {k:<16} {n:>6} 格")

    print()
    print("=" * 96)
    print("② 销量=0 却被染色的格子")
    print("=" * 96)
    zc = Counter(t[7] for t in zero_red)
    print(f"   合计 {len(zero_red)} 格，颜色分布：{dict(zc)}")
    print(f"   其中被【纯红】染的：{sum(1 for t in zero_red if '纯红' in t[7])} 格")
    print(f"   其中被【深粉/浅粉】染的：{sum(1 for t in zero_red if '粉' in t[7])} 格")

    print()
    print("   示例（前 15 格）：")
    print(f"   {'行':>4} {'SKU':<8}{'列':<4}{'日均':>7}{'可售':>6}{'在途':>5}{'调拨':>5}  颜色")
    for t in zero_red[:15]:
        r, a, c, ao, f_, g_, h_, color, rule = t
        print(f"   {r:>4} {a:<8}{c:<4}{(f'{ao:.2f}' if ao is not None else '-'):>7}"
              f"{(f'{f_:.0f}' if f_ is not None else '-'):>6}{(f'{g_:.0f}' if g_ is not None else '-'):>5}"
              f"{(f'{h_:.0f}' if h_ is not None else '-'):>5}  {color}  ({rule})")

    print()
    print("=" * 96)
    print("③ 对照：有销量、被粉色染的格子（证明粉色是断货标记，不是按销量判断）")
    print("=" * 96)
    for s in samples[:8]:
        r, a, c, v, ao, color, rule = s
        print(f"   行{r} {a:<8}{c}  销量={v:.0f} 日均={ao:.2f}  {color} ({rule})")


if __name__ == "__main__":
    main()
