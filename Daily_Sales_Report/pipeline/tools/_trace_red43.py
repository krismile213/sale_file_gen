# -*- coding: utf-8 -*-
"""欧洲表红色口径诊断：为什么 43 不红、0 却红了。

输出：
  ① 用户贴的 17 行对应哪些 SKU、各自上月日均、43 为何不红
  ② 全表日销量分布（用于挑绝对阈值）
  ③ 当前红色规则命中的格子 vs 改成绝对阈值后会命中多少
  ④ 销量=0 却被标红/粉的格子
"""
import os
import re
import sys
import zipfile
from collections import Counter

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _cf_inventory as inv  # noqa: E402

BOOK = "欧盟英国库存销售表.xlsx"
SHEET = "欧洲-每日库存分析表"

USER_ROWS = [
    [18, 16, 12, 24, 43, 32, 22, 13, 21, 14, 16, 9, 18, 22, 23, 16, 23, 9, 20, 16, 16],
    [8, 10, 5, 6, 19, 20, 5, 11, 7, 7, 9, 6, 6, 9, 8, 9, 9, 6, 4, 5, 6],
    [12, 14, 8, 18, 19, 12, 11, 14, 5, 16, 15, 7, 11, 15, 10, 11, 6, 12, 10, 9, 11],
    [11, 14, 13, 12, 23, 32, 16, 13, 12, 17, 19, 13, 15, 8, 10, 9, 13, 10, 14, 11, 14],
    [7, 8, 11, 9, 30, 20, 10, 8, 7, 4, 8, 15, 14, 9, 8, 8, 5, 9, 8, 5, 8],
    [7, 4, 9, 9, 19, 13, 1, 7, 4, 6, 8, 9, 6, 5, 9, 4, 6, 10, 7, 5, 5],
    [5, 5, 4, 3, 3, 6, 5, 2, 5, 3, 5, 7, 4, 4, 3, 1, 1, 2, 5, 4, 2],
    [11, 8, 9, 20, 24, 21, 14, 9, 7, 10, 10, 12, 8, 13, 14, 8, 15, 6, 5, 8, 12],
    [4, 3, 3, 6, 21, 15, 10, 8, 2, 10, 10, 4, 8, 14, 6, 4, 4, 10, 10, 7, 11],
    [5, 11, 10, 7, 26, 22, 7, 14, 7, 8, 8, 10, 18, 14, 13, 14, 13, 9, 12, 16, 4],
    [6, 7, 12, 7, 10, 15, 6, 5, 5, 7, 6, 8, 5, 8, 4, 5, 6, 8, 4, 15, 12],
    [6, 8, 3, 4, 14, 20, 5, 8, 3, 3, 7, 4, 7, 5, 5, 9, 5, 4, 7, 9, 4],
    [4, 2, 3, 6, 10, 11, 6, 6, 9, 3, 9, 5, 7, 3, 5, 4, 3, 6, 3, 2, 7],
    [9, 6, 6, 3, 16, 7, 5, 4, 7, 3, 6, 11, 6, 6, 13, 2, 13, 1, 3, 6, 5],
    [6, 5, 6, 9, 10, 9, 7, 4, 3, 9, 10, 5, 2, 7, 2, 2, 8, 5, 6, 10, 5],
    [2, 1, 6, 6, 16, 17, 3, 3, 4, 6, 3, 5, 3, 4, 7, 2, 6, 6, 4, 3, 6],
    [3, 6, 4, 8, 10, 9, 4, 5, 2, 5, 8, 2, 5, 4, 5, 2, 6, 5, 4, 8, 5],
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


def main():
    z = zipfile.ZipFile(os.path.join(inv.WORK, BOOK))
    ss = shared_strings(z)
    sp = next(p for nm, p in inv.sheet_files(z) if nm == SHEET)
    x = z.read(sp).decode("utf-8", "replace")

    rows = {}
    for rm in re.finditer(r'<row r="(\d+)"[^>]*>(.*?)</row>', x, re.S):
        r = int(rm.group(1))
        cells = {}
        for cm in re.finditer(r'<c r="([A-Z]{1,3})\d+"([^>]*?)(?:/>|>(.*?)</c>)', rm.group(2), re.S):
            cells[cm.group(1)] = (cm.group(3) or "", cm.group(2))
        rows[r] = cells

    # 日列（表头第 5 行为日期序列号）
    daycols = []
    for c, (body, attr) in rows.get(5, {}).items():
        try:
            f = float(cell_val(body, attr, ss))
        except Exception:
            continue
        if 40000 < f < 60000:
            daycols.append(c)
    daycols.sort(key=col2num)

    def num(r, c):
        if c not in rows.get(r, {}):
            return None
        body, attr = rows[r][c]
        try:
            return float(cell_val(body, attr, ss))
        except Exception:
            return None

    def txt(r, c):
        if c not in rows.get(r, {}):
            return ""
        body, attr = rows[r][c]
        return str(cell_val(body, attr, ss)).strip()

    # ---------- ① 用户 17 行定位 ----------
    print("=" * 100)
    print("① 用户贴的 17 行 → SKU / 上月日均 / 43 为何不红")
    print("=" * 100)
    n = len(USER_ROWS[0])
    found = {}
    for r in sorted(rows):
        if r < 7:
            continue
        vals = [num(r, c) for c in daycols]
        for i in range(len(vals) - n + 1):
            win = vals[i:i + n]
            if all(v is not None and int(v) == t for v, t in zip(win, USER_ROWS[0])):
                found[r] = (daycols[i], daycols[i + n - 1])
    print(f"首行匹配行号: {list(found)}  列区间: {list(found.values())[:1]}")

    c0, c1 = list(found.values())[0]
    win_cols = [c for c in daycols if col2num(c0) <= col2num(c) <= col2num(c1)]
    print(f"\n{'行':>4} {'SKU':<10}{'状态':<6}{'日均AO':>8}{'×3阈值':>9}{'区间最大':>9}{'最大是否红':>11}")
    for r, (a, b) in sorted(found.items()):
        ao = num(r, "AO")
        vals = [num(r, c) for c in win_cols]
        mx = max([v for v in vals if v is not None], default=None)
        thr = ao * 3 if ao else None
        isred = (mx > thr) if (mx is not None and thr is not None) else None
        print(f"{r:>4} {txt(r,'A'):<10}{txt(r,'D'):<6}{(f'{ao:.2f}' if ao else '-'):>8}"
              f"{(f'{thr:.1f}' if thr else '-'):>9}{(f'{mx:.0f}' if mx is not None else '-'):>9}"
              f"{('是' if isred else '否') if isred is not None else '-':>11}")

    # ---------- ② 全表日销量分布 ----------
    print()
    print("=" * 100)
    print("② 全表日销量分布（判断绝对阈值该定在哪）")
    print("=" * 100)
    allv = []
    for r in sorted(rows):
        if r < 7:
            continue
        for c in daycols:
            v = num(r, c)
            if v is not None:
                allv.append(v)
    cnt = Counter(int(v) for v in allv)
    tot = len(allv)
    print(f"总格数 {tot}，其中 0 值 {cnt.get(0,0)} 格（{cnt.get(0,0)/tot*100:.1f}%）")
    print("按阈值 K，日销 ≥K 的格子数：")
    for K in (2, 3, 5, 8, 10, 15, 20, 25, 30, 40):
        n_k = sum(1 for v in allv if v >= K)
        print(f"   K={K:<3} → {n_k:>5} 格 ({n_k/tot*100:5.2f}%)")
    print("销量前 15 高的格子：", sorted(allv, reverse=True)[:15])

    # ---------- ③ 当前红色规则命中 ----------
    print()
    print("=" * 100)
    print("③ 当前红色规则命中统计（欧洲表：倍率3，门槛1.5）")
    print("=" * 100)
    hi = lo = 0
    hi_vals, lo_vals = [], []
    for r in sorted(rows):
        if r < 7:
            continue
        ao = num(r, "AO")
        if ao is None:
            continue
        for c in daycols:
            v = num(r, c)
            if v is None:
                continue
            if ao > 1.5 and v > ao * 3:
                hi += 1
                hi_vals.append(v)
            elif 0 < ao <= 1.5 and v >= 2 and v > ao * 4:
                lo += 1
                lo_vals.append(v)
    print(f"红·高销量档 {hi} 格，销量区间 {min(hi_vals) if hi_vals else '-'}~{max(hi_vals) if hi_vals else '-'}")
    print(f"红·低基数档 {lo} 格，销量分布 {sorted(Counter(int(v) for v in lo_vals).items())[:8]}")
    print(f"合计红 {hi+lo} 格；其中低基数档占 {lo/(hi+lo)*100:.0f}%")

    # ---------- ④ 0 值却被标色的格子 ----------
    print()
    print("=" * 100)
    print("④ 销量=0 的格子中有多少落在 粉色/红色 区间内")
    print("=" * 100)
    zero_in_pink = 0
    zero_rows = set()
    for r in sorted(rows):
        if r < 7:
            continue
        f_ = num(r, "F")          # 可售
        if f_ is None or f_ > 0:
            continue
        vals = {c: num(r, c) for c in daycols}
        nz = [(c, v) for c, v in vals.items() if v is not None and v > 0]
        if not nz:
            continue
        last = max(nz, key=lambda t: col2num(t[0]))[0]
        for c in daycols:
            if col2num(c) > col2num(last) and vals.get(c) == 0:
                zero_in_pink += 1
                zero_rows.add(r)
    print(f"可售≤0 且当日销量=0、且在该 SKU 最后出单日之后的格子：{zero_in_pink} 格，涉及 {len(zero_rows)} 个 SKU")
    print("→ 这些格子会被【深粉/浅粉·缺货】规则染色，看起来像'0 也标红'")

    z.close()


if __name__ == "__main__":
    main()
