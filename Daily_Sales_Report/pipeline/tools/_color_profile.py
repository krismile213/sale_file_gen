# -*- coding: utf-8 -*-
"""汇总：多工作簿 × 多 sheet 的染色画像（按实际显示色统计 + 零值格归因）"""
import sys, os, re
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _cf_eval as C
import openpyxl
from openpyxl.utils import get_column_letter
from collections import Counter, defaultdict

COLNAME = {"FFFF0000": "红", "FFFF8FA8": "玫红", "FFFFDDE7": "浅粉",
           "FF61D8FF": "蓝", "FF55D5FF": "蓝", "FF5BD7FF": "蓝"}

TARGETS = [
    ("澳大利亚库存销售表", "澳大利亚-每日库存分析表"),
    ("Interflowd澳大利亚库存销售表", "澳大利亚-每日库存分析表"),
    ("欧盟英国库存销售表", "欧洲-每日库存分析表"),
    ("欧盟英国库存销售表", "英国-每日库存分析表"),
]


def run(name, sh):
    p = os.path.join(C.WORK, name + ".xlsx")
    wb = openpyxl.load_workbook(p, data_only=True)
    if sh not in wb.sheetnames:
        print(f"  !! {name} 没有 {sh}")
        return
    ws = wb[sh]
    hdr = {c: ws.cell(5, c).value for c in range(1, ws.max_column + 1)
           if ws.cell(5, c).value is not None}
    daycols = sorted(c for c, v in hdr.items() if hasattr(v, "year"))
    ucol = next(c for c, v in hdr.items() if isinstance(v, str) and "上月日均" in v)
    fcol = next(c for c, v in hdr.items() if isinstance(v, str) and v.strip() == "可售库存")
    rules = [r for r in C.load_cf(p, sh) if r["type"] == "expression" and r["formula"]]
    dayr = []
    for r in rules:
        if any(C.sqref_covers(r["sqref"], c, 7) for c in daycols):
            r = dict(r)
            r.update(C.classify(r["formula"], r["dxf"], r["color"]))
            dayr.append(r)
    red = [r for r in dayr if r["color"] == "FFFF0000"]
    # 锚点日
    anchor = None
    for r in dayr:
        m = re.search(r"([^\s!(),<>=]+)!\$?([A-Z]{1,3})\$?(\d+)", r["formula"])
        if m and m.group(1) in wb.sheetnames:
            anchor = wb[m.group(1)].cell(int(m.group(3)),
                                         C.column_index_from_string(m.group(2))).value
            break
    print(f"\n{'='*78}\n{name} / {sh}   行7-{ws.max_row}  日列 "
          f"{get_column_letter(daycols[0])}-{get_column_letter(daycols[-1])}({len(daycols)})")
    print(f"  上月日均={get_column_letter(ucol)}  可售={get_column_letter(fcol)}")
    for r in sorted(red, key=lambda x: x["prio"]):
        print(f"    红规则 prio={r['prio']:<5} {r['kind']:<7} "
              f"mult={r.get('mult')} th={r.get('th')}  {r['formula'][:70]}")

    cnt = Counter(); zero = Counter(); ex = defaultdict(list)
    for row in range(7, ws.max_row + 1):
        lastc = None
        for c in daycols:
            v = ws.cell(row, c).value
            if C.isnum(v) and v > 0:
                lastc = c
        total = sum(ws.cell(row, c).value for c in daycols if C.isnum(ws.cell(row, c).value))
        F = ws.cell(row, fcol).value; A = ws.cell(row, 1).value
        G = ws.cell(row, 7).value; H = ws.cell(row, 8).value
        for c in daycols:
            v = ws.cell(row, c).value
            hits = [r for r in dayr
                    if C.sqref_covers(r["sqref"], c, row)
                    and C.hit(r, v, ws.cell(row, ucol).value, total, lastc, c, daycols, A, F, G, H, anchor, ws)]
            if not hits:
                continue
            best = min(hits, key=lambda x: x["prio"])
            k = (best["color"], best["prio"])
            cnt[k] += 1
            if v is None or (C.isnum(v) and v == 0):
                zero[k] += 1
                if len(ex[k]) < 5:
                    ex[k].append(f"{get_column_letter(c)}{row}")
    print("  -- 实际显示色 --")
    for k, n in cnt.most_common():
        print(f"    {COLNAME.get(k[0], k[0])}({k[0]}) prio={k[1]:<5} 共{n:>5}格  "
              f"零值{zero.get(k, 0):>5}格  例 {ex.get(k, [])}")


if __name__ == "__main__":
    for n, s in TARGETS:
        try:
            run(n, s)
        except Exception as e:
            print(f"  !! {n}/{s}: {e}")
