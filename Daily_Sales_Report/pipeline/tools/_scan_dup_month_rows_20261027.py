# -*- coding: utf-8 -*-
"""扫描 13 本簿所有「分析表 + 站点表」12 个月度块的错复制/脏值。只读，不改文件。

精确判据（不同 SKU 不可能巧合命中）：
  D 向量相同 —— 两个以上不同 SKU 行的 12 个月值**完全相同**（且非零月 >=3）
  E 巨值跨行 —— 同一个 >=1000 的值出现在 >=3 个不同 SKU 行
"""
import os, sys, re, glob, zipfile, importlib.util
from collections import defaultdict
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "..", "core"))
import config as C


def load_helpers():
    p = os.path.join(HERE, "maintenance", "rebuild_monthly_from_daily_20261027.py")
    spec = importlib.util.spec_from_file_location("rb", p)
    m = importlib.util.module_from_spec(spec); spec.loader.exec_module(m)
    return m


RB = load_helpers()
CELL = re.compile(r'<c r="([A-Z]+)(\d+)"([^>]*?)(?:/>|>(.*?)</c>)', re.S)
VAL = re.compile(r"<v>(-?[\d.]+)</v>")


def rows_of(x, sst):
    for rm in re.finditer(r'<row r="(\d+)"[^>]*>(.*?)</row>', x, re.S):
        r = int(rm.group(1)); cells = {}
        for c in CELL.finditer(rm.group(2)):
            col, attrs, body = c.group(1), c.group(3), c.group(4) or ""
            if 't="s"' in attrs:
                v = VAL.search(body)
                if v and int(v.group(1)) < len(sst):
                    cells[col] = sst[int(v.group(1))]
                continue
            if 't="inlineStr"' in attrs:
                cells[col] = "".join(re.findall(r"<t[^>]*>(.*?)</t>", body, re.S)); continue
            if re.search(r't="(str|d)"', attrs):
                continue
            v = VAL.search(body)
            if v:
                cells[col] = float(v.group(1))
        yield r, cells


def analyse(path):
    z = zipfile.ZipFile(path); sm = RB.smap(z); sst = RB.sst_of(z)
    res = []
    for nm, sp in sm.items():
        x = z.read(sp).decode("utf-8")
        cols = RB.month_cols(x, sst)
        if not cols or len(cols) < 12:
            continue
        mcols = [cols[m] for m in range(1, 13)]
        data = []
        for r, cells in rows_of(x, sst):
            if r < 7:
                continue
            sku = cells.get("A")
            if not sku or not isinstance(sku, str) or not sku.strip():
                continue
            vals = tuple(int(cells[c]) if isinstance(cells.get(c), float) else 0 for c in mcols)
            data.append((r, sku.strip(), vals))

        # D 向量相同
        grp = defaultdict(list)
        for r, sku, vals in data:
            if sum(1 for v in vals if v > 0) >= 3:
                grp[vals].append((r, sku))
        for vals, rows in grp.items():
            if len(rows) >= 2:
                res.append((nm.strip(), "向量相同×%d" % len(rows),
                            [s for _, s in rows][:10], rows[0][0], list(vals)))
        # E 巨值跨行
        big = defaultdict(set)
        for r, sku, vals in data:
            for v in vals:
                if v >= 1000:
                    big[v].add(r)
        for v, rs in big.items():
            if len(rs) >= 3:
                skus = [s for r, s, _ in data if r in rs]
                res.append((nm.strip(), "巨值 %d ×%d行" % (v, len(rs)), skus[:10],
                            min(rs), None))
    z.close()
    return res


total = 0
for p in sorted(glob.glob(os.path.join(C.WORK_DIR, "*.xlsx"))):
    hits = analyse(p)
    if not hits:
        continue
    print("\n=== %s  (%d 条) ===" % (os.path.basename(p), len(hits)))
    for nm, why, skus, r0, vals in hits[:20]:
        total += 1
        print("  %-16s %-16s 行%-5d %s" % (nm, why, r0, skus))
        if vals:
            print("  %-16s   12个月 = %s" % ("", vals))
print("\n合计: %d" % total)
