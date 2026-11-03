# -*- coding: utf-8 -*-
"""欧洲表红口径退回后，按新（旧）公式逐格复算，确认回到 299 格且 G520 的 43 不红。"""
import os
import re
import sys
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.abspath(os.path.join(HERE, "..", "core")))
import config as C

P = os.path.join(C.DSR, "1.每日销售数据表", "欧盟英国库存销售表.xlsx")
SP = "xl/worksheets/sheet10.xml"


def cn(s):
    n = 0
    for ch in s:
        n = n * 26 + ord(ch) - 64
    return n


def cl(n):
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


z = zipfile.ZipFile(P)
xml = z.read(SP).decode()
SST = [re.sub(r"<[^>]+>", "", m.group(1)) for m in
       re.finditer(r"<si>(.*?)</si>", z.read("xl/sharedStrings.xml").decode(), re.S)]
z.close()

# 逐行取 AO 与 AS..CE 的缓存值
rows = {}
for rm in re.finditer(r'<row r="(\d+)"[^>]*>(.*?)</row>', xml, re.S):
    r = int(rm.group(1))
    if r < 7:
        continue
    d = {}
    for cm in re.finditer(r'<c r="([A-Z]+)%d"([^>]*?)(?:/>|>(.*?)</c>)' % r,
                          rm.group(2), re.S):
        col, body = cm.group(1), cm.group(3) or ""
        v = re.search(r"<v>([-\d.]+)</v>", body)
        if v:
            try:
                d[col] = float(v.group(1))
            except ValueError:
                pass
        f = re.search(r"<f[^>]*>(.*?)</f>|<f[^>]*/>", body, re.S)
        if f and col not in d:
            d[col] = None
    rows[r] = d

c_lo = cn("AS")
c_hi = cn("CE")
daycols = [cl(i) for i in range(c_lo, c_hi + 1)]

red = []
for r, d in rows.items():
    ao = d.get("AO")
    if ao is None:
        continue
    for c in daycols:
        v = d.get(c)
        if v is None:
            continue
        hi = (v > ao * 3 and ao > 1.5)
        lo = (ao > 0 and ao <= 1.5 and v >= 2 and v > ao * 4)
        if hi or lo:
            red.append((r, c, v, ao, "高" if hi else "低基数"))

print("红格总数: %d  （旧口径基线 299）" % len(red))
print("  高销量档: %d" % sum(1 for x in red if x[4] == "高"))
print("  低基数档: %d" % sum(1 for x in red if x[4] == "低基数"))

# G520
print()
g = [r for r, d in rows.items() if d.get("A") is None]
# A 列是 SKU（文本），上面只收了数值，重新扫
sku_rows = {}
for rm in re.finditer(r'<row r="(\d+)"[^>]*>(.*?)</row>', xml, re.S):
    r = int(rm.group(1))
    m = re.search(r'<c r="A%d"[^>]*?t="(?:s|inlineStr)"[^>]*?(?:/>|>(.*?)</c>)' % r,
                  rm.group(2), re.S)
    if not m:
        continue
    body = m.group(1) or ""
    t = re.search(r"<v>(\d+)</v>", body)
    if t:
        i = int(t.group(1))
        sku_rows[r] = SST[i] if i < len(SST) else ""
    else:
        sku_rows[r] = "".join(re.findall(r"<t[^>]*>(.*?)</t>", body, re.S))

g520 = [r for r, s in sku_rows.items() if s.strip() == "G520"]
print("G520 所在行: %s" % g520)
for r in g520:
    d = rows.get(r, {})
    ao = d.get("AO")
    print("   AO 上月日均 = %s" % ao)
    vals = [(c, d.get(c)) for c in daycols if d.get(c)]
    hits = [x for x in vals if x[1] >= 20]
    print("   ≥20 的日格: %s" % [(c, int(v)) for c, v in hits])
    for c, v in hits:
        isred = (v > ao * 3 and ao > 1.5) or (ao > 0 and ao <= 1.5 and v >= 2 and v > ao * 4)
        print("      %s=%d  倍数=%.2f  → %s" % (c, int(v), v / ao, "红" if isred else "不红"))
