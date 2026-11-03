# -*- coding: utf-8 -*-
"""最终复验：openpyxl 落盘后公式缓存为空，月度合计改用「逐行求和」核对。"""
import os, sys, re, zipfile, importlib.util
HERE = os.path.dirname(os.path.abspath(__file__)); sys.path.insert(0, os.path.join(HERE, "..", "core"))
import config as C
p = os.path.join(HERE, "maintenance", "rebuild_monthly_from_daily_20261027.py")
spec = importlib.util.spec_from_file_location("rb", p); RB = importlib.util.module_from_spec(spec); spec.loader.exec_module(RB)
CELL = re.compile(r'<c r="([A-Z]+)(\d+)"([^>]*?)(?:/>|>(.*?)</c>)', re.S)
VAL = re.compile(r"<v>(-?[\d.]+)</v>")

def smap2(z):
    rels = z.read("xl/_rels/workbook.xml.rels").decode(); rid2t = {}
    for rel in re.findall(r"<Relationship\b[^>]*/?>", rels):
        i = re.search(r'Id="([^"]+)"', rel); t = re.search(r'Target="([^"]+)"', rel)
        if i and t:
            t = t.group(1).lstrip("/")
            rid2t[i.group(1)] = t if t.startswith("xl/") else "xl/" + t
    wbx = z.read("xl/workbook.xml").decode()
    return {n: rid2t[r] for n, r in re.findall(r'<sheet [^>]*?name="([^"]+)"[^>]*?r:id="(rId\d+)"', wbx)}

def col_sum(x, col):
    tot = 0; has_f6 = False; n = 0
    for rm in re.finditer(r'<row r="(\d+)"[^>]*>(.*?)</row>', x, re.S):
        r = int(rm.group(1))
        if r == 6:
            has_f6 = ('<f' in rm.group(2)) and (col + '6' in rm.group(2))
            continue
        if r < 7: continue
        c = re.search(r'<c r="%s%d"(?![0-9])([^>]*?)(?:/>|>(.*?)</c>)' % (col, r), rm.group(2), re.S)
        if not c: continue
        if re.search(r't="(s|str|d|inlineStr)"', c.group(1)): continue
        v = VAL.search(c.group(2) or "")
        if v: tot += float(v.group(1)); n += 1
    return int(tot), n, has_f6

X = os.path.join(C.WORK_DIR, "欧盟英国库存销售表.xlsx")
z = zipfile.ZipFile(X)
print("zip:", z.testzip() or "OK", " pivot:", len([n for n in z.namelist() if "pivot" in n.lower()]))
sm = smap2(z); sst = RB.sst_of(z)
cur = {n.strip(): z.read(sp).decode() for n, sp in sm.items()}; z.close()
TBL = ["英国-每日库存分析表","欧洲-每日库存分析表","意大利","法国","西班牙","德国",
       "土耳其","荷兰","瑞典","波兰","比利时","爱尔兰"]
EXP = {"英国-每日库存分析表": (8016, 6854), "欧洲-每日库存分析表": (28319, 24981),
       "意大利": (4873, 4425), "法国": (5351, 4339), "西班牙": (5772, 5425),
       "德国": (8009, 7144), "土耳其": (0, 0), "荷兰": (1081, 793), "瑞典": (789, 679),
       "波兰": (331, 295), "比利时": (1075, 904), "爱尔兰": (1038, 976)}
ok = bad = 0
print("\n%-16s %7s %7s %7s %6s  %s" % ("表","8月","9月","10月","10月格数","行6公式"))
for nm in TBL:
    x = cur[nm]; c = RB.month_cols(x, sst)
    s8, _, f8 = col_sum(x, c[8]); s9, _, f9 = col_sum(x, c[9]); s10, n10, f10 = col_sum(x, c[10])
    e8, e9 = EXP[nm]
    good = (s8 == e8 and s9 == e9)
    ok += 1 if good else 0; bad += 0 if good else 1
    print("%-16s %7d %7d %7d %6d  %s %s" % (nm, s8, s9, s10, n10,
          "有" if (f8 and f9 and f10) else "缺!", "OK" if good else "✗ 期望%d/%d" % (e8, e9)))
print("\n月度合计核对: 通过 %d / 失败 %d" % (ok, bad))
print("\n--- 新行 ---")
for nm, sku, d10, d9 in (("意大利","G598","AX","AY"), ("瑞典","K001","AW","AX")):
    x = cur[nm]; rows = RB.sku_rows(x, sst)
    r = next((k for k, v in rows.items() if v == sku), None)
    print("  %s %s 行%s 10月=%s 9月=%s" % (nm, sku, r,
        RB.get_cells(x, r, [d10]).get(d10) if r else "无行",
        RB.get_cells(x, r, [d9]).get(d9) if r else ""))
print("--- 染色 ---")
for nm in ("意大利","瑞典","德国","爱尔兰"):
    x = cur[nm]
    print("  %-5s 残留*2.5=%d 日列红*%s" % (nm, x.count("*2.5"),
          re.findall(r'AND\(\w7&gt;\$(?:G|F)7\*(\d+),', x)))
