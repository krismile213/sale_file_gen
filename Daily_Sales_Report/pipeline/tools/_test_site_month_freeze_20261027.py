# -*- coding: utf-8 -*-
"""在副本上验证：apply_month_summary 现在能把 10 张站点表一起封存。"""
import os, sys, shutil, tempfile
HERE = os.path.dirname(os.path.abspath(__file__))
CORE = os.path.abspath(os.path.join(HERE, "..", "core"))
sys.path.insert(0, CORE)
import config as C
import runner as R
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter as GL

SRC = os.path.join(C.WORK_DIR, "欧盟英国库存销售表.xlsx")
TMP = tempfile.mkdtemp(prefix="sitemonth_")
copy_p = os.path.join(TMP, "欧盟英国库存销售表.xlsx")
shutil.copy2(SRC, copy_p)

stripped = copy_p + ".strip.xlsx"
n = R.strip_pivot_records(copy_p, stripped)
load_p = stripped if n else copy_p
print("pivot 记录剥离:", n)
wb = load_workbook(load_p, data_only=False)
print("sheet:", len(wb.sheetnames))

R.apply_month_summary(wb, "20261027", fname="欧盟英国库存销售表.xlsx")
out = os.path.join(TMP, "out.xlsx")
wb.save(out)
wb.close()
print("\n已保存副本:", out)

# ---- 复验 ----
# 欧洲分析表 9 月 = 24981（派生地区口径，含 Amazon.ae 1 件 G216），
# 而 10 张站点表合计 = 24980 —— 差 1 是口径本身决定的，不是错。
EXP = {"英国-每日库存分析表": (8016, 6854), "欧洲-每日库存分析表": (28319, 24981),
       "意大利": (4873, 4425), "法国": (5351, 4339), "西班牙": (5772, 5425),
       "德国": (8009, 7144), "土耳其": (0, 0), "荷兰": (1081, 793),
       "瑞典": (789, 679), "波兰": (331, 295), "比利时": (1075, 904), "爱尔兰": (1038, 976)}
wb2 = load_workbook(out, data_only=False)
ok = bad = 0
for sn, (e8, e9) in EXP.items():
    ws = wb2[sn]
    lab = {}
    for c in range(1, (ws.max_column or 0) + 1):
        v = ws.cell(5, c).value
        if isinstance(v, str) and v.strip().endswith("月") and v.strip()[:-1].isdigit():
            lab.setdefault(int(v.strip()[:-1]), []).append(c)
    c8, c9 = min(lab[8]), min(lab[9])
    s8 = sum(ws.cell(r, c8).value or 0 for r in range(7, 7 + 3000)
             if ws.cell(r, 1).value is not None and isinstance(ws.cell(r, c8).value, (int, float)))
    s9 = sum(ws.cell(r, c9).value or 0 for r in range(7, 7 + 3000)
             if ws.cell(r, 1).value is not None and isinstance(ws.cell(r, c9).value, (int, float)))
    flag = "OK " if (int(s8), int(s9)) == (e8, e9) else "✗  "
    if flag.strip() == "OK": ok += 1
    else: bad += 1
    g = next((c for c in range(1, 20) if str(ws.cell(5, c).value).strip() == "上月日均"), None)
    gv = ws.cell(7, g).value if g else None
    print("  %s %-16s 8月=%-6d(期望%d) 9月=%-6d(期望%d)  %s7=%s" % (
        flag, sn, s8, e8, s9, e9, GL(g) if g else "-", str(gv)[:44]))
wb2.close()
print("\n通过 %d / 失败 %d" % (ok, bad))
print("临时目录:", TMP)
