# -*- coding: utf-8 -*-
"""在副本上验证：站点表能否用 append_sku_rows_auto 补行（G598→意大利、K001→瑞典）。"""
import os, sys, shutil, tempfile
HERE = os.path.dirname(os.path.abspath(__file__))
CORE = os.path.abspath(os.path.join(HERE, "..", "core"))
sys.path.insert(0, CORE)
import config as C
import runner as R
import analysis as An
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter as GL

SRC = os.path.join(C.WORK_DIR, "欧盟英国库存销售表.xlsx")
TMP = tempfile.mkdtemp(prefix="siteappend_")
cp = os.path.join(TMP, "欧盟英国库存销售表.xlsx")
shutil.copy2(SRC, cp)
st = cp + ".strip.xlsx"
n = R.strip_pivot_records(cp, st)
wb = load_workbook(st if n else cp, data_only=False)
print("pivot 剥离:", n)

# 品名从欧洲分析表取
eu = wb["欧洲-每日库存分析表"]
nm = {}
for r in range(7, eu.max_row + 1):
    v = eu.cell(r, 1).value
    if v in ("G598", "K001"):
        nm[v] = eu.cell(r, 2).value
print("品名:", nm)

PLAN = [("意大利", ["G598"]), ("瑞典", ["K001"])]
for sn, skus in PLAN:
    ws = wb[sn]
    last = An._last_data_row(ws)
    tpl = An.pick_template_row(ws, last)
    print("\n--- %s --- last=%d 模板行=%d" % (sn, last, tpl))
    e6 = ws.cell(6, 5).value
    print("  E6 公式: %s" % str(e6)[:60])
    added = An.append_sku_rows_auto(ws, skus, name_map=nm, state_map={s: "在售" for s in skus})
    print("  追加:", added)
    r = added[0][0]
    for col in ("A", "B", "C", "D", "E", "F", "G", "H", "I", "AU", "AV", "BH", "BI", "BK", "BZ"):
        print("    %s%d = %s" % (col, r, str(ws["%s%d" % (col, r)].value)[:72]))
    # 日列 SUMIFS 应该指向新行
    print("    日列头 I5 = %s" % str(ws["I5"].value)[:40])

out = os.path.join(TMP, "out.xlsx")
wb.save(out); wb.close()
print("\n已保存:", out)
print("临时目录:", TMP)
