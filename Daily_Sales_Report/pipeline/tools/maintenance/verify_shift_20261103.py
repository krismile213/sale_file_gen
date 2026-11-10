# -*- coding: utf-8 -*-
"""验证平移结果：日期、星期、公式、折叠、条件格式、zip 完整性。"""
import os, sys, datetime, zipfile
HERE = os.path.dirname(os.path.abspath(__file__))
DSR = os.path.dirname(os.path.dirname(os.path.dirname(HERE)))
import openpyxl
from openpyxl.utils import get_column_letter

BK = os.path.join(DSR, "1.每日销售数据表", "印度库存销售表.xlsx")
BAK = os.path.join(DSR, ".workbuddy", "backup_shift_daycols_20261103", "印度库存销售表.xlsx")
ok = True

def chk(name, cond, detail=""):
    global ok
    print(("  PASS " if cond else "  FAIL ") + name, detail)
    if not cond:
        ok = False

# zip 完整
zz = zipfile.ZipFile(BK)
chk("zip完整", zz.testzip() is None)
zz.close()

wb = openpyxl.load_workbook(BK, data_only=False)
ws = wb["印度-每日库存分析表"]
# 1. 行5 日期
days = []
for row in ws.iter_rows(min_row=5, max_row=5, min_col=20, max_col=130):
    for c in row:
        if hasattr(c.value, "strftime"):
            days.append((c.column, c.value.date()))
chk("日列数=68", len(days) == 68, str(len(days)))
chk("首日=2026-10-25", days[0][1] == datetime.date(2026, 10, 25), str(days[0]))
chk("末日=2026-12-31", days[-1][1] == datetime.date(2026, 12, 31), str(days[-1]))
exp = [datetime.date(2026, 10, 25) + datetime.timedelta(days=i) for i in range(68)]
chk("日期连续无断档", [d for _, d in days] == exp)
# 2. 行1 星期
WEEK = ["周一", "周二", "周三", "周四", "周五", "周六", "周日"]
wbad = []
for c, d in days:
    v = ws.cell(1, c).value
    if v not in (None, "") and v != WEEK[d.weekday()]:
        wbad.append((get_column_letter(c), v, str(d)))
chk("行1星期与日期匹配", not wbad, str(wbad[:4]))
# 3. SUMIFS 公式未变 + 引用本列$5
Y = get_column_letter(days[0][0])
f_new = ws["%s7" % Y].value
f_bak_ws = None
wb2 = openpyxl.load_workbook(BAK, data_only=False)
ws2 = wb2["印度-每日库存分析表"]
f_old = ws2["%s7" % Y].value
wb2.close()
chk("SUMIFS公式文本未变", f_new == f_old)
chk("公式引用本列$5", ("%s$5" % Y) in str(f_new), str(f_new)[:80])
# 4. 折叠
hidden_after = [d for c, d in days if ws.column_dimensions[get_column_letter(c)].hidden]
vis = [d for c, d in days if not ws.column_dimensions[get_column_letter(c)].hidden]
chk("可见=10/25~11/3", vis and vis[0] == datetime.date(2026,10,25) and vis[-1] == datetime.date(2026,11,3) and len(vis) == 10,
    "%s~%s n=%d" % (vis[0] if vis else '-', vis[-1] if vis else '-', len(vis)))
chk("隐藏=11/4~12/31", hidden_after and hidden_after[0] == datetime.date(2026,11,4) and hidden_after[-1] == datetime.date(2026,12,31) and len(hidden_after) == 58)
# 5. 条件格式
cf_new = len(list(ws.conditional_formatting))
ws2 = None
wb2 = openpyxl.load_workbook(BAK)
ws2 = wb2["印度-每日库存分析表"]
cf_old = len(list(ws2.conditional_formatting))
wb2.close()
chk("条件格式规则数不变", cf_new == cf_old, "%d vs %d" % (cf_old, cf_new))
wb.close()
print("RESULT:", "ALL PASS" if ok else "HAS FAILURES")
