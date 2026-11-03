# -*- coding: utf-8 -*-
"""就地修正 3 个次表交付件的月份表头（Aug→Sep / Jan-Aug→Jan-Sep），数据零改动。

背景：southeast_asia_sale_file.py L620-621 把 "Aug"/"Jan-Aug" 写死，不随归属月变，
202609 期生成的三个次表数据是 9 月 + 1~9 月累计，但表头还是 8 月字样。
本探针按用户要求：不改脚本、不改文件名，只把表头单元格写对。
改前自动备份到 D:\\dev-cache\\sale_file_gen\\backup\\东南亚次表_表头修正_20261015\\。
"""
import os
import shutil
import datetime
from openpyxl import load_workbook

D = r"D:\Desktop\work-main\sale_file_gen\trunk\result\东南亚销量表\202609"
BAK = r"D:\dev-cache\sale_file_gen\backup\东南亚次表_表头修正_20261015"
FILES = ["Philippines_202609_sales.xlsx", "Indonesia_202609_sales.xlsx", "Thailand_202609_sales.xlsx"]

RENAME = {
    "Aug 2026 Total": "Sep 2026 Total",
    "Jan-Aug 2026 Cumulative": "Jan-Sep 2026 Cumulative",
}


def num(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return 0.0


def snapshot(ws):
    """除行1表头外的全部数据格快照（值）。"""
    snap = {}
    for row in ws.iter_rows(min_row=2):
        for c in row:
            if c.value is not None:
                snap[(c.row, c.column)] = c.value
    return snap


os.makedirs(BAK, exist_ok=True)
print(f"备份目录：{BAK}")
for f in FILES:
    src = os.path.join(D, f)
    if not os.path.exists(os.path.join(BAK, f)):
        shutil.copy2(src, os.path.join(BAK, f))
    print(f"  已备份 {f}")

print("\n" + "=" * 78)
for f in FILES:
    p = os.path.join(D, f)
    wb = load_workbook(p)  # 非 read_only，保留全部内容
    ws = wb["Sheet1"]

    before_head = [c.value for c in ws[1] if c.value is not None]
    before_data = snapshot(ws)

    # 1) 改表头（只动行1 的指定两个单元格）
    changed = []
    for c in ws[1]:
        if c.value in RENAME:
            changed.append((c.coordinate, c.value, RENAME[c.value]))
            c.value = RENAME[c.value]

    wb.save(p)
    wb.close()

    # 2) 重开验证：表头已改 + 数据零变化
    wb2 = load_workbook(p, read_only=True, data_only=True)
    ws2 = wb2["Sheet1"]
    after_head = [c.value for c in ws2[1] if c.value is not None]
    after_data = snapshot(ws2)
    # 合计/累计列求和
    hi = next(i for i, h in enumerate(after_head) if "Total" in str(h))
    ci = next(i for i, h in enumerate(after_head) if "Cumulative" in str(h))
    tot = sum(num(r[hi]) for r in ws2.iter_rows(min_row=2, values_only=True) if r[0] is not None)
    cum = sum(num(r[ci]) for r in ws2.iter_rows(min_row=2, values_only=True) if r[0] is not None)
    wb2.close()

    data_same = before_data == after_data
    print(f"\n[{f}]")
    print(f"  改动单元格：{changed if changed else '（无）'}")
    print(f"  改后表头：{after_head}")
    print(f"  数据区（行2起全部非空格）零变化：{'✅ 是' if data_same else '★ 否！'}")
    print(f"  当月合计列求和={tot:.0f}   累计列求和={cum:.0f}")
    st = os.stat(p)
    print(f"  文件 mtime={datetime.datetime.fromtimestamp(st.st_mtime):%m-%d %H:%M:%S}  大小={st.st_size:,}B")

print("\n完成：三个次表表头已就地修正，数据未动，原件已备份。")
