# -*- coding: utf-8 -*-
"""临时诊断：每日销售数据 sheet 的日列 / 日销是否被限定了窗口

读第 2 行表头之后，看辅助列（AN..AR 附近）的公式，确认日销 SUMIFS 的日期口径，
从而判断这张 sheet 里 9 月早期订单是否还有用。
"""
import zipfile, re

p = r"C:/Users/dell/Desktop/work-main/sale_file_gen/Daily_Sales_Report/1.每日销售数据表/欧盟英国库存销售表.xlsx"
z = zipfile.ZipFile(p)
x = z.read("xl/worksheets/sheet3.xml").decode()

m = re.search(r'<row r="2"[^>]*>(.*?)</row>', x, re.S)
hdr = {}
for mc in re.finditer(r'<c r="([A-Z]+)2"[^>]*>(?:<is><t>(.*?)</t></is>|<v>(.*?)</v>)', m.group(1), re.S):
    hdr[mc.group(1)] = (mc.group(2) or mc.group(3) or "")
print("第2行 27 列之后:", [(k, v) for k, v in hdr.items() if len(k) >= 2][26:])

for rn in (3, 4):
    m = re.search(r'<row r="%d"[^>]*>(.*?)</row>' % rn, x, re.S)
    if not m:
        continue
    print("--- row", rn)
    for mc in re.finditer(r'<c r="([A-Z]+)%d"[^>]*>(.*?)</c>' % rn, m.group(1), re.S):
        col = mc.group(1)
        body = mc.group(2)
        f = re.search(r"<f[^>]*>(.*?)</f>", body, re.S)
        if f:
            print("  %-4s %s" % (col, f.group(1)[:150]))
