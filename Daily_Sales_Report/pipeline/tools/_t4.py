# -*- coding: utf-8 -*-
"""临时诊断：每日销售数据 sheet 的真实数据范围与日列 / 日销统计

用来判断能不能把这张 53MB 的 sheet 按「订单日期窗口」裁短（瘦身），
以免 openpyxl 普通模式加载卡死（整个簿 137MB 解压后）。
"""
import os, sys, zipfile, re
from collections import Counter

p = r"C:/Users/dell/Desktop/work-main/sale_file_gen/Daily_Sales_Report/1.每日销售数据表/欧盟英国库存销售表.xlsx"
z = zipfile.ZipFile(p)
x = z.read("xl/worksheets/sheet3.xml").decode()

dim = re.search(r'<dimension ref="([^"]+)"', x)
print("dimension:", dim.group(1) if dim else None)

# 表头
m = re.search(r'<row r="2"[^>]*>(.*?)</row>', x, re.S)
hdr = {}
for mc in re.finditer(r'<c r="([A-Z]+)2"[^>]*>(?:<is><t>(.*?)</t></is>|<v>(.*?)</v>)', m.group(1), re.S):
    hdr[mc.group(1)] = (mc.group(2) or mc.group(3) or "")
print("表头前 26 列:", [(k, v) for k, v in list(hdr.items())[:26]])

# purchase-date 列位置
pdc = next((k for k, v in hdr.items() if v.strip() == "purchase-date"), None)
skuc = next((k for k, v in hdr.items() if v.strip() == "sku"), None)
print("purchase-date 列:", pdc, " sku 列:", skuc)

# 逐行取日期，统计每月行数
months = Counter()
minmax = [None, None]
rows = re.findall(r'<row r="(\d+)"[^>]*>(.*?)</row>', x, re.S)
for rn, body in rows:
    if int(rn) <= 2:
        continue
    mc = re.search(r'<c r="%s%s"[^>]*>(?:<is><t>(.*?)</t></is>|<v>(.*?)</v>)' % (pdc, rn), body, re.S)
    if not mc:
        continue
    d = (mc.group(1) or mc.group(2) or "")[:7]
    if not d:
        continue
    months[d] += 1
    if minmax[0] is None or d < minmax[0]:
        minmax[0] = d
    if minmax[1] is None or d > minmax[1]:
        minmax[1] = d
print("purchase-date 范围:", minmax)
print("按月行数(最近12):", sorted(months.items())[-12:])
print("总行数(解析到日期):", sum(months.values()))
print("sheet 行数:", len(rows))
