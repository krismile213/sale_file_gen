# -*- coding: utf-8 -*-
"""诊断：openpyxl 加载欧盟簿到哪一步卡住（逐 sheet 计时）"""
import os, sys, time, zipfile
import openpyxl
from openpyxl import load_workbook

p = r"C:/Users/dell/Desktop/work-main/sale_file_gen/Daily_Sales_Report/1.每日销售数据表/欧盟英国库存销售表.xlsx"
t = time.time()

# 1. 看 zip 内各部分大小
z = zipfile.ZipFile(p)
items = sorted(z.infolist(), key=lambda i: -i.file_size)
print("=== zip 最大 10 个部件 ===", flush=True)
for i in items[:10]:
    print("  %-45s raw=%8.1fMB comp=%6.1fMB" % (i.filename, i.file_size / 1e6, i.compress_size / 1e6), flush=True)
tot = sum(i.file_size for i in items)
print("  total raw %.1fMB" % (tot / 1e6), flush=True)
z.close()
print("zip 扫描 %.1fs" % (time.time() - t), flush=True)

# 2. 逐 sheet 计时（read_only 快，普通模式慢）
print("=== 普通模式逐 sheet ===", flush=True)
t = time.time()
wb = load_workbook(p, read_only=True)
print("  read_only 加载 %.1fs" % (time.time() - t), flush=True)
for sn in wb.sheetnames:
    t0 = time.time()
    ws = wb[sn]
    n = ws.max_row
    print("  %-24s max_row=%-7s %.2fs" % (sn, n, time.time() - t0), flush=True)
wb.close()
print("完成 %.1fs" % (time.time() - t), flush=True)
