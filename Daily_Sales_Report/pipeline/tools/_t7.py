# -*- coding: utf-8 -*-
"""临时实验：读入 → 立即写回（不编辑任何内容），看是否可行 + 耗时

避免 openpyxl 层面的大对象（不用 load_workbook 全量），改用「zip 级复制」：
这正是 restore_parts 的做法。目的是验证：如果我们**不加载整簿**，
是否可以用 XML 手术完成补行修复（快、省内存）。
"""
import os, sys, time, zipfile, shutil, re

SRC = r"C:/Users/dell/Desktop/work-main/sale_file_gen/Daily_Sales_Report/1.每日销售数据表/欧盟英国库存销售表.xlsx"
DST = r"C:/Users/dell/Desktop/work-main/sale_file_gen/_archive/_t7_out.xlsx"

t = time.time()
z = zipfile.ZipFile(SRC)
with zipfile.ZipFile(DST, "w", zipfile.ZIP_DEFLATED) as zo:
    for i in z.infolist():
        data = z.read(i.filename)
        zo.writestr(i, data)
z.close()
print("zip 复制耗时 %.1fs, 大小 %.1fMB" % (time.time() - t, os.path.getsize(DST) / 1e6), flush=True)

# 只读 + 只解析分析表 XML（6MB）
t = time.time()
z = zipfile.ZipFile(DST)
x = z.read("xl/worksheets/sheet10.xml").decode()
print("读 sheet10 耗时 %.2fs, %.1fMB" % (time.time() - t, len(x) / 1e6), flush=True)

t = time.time()
rows = re.findall(r'<row r="(\d+)"[^>]*>.*?</row>', x, re.S)
print("正则提取行 %.2fs, %d 行" % (time.time() - t, len(rows)), flush=True)
