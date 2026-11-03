# -*- coding: utf-8 -*-
"""排查 _verify_legend_layout 的三处失败。"""
import os
import re
import zipfile
from xml.etree import ElementTree as ET

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(
    os.path.abspath(__file__)))))
BOOK = os.path.join(ROOT, "Daily_Sales_Report", "1.每日销售数据表",
                    "欧盟英国库存销售表.xlsx")
BAK = os.path.join(ROOT, "Daily_Sales_Report", ".workbuddy",
                   "backup_legend_layout_20261026", "欧盟英国库存销售表.xlsx")
SHEET = "xl/worksheets/sheet10.xml"
SST = "xl/sharedStrings.xml"


def rd(p, part):
    with zipfile.ZipFile(p) as z:
        return z.read(part).decode("utf-8")


print("=" * 70)
print("[1] sheet10.xml 顶层标签序列")
s = rd(BOOK, SHEET)
# 只取根元素下的一级子元素：用 ET 遍历
root = ET.fromstring(s.encode("utf-8"))
print("  根:", root.tag)
for ch in root:
    print("   -", ch.tag, "(%d 子)" % len(list(ch)))

print()
print("=" * 70)
print("[2] sharedStrings 良构对比")
for label, p in (("当前", BOOK), ("备份", BAK)):
    t = rd(p, SST)
    try:
        ET.fromstring(t.encode("utf-8"))
        print("  %s: OK  len=%d" % (label, len(t)))
    except Exception as e:
        print("  %s: FAIL %s  len=%d" % (label, e, len(t)))
        m = re.search(r"column (\d+)", str(e))
        if m:
            col = int(m.group(1))
            lo = max(0, col - 160)
            print("    片段: ...%r..." % t[lo:col + 80])
        # 找非法控制字符
        bad = [(i, hex(ord(c))) for i, c in enumerate(t)
               if ord(c) < 0x20 and c not in "\t\n\r"]
        print("    非法控制字符 %d 个: %s" % (len(bad), bad[:5]))

print()
print("=" * 70)
print("[3] B1 图例实际字数")
t = rd(BOOK, SST)
sis = list(re.finditer(r"<si>(.*?)</si>", t, re.S))
raw = "".join(re.findall(r"<t[^>]*>(.*?)</t>", sis[152553].group(1), re.S))
plain = (raw.replace("&gt;", ">").replace("&lt;", "<")
            .replace("&amp;", "&").replace("&#10;", "\n"))
print("  原始 XML 片段长度:", len(raw))
print("  还原后长度:", len(plain.replace("\n", "")))
print("  逐行字数:", [len(x) for x in plain.split("\n")])
print("  行数:", len(plain.split("\n")))
print("  原始片段:", repr(raw[:80]))
