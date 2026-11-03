# -*- coding: utf-8 -*-
"""跑批后复查：图例排版 / 红色口径 / 橙色规则 是否被 openpyxl 重写冲掉。"""
import os
import re
import sys
import zipfile
from xml.etree import ElementTree as ET

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.abspath(os.path.join(HERE, "..", "core")))
sys.path.insert(0, os.path.join(HERE, "maintenance"))
import config as C
from patch_legend_all_20261026 import sheet_map, b1_info, cf_rules, span_cols

WORK = os.path.join(C.DSR, "1.每日销售数据表")
ok = bad = 0
rows = []


def chk(c, msg):
    global ok, bad
    if c:
        ok += 1
    else:
        bad += 1
        rows.append("  [!!] " + msg)


for f in sorted(os.listdir(WORK)):
    if not f.endswith(".xlsx"):
        continue
    p = os.path.join(WORK, f)
    z = zipfile.ZipFile(p)
    ss = z.read("xl/sharedStrings.xml").decode() if "xl/sharedStrings.xml" in z.namelist() else None
    for sn, sp in sheet_map(z):
        if "分析表" not in sn:
            continue
        tag = "%s::%s" % (f.replace(".xlsx", ""), sn.strip())
        xml = z.read(sp).decode()
        bi = b1_info(xml, ss)
        chk(bi is not None, "%s 无 B1" % tag)
        if not bi:
            continue
        txt = bi[3]
        nline = txt.count("\n") + 1
        chk(nline in (5, 6), "%s 图例 %d 行（应 5/6）" % (tag, nline))
        chk("粉红=缺货" in txt and "橙=" in txt, "%s 图例内容完整" % tag)
        # 合并
        mm = re.search(r"<mergeCells[^>]*>(.*?)</mergeCells>", xml, re.S)
        chk(mm and 'ref="B1:H1"' in mm.group(1), "%s B1:H1 合并丢失" % tag)
        # 行高
        r1 = re.search(r'<row r="1"[^>]*>', xml)
        chk(r1 and 'ht="92"' in r1.group(0), "%s row1 ht 丢失（%s）"
            % (tag, r1.group(0)[:60] if r1 else "无"))
        # 红色口径：日列必须还是相对倍数
        reds = []
        for sq, dx, pri, fo in cf_rules(xml):
            if span_cols(sq) > 20 and "*" in fo and "&lt;" not in fo:
                reds.append((sq, fo))
        chk(len(reds) >= 1, "%s 日列红规则缺失" % tag)
        for sq, fo in reds:
            m = re.search(r"\*([\d.]+),\$?[A-Z]+7&gt;([\d.]+)\)$", fo)
            chk(m is not None, "%s 红公式形态异常: %s" % (tag, fo[:60]))
            if m:
                mult = float(m.group(1))
                chk(mult in (2.0, 2.5, 2.8, 3.0),
                    "%s 红倍率异常 %g" % (tag, mult))
                chk("&gt;=20" not in fo, "%s 仍有绝对阈值 ≥20" % tag)
        # 橙色规则
        has_orange = any('$D7=&quot;在途&quot;' in fo or '"在途"' in fo
                         for _, _, _, fo in cf_rules(xml))
        chk(has_orange, "%s 橙色在途规则丢失" % tag)
        # XML 良构
        try:
            ET.fromstring(xml.encode("utf-8"))
        except Exception as e:
            chk(False, "%s XML 不良构 %s" % (tag, e))
    z.close()

print("通过 %d 项" % ok)
if bad:
    print("失败 %d 项：" % bad)
    for r in rows:
        print(r)
else:
    print("=== 图例排版 / 红口径 / 橙色规则 全部完好 ===")
sys.exit(1 if bad else 0)
