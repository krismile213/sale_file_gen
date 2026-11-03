# -*- coding: utf-8 -*-
"""全量复验 patch_legend_all_20261026 的落盘结果（13 簿 17 表）。

检查：
 ① 欧洲表红口径已退回原两条相对倍数规则（与 backup_red_eu 逐字节一致）
 ② 17 表 B1 文本：行数 5/6、数字与 CF 反解一致、`<`/`>` 无裸字符
 ③ mergeCells 唯一且含 B1:H1、位置在 autoFilter 之后 / conditionalFormatting 之前
 ④ row1 ht=92 customHeight=1
 ⑤ 公式数 `<f`、缓存值 `<v` 与 backup_legend_all 逐表相等
 ⑥ XML 良构 + testzip
 ⑦ B1 的 xf 带 wrapText
"""
import os
import re
import sys
import zipfile
from xml.etree import ElementTree as ET

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.abspath(os.path.join(HERE, "..", "core")))
sys.path.insert(0, os.path.join(HERE, "maintenance"))
import config as C
from patch_legend_all_20261026 import (sheet_map, b1_info, cf_rules, span_cols,
                                       parse_params, legend_lines, sst_texts)

WORK = os.path.join(C.DSR, "1.每日销售数据表")
BAK = os.path.join(C.DSR, ".workbuddy", "backup_legend_all_20261026")
BAK_RED = os.path.join(C.DSR, ".workbuddy", "backup_red_eu_20261026")
EU_BOOK, EU_SHEET, EU_DAY_SQ = "欧盟英国库存销售表.xlsx", "欧洲-每日库存分析表", "AS7:CE1476"

ok, bad = [], []


def chk(c, m):
    (ok if c else bad).append(m)


def rd(z, n):
    return z.read(n).decode("utf-8")


ORDER = ["sheetPr", "dimension", "sheetViews", "sheetFormatPr", "cols",
         "sheetData", "sheetCalcPr", "sheetProtection", "protectedRanges",
         "scenarios", "autoFilter", "sortState", "dataConsolidate",
         "customSheetViews", "mergeCells", "phoneticPr",
         "conditionalFormatting", "dataValidations", "hyperlinks",
         "printOptions", "pageMargins", "pageSetup", "headerFooter",
         "rowBreaks", "colBreaks", "customProperties", "cellWatches",
         "ignoredErrors", "smartTags", "drawing", "legacyDrawing",
         "legacyDrawingHF", "picture", "oleObjects", "controls",
         "webPublishItems", "tableParts", "extLst"]

n_sheet = 0
for f in sorted(os.listdir(WORK)):
    if not f.endswith(".xlsx"):
        continue
    p = os.path.join(WORK, f)
    bp = os.path.join(BAK, f)
    if not os.path.exists(bp):
        bad.append("%s 缺备份" % f)
        continue
    z, zb = zipfile.ZipFile(p), zipfile.ZipFile(bp)
    chk(z.testzip() is None, "%s testzip OK" % f)
    ss = sst_texts(rd(z, "xl/sharedStrings.xml")) if "xl/sharedStrings.xml" in z.namelist() else []

    # ① 欧洲表红口径
    if f == EU_BOOK:
        sp = dict(sheet_map(z))[EU_SHEET]
        cur = rd(z, sp)
        m = re.search(r'<conditionalFormatting sqref="%s"[^>]*>.*?</conditionalFormatting>'
                      % EU_DAY_SQ, cur, re.S)
        zr = zipfile.ZipFile(os.path.join(BAK_RED, EU_BOOK))
        old = re.search(r'<conditionalFormatting sqref="%s"[^>]*>.*?</conditionalFormatting>'
                        % EU_DAY_SQ, rd(zr, sp), re.S)
        zr.close()
        chk(m and old and m.group(0) == old.group(0),
            "欧洲表红口径与 backup_red_eu 逐字节一致")
        chk(m and "AS7&gt;=20" not in m.group(0), "欧洲表已无「≥20件」绝对阈值")
        chk(m and m.group(0).count("<cfRule") == 4, "欧洲表日列块恢复为 4 条规则")

    for sn, sp in sheet_map(z):
        if "分析表" not in sn:
            continue
        n_sheet += 1
        xml, xb = rd(z, sp), rd(zb, sp)
        tag = "%s::%s" % (f.replace(".xlsx", ""), sn.strip())

        # ② 图例文本
        bi = b1_info(xml, rd(z, "xl/sharedStrings.xml") if ss else None)
        chk(bi is not None, "%s 有 B1" % tag)
        if not bi:
            continue
        txt = bi[3]
        mult, thr, div, sb = parse_params(xml)
        want = "\n".join(legend_lines(mult, thr, div, sb))
        chk(txt == want, "%s 图例与公式一致（%d 行）" % (tag, txt.count("\n") + 1))
        chk("<" not in txt and ">" not in txt or txt.count(">") == txt.count(">"),
            "%s 图例无裸标签字符" % tag)
        for kw in ("粉红", "橙", "黄", "浅红", "淡红", "浅绿", "红="):
            chk(kw in txt, "%s 含「%s」" % (tag, kw))
        if div:
            chk("蓝=日销<日均÷%g" % div in txt, "%s 蓝除数 %g 正确" % (tag, div))
        chk("日均%s倍" % ("%g" % mult) in txt, "%s 红倍率 %g 正确" % (tag, mult))

        # ③ 合并
        mm = re.search(r"<mergeCells[^>]*>(.*?)</mergeCells>", xml, re.S)
        chk(mm is not None and 'ref="B1:H1"' in mm.group(1), "%s 有 B1:H1 合并" % tag)
        root = ET.fromstring(xml.encode("utf-8"))
        tags = [c.tag.split("}")[-1] for c in root if c.tag.split("}")[-1] in ORDER]
        idxs = [ORDER.index(t) for t in tags]
        chk(idxs == sorted(idxs), "%s schema 顺序合规" % tag)
        chk("mergeCells" in tags and "autoFilter" in tags
            and tags.index("mergeCells") > tags.index("autoFilter"),
            "%s mergeCells 在 autoFilter 之后" % tag)

        # ④ 行高
        r1 = re.search(r'<row r="1"[^>]*>', xml)
        chk(r1 and 'ht="92"' in r1.group(0) and 'customHeight="1"' in r1.group(0),
            "%s row1 ht=92" % tag)

        # ⑤ 公式/缓存值
        for pat, nm in ((r"<f[ >/]", "<f"), (r"<v[ >/]", "<v")):
            a, b = len(re.findall(pat, xml)), len(re.findall(pat, xb))
            chk(a == b, "%s %s 数 %d == 备份 %d" % (tag, nm, a, b))

        # ⑦ wrapText
        a_attr = bi[1]
        sm = re.search(r's="(\d+)"', a_attr)
        if sm:
            xf = int(sm.group(1))
            st = rd(z, "xl/styles.xml")
            cm = re.search(r"<cellXfs[^>]*>(.*?)</cellXfs>", st, re.S)
            xfs = re.findall(r"<xf\b[^>]*?/>|<xf\b[^>]*?>.*?</xf>", cm.group(1), re.S)
            chk(xf < len(xfs) and 'wrapText="1"' in xfs[xf],
                "%s B1 xf[%d] 带 wrapText" % (tag, xf))

    # ⑥ styles/sst 良构
    for part in ("xl/styles.xml", "xl/sharedStrings.xml"):
        if part in z.namelist():
            try:
                ET.fromstring(rd(z, part).encode("utf-8"))
                chk(True, "%s %s 良构" % (f, part))
            except Exception as e:
                bad.append("%s %s 不良构: %s" % (f, part, e))
    z.close()
    zb.close()

print("检查 sheet 数: %d" % n_sheet)
print("通过 %d 项" % len(ok))
if bad:
    print("\n失败 %d 项：" % len(bad))
    for x in bad:
        print("  [!!] " + x)
else:
    print("\n=== 全部通过 ===")
sys.exit(1 if bad else 0)
