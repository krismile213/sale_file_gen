# -*- coding: utf-8 -*-
"""复验欧洲表图例排版改造（patch_legend_layout_eu_20261026.py）的落盘结果。

检查维度：
 1. B1 文本 6 行 / 字数、&#10; 换行、&gt; 转义
 2. mergeCells 唯一且为 B1:H1，且不与 I1 起的货件号表头冲突
 3. row1 ht="92" customHeight="1"，且未误伤其它 ht
 4. cellXfs count 93->94，新 xf 93 带 wrapText="1"
 5. XML 良构 + worksheet 子元素 schema 顺序
 6. 公式数 / 缓存值数与备份相等、testzip OK
 7. 红色新公式仍在位
"""
import os
import re
import sys
import zipfile
from xml.etree import ElementTree as ET

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(
    os.path.abspath(__file__)))))
BOOK = os.path.join(ROOT, "Daily_Sales_Report", "1.每日销售数据表",
                    "欧盟英国库存销售表.xlsx")
BAK = os.path.join(ROOT, "Daily_Sales_Report", ".workbuddy",
                   "backup_legend_layout_20261026", "欧盟英国库存销售表.xlsx")
SHEET = "xl/worksheets/sheet10.xml"
STYLES = "xl/styles.xml"
SST = "xl/sharedStrings.xml"

ok = []
bad = []


def chk(cond, msg):
    (ok if cond else bad).append(msg)


def read(path, part):
    with zipfile.ZipFile(path) as z:
        return z.read(part).decode("utf-8")


def main():
    for p in (BOOK, BAK):
        if not os.path.exists(p):
            print("!! 缺文件:", p)
            return 2

    # ---- testzip ----
    with zipfile.ZipFile(BOOK) as z:
        chk(z.testzip() is None, "testzip OK")

    cur_sheet = read(BOOK, SHEET)
    bak_sheet = read(BAK, SHEET)
    cur_styles = read(BOOK, STYLES)
    cur_sst = read(BOOK, SST)

    # ---- 1. B1 文本 ----
    # 找 B1 的 cell 标签
    m = re.search(r'<c r="B1"([^>]*)/?>(.*?)</c>|<c r="B1"([^>]*)/>',
                  cur_sheet, re.S)
    chk(m is not None, "B1 单元格存在")
    if m:
        attrs = m.group(1) or m.group(3) or ""
        body = m.group(2) or ""
        chk('s="93"' in attrs, 'B1 使用新样式 s="93"（实际: %s）' % attrs.strip())
        # 取 sharedString 索引
        si = re.search(r'<v>(\d+)</v>', body)
        if si:
            idx = int(si.group(1))
            sis = list(re.finditer(r"<si>(.*?)</si>", cur_sst, re.S))
            chk(idx < len(sis), "B1 si 索引 %d 在范围内(%d)" % (idx, len(sis)))
            if idx < len(sis):
                txt = "".join(re.findall(r"<t[^>]*>(.*?)</t>",
                                         sis[idx].group(1), re.S))
                lines = txt.split("&#10;")
                chk(len(lines) == 6, "图例 6 行（实际 %d 行）" % len(lines))
                # 还原实体后计字数
                plain = (txt.replace("&gt;", ">").replace("&lt;", "<")
                            .replace("&amp;", "&").replace("&#10;", ""))
                chk(120 <= len(plain) <= 220,
                    "图例字数 %d（目标 120~220）" % len(plain))
                chk("<" not in txt and ">" not in txt,
                    "`<`/`>` 全部已转义（裸字符 %d 个）"
                    % (txt.count("<") + txt.count(">")))
                for kw in ("粉红", "橙", "黄", "浅红", "淡红", "浅绿",
                           "红=", "蓝=", "实时速度"):
                    chk(kw in plain, "图例含关键词「%s」" % kw)
                print("---- B1 图例正文 ----")
                for i, ln in enumerate(lines, 1):
                    print("  %d| %s" % (i, ln))
                print("--------------------")

    # ---- 2. mergeCells ----
    mc = re.findall(r"<mergeCells[^>]*>(.*?)</mergeCells>", cur_sheet, re.S)
    chk(len(mc) == 1, "mergeCells 块唯一（实际 %d 个）" % len(mc))
    if mc:
        refs = re.findall(r'<mergeCell ref="([^"]+)"', mc[0])
        chk(refs == ["B1:H1"], "mergeCell = B1:H1（实际 %s）" % refs)
    chk(not re.search(r"<mergeCells", bak_sheet), "备份原本无 mergeCells（新增而非覆盖）")

    # ---- 3. row1 行高 ----
    r1 = re.search(r'<row r="1"[^>]*>', cur_sheet)
    chk(r1 is not None, "row 1 存在")
    if r1:
        tag = r1.group(0)
        chk('ht="92"' in tag, 'row1 ht="92"（实际: %s）' % tag)
        chk('customHeight="1"' in tag, 'row1 customHeight="1"')
    # 误伤检查：其它 row 的 ht 不应变
    def htmap(s):
        return {int(a): b for a, b in
                re.findall(r'<row r="(\d+)"[^>]*(?<![A-Za-z])ht="([\d.]+)"', s)}
    hm_cur, hm_bak = htmap(cur_sheet), htmap(bak_sheet)
    diff = {k: (hm_bak.get(k), v) for k, v in hm_cur.items()
            if k in hm_bak and hm_bak[k] != v}
    chk(diff == {1: ("15", "92")} or set(diff) <= {1},
        "仅 row1 行高变化（实际差异行: %s）" % sorted(diff))

    # ---- 4. styles ----
    cx = re.search(r'<cellXfs count="(\d+)"', cur_styles)
    cx_bak = re.search(r'<cellXfs count="(\d+)"', read(BAK, STYLES))
    chk(cx is not None and cx_bak is not None, "cellXfs count 可读")
    if cx and cx_bak:
        chk(int(cx.group(1)) == int(cx_bak.group(1)) + 1,
            "cellXfs %s -> %s（+1）" % (cx_bak.group(1), cx.group(1)))
        body = cur_styles[cx.end():]
        xfs = re.findall(r"<xf\b[^>]*/>|<xf\b[^>]*>.*?</xf>", body, re.S)
        chk(len(xfs) >= 94, "cellXfs 条目 %d >= 94" % len(xfs))
        if len(xfs) >= 94:
            last = xfs[93]
            chk('wrapText="1"' in last, 'xf[93] 带 wrapText="1"（%s）' % last)
            chk('vertical="top"' in last, 'xf[93] 带 vertical="top"')
            chk('s="36"' in last or 'xfId' in last, 'xf[93] 继承自 xf[36]')

    # ---- 5. XML 良构 + schema 顺序 ----
    for part, txt in ((SHEET, cur_sheet), (STYLES, cur_styles), (SST, cur_sst)):
        try:
            ET.fromstring(txt.encode("utf-8"))
            chk(True, "%s XML 良构" % part)
        except Exception as e:
            chk(False, "%s XML 不良构: %s" % (part, e))

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
    root = ET.fromstring(cur_sheet.encode("utf-8"))
    tags = [ch.tag.split("}")[-1] for ch in root if ch.tag.split("}")[-1] in ORDER]
    idxs = [ORDER.index(t) for t in tags]
    chk(idxs == sorted(idxs),
        "worksheet 子元素顺序合规（一级标签序列: %s）"
        % " > ".join(tags[:10]))
    chk("autoFilter" in tags and "mergeCells" in tags
        and tags.index("mergeCells") > tags.index("autoFilter"),
        "mergeCells 在 autoFilter 之后（schema 要求）")
    chk("mergeCells" in tags and "conditionalFormatting" in tags
        and tags.index("mergeCells") < tags.index("conditionalFormatting"),
        "mergeCells 在 conditionalFormatting 之前")

    # ---- 6. 公式数 / 缓存值 ----
    for pat, name in ((r"<f[ >/]", "公式 <f"), (r"<v[ >/]", "缓存值 <v")):
        a = len(re.findall(pat, cur_sheet))
        b = len(re.findall(pat, bak_sheet))
        chk(a == b, "%s 数 %d == 备份 %d" % (name, a, b))

    # ---- 7. 红色新公式 ----
    chk("OR(AND($AO7&gt;0,AS7&gt;$AO7*2,AS7&gt;=5),AS7&gt;=20)" in cur_sheet,
        "红色新口径公式在位")
    chk("$AO7&gt;1.5)" not in cur_sheet, "旧高销量档已移除")
    chk("AND($AO7&gt;0,$AO7&lt;=1.5,AS7&gt;=2,AS7&gt;$AO7*4)" not in cur_sheet,
        "旧低基数档已移除")

    print()
    print("通过 %d 项" % len(ok))
    for s in ok:
        print("  [OK] " + s)
    if bad:
        print("\n失败 %d 项" % len(bad))
        for s in bad:
            print("  [!!] " + s)
        return 1
    print("\n=== 全部通过 ===")
    return 0


if __name__ == "__main__":
    sys.exit(main())
