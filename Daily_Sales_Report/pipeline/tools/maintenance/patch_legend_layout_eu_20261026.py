# -*- coding: utf-8 -*-
"""欧洲表图例：更新红色描述 + 排版（合并 B1:H1、自动换行、行高 92）。

背景：B1 的图例文本到 H 列就被截断（I1 起是货件号表头，Excel 不往有内容的单元格溢出），
      所以 233 字的图例实际只显示前 ~28 个汉字 → 用户反馈「拥挤」。

改动：
  ① 红色描述换成新口径：日销≥20件，或>日均2倍且≥5件
  ② 文本压成 6 行，用 &#10; 换行
  ③ 合并 B1:H1（C1~H1 为空），开启 wrapText
  ④ 行 1 高 15 → 92

用法：
  python patch_legend_layout_eu_20261026.py            # 干跑
  python patch_legend_layout_eu_20261026.py --apply    # 落盘（先备份）

修订 2026-10-26 19:05（v2）：
  - 修 BUG1：原文只转义了 `>`，`蓝=日销<日均÷3` 的裸 `<` 让 sharedStrings.xml 不良构。
            改为按 `&` → `<` → `>` 顺序完整转义。
  - 修 BUG2：mergeCells 原插在 </sheetData> 后，落在 autoFilter 之前，违反 schema 顺序。
            改为插在第一个 conditionalFormatting 之前（autoFilter 之后）。
  - 新增落盘前 XML 良构校验（这正是 v1 没拦住 BUG1 的缺口）。
"""
import os
import re
import shutil
import sys
import zipfile
from xml.etree import ElementTree as ET

HERE = os.path.dirname(os.path.abspath(__file__))
DSR = HERE
for _ in range(6):
    if os.path.isdir(os.path.join(DSR, "1.每日销售数据表")):
        break
    DSR = os.path.dirname(DSR)
WORK = os.path.join(DSR, "1.每日销售数据表")
BAK = os.path.join(DSR, ".workbuddy", "backup_legend_layout_20261026")
APPLY = "--apply" in sys.argv

BOOK = "欧盟英国库存销售表.xlsx"
SHEET = "欧洲-每日库存分析表"

LINES = [
    "粉红=缺货(深粉:无补给　浅粉:补给在途)　橙=在途未上架",
    "黄=保底　浅红=SKU重复或单日>20　淡红=命中透视表",
    "浅绿=保底+机会产品且月内最高>4",
    "红=日销≥20件，或>日均2倍且≥5件（异常放量）",
    "蓝=日销<日均÷3",
    "实时速度列：红>日均1.5倍，蓝<日均1/2(日均>2)",
]


def xml_escape(s):
    """完整转义，顺序必须是 & 先行，否则会把 &lt; 再转成 &amp;lt;。"""
    return (s.replace("&", "&amp;")
             .replace("<", "&lt;")
             .replace(">", "&gt;"))


NEW_TEXT = "&#10;".join(xml_escape(x) for x in LINES)
ROW_HT = 92
MERGE = "B1:H1"


def sheet_path(z, want):
    wb = z.read("xl/workbook.xml").decode("utf-8", "replace")
    rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8", "replace")
    rid2t = {}
    for tag in re.findall(r"<Relationship\b[^>]*/?>", rels):
        rid = re.search(r'\bId="([^"]+)"', tag)
        tgt = re.search(r'\bTarget="([^"]+)"', tag)
        if rid and tgt:
            rid2t[rid.group(1)] = tgt.group(1)
    for tag in re.findall(r"<sheet\b[^>]*/?>", wb):
        nm = re.search(r'\bname="([^"]+)"', tag)
        rid = re.search(r'\br:id="([^"]+)"', tag)
        if nm and rid and nm.group(1) == want:
            t = rid2t.get(rid.group(1), "")
            sp = t if t.startswith("xl/") else "xl/" + t.lstrip("/")
            return sp.replace("xl/xl/", "xl/")
    return None


def insert_merge(sheet):
    """把 mergeCells 插到 schema 正确位置：autoFilter 之后、conditionalFormatting 之前。"""
    block = f'<mergeCells count="1"><mergeCell ref="{MERGE}"/></mergeCells>'
    for anchor in ("<phoneticPr", "<conditionalFormatting", "<dataValidations",
                   "<hyperlinks", "<printOptions", "<pageMargins",
                   "<pageSetup", "<headerFooter", "<drawing"):
        i = sheet.find(anchor)
        if i != -1:
            return sheet[:i] + block + sheet[i:], anchor
    return None, None


def main():
    p = os.path.join(WORK, BOOK)
    z = zipfile.ZipFile(p)
    sp = sheet_path(z, SHEET)
    sheet = z.read(sp).decode("utf-8")
    styles = z.read("xl/styles.xml").decode("utf-8")
    sst = z.read("xl/sharedStrings.xml").decode("utf-8")
    z.close()

    # --- B1 的 si 索引与 xf ---
    m = re.search(r'<c r="B1"([^>]*)>', sheet)
    if not m:
        print("!! 未找到 B1")
        return
    attr = m.group(1)
    si_idx = int(re.search(r"<v>(\d+)</v>", sheet[m.end():m.end() + 80]).group(1))
    xf_old = int(re.search(r's="(\d+)"', attr).group(1))
    print(f"B1: xf={xf_old}  si={si_idx}")

    if "<mergeCells" in sheet:
        print("!! 已有 mergeCells —— 请先从 backup_legend_layout_20261026 恢复再跑")
        return

    # --- 新 xf（copy + wrapText）---
    cm = re.search(r"<cellXfs[^>]*>(.*?)</cellXfs>", styles, re.S)
    xfs = re.findall(r"<xf\b[^>]*?/>|<xf\b[^>]*?>.*?</xf>", cm.group(1), re.S)
    base = xfs[xf_old]
    new_xf = re.sub(r"<alignment[^>]*/>", "", base)
    new_xf = new_xf.replace("</xf>", '<alignment vertical="top" wrapText="1"/></xf>')
    if "<alignment" not in new_xf:
        print("!! xf 结构异常")
        return
    xf_new = len(xfs)
    print(f"新 xf 索引={xf_new}")
    print(f"   {new_xf}")

    # --- 文本 ---
    sis = list(re.finditer(r"<si>(.*?)</si>", sst, re.S))
    if si_idx >= len(sis):
        print("!! si 索引越界")
        return
    old_plain = "".join(re.findall(r"<t[^>]*>(.*?)</t>", sis[si_idx].group(1), re.S))
    print(f"\n原图例 ({len(old_plain)} 字):\n   {old_plain[:120]}...")
    print(f"新图例 ({len(NEW_TEXT)} 字，{len(LINES)} 行)")
    for i, ln in enumerate(LINES, 1):
        print(f"   {i}| {ln}")

    # --- 组装 ---
    styles2 = (styles[: cm.end(1)] + new_xf + styles[cm.end(1):])
    styles2 = re.sub(r'(<cellXfs count=")(\d+)(")',
                     lambda mm: mm.group(1) + str(int(mm.group(2)) + 1) + mm.group(3),
                     styles2, count=1)

    sst2 = (sst[: sis[si_idx].start(1)]
            + f'<t xml:space="preserve">{NEW_TEXT}</t>'
            + sst[sis[si_idx].end(1):])

    sheet2 = sheet.replace(attr, attr.replace(f's="{xf_old}"', f's="{xf_new}"'), 1)

    # 行 1 高度
    r1 = re.search(r'<row r="1"[^>]*>', sheet2)
    if not r1:
        print("!! 未找到 row 1")
        return
    tag = r1.group(0)
    newtag = re.sub(r'(?<![A-Za-z])ht="[\d.]+"', f'ht="{ROW_HT}"', tag)
    if 'customHeight' not in newtag:
        newtag = newtag.replace(">", ' customHeight="1">')
    sheet2 = sheet2.replace(tag, newtag, 1)

    # 合并区（正确 schema 位置）
    sheet2, anchor = insert_merge(sheet2)
    if sheet2 is None:
        print("!! 找不到 mergeCells 插入锚点")
        return

    print(f"\nrow1: {tag}  →  {newtag}")
    print(f"mergeCells: +{MERGE}  (插在 {anchor} 之前)")

    # --- 落盘前 XML 良构校验（v1 缺这一步）---
    problems = []
    for part, txt in ((sp, sheet2), ("xl/styles.xml", styles2),
                      ("xl/sharedStrings.xml", sst2)):
        try:
            ET.fromstring(txt.encode("utf-8"))
        except Exception as e:
            problems.append(f"{part}: {e}")
    if problems:
        print("\n!! XML 校验失败，放弃落盘：")
        for x in problems:
            print("   " + x)
        # 定位第一个非法字符
        mm = re.search(r"column (\d+)", problems[0])
        if mm and "sharedStrings" in problems[0]:
            c = int(mm.group(1))
            print("   片段: ...%r..." % sst2[max(0, c - 100):c + 60])
        return
    print("XML 良构校验：3 个部件全部通过")

    if not APPLY:
        print("\n（干跑，未落盘。加 --apply 执行）")
        return

    os.makedirs(BAK, exist_ok=True)
    bak = os.path.join(BAK, BOOK)
    if not os.path.exists(bak):
        shutil.copy2(p, bak)

    repl = {sp: sheet2.encode("utf-8"),
            "xl/styles.xml": styles2.encode("utf-8"),
            "xl/sharedStrings.xml": sst2.encode("utf-8")}

    tmp = p + ".tmp_layout"
    zin = zipfile.ZipFile(p)
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            data = repl.get(info.filename)
            if data is None:
                data = zin.read(info.filename)
            zi = zipfile.ZipInfo(info.filename, date_time=info.date_time)
            zi.compress_type = zipfile.ZIP_DEFLATED
            zi.external_attr = info.external_attr
            zout.writestr(zi, data)
    zin.close()

    with zipfile.ZipFile(tmp) as zt:
        bad = zt.testzip()
        if bad:
            os.remove(tmp)
            print(f"!! testzip 失败于 {bad}，已放弃")
            return
    try:
        os.replace(tmp, p)
    except PermissionError:
        with open(tmp, "rb") as f:
            blob = f.read()
        with open(p, "wb") as f:
            f.write(blob)
        os.remove(tmp)
    print(f"\n已落盘 {BOOK}  (备份 {bak})")


if __name__ == "__main__":
    main()
