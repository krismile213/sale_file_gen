# -*- coding: utf-8 -*-
"""17 张分析表图例统一：退回欧洲表红口径 + 全表排版 + 数字与公式对齐。

做三件事：
  ① 欧洲表红口径退回原两条相对倍数规则（从 backup_red_eu_20261026 整块还原）
  ② 17 张表 B1 图例改成多行紧凑文本，红色/蓝色倍数**从各自 CF 公式反解**后写入
     （不靠记忆，避免图例与公式不符）
  ③ 排版：合并 B1:H1 + wrapText + 顶对齐 + 行高 92

用法：
  python patch_legend_all_20261026.py            # 干跑
  python patch_legend_all_20261026.py --apply    # 落盘
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
BAK = os.path.join(DSR, ".workbuddy", "backup_legend_all_20261026")
BAK_RED = os.path.join(DSR, ".workbuddy", "backup_red_eu_20261026")

APPLY = "--apply" in sys.argv
EU_BOOK = "欧盟英国库存销售表.xlsx"
EU_SHEET = "欧洲-每日库存分析表"
EU_DAY_SQ = "AS7:CE1476"
ROW_HT = 92
MERGE = "B1:H1"


# ────────────────────────── 基础工具 ──────────────────────────
def sheet_map(z):
    wb = z.read("xl/workbook.xml").decode("utf-8", "replace")
    rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8", "replace")
    rid2t = {}
    for tag in re.findall(r"<Relationship\b[^>]*/?>", rels):
        rid = re.search(r'\bId="([^"]+)"', tag)
        tgt = re.search(r'\bTarget="([^"]+)"', tag)
        if rid and tgt:
            rid2t[rid.group(1)] = tgt.group(1)
    out = []
    for tag in re.findall(r"<sheet\b[^>]*/?>", wb):
        nm = re.search(r'\bname="([^"]+)"', tag)
        rid = re.search(r'\br:id="([^"]+)"', tag)
        if not (nm and rid):
            continue
        t = rid2t.get(rid.group(1), "")
        sp = t if t.startswith("xl/") else "xl/" + t.lstrip("/")
        out.append((nm.group(1), sp.replace("xl/xl/", "xl/")))
    return out


def xml_escape(s):
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def _unesc(s):
    """还原 XML 实体（&#10; → 换行）。顺序：&amp; 必须最后。"""
    return (s.replace("&#10;", "\n").replace("&#13;", "\r")
             .replace("&lt;", "<").replace("&gt;", ">")
             .replace("&quot;", '"').replace("&apos;", "'")
             .replace("&amp;", "&"))


def sst_texts(sst):
    return [_unesc(re.sub(r"<[^>]+>", "", m.group(1)))
            for m in re.finditer(r"<si>(.*?)</si>", sst, re.S)]


def b1_info(xml, sst=None):
    """返回 (完整 <c> 标签, 属性串, si 索引或 None, 当前文本)。

    两种存储形式都要支持：
      · t="s"           → <v>索引</v>，文本在 sharedStrings.xml
      · t="inlineStr"   → <is><t>文本</t></is>（12 本簿都是这种）
    """
    m = re.search(r'<c r="B1"([^>]*?)(?:/>|>(.*?)</c>)', xml, re.S)
    if not m:
        return None
    attrs, body = m.group(1), m.group(2) or ""
    si = None
    v = re.search(r"<v>(\d+)</v>", body)
    if v:
        si = int(v.group(1))
    if 't="inlineStr"' in attrs or "<is>" in body:
        txt = "".join(re.findall(r"<t[^>]*>(.*?)</t>", body, re.S))
        txt = (txt.replace("&#10;", "\n").replace("&gt;", ">")
                  .replace("&lt;", "<").replace("&amp;", "&"))
        return m.group(0), attrs, None, txt
    txt = ""
    if si is not None and sst:
        texts = sst_texts(sst)
        txt = texts[si] if si < len(texts) else ""
    return m.group(0), attrs, si, txt


def set_b1_text(xml, text):
    """把 B1 写成 inlineStr。返回新 xml。"""
    m = re.search(r'<c r="B1"([^>]*?)(?:/>|>(.*?)</c>)', xml, re.S)
    if not m:
        return xml
    attrs = m.group(1)
    attrs = re.sub(r'\s*t="[^"]*"', "", attrs) + ' t="inlineStr"'
    esc = xml_escape(text).replace("\n", "&#10;")
    new = '<c r="B1"%s><is><t xml:space="preserve">%s</t></is></c>' % (attrs, esc)
    return xml[:m.start()] + new + xml[m.end():]


# ────────────────────────── CF 反解 ──────────────────────────
def cf_rules(xml):
    """[(sqref, dxfId, priority, formula)]"""
    out = []
    for cf in re.finditer(
            r'<conditionalFormatting sqref="([^"]+)"[^>]*>(.*?)</conditionalFormatting>',
            xml, re.S):
        sq = cf.group(1)
        for r_ in re.finditer(r"<cfRule([^>]*)>(.*?)</cfRule>", cf.group(2), re.S):
            attrs, inner = r_.group(1), r_.group(2)
            f = re.search(r"<formula>(.*?)</formula>", inner, re.S)
            if not f:
                continue
            dx = re.search(r'dxfId="(\d+)"', attrs)
            pri = re.search(r'priority="(\d+)"', attrs)
            out.append((sq, dx.group(1) if dx else None,
                        int(pri.group(1)) if pri else 0, f.group(1)))
    return out


def span_cols(sq):
    """sqref 的列跨度（只取第一个区域）。"""
    a = sq.split()[0]
    m = re.match(r"\$?([A-Z]+)\$?\d+(?::\$?([A-Z]+)\$?\d+)?", a)
    if not m:
        return 1
    if not m.group(2):
        return 1

    def cn(x):
        n = 0
        for ch in x:
            n = n * 26 + ord(ch) - 64
        return n
    return cn(m.group(2)) - cn(m.group(1)) + 1


def parse_params(xml):
    """从 CF 公式反解：(红倍率, 红阈值, 蓝除数 或 None, 实时速度列有蓝?)"""
    mult = thr = div = None
    speed_blue = False
    for sq, dx, pri, fo in cf_rules(xml):
        wide = span_cols(sq) > 20            # 日列（Y..BK / AI..BU / AS..CE）
        # 红·高销量：AND(X7>$Y7*N,$Y7>TH)
        m = re.search(r"\*([\d.]+),\$?[A-Z]+7&gt;([\d.]+)\)$", fo)
        if wide and m and "&lt;" not in fo:
            mult, thr = float(m.group(1)), float(m.group(2))
        # 蓝·日列：AND(X7<$Y7/D,X7>0)
        m = re.search(r"&lt;\$?[A-Z]+7/([\d.]+)", fo)
        if wide and m:
            div = float(m.group(1))
        # 实时速度列蓝：AND($S7<$U7/2,$U7>2)
        if not wide and "/2," in fo.replace("/2.0,", "/2,") and "&gt;2)" in fo:
            speed_blue = True
    return mult, thr, div, speed_blue


def legend_lines(mult, thr, div, speed_blue):
    th = ("%g" % thr)
    lines = [
        "粉红=缺货(深粉:无补给　浅粉:补给在途)　橙=在途未上架",
        "黄=保底　浅红=SKU重复或单日>20　淡红=命中透视表",
        "浅绿=保底+机会产品且月内最高>4",
        "红=日销>日均%s倍且日均>%s；日均≤%s时≥2件且>4倍" % ("%g" % mult, th, th),
    ]
    if div:
        lines.append("蓝=日销<日均÷%g" % div)
    if speed_blue:
        lines.append("实时速度列：红>日均1.5倍，蓝<日均1/2(日均>2)")
    else:
        lines.append("实时速度列：红>日均1.5倍")
    return lines


# ────────────────────────── 样式 / 合并 / 行高 ──────────────────────────
def ensure_wrap_xf(styles, base_xf):
    """返回 (新 styles, 可用的 xf 索引)。若 base 已带 wrapText 则直接复用。"""
    cm = re.search(r"<cellXfs[^>]*>(.*?)</cellXfs>", styles, re.S)
    xfs = re.findall(r"<xf\b[^>]*?/>|<xf\b[^>]*?>.*?</xf>", cm.group(1), re.S)
    if 'wrapText="1"' in xfs[base_xf]:
        return styles, base_xf
    # 已有同款（numFmt/font/fill/border 相同 + wrapText）就复用
    base = xfs[base_xf]
    key = re.sub(r"\s", "", base.split(">")[0])
    for i, x in enumerate(xfs):
        if 'wrapText="1"' in x and re.sub(r"\s", "", x.split(">")[0]) == key:
            return styles, i
    new = re.sub(r"<alignment[^>]*/>", "", base)
    if new.endswith("/>"):
        new = new[:-2] + '><alignment vertical="top" wrapText="1"/></xf>'
    else:
        new = new.replace("</xf>", '<alignment vertical="top" wrapText="1"/></xf>')
    styles2 = styles[: cm.end(1)] + new + styles[cm.end(1):]
    styles2 = re.sub(r'(<cellXfs count=")(\d+)(")',
                     lambda m: m.group(1) + str(int(m.group(2)) + 1) + m.group(3),
                     styles2, count=1)
    return styles2, len(xfs)


def ensure_merge(xml):
    """确保 B1:H1 合并，插在 schema 正确位置。返回 (新 xml, 是否改动)。"""
    mm = re.search(r"<mergeCells([^>]*)>(.*?)</mergeCells>", xml, re.S)
    if mm:
        if 'ref="%s"' % MERGE in mm.group(2):
            return xml, False
        refs = re.findall(r'<mergeCell ref="([^"]+)"', mm.group(2))
        body = mm.group(2) + '<mergeCell ref="%s"/>' % MERGE
        new = '<mergeCells count="%d">%s</mergeCells>' % (len(refs) + 1, body)
        return xml[:mm.start()] + new + xml[mm.end():], True
    block = '<mergeCells count="1"><mergeCell ref="%s"/></mergeCells>' % MERGE
    for anchor in ("<phoneticPr", "<conditionalFormatting", "<dataValidations",
                   "<hyperlinks", "<printOptions", "<pageMargins",
                   "<pageSetup", "<headerFooter", "<drawing", "<legacyDrawing"):
        i = xml.find(anchor)
        if i != -1:
            return xml[:i] + block + xml[i:], True
    return None, False


def ensure_row_ht(xml, ht=ROW_HT):
    r1 = re.search(r'<row r="1"[^>]*>', xml)
    if not r1:
        return xml, False
    tag = r1.group(0)
    new = re.sub(r'(?<![A-Za-z])ht="[\d.]+"', 'ht="%d"' % ht, tag)
    if "customHeight" not in new:
        new = new.replace(">", ' customHeight="1">')
    if new == tag:
        return xml, False
    return xml.replace(tag, new, 1), True


# ────────────────────────── 主流程 ──────────────────────────
def revert_eu_red(book_path):
    """把欧洲表 AS7:CE1476 整块还原成备份里的原样。"""
    bp = os.path.join(BAK_RED, EU_BOOK)
    if not os.path.exists(bp):
        print("  !! 找不到红口径备份，跳过还原:", bp)
        return None
    bak_xml = zipfile.ZipFile(bp).read("xl/worksheets/sheet10.xml").decode("utf-8")
    m = re.search(r'<conditionalFormatting sqref="%s"[^>]*>.*?</conditionalFormatting>'
                  % EU_DAY_SQ, bak_xml, re.S)
    if not m:
        print("  !! 备份里找不到 %s 块" % EU_DAY_SQ)
        return None
    orig_block = m.group(0)
    return orig_block


def patch_book(p, revert_block=None):
    fname = os.path.basename(p)
    zin = zipfile.ZipFile(p)
    smap = sheet_map(zin)
    parts = {n: zin.read(n).decode("utf-8") for n in
             ("xl/styles.xml", "xl/sharedStrings.xml") if n in zin.namelist()}
    targets = [(sn, sp) for sn, sp in smap if "分析表" in sn]
    if not targets:
        zin.close()
        return 0

    changed = {}
    report = []

    for sn, sp in targets:
        xml = zin.read(sp).decode("utf-8")
        orig = xml

        # ① 欧洲表红口径还原
        if fname == EU_BOOK and sn == EU_SHEET and revert_block:
            cur = re.search(r'<conditionalFormatting sqref="%s"[^>]*>.*?</conditionalFormatting>'
                            % EU_DAY_SQ, xml, re.S)
            if cur:
                xml = xml[:cur.start()] + revert_block + xml[cur.end():]

        # ② 图例文本
        bi = b1_info(xml, parts.get("xl/sharedStrings.xml"))
        if not bi:
            report.append((sn, "无 B1，跳过", 0))
            continue
        full, attrs, si, cur_txt = bi
        mult, thr, div, speed_blue = parse_params(xml)
        if mult is None or thr is None:
            report.append((sn, "反解失败 mult=%s thr=%s" % (mult, thr), 0))
            continue
        lines = legend_lines(mult, thr, div, speed_blue)
        text = "\n".join(lines)

        if cur_txt != text:
            if si is not None and "xl/sharedStrings.xml" in parts:
                sst = parts["xl/sharedStrings.xml"]
                n_si = len(re.findall(r"<si>", sst))
                sst = sst.replace(
                    "</sst>",
                    '<si><t xml:space="preserve">%s</t></si></sst>'
                    % xml_escape(text).replace("\n", "&#10;"), 1)
                sst = re.sub(r'(<sst[^>]*uniqueCount=")(\d+)(")',
                             lambda m: m.group(1) + str(n_si + 1) + m.group(3),
                             sst, count=1)
                parts["xl/sharedStrings.xml"] = sst
                m2 = re.search(r'<c r="B1"([^>]*?)(?:/>|>(.*?)</c>)', xml, re.S)
                a2, b2 = m2.group(1), m2.group(2) or ""
                b2 = re.sub(r"<v>\d+</v>", "<v>%d</v>" % n_si, b2)
                if "<v>" not in b2:
                    b2 = "<v>%d</v>" % n_si
                xml = xml[:m2.start()] + '<c r="B1"%s>%s</c>' % (a2, b2) + xml[m2.end():]
                if 't="s"' not in a2:
                    xml = xml.replace('<c r="B1"%s>' % a2,
                                      '<c r="B1"%s t="s">' % a2, 1)
            else:
                xml = set_b1_text(xml, text)

        # ③ 排版：wrapText
        bi = b1_info(xml)
        attrs = bi[1]
        sm = re.search(r's="(\d+)"', attrs)
        if sm:
            base_xf = int(sm.group(1))
            styles2, xf_use = ensure_wrap_xf(parts["xl/styles.xml"], base_xf)
            parts["xl/styles.xml"] = styles2
            if xf_use != base_xf:
                xml = xml.replace('<c r="B1"%s' % attrs,
                                  '<c r="B1"%s' % attrs.replace(
                                      's="%d"' % base_xf, 's="%d"' % xf_use), 1)

        # ④ 合并 + 行高
        xml, _m = ensure_merge(xml)
        if xml is None:
            report.append((sn, "mergeCells 无锚点", 0))
            continue
        xml, _h = ensure_row_ht(xml)

        # ⑤ 良构
        try:
            ET.fromstring(xml.encode("utf-8"))
        except Exception as e:
            report.append((sn, "XML 不良构: %s" % e, 0))
            continue
        if xml != orig:
            changed[sp] = xml.encode("utf-8")
        report.append((sn, "mult=%g thr=%g div=%s 蓝速度=%s %d行"
                       % (mult, thr, div, speed_blue, len(lines)), len(text)))
    zin.close()

    for part, txt in parts.items():
        try:
            ET.fromstring(txt.encode("utf-8"))
        except Exception as e:
            print("  !! %s XML 不良构: %s" % (part, e))
            return -1

    print("=" * 74)
    print(fname)
    for sn, msg, n in report:
        print("   %-26s %s" % (sn, msg))
    if not changed:
        print("   （无改动）")
        return 0
    if not APPLY:
        print("   → 将改动 %d 个 sheet part（干跑）" % len(changed))
        return len(changed)

    os.makedirs(BAK, exist_ok=True)
    bak = os.path.join(BAK, fname)
    if not os.path.exists(bak):
        shutil.copy2(p, bak)

    repl = dict(changed)
    repl.update({k: v.encode("utf-8") for k, v in parts.items()})

    tmp = p + ".tmp_legend"
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
            print("   !! testzip 失败于 %s，放弃" % bad)
            return -1
    try:
        os.replace(tmp, p)
    except PermissionError:
        with open(tmp, "rb") as f:
            blob = f.read()
        with open(p, "wb") as f:
            f.write(blob)
        os.remove(tmp)
    print("   → 已落盘（备份 %s）" % bak)
    return len(changed)


def main():
    files = sorted(f for f in os.listdir(WORK) if f.endswith(".xlsx"))
    total = 0
    for f in files:
        p = os.path.join(WORK, f)
        rb = None
        if f == EU_BOOK:
            rb = revert_eu_red(p)
        n = patch_book(p, rb)
        if n < 0:
            print("!! %s 失败" % f)
            return 1
        total += n
    print()
    print("合计改动 sheet part: %d%s" % (total, "" if APPLY else "（干跑）"))
    return 0


if __name__ == "__main__":
    sys.exit(main())
