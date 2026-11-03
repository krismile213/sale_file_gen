# -*- coding: utf-8 -*-
"""全量条件格式清单：解析 dxfs + theme 颜色，输出每张分析表的 CF 规则全貌。

用途：核查图例与实际 CF 是否一致（一次性诊断，不修改任何文件）。
"""
import colorsys
import os
import re
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
WORK = os.path.join(os.path.dirname(os.path.dirname(HERE)), "1.每日销售数据表")

# OOXML theme 索引 -> clrScheme 子元素顺序
THEME_ORDER = [1, 0, 3, 2, 4, 5, 6, 7, 8, 9, 10, 11]


def parse_theme(z):
    if "xl/theme/theme1.xml" not in z.namelist():
        return []
    s = z.read("xl/theme/theme1.xml").decode("utf-8", "replace")
    m = re.search(r"<a:clrScheme[^>]*>(.*?)</a:clrScheme>", s, re.S)
    if not m:
        return []
    cols = []
    for blk in re.findall(r"<a:(dk1|lt1|dk2|lt2|accent\d|hlink|folHlink)>(.*?)</a:\1>", m.group(1), re.S):
        rgb = re.search(r'(?:srgbClr val|lastClr)="([0-9A-Fa-f]{6})"', blk[1])
        cols.append(rgb.group(1).upper() if rgb else "000000")
    return cols


def apply_tint(rgb, tint):
    if not tint:
        return rgb
    r, g, b = (int(rgb[i:i + 2], 16) / 255.0 for i in (0, 2, 4))
    h, l, s = colorsys.rgb_to_hls(r, g, b)
    l = l * (1 + tint) if tint < 0 else l * (1 - tint) + tint
    l = max(0.0, min(1.0, l))
    r, g, b = colorsys.hls_to_rgb(h, l, s)
    return "%02X%02X%02X" % (round(r * 255), round(g * 255), round(b * 255))


def color_of(tag, theme):
    """从 <fgColor .../> 或 <bgColor .../> 标签解析出 RRGGBB"""
    if not tag:
        return ""
    rgb = re.search(r'rgb="([0-9A-Fa-f]{8})"', tag)
    if rgb:
        v = rgb.group(1).upper()
        return v[2:] if v.startswith("FF") else v
    th = re.search(r'theme="(\d+)"', tag)
    if th and theme:
        idx = int(th.group(1))
        base = theme[THEME_ORDER[idx]] if idx < len(THEME_ORDER) and THEME_ORDER[idx] < len(theme) else ""
        tint = re.search(r'tint="(-?[\d.]+)"', tag)
        if not base:
            return ""
        return apply_tint(base, float(tint.group(1))) if tint else base
    idx = re.search(r'indexed="(\d+)"', tag)
    if idx:
        return "indexed:%s" % idx.group(1)
    return ""


def parse_dxfs(z, theme):
    if "xl/styles.xml" not in z.namelist():
        return []
    s = z.read("xl/styles.xml").decode("utf-8", "replace")
    m = re.search(r"<dxfs[^>]*>(.*?)</dxfs>", s, re.S)
    if not m:
        return []
    out = []
    for blk in re.findall(r"<dxf>(.*?)</dxf>", m.group(1), re.S):
        fill = re.search(r"<fill>(.*?)</fill>", blk, re.S)
        fg = bg = ""
        if fill:
            f = re.search(r"<fgColor[^>]*/?>", fill.group(1))
            b = re.search(r"<bgColor[^>]*/?>", fill.group(1))
            fg = color_of(f.group(0) if f else "", theme)
            bg = color_of(b.group(0) if b else "", theme)
        out.append((fg, bg))
    return out


def sheet_files(z):
    wb = z.read("xl/workbook.xml").decode("utf-8", "replace")
    rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8", "replace")
    # 属性顺序在不同工作簿里不一致，必须逐标签独立取属性
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
        if not nm or not rid:
            continue
        t = rid2t.get(rid.group(1), "")
        if not t:
            continue
        sp = t if t.startswith("xl/") else "xl/" + t.lstrip("/")
        out.append((nm.group(1), sp.replace("xl/xl/", "xl/")))
    return out


NAMED = {
    "FFFF0000": "纯红", "FFFFC7CE": "浅红(标准bad)", "FFFF8FA8": "深粉",
    "FFFFDDE7": "浅粉", "FFFF9900": "橙", "FFFFFF00": "黄",
    "FFFBB2B2": "淡红", "FF4BD9FF": "浅蓝", "FF0070C0": "深蓝", "FF55D5FF": "天蓝",
}


def main():
    names = sorted(f for f in os.listdir(WORK) if f.endswith(".xlsx") and not f.startswith("~"))
    for fn in names:
        p = os.path.join(WORK, fn)
        try:
            z = zipfile.ZipFile(p)
        except Exception as e:
            print(fn, "OPEN FAIL", e)
            continue
        theme = parse_theme(z)
        dxf = parse_dxfs(z, theme)
        print("=" * 100)
        print("###", fn, f"(dxfs={len(dxf)}, theme={len(theme)})")
        for nm, sp in sheet_files(z):
            if "每日库存分析表" not in nm or sp not in z.namelist():
                continue
            x = z.read(sp).decode("utf-8", "replace")
            print(f"  -- {nm}")
            for sq, body in re.findall(
                r'<conditionalFormatting[^>]*sqref="([^"]*)"[^>]*>(.*?)</conditionalFormatting>', x, re.S
            ):
                for rule in re.findall(r"<cfRule[^>]*?/>|<cfRule[^>]*?>.*?</cfRule>", body, re.S):
                    typ = re.search(r'type="([^"]+)"', rule)
                    op = re.search(r'operator="([^"]+)"', rule)
                    did = re.search(r'dxfId="(\d+)"', rule)
                    pri = re.search(r'priority="(\d+)"', rule)
                    fml = re.findall(r"<formula>(.*?)</formula>", rule, re.S)
                    txt = re.search(r'<text>(.*?)</text>', rule, re.S)
                    fg, bg = ("", "")
                    if did and int(did.group(1)) < len(dxf):
                        fg, bg = dxf[int(did.group(1))]
                    col = bg or fg
                    label = NAMED.get("FF" + col, col)
                    f0 = (fml[0] if fml else "").replace("&quot;", '"').replace("&gt;", ">").replace("&lt;", "<")
                    print(f"     p={pri.group(1) if pri else '?':>3} {typ.group(1) if typ else '?':<18}"
                          f" op={op.group(1) if op else '-':<12} 色={label:<12} sqref={sq[:32]:<32}"
                          + (f" 文本={txt.group(1)}" if txt else ""))
                    if f0:
                        print(f"          f: {f0[:120]}")
        z.close()


if __name__ == "__main__":
    main()
