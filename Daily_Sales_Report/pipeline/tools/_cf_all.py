# -*- coding: utf-8 -*-
"""提取全部13本工作簿×17张分析表的CF规则，纯XML读取，不加载openpyxl。"""
import os, re, zipfile, glob
from collections import Counter

HERE = os.path.dirname(os.path.abspath(__file__))
WORK = os.path.join(os.path.dirname(os.path.dirname(HERE)), "1.每日销售数据表")

def _rels_map(rels_xml):
    out = {}
    for tag in re.findall(r"<Relationship\b[^>]*/?>", rels_xml):
        rid = re.search(r'\bId="([^"]+)"', tag)
        tgt = re.search(r'\bTarget="([^"]+)"', tag)
        if rid and tgt:
            out[rid.group(1)] = tgt.group(1)
    return out

def extract_cf(path):
    """返回 {sheet_name: [rule_dict, ...]}"""
    z = zipfile.ZipFile(path)
    wbxml = z.read("xl/workbook.xml").decode("utf-8")
    rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    rid_map = {}
    for m in re.finditer(r'<sheet[^>]*name="([^"]+)"[^>]*r:id="([^"]+)"', wbxml):
        rid_map[m.group(2)] = m.group(1)
    tgt = _rels_map(rels)

    # dxfId → color
    dxfcol = {}
    st = z.read("xl/styles.xml").decode("utf-8")
    m = re.search(r"<dxfs[^>]*>(.*?)</dxfs>", st, re.S)
    if m:
        for i, d in enumerate(re.findall(r"<dxf>(.*?)</dxf>", m.group(1), re.S)):
            fill = re.search(r"<fill>.*?</fill>", d, re.S)
            col = None
            if fill:
                c = re.search(r'(?:fgColor|bgColor)[^>]*rgb="([0-9A-Fa-f]{8})"', fill.group(0))
                if c:
                    col = c.group(1).upper()
            dxfcol[i] = col

    result = {}
    for rid, sheetname in rid_map.items():
        t = tgt.get(rid)
        if not t:
            continue
        t = t.lstrip("/")
        if not t.startswith("xl/"):
            t = "xl/" + t
        try:
            xml = z.read(t).decode("utf-8")
        except KeyError:
            continue
        rules = []
        for m in re.finditer(r"<conditionalFormatting([^>]*)>(.*?)</conditionalFormatting>", xml, re.S):
            sqref = re.search(r'sqref="([^"]+)"', m.group(1))
            if not sqref:
                continue
            sqref = sqref.group(1)
            for r2 in re.finditer(r"<cfRule([^>]*?)(?:/>|>(.*?)</cfRule>)", m.group(2), re.S):
                a, body = r2.group(1), (r2.group(2) or "")
                prio = re.search(r'priority="(\d+)"', a)
                typ = re.search(r'type="([^"]+)"', a)
                dx = re.search(r'dxfId="(\d+)"', a)
                f = re.search(r"<formula>(.*?)</formula>", body, re.S)
                fo = f.group(1) if f else ""
                fo = (fo.replace("&quot;", '"').replace("&gt;", ">")
                      .replace("&lt;", "<").replace("&amp;", "&"))
                rules.append(dict(
                    sqref=sqref,
                    prio=int(prio.group(1)) if prio else 0,
                    type=typ.group(1) if typ else "",
                    dxfId=int(dx.group(1)) if dx else None,
                    color=dxfcol.get(int(dx.group(1))) if dx else None,
                    formula=fo
                ))
        if rules:
            result[sheetname] = rules
    return result

def classify(f):
    f2 = f.replace(" ", "")
    if "LOOKUP" in f2 or "NOT(ISERROR" in f2:
        return "pink"
    if ">=2" in f2:
        return "red_lo"
    mhi = re.search(r">\$?[A-Z]{1,3}\d+\*([\d.]+)", f2)
    if mhi:
        return f"red_hi({mhi.group(1)})"
    mbl = re.search(r"/\$?[A-Z]{1,3}\d+/([\d.]+)", f2) or re.search(r"/([\d.]+)", f2)
    if mbl and "<" in f2:
        return f"blue({mbl.group(1)})"
    return "other"

def main():
    xlsx_files = sorted(glob.glob(os.path.join(WORK, "*.xlsx")))
    print(f"找到 {len(xlsx_files)} 本工作簿\n")
    
    for path in xlsx_files:
        name = os.path.basename(path)
        # 跳过临时文件
        if name.startswith("~$") or name.startswith("."):
            continue
        try:
            sheets = extract_cf(path)
        except Exception as e:
            print(f"!! {name}: {e}")
            continue
        
        # 只看分析表
        analysis_sheets = {k: v for k, v in sheets.items() if "每日库存分析表" in k}
        if not analysis_sheets:
            continue
        
        print("=" * 90)
        print(f"📄 {name}")
        print("=" * 90)
        for sname, rules in sorted(analysis_sheets.items()):
            # 过滤出表达式规则（染色用）
            expr_rules = [r for r in rules if r["type"] == "expression" and r["formula"]]
            other_rules = [r for r in rules if r not in expr_rules]
            print(f"\n  ─ {sname} ({len(rules)} 条规则, {len(expr_rules)} 表达式)")
            for r in sorted(expr_rules, key=lambda x: x["prio"]):
                kind = classify(r["formula"])
                color = r["color"] or "—"
                # 截取 sqref 的前 40 字符
                sq = r["sqref"][:50] + ("..." if len(r["sqref"]) > 50 else "")
                print(f"    prio={r['prio']:<4} color={color:<10} kind={kind:<12} sqref={sq}")
                print(f"           formula: {r['formula'][:120]}")
            if other_rules:
                print(f"    (+ {len(other_rules)} 非表达式规则: {[r['type'] for r in other_rules]})")

if __name__ == "__main__":
    main()
