# -*- coding: utf-8 -*-
"""B1 图例普查：列出每张分析表的 B1 原文、存储形态（shared/inline）、字符串索引、引用次数。
用法：python _legend_survey.py
"""
import os, re, zipfile, glob
from collections import Counter

HERE = os.path.dirname(os.path.abspath(__file__))
WORK = os.path.join(os.path.dirname(os.path.dirname(HERE)), "1.每日销售数据表")


def sheet_map(z):
    wb = z.read("xl/workbook.xml").decode("utf-8")
    rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    rid2t = {}
    for m in re.finditer(r"<Relationship\b[^>]*/?>", rels):
        a = dict(re.findall(r'([\w:]+)="([^"]*)"', m.group(0)))
        rid2t[a.get("Id", "")] = a.get("Target", "")
    out = []
    for m in re.finditer(r"<sheet\b[^>]*/?>", wb):
        a = dict(re.findall(r'([\w:]+)="([^"]*)"', m.group(0)))
        t = rid2t.get(a.get("r:id", "") or a.get("id", ""), "").lstrip("/")
        if not t.startswith("xl/"):
            t = "xl/" + t
        out.append((a.get("name", ""), t))
    return out


def shared_strings(z):
    try:
        sx = z.read("xl/sharedStrings.xml").decode("utf-8")
    except KeyError:
        return [], ""
    items = re.findall(r"<si>(.*?)</si>", sx, re.S)
    return items, sx


def plain(si_body):
    txt = "".join(re.findall(r"<t[^>]*>(.*?)</t>", si_body, re.S))
    return (txt.replace("&lt;", "<").replace("&gt;", ">")
               .replace("&quot;", '"').replace("&apos;", "'").replace("&amp;", "&"))


def main():
    files = sorted(glob.glob(os.path.join(WORK, "*.xlsx")))
    for p in files:
        n = os.path.basename(p)
        if n.startswith("~$") or n.endswith(".bak"):
            continue
        z = zipfile.ZipFile(p)
        ss, ssxml = shared_strings(z)
        # 统计每个 sharedString 被引用次数
        ref = Counter()
        sheets = sheet_map(z)
        sheet_xml = {}
        for sname, t in sheets:
            try:
                xml = z.read(t).decode("utf-8")
            except KeyError:
                continue
            sheet_xml[sname] = xml
            for m in re.finditer(r'<c r="[A-Z]+\d+"([^>]*?)(?:/>|>)', xml):
                a = m.group(1)
                if 't="s"' in a:
                    v = re.search(r'<v>(\d+)</v>', xml[m.end():m.end()+120])
                    if v:
                        ref[int(v.group(1))] += 1
        # 只统计 B1
        for sname, xml in sheet_xml.items():
            if "每日库存分析表" not in sname:
                continue
            m = re.search(r'<c r="B1"([^>]*?)(?:/>|>(.*?)</c>)', xml, re.S)
            if not m:
                print(f"[{n}] {sname}: 无 B1")
                continue
            attrs, body = m.group(1), (m.group(2) or "")
            typ = "s" if 't="s"' in attrs else ("inlineStr" if 't="inlineStr"' in attrs else "other")
            idx = None
            if typ == "s":
                v = re.search(r"<v>(\d+)</v>", body)
                if v:
                    idx = int(v.group(1))
                    text = plain(ss[idx]) if idx < len(ss) else "??"
                else:
                    text = "??"
            elif typ == "inlineStr":
                text = plain(body)
            else:
                text = body[:200]
            print("=" * 100)
            print(f"[{n}]  sheet={sname}")
            print(f"  存储={typ}  sharedIndex={idx}  该串被引用={ref.get(idx, 0) if idx is not None else '-'} 次")
            print(f"  原文: {text}")
            # 找蓝色分句
            bl = re.search(r"蓝色[：:]\s*([^；;]*)", text)
            if bl:
                print(f"  蓝色分句: {bl.group(0)}")
            else:
                print("  蓝色分句: （无）")
        z.close()


if __name__ == "__main__":
    main()
