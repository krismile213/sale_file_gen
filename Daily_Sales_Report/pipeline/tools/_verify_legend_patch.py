# -*- coding: utf-8 -*-
"""图例补丁后完整性核查：图例文本 / CF 规则数 / 公式缓存值 / zip 完整性"""
import os, re, zipfile, glob, sys

HERE = os.path.dirname(os.path.abspath(__file__))
DSR = HERE
for _ in range(6):
    if os.path.isdir(os.path.join(DSR, "1.每日销售数据表")):
        break
    DSR = os.path.dirname(DSR)
WORK = os.path.join(DSR, "1.每日销售数据表")
BAK = os.path.join(DSR, ".workbuddy", "backup_legend_20261026")


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


def shared(z):
    try:
        sx = z.read("xl/sharedStrings.xml").decode("utf-8")
    except KeyError:
        return []
    return re.findall(r"<si>(.*?)</si>", sx, re.S)


def plain(s):
    t = "".join(re.findall(r"<t[^>]*>(.*?)</t>", s, re.S))
    return (t.replace("&lt;", "<").replace("&gt;", ">")
             .replace("&quot;", '"').replace("&amp;", "&"))


def stats(z, xmls):
    """返回 (CF规则数, 公式数, 有缓存值的公式单元格数)"""
    cf = 0
    nf = 0
    cached = 0
    for xml in xmls:
        cf += len(re.findall(r"<cfRule\b", xml))
        for m in re.finditer(r"<c\b[^>]*>(.*?)</c>", xml, re.S):
            body = m.group(1)
            if "<f" in body:
                nf += 1
                if "<v>" in body:
                    cached += 1
    return cf, nf, cached


def main():
    ok = True
    print("=" * 104)
    print("一、17 张分析表 B1 图例（改后）")
    print("=" * 104)
    for p in sorted(glob.glob(os.path.join(WORK, "*.xlsx"))):
        n = os.path.basename(p)
        if n.startswith("~$") or n.endswith(".bak"):
            continue
        z = zipfile.ZipFile(p)
        ss = shared(z)
        for sname, t in sheet_map(z):
            if "每日库存分析表" not in sname:
                continue
            xml = z.read(t).decode("utf-8")
            m = re.search(r'<c r="B1"([^>]*?)(?:/>|>(.*?)</c>)', xml, re.S)
            if not m:
                continue
            attrs, body = m.group(1), (m.group(2) or "")
            if 't="s"' in attrs:
                v = re.search(r"<v>(\d+)</v>", body)
                txt = plain(ss[int(v.group(1))]) if v else "?"
            elif 't="inlineStr"' in attrs:
                txt = plain(body)
            else:
                txt = body[:80]
            bl = re.search(r"蓝色：([^；;]*)", txt)
            mark = "✓" if (not bl or "1/" in bl.group(1)) else "✗"
            if mark == "✗":
                ok = False
            print(f"{mark} {n[:26]:<28} {sname.replace('-每日库存分析表',''):<8} "
                  f"{('蓝色：' + bl.group(1)) if bl else '（无蓝色分句）'}")
        z.close()

    print()
    print("=" * 104)
    print("二、CF 规则数 / 公式数 / 缓存值 —— 当前 vs 备份")
    print("=" * 104)
    print(f"{'工作簿':<30} {'CF当前':>7} {'CF备份':>7} {'公式当前':>9} {'公式备份':>9} {'缓存当前':>9} {'缓存备份':>9}  判定")
    for p in sorted(glob.glob(os.path.join(WORK, "*.xlsx"))):
        n = os.path.basename(p)
        if n.startswith("~$") or n.endswith(".bak"):
            continue
        bak = os.path.join(BAK, n)
        z = zipfile.ZipFile(p)
        xmls = [z.read(t).decode("utf-8") for s, t in sheet_map(z) if "每日库存分析表" in s]
        cur = stats(z, xmls)
        z.close()
        if os.path.exists(bak):
            zb = zipfile.ZipFile(bak)
            xmlsb = [zb.read(t).decode("utf-8") for s, t in sheet_map(zb) if "每日库存分析表" in s]
            old = stats(zb, xmlsb)
            zb.close()
        else:
            old = ("-", "-", "-")
        same = "OK" if cur == old else ("改了(仅图例)" if old[0] == "-" else "!! 不一致")
        if cur != old and old[0] != "-":
            ok = False
        print(f"{n[:30]:<30} {cur[0]:>7} {old[0]:>7} {cur[1]:>9} {old[1]:>9} {cur[2]:>9} {old[2]:>9}  {same}")

    print()
    print("=" * 104)
    print("三、zip 完整性 + 缓存值抽查（欧洲表 AO7/AN7 应为 17.5/19.4）")
    print("=" * 104)
    for p in sorted(glob.glob(os.path.join(WORK, "*.xlsx"))):
        n = os.path.basename(p)
        if n.startswith("~$") or n.endswith(".bak"):
            continue
        z = zipfile.ZipFile(p)
        bad = z.testzip()
        z.close()
        print(f"  {'✓' if bad is None else '✗'} {n[:40]:<42} testzip={'OK' if bad is None else bad}")
        if bad:
            ok = False

    p = os.path.join(WORK, "欧盟英国库存销售表.xlsx")
    z = zipfile.ZipFile(p)
    xml = z.read("xl/worksheets/sheet10.xml").decode("utf-8")
    for rc in ["AO7", "AN7", "AO452", "AN452"]:
        m = re.search(r'<c r="' + rc + r'"([^>]*?)(?:/>|>(.*?)</c>)', xml, re.S)
        body = m.group(2) if m else ""
        v = re.search(r"<v>([^<]*)</v>", body)
        f = re.search(r"<f", body)
        print(f"  {rc}: 公式={'有' if f else '无'}  缓存值={v.group(1) if v else 'NONE'}")
    z.close()

    print()
    print("总判定:", "全部通过" if ok else "有异常，见上")


if __name__ == "__main__":
    main()
