# -*- coding: utf-8 -*-
"""临时诊断：库存源 sheet 的列结构与值覆盖（只读，不写文件）

分别看「英国可售+在途库存」「欧洲可售+在途库存」：
  行1/行2 表头、每列标签、有多少行、每列非空单元格数
用来定位 F/G/H 公式引用的 Y / AA / Z 列到底有没有数据。
"""
import os, sys, zipfile, re
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

BOOK = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))),
                    "1.每日销售数据表", "欧盟英国库存销售表.xlsx")


def sheet_files(path):
    z = zipfile.ZipFile(path)
    wbxml = z.read("xl/workbook.xml").decode()
    rels = z.read("xl/_rels/workbook.xml.rels").decode()
    rmap = {}
    for tag in re.findall(r"<Relationship\b[^>]*/?>", rels):
        rid = re.search(r'\bId="([^"]+)"', tag)
        tgt = re.search(r'\bTarget="([^"]+)"', tag)
        if rid and tgt:
            rmap[rid.group(1)] = tgt.group(1)
    out = []
    for m in re.finditer(r'<sheet[^>]*name="([^"]+)"[^>]*r:id="([^"]+)"', wbxml):
        t = rmap.get(m.group(2), "")
        t = t.lstrip("/")
        if not t.startswith("xl/"):
            t = "xl/" + t
        out.append((m.group(1), t))
    return z, out


def cells_of_row(x, rn):
    m = re.search(r'<row r="%d"[^>]*>(.*?)</row>' % rn, x, re.S)
    if not m:
        return {}
    out = {}
    for cr, attr, body in re.findall(r'<c r="([A-Z]+%d)"([^>]*)>(.*?)</c>' % rn, m.group(1), re.S):
        col = re.match(r"[A-Z]+", cr).group(0)
        f = re.search(r"<f[^>]*>(.*?)</f>", body, re.S)
        t = re.search(r"<t[^>]*>(.*?)</t>", body, re.S)
        v = re.search(r"<v>(.*?)</v>", body, re.S)
        out[col] = dict(f=(f.group(1) if f else ""), t=(t.group(1) if t else ""), v=(v.group(1) if v else ""))
    return out


def main():
    z, sheets = sheet_files(BOOK)
    for name, tgt in sheets:
        if name not in ("英国可售+在途库存", "欧洲可售+在途库存"):
            continue
        x = z.read(tgt).decode()
        rows = [int(m) for m in re.findall(r'<row r="(\d+)"', x)]
        print("=" * 70)
        print(name, "|", tgt, "| rows:", min(rows) if rows else None, "-", max(rows) if rows else None)
        r1 = cells_of_row(x, 1)
        r2 = cells_of_row(x, 2)
        # 统计每列非空
        cnt = {}
        for rn in rows:
            if rn < 3:
                continue
            for col, c in cells_of_row(x, rn).items():
                if c["v"] or c["t"] or c["f"]:
                    cnt[col] = cnt.get(col, 0) + 1
        # 按列名排序输出
        def key(c):
            n = 0
            for ch in c:
                n = n * 26 + (ord(ch) - 64)
            return n
        for col in sorted(set(list(r1) + list(r2) + list(cnt)), key=key):
            a = r1.get(col, {})
            b = r2.get(col, {})
            lab1 = (a.get("t") or a.get("v") or "")
            lab2 = (b.get("t") or b.get("v") or "")
            print("  %-4s r1=%-28s r2=%-32s nonempty=%s" % (col, lab1[:28], lab2[:32], cnt.get(col, 0)))


if __name__ == "__main__":
    main()
