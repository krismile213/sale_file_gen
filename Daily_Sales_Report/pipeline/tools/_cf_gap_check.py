# -*- coding: utf-8 -*-
"""检查每张分析表的实际数据行 vs 粉色CF范围 vs 红色CF范围"""
import os, re, zipfile, glob
from openpyxl.utils import column_index_from_string

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

def get_sheet_info(path):
    """返回 {sheet_name: dict(max_row, pink_rows, red_rows, gap)}"""
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
        if "每日库存分析表" not in sheetname:
            continue
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

        # sheetData max row (look at the last row tag)
        row_nums = [int(m.group(1)) for m in re.finditer(r'<row r="(\d+)"', xml)]
        max_row = max(row_nums) if row_nums else 0
        # 找最后一个有内容的行 (非空单元格)
        # 简单方法：找最大的行号

        # Extract CF rules
        pink_rows = 0
        red_rows = 0
        for m in re.finditer(r"<conditionalFormatting([^>]*)>(.*?)</conditionalFormatting>", xml, re.S):
            sqref = re.search(r'sqref="([^"]+)"', m.group(1))
            if not sqref:
                continue
            sqref = sqref.group(1)
            for r2 in re.finditer(r"<cfRule([^>]*?)(?:/>|>(.*?)</cfRule>)", m.group(2), re.S):
                a = r2.group(1)
                body = (r2.group(2) or "")
                dx = re.search(r'dxfId="(\d+)"', a)
                color = dxfcol.get(int(dx.group(1))) if dx else None
                f = re.search(r"<formula>(.*?)</formula>", body, re.S)
                fo = f.group(1) if f else ""

                # 只看日列范围上的规则 (sqref包含Y或AI或AS开头)
                if color in ("FFFF8FA8", "FFFFDDE7"):  # pink
                    # 找 sqref 中最大的行号
                    for part in sqref.split():
                        if ":" in part:
                            b = part.split(":")[1].replace("$", "")
                            mm = re.match(r"([A-Z]+)(\d+)", b)
                            if mm:
                                pink_rows = max(pink_rows, int(mm.group(2)))
                if color == "FFFF0000" and "LOOKUP" not in fo and "$" not in fo[:3]:
                    # red_hi on daily columns
                    if re.search(r'>\$?[A-Z]{1,3}\d+\*[\d.]+', fo):
                        for part in sqref.split():
                            if ":" in part:
                                b = part.split(":")[1].replace("$", "")
                                mm = re.match(r"([A-Z]+)(\d+)", b)
                                if mm:
                                    red_rows = max(red_rows, int(mm.group(2)))

        gap = max_row - pink_rows if pink_rows > 0 else -1
        result[sheetname] = dict(
            max_row=max_row,
            pink_rows=pink_rows,
            red_rows=red_rows,
            gap=gap,
            gap_pct=round(gap / max_row * 100, 1) if gap > 0 and max_row > 0 else 0
        )
    return result

def main():
    xlsx_files = sorted(glob.glob(os.path.join(WORK, "*.xlsx")))
    print(f"{'表名':<30} {'实际行':>6} {'粉红行':>6} {'红色行':>6} {'缺口':>6} {'缺口%':>6}")
    print("-" * 70)
    for path in xlsx_files:
        name = os.path.basename(path)
        if name.startswith("~$") or name.startswith(".") or name.endswith(".bak"):
            continue
        try:
            info = get_sheet_info(path)
        except Exception as e:
            print(f"!! {name}: {e}")
            continue
        for sname, d in sorted(info.items()):
            short = sname.replace("-每日库存分析表", "")
            gap_str = f"{d['gap']}" if d['gap'] >= 0 else "N/A"
            pct_str = f"{d['gap_pct']}%" if d['gap_pct'] > 0 else "—"
            print(f"{short:<30} {d['max_row']:>6} {d['pink_rows']:>6} {d['red_rows']:>6} {gap_str:>6} {pct_str:>6}")

if __name__ == "__main__":
    main()
