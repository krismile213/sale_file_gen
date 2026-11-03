# -*- coding: utf-8 -*-
"""CF 复算器 v2：按规则形态显式判定，Excel 语义（空=0，空串是文本）。
用法：
  python _cf_eval.py <工作簿> <sheet>             # 汇总
  python _cf_eval.py <工作簿> <sheet> zero        # 列出值=0 却标色的格
  python _cf_eval.py <工作簿> <sheet> big 20      # 列出值>=20 却没标红的格
"""
import sys, os, re, zipfile
import openpyxl
from openpyxl.utils import get_column_letter, column_index_from_string

HERE = os.path.dirname(os.path.abspath(__file__))
WORK = os.path.join(os.path.dirname(os.path.dirname(HERE)), "1.每日销售数据表")


def _rels_map(rels_xml):
    """rId → Target。属性顺序不固定（openpyxl 写的是 Type/Target/Id），逐标签提取。"""
    out = {}
    for tag in re.findall(r"<Relationship\b[^>]*/?>", rels_xml):
        rid = re.search(r'\bId="([^"]+)"', tag)
        tgt = re.search(r'\bTarget="([^"]+)"', tag)
        if rid and tgt:
            out[rid.group(1)] = tgt.group(1)
    return out


def load_cf(path, sheetname):
    z = zipfile.ZipFile(path)
    wbxml = z.read("xl/workbook.xml").decode("utf-8")
    rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    rid = {m.group(1): m.group(2) for m in
           re.finditer(r'<sheet[^>]*name="([^"]+)"[^>]*r:id="([^"]+)"', wbxml)}
    tgt = _rels_map(rels)
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
                    col = c.group(1)
            dxfcol[i] = col
    out = []
    t = tgt.get(rid.get(sheetname))
    if not t:
        return out
    t = t.lstrip("/")
    if not t.startswith("xl/"):
        t = "xl/" + t
    xml = z.read(t).decode("utf-8")
    for m in re.finditer(r"<conditionalFormatting([^>]*)>(.*?)</conditionalFormatting>", xml, re.S):
        sqref = re.search(r'sqref="([^"]+)"', m.group(1)).group(1)
        for r2 in re.finditer(r"<cfRule([^>]*?)(?:/>|>(.*?)</cfRule>)", m.group(2), re.S):
            a, body = r2.group(1), (r2.group(2) or "")
            prio = re.search(r'priority="(\d+)"', a)
            typ = re.search(r'type="([^"]+)"', a)
            dx = re.search(r'dxfId="(\d+)"', a)
            f = re.search(r"<formula>(.*?)</formula>", body, re.S)
            fo = f.group(1) if f else ""
            fo = (fo.replace("&quot;", '"').replace("&gt;", ">")
                  .replace("&lt;", "<").replace("&amp;", "&"))
            out.append(dict(sqref=sqref, prio=int(prio.group(1)) if prio else 0,
                            type=typ.group(1) if typ else "",
                            dxf=int(dx.group(1)) if dx else None,
                            color=dxfcol.get(int(dx.group(1))) if dx else None,
                            formula=fo))
    return out


def sqref_covers(sqref, col, row):
    for part in sqref.split():
        if ":" in part:
            a, b = [x.replace("$", "") for x in part.split(":")]
            ca = re.match(r"([A-Z]+)(\d*)", a)
            cb = re.match(r"([A-Z]+)(\d*)", b)
            c1, c2 = column_index_from_string(ca.group(1)), column_index_from_string(cb.group(1))
            r1 = int(ca.group(2)) if ca.group(2) else 1
            r2 = int(cb.group(2)) if cb.group(2) else 1048576
            if c1 <= col <= c2 and r1 <= row <= r2:
                return True
        else:
            p = part.replace("$", "")
            mm = re.match(r"([A-Z]+)(\d*)", p)
            if column_index_from_string(mm.group(1)) == col:
                if not mm.group(2) or int(mm.group(2)) == row:
                    return True
    return False


def isnum(x):
    return isinstance(x, (int, float)) and not isinstance(x, bool)


def num0(x):
    """Excel 算术：空=0"""
    return 0 if x is None else x


def classify(formula, dxf, color):
    """把规则归类成可执行形态（倍率/阈值都从公式里抽，不写死）"""
    f = formula.replace(" ", "")
    if "LOOKUP" in f:
        return dict(kind="pink", deep=(color == "FFFF8FA8"))
    if ">=2" in f:                                   # 低基数补偿规则
        th = re.search(r"<=([\d.]+)", f)
        mult = re.search(r"\*([\d.]+)\)?$", f)
        return dict(kind="red_lo", th=float(th.group(1)) if th else 1.0,
                    mult=float(mult.group(1)) if mult else 4.0)
    mhi = re.search(r">\$?[A-Z]{1,3}\d+\*([\d.]+)", f)
    if mhi:                                          # 高销量红：cell > 上月日均 * N
        th = re.search(r">([\d.]+)\)$", f) or re.search(r">([\d.]+)\)", f)
        return dict(kind="red_hi", mult=float(mhi.group(1)),
                    th=float(th.group(1)) if th else 1.0)
    mbl = re.search(r"/\$?[A-Z]{1,3}\d+/([\d.]+)", f) or re.search(r"/([\d.]+)", f)
    if mbl and "<" in f:                             # 低销量蓝
        return dict(kind="blue", mult=float(mbl.group(1)))
    return dict(kind="other")


def main():
    name, sheet = sys.argv[1], sys.argv[2]
    mode = sys.argv[3] if len(sys.argv) > 3 else ""
    big = float(sys.argv[4]) if len(sys.argv) > 4 else 20
    p = os.path.join(WORK, name + ".xlsx")
    wb = openpyxl.load_workbook(p, data_only=True)
    ws = wb[sheet]

    hdr = {c: ws.cell(5, c).value for c in range(1, ws.max_column + 1)
           if ws.cell(5, c).value is not None}
    daycols = sorted(c for c, v in hdr.items() if hasattr(v, "year"))
    ucol = next((c for c, v in hdr.items() if isinstance(v, str) and "上月日均" in v), None)
    fcol = next((c for c, v in hdr.items() if isinstance(v, str) and v.strip() == "可售库存"), None)

    print("=" * 78)
    print(f"{name} / {sheet}  行7-{ws.max_row}")
    print(f"日列 {get_column_letter(daycols[0])}..{get_column_letter(daycols[-1])} "
          f"共{len(daycols)}列  {str(hdr[daycols[0]])[:10]} ~ {str(hdr[daycols[-1]])[:10]}")
    print(f"上月日均={get_column_letter(ucol)}  可售库存={get_column_letter(fcol) if fcol else '?'}")

    rules = [r for r in load_cf(p, sheet) if r["type"] == "expression" and r["formula"]]
    dayr = []
    for r in rules:
        if any(sqref_covers(r["sqref"], c, 7) for c in daycols):
            r = dict(r)
            r.update(classify(r["formula"], r["dxf"], r["color"]))
            dayr.append(r)
    print(f"\n-- 日列上的表达式规则 {len(dayr)} 条 --")
    for r in sorted(dayr, key=lambda x: x["prio"]):
        print(f"  prio={r['prio']:<5} {r['color']} {r['kind']:<8}"
              f"{' th=' + str(r.get('th')) if 'th' in r else ''}  {r['formula'][:95]}")

    # 上月日均分布
    uvals = [ws.cell(r, ucol).value for r in range(7, ws.max_row + 1)]
    num = [v for v in uvals if isnum(v)]
    blk = [v for v in uvals if v is None]
    txt = [v for v in uvals if isinstance(v, str)]
    print(f"\n-- 上月日均 分布 -- 数值{len(num)} 空{len(blk)} 文本{len(txt)}")
    if num:
        s = sorted(num)
        print(f"   min={s[0]} p25={s[len(s)//4]} 中位={s[len(s)//2]} p75={s[3*len(s)//4]} max={s[-1]}")
    if txt:
        from collections import Counter
        print(f"   文本分布 {Counter(txt).most_common(5)}  样例行 "
              f"{[r for r in range(7, ws.max_row+1) if isinstance(ws.cell(r,ucol).value,str)][:8]}")

    # 锚点日（粉色规则里的跨表引用，如 数据透视表!$B$11）
    anchor = None
    for r in dayr:
        m = re.search(r"([^\s!(),<>=]+)!\$?([A-Z]{1,3})\$?(\d+)", r["formula"])
        if m and m.group(1) in wb.sheetnames:
            anchor = wb[m.group(1)].cell(int(m.group(3)), column_index_from_string(m.group(2))).value
            print(f"\n-- 锚点日 = {m.group(1)}!{m.group(2)}{m.group(3)} = {anchor} --")
            break

    # 逐格复算
    from collections import Counter
    stat = Counter()
    zero_hit, big_miss = [], []
    last_day_of_row = {}
    for r in range(7, ws.max_row + 1):
        # 该行最后一个有销量的日列（供粉色规则用）
        lastc = None
        for c in daycols:
            v = ws.cell(r, c).value
            if isnum(v) and v > 0:
                lastc = c
        last_day_of_row[r] = lastc
        total = sum(ws.cell(r, c).value for c in daycols if isnum(ws.cell(r, c).value))
        F = ws.cell(r, fcol).value if fcol else None
        A = ws.cell(r, 1).value
        G = ws.cell(r, 7).value      # 在途
        H = ws.cell(r, 8).value      # 调拨在途合计

        for c in daycols:
            v = ws.cell(r, c).value
            hits = []
            for rule in dayr:
                if not sqref_covers(rule["sqref"], c, r):
                    continue
                if hit(rule, v, ws.cell(r, ucol).value, total, lastc, c, daycols, A, F, G, H, anchor, ws):
                    hits.append(rule)
            red = [h for h in hits if h["color"] == "FFFF0000"]
            stat["red" if red else ("other" if hits else "none")] += 1

            zero = (v is None or (isnum(v) and v == 0))
            if zero and hits and len(zero_hit) < 30:
                zero_hit.append((r, get_column_letter(c), v,
                                 [(h["prio"], h["color"], h["kind"]) for h in hits]))
            if isnum(v) and v >= big and not red and len(big_miss) < 30:
                big_miss.append((r, get_column_letter(c), v, ws.cell(r, ucol).value,
                                 [(h["prio"], h["color"], h["kind"]) for h in hits]))

    print(f"\n-- 复算结果 --  红={stat['red']}  其它色={stat['other']}  无色={stat['none']}")

    if mode in ("", "zero"):
        print(f"\n-- 值=0/空 但被标色（{len(zero_hit)} 例，最多30）--")
        for s in zero_hit:
            print(f"  行{s[0]:<4} {s[1]:<3} 值={s[2]!r:<5} 命中={s[3]}")
        if not zero_hit:
            print("  （无）")
    if mode in ("", "big"):
        print(f"\n-- 值>={big:g} 但没标红（{len(big_miss)} 例，最多30）--")
        for s in big_miss:
            print(f"  行{s[0]:<4} {s[1]:<3} 值={s[2]!r:<5} 上月日均={s[3]!r:<8} 命中={s[4]}")
        if not big_miss:
            print("  （无）")


def hit(rule, v, u, total, lastc, c, daycols, A, F, G, H, anchor=None, ws=None):
    k = rule["kind"]
    try:
        if k == "red_hi":
            th, mult = rule["th"], rule["mult"]
            if not isnum(u):
                # 空 → 0：0>th 假；空串是文本：''*N = #VALUE! → 整条规则静默失效
                return False
            return (num0(v) > u * mult) and (u > th)
        if k == "red_lo":
            th, mult = rule["th"], rule["mult"]
            if not isnum(u):
                return False          # ''>0 为 TRUE 但 ''<=th 为 FALSE
            return (u > 0) and (u <= th) and (num0(v) >= 2) and (num0(v) > u * mult)
        if k == "blue":
            if not isnum(u):
                return False
            return (num0(v) < u / rule["mult"]) and (num0(v) > 0)
        if k == "pink":
            if A is None or A == "":
                return False
            if not (isnum(F) and F <= 0):
                return False
            if not (total > 0):
                return False
            if lastc is None:
                return False
            # 断货持续 >=3 天：最后一个销量日距今 >=3
            days = daycols.index(c) - daycols.index(lastc)
            if days < 3:
                return False
            if c <= lastc:
                return False          # 只标断货之后的格
            # 上界：该日列日期 < 锚点日（锚点当天及以后不标）
            if anchor is not None:
                head = ws.cell(5, c).value
                if hasattr(head, "date"):
                    head = head.date()
                elif hasattr(head, "year"):
                    head = head
                anc = anchor.date() if hasattr(anchor, "date") else anchor
                try:
                    if head >= anc:
                        return False
                except TypeError:
                    return False
            if rule["deep"]:
                return (num0(G) == 0) and (num0(H) == 0)
            return (num0(G) > 0) or (num0(H) > 0)
    except Exception:
        return False
    return False


if __name__ == "__main__":
    main()
