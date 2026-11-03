# -*- coding: utf-8 -*-
"""一次性修复 7 张「库存记录」表：
  ① 在 AI(35) 处插 10 列，把「在售块」从 10/02 补到 10/12（右侧「在途块」整体右移 10）
  ② 清掉两块里"无日期条件"的假公式列（在售 SUMIF / 在途 VLOOKUP）
  ③ 用 10/06、10/10、10/12 三份报告写入**固化字面值**（无报告日留空）
纯 XML 手术。用法： --dry / --apply
"""
# --- pipeline 路径引导（2026-10-14 目录收敛自动注入）---
import os as _os, sys as _sys
_bp = _os.path.dirname(_os.path.abspath(__file__))
for _i in range(6):
    if _os.path.isdir(_os.path.join(_bp, "1.每日销售数据表")):
        break
    _bp = _os.path.dirname(_bp)
DSR = _bp
_bcore = _os.path.join(DSR, "pipeline", "core")
if _bcore not in _sys.path:
    _sys.path.insert(0, _bcore)
# --- 引导结束 ---
import os, re, sys, zipfile, shutil, datetime

# [normalize] DSR 由上方引导块赋值
WORK = os.path.join(DSR, "1.每日销售数据表")
TRUNK = r"C:\Users\dell\Desktop\work-main\sale_file_gen\trunk\data\数据归属周期"
pass  # [normalize] 路径已由引导块设置

INSERT_AT = 35          # AI
NCOL = 10               # 10-03 .. 10-12
NEW_DATES = [datetime.date(2026, 10, 3) + datetime.timedelta(days=i) for i in range(NCOL)]
FILL_DATES = [datetime.date(2026, 10, 6), datetime.date(2026, 10, 10), datetime.date(2026, 10, 12)]
EPOCH = datetime.date(1899, 12, 30)
SNAP0606 = os.path.join(TRUNK, "202609", "快照日期", "20261006")
SNAP0610 = os.path.join(TRUNK, "202610", "快照日期", "20261010")
SNAP0612 = os.path.join(TRUNK, "202610", "快照日期", "20261010", "20261012")
DAYDIR = {datetime.date(2026, 10, 6): (SNAP0606, "20261006", "1006"),
          datetime.date(2026, 10, 10): (SNAP0610, "20261010", "1010"),
          datetime.date(2026, 10, 12): (SNAP0612, "20261012", "1012")}

SHEETCFG = [
    ("欧盟英国库存销售表.xlsx", "英国库存记录 ", "亚马逊英国管理库存报告{d}.csv", "inv",
     "亚马逊英国管理库存报告{d}.csv", "inv"),
    ("欧盟英国库存销售表.xlsx", "欧洲库存记录", "欧洲库存{d4}.csv", "age",
     "欧洲库存{d4}.csv", "age"),
    ("北美加拿大库存销售表.xlsx", "美国库存记录", "亚马逊美国管理库存报告{d}.csv", "inv",
     "亚马逊美国管理库存报告{d}.csv", "inv"),
    ("北美加拿大库存销售表.xlsx", "加拿大库存记录", "亚马逊加拿大管理库存报告{d}.csv", "inv",
     "亚马逊加拿大管理库存报告{d}.csv", "inv"),
    ("北美加拿大库存销售表.xlsx", "墨西哥库存记录", "亚马逊墨西哥管理库存报告{d}.csv", "balance",
     "亚马逊墨西哥管理库存报告{d}.csv", "inv"),
    ("Interflowd北美加拿大库存销售表.xlsx", "美国库存记录", "interflow美国管理库存报告{d}.csv", "inv",
     "interflow美国管理库存报告{d}.csv", "inv"),
    ("Interflowd北美加拿大库存销售表.xlsx", "加拿大库存记录", "interflow加拿大管理库存报告{d}.csv", "inv",
     "interflow加拿大管理库存报告{d}.csv", "inv"),
]

L = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
def colname(n):
    s = ""
    while n:
        n, k = divmod(n - 1, 26)
        s = L[k] + s
    return s
def colnum(c):
    n = 0
    for ch in c:
        n = n * 26 + (ord(ch) - 64)
    return n

QUAL = r"(?:'(?P<q1>[^']+)'|(?P<q2>[A-Za-z0-9_\u4e00-\u9fff\.~]+))!"
RANGE_RE = re.compile(r"(?P<q>(?:" + QUAL + r")?)(?P<a>\$?[A-Z]{1,3}\$?\d{1,7}):(?P<b>\$?[A-Z]{1,3}\$?\d{1,7})(?![A-Za-z0-9_(])")
CELL_RE = re.compile(r"(?<![A-Za-z0-9_.$!'~])(?P<q>(?:" + QUAL + r")?)(?P<dc>\$?)(?P<col>[A-Z]{1,3})(?P<dr>\$?)(?P<row>\d{1,7})(?![A-Za-z0-9_(])")

def _bump_ref(txt, n):
    m = re.match(r"(\$?)([A-Z]{1,3})(\$?)(\d{1,7})$", txt)
    return m.group(1) + colname(colnum(m.group(2)) + n) + m.group(3) + m.group(4)

def _q_ok(m, sheet_name):
    q = m.group("q1") or m.group("q2")
    return True if q is None else (q.strip() == sheet_name.strip())

def _bump_cells(text, sheet_name, n):
    def rep(m):
        if not _q_ok(m, sheet_name):
            return m.group(0)
        cn = colnum(m.group("col"))
        if cn < INSERT_AT:
            return m.group(0)
        pre = m.group(0)[:m.start("dc") - m.start(0)]
        return pre + m.group("dc") + colname(cn + n) + m.group("dr") + m.group("row")
    return CELL_RE.sub(rep, text)

def bump_formula(f, sheet_name, n=NCOL):
    out, pos = [], 0
    for m in RANGE_RE.finditer(f):
        out.append(_bump_cells(f[pos:m.start()], sheet_name, n))
        ca = colnum(re.match(r"\$?([A-Z]{1,3})", m.group("a")).group(1))
        cb = colnum(re.match(r"\$?([A-Z]{1,3})", m.group("b")).group(1))
        if _q_ok(m, sheet_name) and (ca >= INSERT_AT or cb >= INSERT_AT):
            out.append((m.group("q") or "") + (_bump_ref(m.group("a"), n) if ca >= INSERT_AT else m.group("a"))
                       + ":" + (_bump_ref(m.group("b"), n) if cb >= INSERT_AT else m.group("b")))
        else:
            out.append(m.group(0))
        pos = m.end()
    out.append(_bump_cells(f[pos:], sheet_name, n))
    return "".join(out)

import config as C          # noqa
import runner as R          # noqa

def _report_path(tpl, day):
    snap, d8, d4 = DAYDIR[day]
    return os.path.join(DSR, tpl.replace("{d4}", d4)) if "{d4}" in tpl else os.path.join(snap, tpl.replace("{d}", d8))

def load_values(day, ins_tpl, ins_kind, int_tpl, int_kind):
    p1 = _report_path(ins_tpl, day)
    d1, _ = R.read_report(p1, ins_kind, sep=("," if p1.lower().endswith(".csv") else "\t"))
    ins = {s: (d.get("afn-fulfillable-quantity") or 0) if ins_kind == "balance"
           else (d.get("afn-onhand-buyable-quantity") or 0) for s, d in d1.items()}
    p2 = _report_path(int_tpl, day)
    d2, _ = R.read_report(p2, int_kind, sep=("," if p2.lower().endswith(".csv") else "\t"))
    itr = {s: ((d.get("afn-inbound-working-quantity") or 0) + (d.get("afn-inbound-shipped-quantity") or 0)
               + (d.get("afn-inbound-receiving-quantity") or 0)) for s, d in d2.items()}
    return ins, itr

def wb_sheet_map(z):
    wb = z.read("xl/workbook.xml").decode("utf-8")
    rel = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    rm = {}
    for tag in re.findall(r"<Relationship\b[^>]*/>", rel):
        i, t = re.search(r'Id="([^"]*)"', tag), re.search(r'Target="([^"]*)"', tag)
        if i and t:
            rm[i.group(1)] = t.group(1)
    return {n: (rm[r].lstrip("/") if rm[r].startswith("/") else "xl/" + rm[r])
            for n, r in re.findall(r'<sheet name="([^"]*)"[^>]*r:id="(rId\d+)"', wb)}

def parse_cells(body):
    cells = []
    for m in re.finditer(r"<c\b([^>]*?)(/>|>(.*?)</c>)", body, re.S):
        attrs, inner, sc = m.group(1), (m.group(3) or ""), m.group(2) == "/>"
        rm = re.search(r'\br="([A-Z]+\d+)"', attrs)
        if rm:
            cells.append((rm.group(1), re.sub(r'\s*r="[A-Z]+\d+"', "", attrs), inner, sc))
    return cells

def get_formula(inner):
    m = re.search(r"<f\b[^>]*>(.*?)</f>", inner, re.S)
    return m.group(1) if m else None

def is_fake(f, kind):
    return bool(f) and "IFERROR(IF(OR(" in f and (("SUMIF(" in f) if kind == "insale" else ("VLOOKUP(" in f))

def find_runs(xml):
    m = re.search(r'<row\b[^>]*\br="4"(?!\d)[^>]*>(.*?)</row>', xml, re.S)
    cols = []
    for ref, at, inner, sc in parse_cells(m.group(1)):
        if re.fullmatch(r"[A-Z]+4", ref) and re.search(r"<v>\d{4,6}</v>", inner or ""):
            cols.append(colnum(re.match(r"[A-Z]+", ref).group(0)))
    cols.sort()
    runs = []
    for c in cols:
        if runs and c == runs[-1][1] + 1:
            runs[-1][1] = c
        else:
            runs.append([c, c])
    return runs

def transform(xml, sheet_name, runs, fill):
    intra_s, intra_e = runs[1]
    intra_new_s, intra_new_e = intra_s + NCOL, intra_e + NCOL
    # 行1 的 SUM 结束行：取本表任一日期列 r1 公式的收尾行号（各表不同：19957/19761/19991/19644/20126）
    r1end = None
    m1 = re.search(r'<row\b[^>]*\br="1"(?!\d)[^>]*>(.*?)</row>', xml, re.S)
    if m1:
        for ref, at, inn, sc in parse_cells(m1.group(1)):
            f = get_formula(inn)
            if f and f.upper().startswith("SUM("):
                m2 = re.search(r"\$5:[^)]*?(\d+)", f)
                if m2:
                    r1end = m2.group(1)
                    break
    r1end = r1end or "20126"
    # 日期 → 新列号
    def new_col_for(block, date):
        if block == "insale":
            return INSERT_AT + (date - NEW_DATES[0]).days
        return intra_new_s + (date - datetime.date(2026, 9, 1)).days
    m = re.search(r"<sheetData>(.*?)</sheetData>", xml, re.S)
    body = m.group(1)
    st = dict(drop_in=0, drop_it=0, shifted=0, added=0)
    rows_out = []
    for rm in re.finditer(r"<row\b([^>]*?)(/>|>(.*?)</row>)", body, re.S):
        attrs, inner = rm.group(1), (rm.group(3) or "")
        rr = re.search(r'\br="(\d+)"', attrs)
        if not rr:
            continue
        rnum = int(rr.group(1))
        bycol = {}
        for ref, at, inn, sc in parse_cells(inner):
            bycol[colnum(re.match(r"[A-Z]+", ref).group(0))] = (at, inn, sc)
        s_in = None
        if 34 in bycol:
            mm = re.search(r'\bs="(\d+)"', bycol[34][0])
            s_in = mm.group(1) if mm else None
        s_it = None
        if intra_e in bycol:
            mm = re.search(r'\bs="(\d+)"', bycol[intra_e][0])
            s_it = mm.group(1) if mm else None
        if s_it is None and (intra_e - 1) in bycol:
            mm = re.search(r'\bs="(\d+)"', bycol[intra_e - 1][0])
            s_it = mm.group(1) if mm else None
        kept = {}
        for cn, (at, inn, sc) in bycol.items():
            f = get_formula(inn)
            if 3 <= cn <= 34 and is_fake(f, "insale"):
                st["drop_in"] += 1
                continue
            if intra_s <= cn <= intra_e and is_fake(f, "intra"):
                st["drop_it"] += 1
                continue
            if cn >= INSERT_AT:
                ncn = cn + NCOL
                st["shifted"] += 1
                if f:
                    nf = bump_formula(f, sheet_name)
                    inn = re.sub(r"(<f\b[^>]*>).*?(</f>)",
                                 lambda mm: mm.group(1) + nf + mm.group(2), inn, count=1, flags=re.S)
                kept[ncn] = (at, inn, sc)
            else:
                kept[cn] = (at, inn, sc)
        if rnum >= 5:
            for date in FILL_DATES:
                for block, sref, dct in (("insale", s_in, fill["insale"]), ("intra", s_it, fill["intra"])):
                    cn = new_col_for(block, date)
                    if cn in kept:
                        continue
                    val = dct.get(date, {}).get(fill["sku_by_row"].get(rnum), 0)
                    kept[cn] = (f' s="{sref}"' if sref else "", f"<v>{int(val)}</v>", False)
                    st["added"] += 1
        for i in range(NCOL):
            cn = INSERT_AT + i
            if cn in kept:
                continue
            ins = None
            if rnum == 1:
                ins = f"<f>SUM({colname(cn)}$5:{colname(cn)}${r1end})</f>"
            elif rnum == 3:
                ins = "<is><t>%s</t></is>" % ["周一", "周二", "周三", "周四", "周五", "周六", "周日"][NEW_DATES[i].weekday()]
            elif rnum == 4:
                ins = f"<v>{(NEW_DATES[i] - EPOCH).days}</v>"
            if ins is not None:
                t_attr = ' t="inlineStr"' if rnum == 3 else ""
                kept[cn] = ((f' s="{s_in}"' if s_in else "") + t_attr, ins, False)
                st["added"] += 1
        if not kept:
            continue
        cols = sorted(kept)
        parts = []
        for cn in cols:
            at, inn, sc = kept[cn]
            at = re.sub(r'\s*r="[A-Z]+\d+"', "", at)
            at = f' r="{colname(cn)}{rnum}"' + at
            parts.append(f"<c{at}>{inn}</c>" if inn else f"<c{at}/>")
        a2 = re.sub(r'\s*spans="[^"]*"', "", attrs)
        rows_out.append(f'<row{a2} spans="1:{max(cols)}">{"".join(parts)}</row>')
    new_xml = xml[:m.start(1)] + "".join(rows_out) + xml[m.end(1):]

    dm = re.search(r'<dimension ref="([^"]*)"\s*/>', new_xml)
    if dm:
        mm = re.match(r"([A-Z]+)(\d+):([A-Z]+)(\d+)", dm.group(1))
        if mm:
            c2 = colnum(mm.group(3))
            if c2 >= INSERT_AT:
                c2 += NCOL
            new_xml = new_xml[:dm.start()] + f'<dimension ref="{mm.group(1)}{mm.group(2)}:{colname(c2)}{mm.group(4)}" />' + new_xml[dm.end():]

    cm = re.search(r"<cols>(.*?)</cols>", new_xml, re.S)
    if cm:
        ent = []
        for em in re.finditer(r"<col\b([^>]*)/>", cm.group(1)):
            a = em.group(1)
            ent.append((int(re.search(r'\bmin="(\d+)"', a).group(1)),
                        int(re.search(r'\bmax="(\d+)"', a).group(1)), a))
        ne = []
        for mn, mx, a in ent:
            if mx < INSERT_AT - 1:
                ne.append((mn, mx, a))
            elif mx == INSERT_AT - 1:
                newmx = INSERT_AT + NCOL - 1
                ne.append((mn, newmx, re.sub(r'\bmax="\d+"', f'max="{newmx}"', a)))
            else:
                n2 = mx + NCOL
                a2 = re.sub(r'\bmax="\d+"', f'max="{n2}"', a)
                if mn >= INSERT_AT:
                    m2 = mn + NCOL
                    a2 = re.sub(r'\bmin="\d+"', f'min="{m2}"', a2)
                    ne.append((m2, n2, a2))
                else:
                    ne.append((mn, n2, a2))
        ne.sort(key=lambda t: t[0])
        new_xml = new_xml[:cm.start()] + "<cols>" + "".join(f"<col{x}/>" for _, _, x in ne) + "</cols>" + new_xml[cm.end():]

    def fix_merge(mm):
        c1, r1, c2, r2 = mm.group(1), mm.group(2), mm.group(3), mm.group(4)
        n1, n2 = colnum(c1), colnum(c2)
        if r1 == "2" and n1 == 3 and n2 == 34:
            n2 = INSERT_AT + NCOL - 1
        else:
            if n1 >= INSERT_AT:
                n1 += NCOL
            if n2 >= INSERT_AT:
                n2 += NCOL
            if r1 == "2" and n1 == intra_new_s:
                n2 = intra_new_e
        return f'<mergeCell ref="{colname(n1)}{r1}:{colname(n2)}{r2}" />'
    new_xml = re.sub(r'<mergeCell ref="([A-Z]+)(\d+):([A-Z]+)(\d+)"\s*/>', fix_merge, new_xml)

    def fix_cf(mm):
        head, sq, tail = mm.group(1), mm.group(2), mm.group(3)
        parts = []
        for rg in sq.split():
            q = re.match(r"([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$", rg)
            n1 = colnum(q.group(1))
            n2 = colnum(q.group(3)) if q.group(3) else None
            if n2 is None:
                parts.append(f"{colname(n1 + NCOL if n1 >= INSERT_AT else n1)}{q.group(2)}")
            else:
                n2n = (INSERT_AT + NCOL - 1) if n2 == 34 else (n2 + NCOL if n2 >= INSERT_AT else n2)
                n1n = n1 + NCOL if n1 >= INSERT_AT else n1
                parts.append(f"{colname(n1n)}{q.group(2)}:{colname(n2n)}{q.group(4)}")
        def fr(m2):
            raw = m2.group(2).replace("&lt;", "<").replace("&gt;", ">").replace("&amp;", "&")
            return m2.group(1) + bump_formula(raw, sheet_name).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;") + m2.group(3)
        tail2 = re.sub(r"(<formula>)(.*?)(</formula>)", fr, tail, flags=re.S)
        return head + " ".join(parts) + tail2
    new_xml = re.sub(r'(<conditionalFormatting sqref=")([^"]*)(">.*?</conditionalFormatting>)', fix_cf, new_xml, flags=re.S)
    return new_xml, st

def bump_sheet_refs(xml, names):
    """把其它 sheet 里指向记录表的公式列号整体 +N"""
    def rep(m):
        tag, inner = m.group(1), m.group(2)
        raw = inner.replace("&lt;", "<").replace("&gt;", ">").replace("&amp;", "&")
        new = raw
        for nm in names:
            if nm in new:
                new = bump_formula(new, nm)
        if new == raw:
            return m.group(0)
        return tag + new.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;") + "</f>"
    return re.sub(r"(<f[^>]*>)(.*?)</f>", rep, xml, flags=re.S)


def main():
    apply_mode = "--apply" in sys.argv
    only = None
    for a in sys.argv[1:]:
        if a.startswith("--only="):
            only = a.split("=", 1)[1]
    out = []
    for book, sheet, ins_tpl, ins_kind, int_tpl, int_kind in SHEETCFG:
        if only and not (book == only or book.startswith(only)):
            continue
        path = os.path.join(WORK, book)
        z = zipfile.ZipFile(path)
        smap = wb_sheet_map(z)
        shp = smap[sheet]
        xml = z.read(shp).decode("utf-8")
        runs = find_runs(xml)
        assert runs[0][0] == 3 and runs[0][1] == 34, (book, sheet, runs)
        rows = re.findall(r'<row[^>]*r="(\d+)"[^>]*>(.*?)</row>', xml, re.S)
        sku_by_row = {}
        for rn, body in rows:
            if int(rn) < 5:
                continue
            mm = re.search(r'<c\b[^>]*r="B\d+"[^>]*>(.*?)</c>', body, re.S)
            if mm:
                t = re.search(r"<t[^>]*>(.*?)</t>", mm.group(1), re.S)
                if t:
                    sku_by_row[int(rn)] = t.group(1)
        insale, intra = {}, {}
        for day in FILL_DATES:
            i1, i2 = load_values(day, ins_tpl, ins_kind, int_tpl, int_kind)
            insale[day], intra[day] = i1, i2
        fill = dict(sku_by_row=sku_by_row, insale=insale, intra=intra)
        new_xml, st = transform(xml, sheet, runs, fill)
        out.append(f"{book} | {sheet}: runs={runs} SKU行={len(sku_by_row)} "
                   f"移位={st['shifted']} 删在售假公式={st['drop_in']} 删在途假公式={st['drop_it']} 新写格={st['added']}")
        if apply_mode:
            allnames = z.namelist()
            newpath = path + ".new.xlsx"
            # 本簿全部记录表名（跨表引用需要一起平移）
            rec_names = [sheet]   # 只平移本表引用，避免多次迭代重复 +10
            changed = {}
            with zipfile.ZipFile(newpath, "w", zipfile.ZIP_DEFLATED) as zo:
                for n in allnames:
                    if n == shp:
                        zo.writestr(n, new_xml.encode("utf-8"))
                    elif n in smap.values():
                        raw = z.read(n).decode("utf-8")
                        if any(rn in raw for rn in rec_names):
                            out2 = bump_sheet_refs(raw, rec_names)
                            if out2 != raw:
                                changed[n] = True
                                zo.writestr(n, out2.encode("utf-8"))
                                continue
                        zo.writestr(n, z.read(n))
                    else:
                        zo.writestr(n, z.read(n))
            if changed:
                out.append("    跨表公式平移: " + ", ".join(sorted(changed)))
            z.close()
            zz = zipfile.ZipFile(newpath)
            bad = zz.testzip(); zz.close()
            if bad:
                raise RuntimeError("zip 坏: " + bad)
            shutil.copyfile(newpath, path)
            os.remove(newpath)
            out.append("    -> 写回完成")
        else:
            z.close()
    print("\n".join(out))
    print("DONE", "APPLIED" if apply_mode else "DRY")

if __name__ == "__main__":
    main()
