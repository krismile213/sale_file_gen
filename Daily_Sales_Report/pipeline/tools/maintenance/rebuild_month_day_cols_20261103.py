# -*- coding: utf-8 -*-
"""给分析表补建 11 月日销列（XML 手术，不改其它部件）。

背景（2026-11-02 实测）：
  分析表的日销列不是每天自动加，而是「上月最后7天 + 当月全月 + 次月1天」整块存在。
  9 月终态 = 8/25~10/01（38 列），10 月 = 9/24~11/01（39 列），
  10→11 换月只裁行不建列 → 当前簿仍是 9/24~11/01，**缺 11/2~11/30**，
  明天导入的 11-02 订单在明细里，却因为没有日列而不显示。

做法：在最后一天（11-01）之后插入 N 列，整块右侧右移，并同步
  单元格引用 / 公式（本表 + 跨表）/ dimension / cols / mergeCells /
  conditionalFormatting / dataValidation / 行1 公式。

用法：
  python rebuild_month_day_cols_20261103.py --dry-run
  python rebuild_month_day_cols_20261103.py --apply [--only 阿联酋]
"""
import os
import re
import sys
import shutil
import zipfile
import datetime
import argparse

_HERE = os.path.dirname(os.path.abspath(__file__))
_CORE = os.path.abspath(os.path.join(_HERE, "..", "..", "core"))
sys.path.insert(0, _CORE)
import record_extend as RE                                    # noqa: E402

DSR = os.path.abspath(os.path.join(_HERE, "..", "..", ".."))
WORK = os.path.join(DSR, "1.每日销售数据表")
BACKUP = os.path.join(DSR, ".workbuddy", "backup_daycols_20261103")

# 本月最后一天 / 次月 1 号（与 10 月模板口径一致：当月全月 + 次月1天）
MONTH_END = datetime.date(2026, 11, 30)
NEXT_1ST = datetime.date(2026, 12, 1)
SNAP_MAX = datetime.date(2026, 11, 30)      # 只补到 11/30

L, CN = RE.colname, RE.colnum


# ---------------- 通用：列右移 ----------------
def _shift_range_text(txt, at, n):
    """'Y7:BK500' / 'BK7' → 列 >= at 的 +n（支持多段空格分隔的 sqref）"""
    if not txt:
        return txt
    out = []
    for seg in txt.split():
        m = re.fullmatch(r"(\$?)([A-Z]{1,3})(\$?\d*):(\$?)([A-Z]{1,3})(\$?\d*)", seg)
        if m:
            a = CN(m.group(2))
            b = CN(m.group(5))
            na = a + n if a >= at else a
            nb = b + n if b >= at else b
            out.append("%s%s%s:%s%s%s" % (m.group(1), L(na), m.group(3),
                                          m.group(4), L(nb), m.group(6)))
            continue
        m = re.fullmatch(r"(\$?)([A-Z]{1,3})(\$?\d+)", seg)
        if m:
            c = CN(m.group(2))
            nc = c + n if c >= at else c
            out.append("%s%s%s" % (m.group(1), L(nc), m.group(3)))
            continue
        out.append(seg)
    return " ".join(out)


def _shift_attrs_ref(m, at, n):
    """REF_ATTR_RE 组位：1=前缀 2=$ 3=列A 4=$ 5=行A 6=$ 7=列B 8=$ 9=行B 10=收尾引号"""
    pre = m.group(1)
    da, ca, dra, ra = m.group(2), m.group(3), m.group(4), m.group(5)
    na = CN(ca) + n if CN(ca) >= at else CN(ca)
    out = pre + da + L(na) + dra + ra
    if m.group(7):
        db, cb, drb, rb = m.group(6), m.group(7), m.group(8), m.group(9)
        nb = CN(cb) + n if CN(cb) >= at else CN(cb)
        out += ":" + db + L(nb) + drb + rb
    return out + m.group(10)


def shift_sheet_xml(xml, sheet_name, at, n, other_sheets):
    """把 >= at 的列右移 n（单元格 + 公式 + dimension + cols + merge + CF + DV）。"""
    if n <= 0:
        return xml, 0

    def delta(c):
        return c + n if c >= at else c

    # 1) sheetData 里的单元格（原样保留标签，只改 r 里的列号 + 公式）
    m = re.search(r"<sheetData>(.*?)</sheetData>", xml, re.S)
    n_cell = 0
    if m:
        body = m.group(1)

        def cell_rep(mm):
            nonlocal n_cell
            attrs, tail = mm.group(1), mm.group(2)
            inn = mm.group(3) or ""
            rm = re.search(r'\br="([A-Z]+)(\d+)"', attrs)
            if not rm:
                return mm.group(0)
            c = CN(rm.group(1))
            nc = delta(c)
            newattrs = attrs
            if nc != c:
                n_cell += 1
                newattrs = attrs.replace(rm.group(0), 'r="%s%s"' % (L(nc), rm.group(2)), 1)
            if inn and "<f" in inn:
                inn = _shift_formula(inn, sheet_name, at, n, other_sheets)
                inn = RE.REF_ATTR_RE.sub(lambda x: _shift_attrs_ref(x, at, n), inn)
            if tail == "/>":
                return "<c%s/>" % newattrs
            return "<c%s>%s</c>" % (newattrs, inn)

        body = re.sub(r"<c\b([^>]*?)(/>|>(.*?)</c>)", cell_rep, body, flags=re.S)
        xml = xml[:m.start(1)] + body + xml[m.end(1):]

    # 2) dimension
    def dim_rep(mm):
        return '<dimension ref="%s"/>' % _shift_range_text(mm.group(1), at, n)
    xml = re.sub(r'<dimension ref="([^"]+)"/>', dim_rep, xml)

    # 3) cols 定义（跨界的拆开）
    #    ⚠️ 兼容属性顺序：col 元素的 min/max 可能在前（openpyxl 格式）也可能在尾
    #    （历史 XML 手术格式 <col outlineLevel="1" ... min="9" max="29"/>）。
    #    旧正则只认 min/max 在前 → min/max 在尾的元素不右移 → 货件列 outline 丢失/错位。
    def cols_rep(mm):
        blk = mm.group(0)

        def one(cm):
            a = cm.group(1)
            gm, gx = re.search(r'\bmin="(\d+)"', a), re.search(r'\bmax="(\d+)"', a)
            if not (gm and gx):
                return cm.group(0)
            mn, mx = int(gm.group(1)), int(gx.group(1))
            rest = re.sub(r'\b(min|max)="(\d+)"\s*', "", a)
            rest = re.sub(r"\s+", " ", rest).strip()
            tail = (" " + rest) if rest else ""
            if mx < at:
                return cm.group(0)
            if mn >= at:
                return '<col min="%d" max="%d"%s/>' % (mn + n, mx + n, tail)
            left = '<col min="%d" max="%d"%s/>' % (mn, at - 1, tail)
            right = '<col min="%d" max="%d"%s/>' % (at + n, mx + n, tail)
            return left + right

        return re.sub(r"<col\b([^>]*?)/>", one, blk)
    xml = re.sub(r"<cols>.*?</cols>", cols_rep, xml, flags=re.S)

    # 4) mergeCells / CF / DV 的 sqref
    def sq_rep(mm):
        return mm.group(1) + _shift_range_text(mm.group(2), at, n) + mm.group(3)
    xml = re.sub(r'(<mergeCell ref=")([^"]+)(")', sq_rep, xml)
    xml = re.sub(r'(<conditionalFormatting sqref=")([^"]+)(")', sq_rep, xml)
    xml = re.sub(r'(<dataValidation[^>]*?sqref=")([^"]+)(")', sq_rep, xml)

    return xml, n_cell


def _shift_formula(inn, sheet_name, at, n, other_sheets):
    """只处理 <f> 文本；本表非限定引用 + 限定本表/其它被右移表的引用都按列移。"""
    def f_rep(mm):
        f = mm.group(2)
        f2 = RE._shift_cells(f, sheet_name, lambda c: c + n if c >= at else c,
                             qualified_only=False)
        for osh in other_sheets:
            f2 = RE._shift_cells(f2, osh, lambda c: c + n if c >= at else c,
                                 qualified_only=True)
        return mm.group(1) + f2 + mm.group(3)
    return re.sub(r"(<f[^>]*>)(.*?)(</f>)", f_rep, inn, flags=re.S)


# ---------------- 主流程 ----------------
def find_day_block(rows5):
    """返回 (起列, 末列, [(列, date)])，取连续的 datetime 段里最长的一段。"""
    ds = [(i + 1, v.date() if isinstance(v, datetime.datetime) else v)
          for i, v in enumerate(rows5) if isinstance(v, (datetime.datetime, datetime.date))]
    if not ds:
        return None
    blocks, cur = [], [ds[0]]
    for p, q in zip(ds, ds[1:]):
        if q[0] - p[0] == 1 and (q[1] - p[1]).days == 1:
            cur.append(q)
        else:
            blocks.append(cur); cur = [q]
    blocks.append(cur)
    best = max(blocks, key=len)
    return best[0][0], best[-1][0], best


SAME_COL_SUM = re.compile(r"^SUM\(([A-Z]+)\d+:(\1)\d+\)$")


def _widen_formula(f, old_last, new_last):
    """把范围末端 ':<old_last>' 换成 ':<new_last>'。
    跳过 SUM(BK7:BK2727) 这类「同行同列」的区间（那是行6 合计，不是跨列范围）。"""
    if SAME_COL_SUM.match(f.strip()):
        return f
    return re.sub(r"(:\$?)%s(?=\$?\d|\b)(?!:)" % old_last, r"\1" + new_last, f)


def widen_parts(parts, smap, targets_info):
    """targets_info: [(sheet_name, old_last, new_last)]。扫全部 part 的公式。"""
    by_sheet = {t[0]: (t[1], t[2]) for t in targets_info}
    n = 0
    for name, part in smap.items():
        xml = parts.get(part)
        if xml is None:
            continue
        txt = xml.decode("utf-8")
        hit = False
        for sname, (ol, nl) in by_sheet.items():
            if sname == name:
                pass
            elif ("'%s'!" % sname) not in txt and ("%s!" % sname) not in txt:
                continue
            before = txt
            txt = re.sub(r"(<f[^>]*>)(.*?)(</f>)",
                         lambda m: m.group(1) + _widen_formula(m.group(2), ol, nl) + m.group(3),
                         txt, flags=re.S)
            # 条件格式 sqref / 规则公式里的日列末端也要扩
            txt = re.sub(r'(sqref=")([^"]+)(")',
                         lambda m: m.group(1) + _widen_formula(m.group(2), ol, nl) + m.group(3), txt)
            txt = re.sub(r"(<formula[^>]*>)(.*?)(</formula>)",
                         lambda m: m.group(1) + _widen_formula(m.group(2), ol, nl) + m.group(3),
                         txt, flags=re.S)
            if txt != before:
                hit = True
                n += 1
        if hit:
            parts[part] = txt.encode("utf-8")
    return n


def process(path, dry=True, month_end=None):
    """month_end：日列补到哪天（默认=模块常量 2026-11-30，即 11-03 那次补列）。
    任务㉓（2026-11-03）：month_rollover 阶段2.5 传 TO_PERIOD 月末调用本函数，
    实现「换月即建列」，不再需要事后手工渲染。"""
    import openpyxl
    if month_end is None:
        month_end = MONTH_END
    wb = openpyxl.load_workbook(path, read_only=True, data_only=False)
    targets = []
    for ws in wb.worksheets:
        if "每日库存分析表" not in ws.title:
            continue
        r5 = next(ws.iter_rows(min_row=5, max_row=5, max_col=200, values_only=True))
        info = find_day_block(r5)
        if not info:
            print("    [跳过] %s: 行5 无日列" % ws.title)
            continue
        s, e, blk = info
        last = blk[-1][1]
        if last >= month_end:
            print("    [跳过] %s: 日列已到 %s" % (ws.title, last))
            continue
        new_dates = []
        d = last + datetime.timedelta(days=1)
        while d <= month_end:
            new_dates.append(d)
            d += datetime.timedelta(days=1)
        targets.append((ws.title, s, e, last, new_dates))
    wb.close()

    if not targets:
        print("  → 无需处理")
        return 0

    for t in targets:
        print("    %s: 日列 %s → %s，补 %d 列（%s ~ %s）" % (
            t[0], L(t[1]), L(t[2]), len(t[4]), t[4][0], t[4][-1]))

    if dry:
        return sum(len(t[4]) for t in targets)

    # ---- 非 dry：XML 手术 ----
    z = zipfile.ZipFile(path)
    names = [i.filename for i in z.infolist()]
    parts = {nm: z.read(nm) for nm in names}
    z.close()

    wbx = openpyxl.load_workbook(path, read_only=True)
    smap = {}
    for ws in wbx.worksheets:
        smap[ws.title] = ws
    wbx.close()
    part_map = RE.sheet_part_map(zipfile.ZipFile(path))

    all_sheets = [t[0] for t in targets]
    total = 0
    for title, s, e, last, new_dates in targets:
        part = part_map[title]
        xml = parts[part].decode("utf-8")
        n = len(new_dates)
        at = e + 1
        # 1) 右移
        xml, nc = shift_sheet_xml(xml, title, at, n, all_sheets)
        # 2) 插新列
        xml = insert_cols(xml, title, e, new_dates)
        parts[part] = xml.encode("utf-8")
        print("    [手术] %s: 单元格右移 %d，插 %d 列" % (title, nc, n))
        total += n

    # 3) 硬编码「日列末列」的范围扩到新末列（$Y$5:$BK$5 / Y7:BK7 / A:BK …）
    info = [(t[0], L(t[2]), L(t[2] + len(t[4]))) for t in targets]
    widen_parts(parts, part_map, info)
    for sname, ol, nl in info:
        print("    [范围] %s: 日列区间末端 %s → %s" % (sname, ol, nl))

    tmp = path + ".daycols.tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zo:
        for nm in names:
            zo.writestr(nm, parts[nm])
    zz = zipfile.ZipFile(tmp)
    bad = zz.testzip()
    zz.close()
    if bad:
        raise RuntimeError("zip 损坏: %s" % bad)
    os.replace(tmp, path)
    return total


def _sum_end_row(row_inner, col_letter):
    """从该行 <col>6 的 =SUM(BK7:BK2727) 里取结束行号 2727。"""
    m = re.search(r'\br="%s6"[^>]*>(.*?)</c>' % col_letter, row_inner, re.S)
    if not m:
        return None
    fm = re.search(r"SUM\([A-Z]+\d+:[A-Z]+(\d+)\)", m.group(1))
    return fm.group(1) if fm else None


def insert_cols(xml, sheet_name, after_col, dates):
    """在 after_col 之后插入 len(dates) 列：行5 日期 + 行6 SUM + 行7+ 公式（复制 after_col 的）。
    只往原文本里插，不重建其它单元格。"""
    m = re.search(r"<sheetData>(.*?)</sheetData>", xml, re.S)
    if not m:
        return xml
    body = m.group(1)
    anchor = L(after_col)

    def row_rep(mm):
        attrs, tail = mm.group(1), mm.group(2)
        inner = mm.group(3) or ""
        if tail == "/>":
            return mm.group(0)
        rr = re.search(r'\br="(\d+)"', attrs)
        if not rr:
            return mm.group(0)
        rnum = int(rr.group(1))
        # 模板 cell（after_col 那一格），取其样式
        tm = re.search(r'<c r="%s%d"([^>]*?)(/>|>(.*?)</c>)' % (anchor, rnum), inner, re.S)
        sattr = ""
        tmpl_in = ""
        if tm:
            sm = re.search(r'\bs="(\d+)"', tm.group(1))
            sattr = ' s="%s"' % sm.group(1) if sm else ""
            tmpl_in = tm.group(3) or ""
        new = []
        for i, d in enumerate(dates):
            c = after_col + 1 + i
            if rnum == 5:
                serial = (d - RE.EPOCH).days
                new.append('<c r="%s5"%s><v>%d</v></c>' % (L(c), sattr, serial))
            elif rnum == 6:
                end = _sum_end_row(inner, anchor) or RE._r1end(xml, after_col, after_col)
                new.append('<c r="%s6"%s><f>SUM(%s7:%s%s)</f></c>' % (
                    L(c), sattr, L(c), L(c), end))
            elif rnum >= 7 and tmpl_in:
                fml = RE.get_formula(tmpl_in)
                if fml:
                    nf = fml.replace("%s$5" % anchor, "%s$5" % L(c))
                    new.append('<c r="%s%d"%s><f>%s</f></c>' % (L(c), rnum, sattr, nf))
                else:
                    new.append('<c r="%s%d"%s/>' % (L(c), rnum, sattr))
            elif rnum in (2, 3, 4):
                new.append('<c r="%s%d"%s/>' % (L(c), rnum, sattr))
        if not new:
            return mm.group(0)
        blob = "".join(new)
        if tm:
            pos = tm.end()
            return "<row%s>%s%s%s</row>" % (attrs, inner[:pos], blob, inner[pos:])
        return "<row%s>%s%s</row>" % (attrs, inner, blob)

    body = re.sub(r"<row\b([^>]*?)(/>|>(.*?)</row>)", row_rep, body, flags=re.S)
    return xml[:m.start(1)] + body + xml[m.end(1):]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--apply", action="store_true")
    ap.add_argument("--only", default="")
    args = ap.parse_args()
    files = sorted(f for f in os.listdir(WORK)
                   if f.endswith("库存销售表.xlsx") and not f.startswith("~$")
                   and (not args.only or args.only in f))
    print("模式:", "APPLY" if args.apply else "DRY-RUN", "| 待处理", len(files), "本")
    if args.apply:
        os.makedirs(BACKUP, exist_ok=True)
    tot = 0
    for f in files:
        p = os.path.join(WORK, f)
        print("\n==== %s ====" % f)
        if args.apply:
            bp = os.path.join(BACKUP, f)
            if not os.path.exists(bp):
                shutil.copy2(p, bp)
        tot += process(p, dry=not args.apply)
    print("\n合计新增列:", tot)


if __name__ == "__main__":
    main()
