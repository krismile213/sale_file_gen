# -*- coding: utf-8 -*-
"""库存记录 sheet：按天加列 + 只写跑批当天固化值（XML 手术，openpyxl 保存后执行）

为什么重写（老 extend_stock_record 的三个坑）：
  1. 老逻辑只往行4最右日期列后追加 ⇒ 只喂最右边的在途块，「在售块」永远长不了（饿死）；
  2. 新列 = 原样复制上一列：上一列是活公式（无日期条件 VLOOKUP 可售+在途表）时，
     新列间接取"当下快照"值，且所有活公式老列每次跑批后被静默刷新成最新值 ⇒ 历史假；
     上一列是固化静态值时，新列原样复制前一天的值 ⇒ 当天报告根本没被读；
  3. 行1 SUM 同样被复制，静态值列会把上一天的合计抄进新列。

本模块口径（用户 2026-10-13 拍板）：
  - 在售块 / 在途块各自按天加列：在售块在块尾（分隔列前）插入，右侧整体右移；
    在途块在表尾追加。两块末日不同的（如在售停 10/02、在途到 10/12）各自补齐。
  - 只有"本次跑批日"写固化字面值（来自当天库存报告，逐 SKU）；
    中间无报告的日期只建列（日期/星期/行1 SUM），数据区留空。
  - 其它列一律不动：历史固化值不被刷新、活公式列也不新增。
  - 同步平移：dimension / cols / mergeCells / conditionalFormatting /
    表内公式 / 跨表公式（数据透视图的 SUMPRODUCT 取数范围等）。

取值口径（与 2026-10-13 在售块改造 verify_full.py 实测一致）：
  - 在售 = afn-onhand-buyable-quantity（墨西哥记录表用 afn-fulfillable-quantity）
  - 在途 = afn-inbound-working + afn-inbound-shipped + afn-inbound-receiving
  -  sheet 里有、当天报告没有的 SKU → 0（与老假公式 IFERROR(...,0) 兜底一致）

用法：extend_record_file(path, sheet_name, "2026-10-13", insale={sku:v}, intra={sku:v})
必须在 openpyxl 保存（wb.save）之后调用——手术直接改 xlsx 里的 sheet XML。
"""
import datetime
import re
import zipfile
import shutil

WEEK_TXT = ["周一", "周二", "周三", "周四", "周五", "周六", "周日"]
EPOCH = datetime.date(1899, 12, 30)          # Excel 日期序列起点
DEFAULT_R1END = "20126"                       # 行1 SUM 收尾行号兜底
L = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"

# 在途三件套（inv 统一字段名；age 报告已由 read_report 映射成同名）
INTRALE_FIELDS = ("afn-inbound-working-quantity",
                  "afn-inbound-shipped-quantity",
                  "afn-inbound-receiving-quantity")


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


# ---------------- 公式引用识别（移植自 fix_onsale_blocks.py，已在生产簿验证） ----------------
QUAL = r"(?:'(?P<q1>[^']+)'|(?P<q2>[A-Za-z0-9_\u4e00-\u9fff\.~]+))!"
RANGE_RE = re.compile(r"(?P<q>(?:" + QUAL + r")?)(?P<a>\$?[A-Z]{1,3}\$?\d{1,7}):(?P<b>\$?[A-Z]{1,3}\$?\d{1,7})(?![A-Za-z0-9_(])")
CELL_RE = re.compile(r"(?<![A-Za-z0-9_.$!'~])(?P<q>(?:" + QUAL + r")?)(?P<dc>\$?)(?P<col>[A-Z]{1,3})(?P<dr>\$?)(?P<row>\d{1,7})(?![A-Za-z0-9_(])")
REF_ATTR_RE = re.compile(r'(<f\b[^>]*?\bref=")(\$?)([A-Z]{1,3})(\$?)(\d{1,7})(?::(\$?)([A-Z]{1,3})(\$?)(\d{1,7}))?(")')


def _bump_ref(txt, n):
    """'A12'/'$A$12' → 列 +n"""
    m = re.match(r"(\$?)([A-Z]{1,3})(\$?)(\d{1,7})$", txt)
    return m.group(1) + colname(colnum(m.group(2)) + n) + m.group(3) + m.group(4)


def _q_ok(m, sheet_name, qualified_only=False):
    q = m.group("q1") or m.group("q2")
    if q is None:
        # 非限定引用：本表公式里视为本表；跨表公式里不动（那是别的表自己的格子）
        return not qualified_only
    return q.strip() == sheet_name.strip()


def _col_of(token):
    return colnum(re.match(r"\$?([A-Z]{1,3})", token).group(1))


def _shift_cells(text, sheet_name, delta_fn, qualified_only=False):
    """把 text 里的单元格引用按 delta_fn(col) 平移（返回 0 不动）。
    qualified_only=True 时只平移"显式带了本表名前缀"的引用（跨表公式用，
    避免误动其它 sheet 自己的非限定引用）；False 时非限定引用视为本表引用。
    区域引用整段处理，避免单元格正则二次命中。"""
    out, pos = [], 0
    for m in RANGE_RE.finditer(text):
        out.append(_shift_plain_cells(text[pos:m.start()], sheet_name, delta_fn, qualified_only))
        if _q_ok(m, sheet_name, qualified_only):
            ca, cb = _col_of(m.group("a")), _col_of(m.group("b"))
            na, nb = delta_fn(ca), delta_fn(cb)
            a = m.group("a") if na == ca else _bump_ref(m.group("a"), na - ca)
            b = m.group("b") if nb == cb else _bump_ref(m.group("b"), nb - cb)
            out.append((m.group("q") or "") + a + ":" + b)
        else:
            out.append(m.group(0))
        pos = m.end()
    out.append(_shift_plain_cells(text[pos:], sheet_name, delta_fn, qualified_only))
    return "".join(out)


def _shift_plain_cells(text, sheet_name, delta_fn, qualified_only=False):
    def rep(m):
        if not _q_ok(m, sheet_name, qualified_only):
            return m.group(0)
        cn = colnum(m.group("col"))
        nn = delta_fn(cn)
        if nn == cn:
            return m.group(0)
        pre = m.group(0)[:m.start("dc") - m.start(0)]
        return pre + m.group("dc") + colname(nn) + m.group("dr") + m.group("row")
    return CELL_RE.sub(rep, text)


# ---------------- XML 基础解析 ----------------
def parse_cells(body):
    """<c .../> 或 <c ...>...</c> → [(ref, attrs, inner, self_closing)]"""
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


def sheet_part_map(z):
    wb = z.read("xl/workbook.xml").decode("utf-8")
    rel = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    rm = {}
    for tag in re.findall(r"<Relationship\b[^>]*/>", rel):
        i, t = re.search(r'Id="([^"]*)"', tag), re.search(r'Target="([^"]*)"', tag)
        if i and t:
            rm[i.group(1)] = t.group(1)
    return {n: (rm[r].lstrip("/") if rm[r].startswith("/") else "xl/" + rm[r])
            for n, r in re.findall(r'<sheet name="([^"]*)"[^>]*r:id="(rId\d+)"', wb)}


def resolve_sheet_name(smap, want):
    """表名容错：config 里可能带尾随空格（'英国库存记录 '）"""
    if want in smap:
        return want
    for n in smap:
        if n.strip() == want.strip():
            return n
    return None


def find_runs(xml):
    """行4 日期列 → 连续块 [[start,end],...]（列号，1-based）。date 序列 4-6 位。

    ⚠️ 必须排除 t="s"（共享串）格：它的 <v> 是**串表索引**而非数值，可能恰好是 4-6 位数字
    而被误判成日期列。实例：英国/欧洲库存记录 行4 的 A4='状态：'(索引 103336)、B4='SKU'
    (索引 223) —— 不排除就会多出一个"日期块"[1,1]，令 extend_record_file 误判
    「行4 日期块 3 个（>2，结构异常）」而静默跳过加列。
    （openpyxl 存盘会把共享串转 inlineStr，所以经 runner 调用时碰巧不会触发；
    但直接对本模块调用、或将来 openpyxl 改回保留 sharedStrings 时就会炸。）"""
    m = re.search(r'<row\b[^>]*\br="4"(?!\d)[^>]*>(.*?)</row>', xml, re.S)
    if not m:
        return []
    cols = []
    for ref, at, inner, sc in parse_cells(m.group(1)):
        if 't="s"' in at:
            continue
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


def serial_to_date(s):
    return EPOCH + datetime.timedelta(days=int(s))


def _row4_dates(xml, runs):
    """每块的日期列表 {块索引: [(col, date), ...]}"""
    m = re.search(r'<row\b[^>]*\br="4"(?!\d)[^>]*>(.*?)</row>', xml, re.S)
    out = {}
    if not m:
        return out
    for ref, at, inner, sc in parse_cells(m.group(1)):
        mm = re.fullmatch(r"([A-Z]+)4", ref)
        if not mm:
            continue
        if 't="s"' in at:          # 共享串索引不是日期（同 find_runs）
            continue
        vm = re.search(r"<v>(\d{4,6})</v>", inner or "")
        if not vm:
            continue
        cn = colnum(mm.group(1))
        for bi, (s, e) in enumerate(runs):
            if s <= cn <= e:
                out.setdefault(bi, []).append((cn, serial_to_date(vm.group(1))))
    for bi in out:
        out[bi].sort()
    return out


def _sku_by_row(xml, shared=None):
    """{行号: SKU}（B 列；inlineStr 直接取，t='s' 共享串查表）"""
    out = {}
    for rm in re.finditer(r"<row\b([^>]*?)(/>|>(.*?)</row>)", xml, re.S):
        attrs, inner = rm.group(1), (rm.group(3) or "")
        rr = re.search(r'\br="(\d+)"', attrs)
        if not rr or int(rr.group(1)) < 5:
            continue
        for ref, at, inn, sc in parse_cells(inner):
            mm = re.fullmatch(r"B(\d+)", ref)
            if not mm:
                continue
            t = re.search(r"<t[^>]*>(.*?)</t>", inn or "", re.S)
            if t:
                out[int(mm.group(1))] = t.group(1)
            else:
                vs = re.search(r"<v>(\d+)</v>", inn or "")
                if vs and shared is not None and 't="s"' in at:
                    idx = int(vs.group(1))
                    if 0 <= idx < len(shared):
                        out[int(mm.group(1))] = shared[idx]
    return out


def _r1end(xml, c_lo, c_hi):
    """行1 SUM 公式收尾行号：优先 [c_lo,c_hi] 块内，其次全表，兜底 DEFAULT_R1END"""
    m = re.search(r'<row\b[^>]*\br="1"(?!\d)[^>]*>(.*?)</row>', xml, re.S)
    if not m:
        return DEFAULT_R1END
    hits = []
    for ref, at, inn, sc in parse_cells(m.group(1)):
        mm = re.fullmatch(r"([A-Z]+)1", ref)
        if not mm:
            continue
        f = get_formula(inn)
        if f and f.upper().startswith("SUM("):
            m2 = re.search(r"\$5:[^)]*?(\d+)", f)
            if m2:
                hits.append((colnum(mm.group(1)), m2.group(1)))
    for cn, end in hits:
        if c_lo <= cn <= c_hi:
            return end
    return hits[0][1] if hits else DEFAULT_R1END


# ---------------- 主变换 ----------------
def _transform(xml, sheet_name, ctx):
    ins_s, ins_e = ctx["ins_s"], ctx["ins_e"]
    int_s, int_e = ctx["int_s"], ctx["int_e"]
    days_ins, days_int = ctx["days_ins"], ctx["days_int"]
    ncol, nint = len(days_ins), len(days_int)
    insert_at = ins_e + 1
    fill_date = ctx["fill_date"]

    # 新列位置：(列号, 日期)
    ins_new = [(insert_at + i, d) for i, d in enumerate(days_ins)]
    int_new = [(int_e + ncol + 1 + i, d) for i, d in enumerate(days_int)]
    ins_fill_col = next((cn for cn, d in ins_new if d == fill_date), None)
    int_fill_col = next((cn for cn, d in int_new if d == fill_date), None)

    def same_sheet_delta(col):
        return col + ncol if col >= insert_at else col

    m = re.search(r"<sheetData>(.*?)</sheetData>", xml, re.S)
    if not m:
        return xml, dict(added=0, filled=0)
    body = m.group(1)
    rows_out = []
    n_added = n_filled = 0
    # 行1 SUM 缓存值必须预算：行1 在行5+ 之前处理，逐行累加会全是 0
    sheet_skus = set(ctx["sku_by_row"].values())
    col_sums = {}
    if ins_fill_col is not None:
        col_sums[ins_fill_col] = sum(int((ctx["insale"] or {}).get(s, 0) or 0) for s in sheet_skus)
    if int_fill_col is not None:
        col_sums[int_fill_col] = sum(int((ctx["intra"] or {}).get(s, 0) or 0) for s in sheet_skus)
    for rm in re.finditer(r"<row\b([^>]*?)(/>|>(.*?)</row>)", body, re.S):
        attrs, inner = rm.group(1), (rm.group(3) or "")
        rr = re.search(r'\br="(\d+)"', attrs)
        if not rr:
            continue
        rnum = int(rr.group(1))
        bycol = {}
        for ref, at, inn, sc in parse_cells(inner):
            bycol[colnum(re.match(r"[A-Z]+", ref).group(0))] = (at, inn, sc)
        # 块尾列样式（逐行取；没有就退一格）
        def style_of(col):
            if col in bycol:
                mm = re.search(r'\bs="(\d+)"', bycol[col][0])
                return mm.group(1) if mm else None
            if col - 1 in bycol:
                mm = re.search(r'\bs="(\d+)"', bycol[col - 1][0])
                return mm.group(1) if mm else None
            return None
        s_in = style_of(ins_e)
        s_it = style_of(int_e)
        kept = {}
        for cn, (at, inn, sc) in bycol.items():
            if cn >= insert_at:
                f = get_formula(inn)
                if f:
                    nf = _shift_cells(f, sheet_name, same_sheet_delta)
                    inn = re.sub(r"(<f\b[^>]*>).*?(</f>)",
                                 lambda mm: mm.group(1) + nf + mm.group(2), inn, count=1, flags=re.S)
                kept[cn + ncol] = (at, inn, sc)
            else:
                kept[cn] = (at, inn, sc)
        # 当天固化值（只写跑批日；sheet 有 SKU 行才写；报告没有的 SKU → 0）
        if rnum >= 5:
            sku = ctx["sku_by_row"].get(rnum)
            if sku is not None:
                for cn, dct in ((ins_fill_col, ctx["insale"]), (int_fill_col, ctx["intra"])):
                    if cn is None:
                        continue
                    val = int(dct.get(sku, 0) or 0)
                    sref = s_in if cn == ins_fill_col else s_it
                    kept[cn] = (f' s="{sref}"' if sref else "", f"<v>{val}</v>", False)
                    n_filled += 1
        # 新列表头四件套（行1 SUM 公式+缓存值 / 行3 星期 / 行4 日期序列）
        for cn, d in ins_new:
            sref = s_in
            if rnum == 1:
                Lc = colname(cn)
                v = col_sums.get(cn, 0)
                kept[cn] = (f' s="{sref}"' if sref else "",
                            f"<f>SUM({Lc}$5:{Lc}${ctx['r1end_in']})</f><v>{v}</v>", False)
                n_added += 1
            elif rnum == 3:
                # ⚠️ 文本格必须显式 t="inlineStr"：漏写时 <is> 不是合法内容，
                # openpyxl 二次存盘（次日跑批/重跑）会把星期格丢成空单元格。
                # 2026-10-13 双跑实测：AS3/CM3 星期全空、行1 缓存值被清。
                kept[cn] = (' t="inlineStr"' + (f' s="{sref}"' if sref else ""),
                            "<is><t>%s</t></is>" % WEEK_TXT[d.weekday()], False)
                n_added += 1
            elif rnum == 4:
                kept[cn] = (f' s="{sref}"' if sref else "", f"<v>{(d - EPOCH).days}</v>", False)
                n_added += 1
        for cn, d in int_new:
            sref = s_it
            if rnum == 1:
                Lc = colname(cn)
                v = col_sums.get(cn, 0)
                kept[cn] = (f' s="{sref}"' if sref else "",
                            f"<f>SUM({Lc}$5:{Lc}${ctx['r1end_it']})</f><v>{v}</v>", False)
                n_added += 1
            elif rnum == 3:
                # ⚠️ 文本格必须显式 t="inlineStr"：漏写时 <is> 不是合法内容，
                # openpyxl 二次存盘（次日跑批/重跑）会把星期格丢成空单元格。
                # 2026-10-13 双跑实测：AS3/CM3 星期全空、行1 缓存值被清。
                kept[cn] = (' t="inlineStr"' + (f' s="{sref}"' if sref else ""),
                            "<is><t>%s</t></is>" % WEEK_TXT[d.weekday()], False)
                n_added += 1
            elif rnum == 4:
                kept[cn] = (f' s="{sref}"' if sref else "", f"<v>{(d - EPOCH).days}</v>", False)
                n_added += 1
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

    # <f ref="A1:B2"> 属性（shared/array 公式的作用域）同步平移
    def _fix_ref_attr(mm):
        # REF_ATTR_RE 组序：head, d1, c1, d1b, r1, [d2, c2, d2b, r2], tail
        # （冒号 ":" 是字面量、不设捕获组；区间不存在时 d2 为 None）
        head, d1, c1, d1b, r1, d2, c2, d2b, r2, tail = mm.groups()
        def bump_col(dollar, col, dollar2, row):
            cn = colnum(col)
            nn = same_sheet_delta(cn)
            return (dollar + colname(nn) + dollar2 + row) if nn != cn else (dollar + col + dollar2 + row)
        s = head + bump_col(d1, c1, d1b, r1)
        if d2 is not None:
            s += ":" + bump_col(d2, c2, d2b, r2)
        return s + tail
    new_xml = REF_ATTR_RE.sub(_fix_ref_attr, new_xml)

    # dimension
    dm = re.search(r'<dimension ref="([^"]*)"\s*/>', new_xml)
    if dm:
        mm = re.match(r"([A-Z]+)(\d+):([A-Z]+)(\d+)", dm.group(1))
        if mm:
            c2 = colnum(mm.group(3))
            if c2 >= int_e:
                c2 = int_e + ncol + nint
            elif c2 >= insert_at:
                c2 += ncol
            new_xml = new_xml[:dm.start()] + f'<dimension ref="{mm.group(1)}{mm.group(2)}:{colname(c2)}{mm.group(4)}" />' + new_xml[dm.end():]

    # cols：>= insert_at 平移；在售块尾条目扩到新块尾；表尾追加列克隆块尾条目
    cm = re.search(r"<cols>(.*?)</cols>", new_xml, re.S)
    if cm:
        ent = []
        for em in re.finditer(r"<col\b([^>]*)/>", cm.group(1)):
            a = em.group(1)
            ent.append([int(re.search(r'\bmin="(\d+)"', a).group(1)),
                        int(re.search(r'\bmax="(\d+)"', a).group(1)), a])
        ne = []
        for mn, mx, a in ent:
            if mx < insert_at:
                ne.append((mn, mx, a))
            elif mx == insert_at - 1:                       # 在售块尾 → 扩到新块尾
                newmx = insert_at + ncol - 1
                ne.append((mn, newmx, re.sub(r'\bmax="\d+"', f'max="{newmx}"', a)))
            else:
                n2 = mx + ncol
                a2 = re.sub(r'\bmax="\d+"', f'max="{n2}"', a)
                if mn >= insert_at:
                    m2 = mn + ncol
                    a2 = re.sub(r'\bmin="\d+"', f'min="{m2}"', a2)
                    ne.append((m2, n2, a2))
                else:
                    ne.append((mn, n2, a2))
        # 追加列（在途块新列）：若没有被已平移条目覆盖，克隆旧块尾条目并后移其后条目
        if nint:
            app_lo, app_hi = int_e + ncol + 1, int_e + ncol + nint
            covered = any(mn <= app_lo and app_hi <= mx for mn, mx, _ in ne)
            if not covered:
                clone = next((a for mn, mx, a in ent if mn <= int_e <= mx), None)
                if clone:
                    shifted = []
                    for mn, mx, a in ne:
                        if mn > int_e:
                            shifted.append((mn + nint, mx + nint, a))
                        else:
                            shifted.append((mn, mx, a))
                    shifted.append((app_lo, app_hi, clone))
                    ne = shifted
        ne.sort(key=lambda t: t[0])
        new_xml = new_xml[:cm.start()] + "<cols>" + "".join(f"<col{x}/>" for _, _, x in ne) + "</cols>" + new_xml[cm.end():]

    # mergeCells：块尾合并格扩到新块尾，其余 >= insert_at 平移
    def cross_delta(col):
        if col == int_e:
            return int_e + ncol + nint
        if col == ins_e:
            return ins_e + ncol
        return col + ncol if col >= insert_at else col

    def fix_merge(mm):
        c1, r1, c2, r2 = mm.group(1), mm.group(2), mm.group(3), mm.group(4)
        n1, n2 = colnum(c1), colnum(c2)
        n1, n2 = cross_delta(n1), cross_delta(n2)
        return f'<mergeCell ref="{colname(n1)}{r1}:{colname(n2)}{r2}" />'
    new_xml = re.sub(r'<mergeCell ref="([A-Z]+)(\d+):([A-Z]+)(\d+)"\s*/>', fix_merge, new_xml)

    # conditionalFormatting：sqref 区域同上；规则公式里的本表引用平移
    def fix_cf(mm):
        head, sq, tail = mm.group(1), mm.group(2), mm.group(3)
        parts = []
        for rg in sq.split():
            q = re.match(r"([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$", rg)
            if not q:
                parts.append(rg)
                continue
            n1 = colnum(q.group(1))
            n2 = colnum(q.group(3)) if q.group(3) else None
            n1n = cross_delta(n1)
            if n2 is None:
                parts.append(f"{colname(n1n)}{q.group(2)}")
            else:
                parts.append(f"{colname(n1n)}{q.group(2)}:{colname(cross_delta(n2))}{q.group(4)}")
        def fr(m2):
            raw = m2.group(2).replace("&lt;", "<").replace("&gt;", ">").replace("&amp;", "&")
            return m2.group(1) + _shift_cells(raw, sheet_name, same_sheet_delta).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;") + m2.group(3)
        tail2 = re.sub(r"(<formula>)(.*?)(</formula>)", fr, tail, flags=re.S)
        return head + " ".join(parts) + tail2
    new_xml = re.sub(r'(<conditionalFormatting sqref=")([^"]*)(">.*?</conditionalFormatting>)', fix_cf, new_xml, flags=re.S)
    return new_xml, dict(added=n_added, filled=n_filled)


def _bump_cross_sheet(xml, sheet_name, cross_delta):
    """其它 sheet 的公式里，指向本记录表的引用按 cross_delta 平移（含块尾范围扩展）"""
    def rep(m):
        tag, inner = m.group(1), m.group(2)
        raw = inner.replace("&lt;", "<").replace("&gt;", ">").replace("&amp;", "&")
        new = _shift_cells(raw, sheet_name, cross_delta, qualified_only=True)
        if new == raw:
            return m.group(0)
        return tag + new.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;") + "</f>"
    return re.sub(r"(<f[^>]*>)(.*?)</f>", rep, xml, flags=re.S)


def _load_shared_strings(z):
    try:
        ss = z.read("xl/sharedStrings.xml").decode("utf-8")
    except KeyError:
        return []
    out = []
    for si in re.findall(r"<si>(.*?)</si>", ss, re.S):
        out.append("".join(re.findall(r"<t[^>]*>(.*?)</t>", si, re.S)))
    return out


def fill_from_report(data, insale_field="afn-onhand-buyable-quantity"):
    """报告 {sku: {field: val}} → (在售 dict, 在途 dict)"""
    insale = {sku: (d.get(insale_field) or 0) for sku, d in data.items()}
    intra = {sku: sum((d.get(f) or 0) for f in INTRALE_FIELDS) for sku, d in data.items()}
    return insale, intra


def _refresh_today_cols(xml, ins_last_col, int_last_col, insale, intra, sku_by_row):
    """当天列已存在（块末日期 == fill_date）时，把当天列数据区刷新为固化值。
    只覆盖报告里有的 SKU（不碰手工补的行）；r1 SUM 公式补缓存值。
    返回 (new_xml, 刷新格数, r1缓存数)。"""
    m = re.search(r"<sheetData>(.*?)</sheetData>", xml, re.S)
    if not m:
        return xml, 0, 0
    body = m.group(1)
    targets = {}
    if insale:
        targets[ins_last_col] = insale
    if intra:
        targets[int_last_col] = intra
    if not targets:
        return xml, 0, 0
    sheet_skus = set(sku_by_row.values())
    sums = {cn: sum(int((dct.get(s, 0) or 0) or 0) for s in sheet_skus) for cn, dct in targets.items()}
    rows_out = []
    n_refresh = n_r1 = 0
    for rm in re.finditer(r"<row\b([^>]*?)(/>|>(.*?)</row>)", body, re.S):
        attrs, tail, inner = rm.group(1), rm.group(2), (rm.group(3) or "")
        rr = re.search(r'\br="(\d+)"', attrs)
        if tail == "/>" or not rr:
            rows_out.append(rm.group(0))
            continue
        rnum = int(rr.group(1))
        if rnum != 1 and rnum < 5:
            rows_out.append(rm.group(0))
            continue
        changed = False
        newcells = []
        for ref, at, inn, sc in parse_cells(inner):
            cn = colnum(re.match(r"[A-Z]+", ref).group(0))
            if rnum >= 5 and cn in targets:
                sku = sku_by_row.get(rnum)
                if sku is not None and sku in targets[cn]:
                    at = re.sub(r'\s+t="[^"]*"', "", at)
                    inn = "<v>%d</v>" % int(targets[cn].get(sku, 0) or 0)
                    changed = True
                    n_refresh += 1
            elif rnum == 1 and cn in targets and "<f" in inn:
                v = sums[cn]
                if re.search(r"<v>.*?</v>", inn, re.S):
                    inn = re.sub(r"<v>.*?</v>", "<v>%d</v>" % v, inn, count=1, flags=re.S)
                else:
                    inn = inn.replace("</f>", "</f><v>%d</v>" % v, 1)
                at = re.sub(r'\s+t="[^"]*"', "", at)
                changed = True
                n_r1 += 1
            newcells.append((ref, at, inn, sc))
        if not changed:
            rows_out.append(rm.group(0))
            continue
        parts = []
        for ref, at, inn, sc in newcells:
            # ⚠️ 必须补 r 属性：记录表是稀疏行（只有部分日期列有格），无 r 的 <c>
            # 会被按连续位置解析 → 整行数据错位（2026-10-13 实测 4 本簿记录表损坏）
            at = re.sub(r'\s*r="[A-Z]+\d+"', "", at)
            at = f' r="{ref}"' + at
            parts.append(f"<c{at}>{inn}</c>" if inn else f"<c{at}/>")
        a2 = re.sub(r'\s*spans="[^"]*"', "", attrs)
        rows_out.append(f'<row{a2}>{"".join(parts)}</row>')
    new_body = "".join(rows_out)
    return xml[:m.start(1)] + new_body + xml[m.end(1):], n_refresh, n_r1


def extend_record_file(path, sheet_name, until_date_str, insale=None, intra=None):
    """把 path 里 sheet_name 的在售/在途两块按天加列到 until_date（'YYYY-MM-DD'），
    只有 until_date 当天写固化值。返回统计 dict；无需加列/跳过返回 None。"""
    until = datetime.datetime.strptime(until_date_str, "%Y-%m-%d").date()
    z = zipfile.ZipFile(path)
    try:
        smap = sheet_part_map(z)
        real = resolve_sheet_name(smap, sheet_name)
        if real is None:
            print(f"    [记录] {sheet_name}: 表中无此 sheet，跳过")
            return None
        part = smap[real]
        xml = z.read(part).decode("utf-8")
        runs = find_runs(xml)
        if len(runs) < 2:
            print(f"    [记录] {real}: 行4 日期块 {len(runs)} 个（无日期列/结构异常），跳过")
            return None
        if len(runs) > 2:
            print(f"    [记录] {real}: 行4 日期块 {len(runs)} 个（>2，结构异常），跳过")
            return None
        blk_dates = _row4_dates(xml, runs)
        ins_s, ins_e = runs[0]
        int_s, int_e = runs[1]
        last_ins = blk_dates.get(0, [])[-1][1] if blk_dates.get(0) else None
        last_int = blk_dates.get(1, [])[-1][1] if blk_dates.get(1) else None
        if last_ins is None or last_int is None:
            print(f"    [记录] {real}: 日期块解析失败，跳过")
            return None

        def days_from(last):
            out, d = [], last + datetime.timedelta(days=1)
            while d <= until:
                out.append(d)
                d += datetime.timedelta(days=1)
            return out
        days_ins, days_int = days_from(last_ins), days_from(last_int)
        if not days_ins and not days_int:
            # 当天列已存在（此前由旧逻辑/手工加的，值可能是旧快照复制来的）。
            # 语义（用户定稿 2026-10-13）：当天列必须 = 当天报告的固化值 → 刷新数据区 + r1 缓存。
            if not insale and not intra:
                print(f"    [记录] {real}: 两块均已到 {until_date_str}，无需加列（未传当天值，不刷新）")
                return None
            shared = _load_shared_strings(z)
            sku_by_row = _sku_by_row(xml, shared)
            ins_last_col = blk_dates.get(0, [])[-1][0]
            int_last_col = blk_dates.get(1, [])[-1][0]
            new_xml, n_ref, n_r1 = _refresh_today_cols(
                xml, ins_last_col, int_last_col, insale or {}, intra or {}, sku_by_row)
            if not n_ref:
                print(f"    [记录] {real}: 两块均已到 {until_date_str}，无匹配 SKU，未刷新")
                return None
            parts = {i.filename: z.read(i.filename) for i in z.infolist()}
            order = [i.filename for i in z.infolist()]
            z.close()
            parts[part] = new_xml.encode("utf-8")
            tmp = path + ".rec.tmp"
            with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zo:
                for n in order:
                    zo.writestr(n, parts[n])
            zz = zipfile.ZipFile(tmp)
            bad = zz.testzip()
            zz.close()
            if bad:
                raise RuntimeError(f"zip 损坏: {bad}")
            shutil.move(tmp, path)
            print(f"    [记录] {real}: 当天列已存在 → 刷新 {until_date_str} 固化值 {n_ref} 格"
                  f" + r1 合计缓存 {n_r1} 处（旧值被当天报告覆盖）")
            return dict(sheet=real, days_ins=[], days_int=[], filled=n_ref, refreshed=True)
        shared = _load_shared_strings(z)
        sku_by_row = _sku_by_row(xml, shared)
        ctx = dict(
            ins_s=ins_s, ins_e=ins_e, int_s=int_s, int_e=int_e,
            days_ins=days_ins, days_int=days_int, fill_date=until,
            insale=insale or {}, intra=intra or {},
            sku_by_row=sku_by_row,
            r1end_in=_r1end(xml, ins_s, ins_e), r1end_it=_r1end(xml, int_s, int_e),
        )
        new_xml, st = _transform(xml, real, ctx)
        ncol, nint = len(days_ins), len(days_int)

        def cross_delta(col):
            if col == int_e:
                return int_e + ncol + nint
            if col == ins_e:
                return ins_e + ncol
            return col + ncol if col >= ins_e + 1 else col
        # 跨表公式平移（数据透视图 SUMPRODUCT 等）
        others = {}
        for n in z.namelist():
            if n == part or not n.startswith("xl/worksheets/") or not n.endswith(".xml"):
                continue
            raw = z.read(n).decode("utf-8")
            if real in raw:
                out = _bump_cross_sheet(raw, real, cross_delta)
                if out != raw:
                    others[n] = out.encode("utf-8")
        parts = {i.filename: z.read(i.filename) for i in z.infolist()}
        order = [i.filename for i in z.infolist()]
    finally:
        z.close()
    parts[part] = new_xml.encode("utf-8")
    parts.update(others)
    tmp = path + ".rec.tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zo:
        for n in order:
            zo.writestr(n, parts[n])
    zz = zipfile.ZipFile(tmp)
    bad = zz.testzip()
    zz.close()
    if bad:
        raise RuntimeError(f"zip 损坏: {bad}")
    shutil.move(tmp, path)
    msg_ins = ",".join(d.strftime("%m-%d") for d in days_ins) or "-"
    msg_int = ",".join(d.strftime("%m-%d") for d in days_int) or "-"
    print(f"    [记录] {real}: 在售加列[{msg_ins}] 在途加列[{msg_int}]"
          f"（{until_date_str} 写固化值 {st['filled']} 格，中间日留空；"
          f"跨表公式平移 {len(others)} 个 sheet）")
    return dict(sheet=real, days_ins=days_ins, days_int=days_int, **st)
