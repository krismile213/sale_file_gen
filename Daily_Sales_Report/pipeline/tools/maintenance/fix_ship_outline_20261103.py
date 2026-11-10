# -*- coding: utf-8 -*-
"""
恢复 13 本分析表「货件列区域」的列组合(outline)——11-02 换月链路把 <col> 的
outlineLevel/hidden 剥掉了（点 1/2 组合按钮消失）。

原理：
  换月前备份 backup_mixed_202611 里每个分析表的货件列 <col> 定义完好
  （E 列 hidden、H + I..AC outlineLevel=1 等，区间式、min/max 属性在尾部）。
  货件列区域（日列首列之前）的列结构自换月前至今未变（行1货件号段一致），
  因此可把 mixed 备份里 max<日列首 的 outline/hidden 区间原样套回当前簿。

做法（纯 XML 手术，幂等，不动单元格/公式/CF/日列区）：
  对每个分析表：
    1) 从 mixed 备份取目标区间 [(mn,mx,attrs_outline_hidden), ...]（仅 mx<日列首）
    2) 当前簿 <cols> 逐元素与目标区间求交：交集段补 outlineLevel/hidden/collapsed，
       其余段原样；目标列当前无定义时新增元素
    3) 日列区（>=日列首）完全不碰

用法：
  python fix_ship_outline_20261103.py            # 预演
  python fix_ship_outline_20261103.py --apply    # 应用
"""
import sys, os, re, glob, zipfile, shutil
sys.stdout.reconfigure(encoding="utf-8")

BASE = r"C:/Users/dell/Desktop/work-main/sale_file_gen/Daily_Sales_Report"
WORK = os.path.join(BASE, "1.每日销售数据表")
BAK_MIXED = os.path.join(BASE, ".workbuddy", "backup_mixed_202611")
BKDIR = os.path.join(BASE, ".workbuddy", "backup_ship_outline_20261103")

COL_EL = re.compile(r"<col\b([^>]*?)/>")
COLS_BLK = re.compile(r"<cols>(.*?)</cols>", re.S)


def attr(el, k, default=None):
    m = re.search(r'\b%s="([^"]*)"' % k, el)
    return m.group(1) if m else default


def parse_col_elems(block):
    """[(mn, mx, whole_elem_text, attrs_text)]"""
    out = []
    for cm in COL_EL.finditer(block):
        a = cm.group(1)
        try:
            mn, mx = int(attr(a, "min")), int(attr(a, "max"))
        except (TypeError, ValueError):
            continue
        out.append((mn, mx, cm.group(0), a))
    return out


def sheet_xml_of(z, sheet_name):
    wbxml = z.read("xl/workbook.xml").decode("utf-8")
    rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    rmap = {}
    for rel in re.findall(r"<Relationship\b[^>]*/>", rels):
        i = re.search(r'Id="(rId\d+)"', rel)
        t = re.search(r'Target="([^"]+)"', rel)
        if i and t:
            rmap[i.group(1)] = t.group(1)
    for nm, rid in re.findall(r'<sheet name="([^"]+)"[^>]*r:id="(rId\d+)"', wbxml):
        if nm == sheet_name:
            t = rmap[rid].lstrip("/")
            return z.read(t if t.startswith("xl/") else "xl/" + t).decode("utf-8")
    return None


def first_daycol_idx(cur_xml):
    """行5 第一个 datetime 序列号列号（40000<n<60000）。"""
    m = re.search(r'<row[^>]*r="5"[^>]*>(.*?)</row>', cur_xml, re.S)
    if not m:
        return None
    best = None
    for cm in re.finditer(r'<c\b([^>]*)>(.*?)</c>', m.group(1), re.S):
        rm = re.search(r'r="([A-Z]+)5"', cm.group(1))
        vm = re.search(r"<v>([^<]*)</v>", cm.group(2))
        if not (rm and vm):
            continue
        try:
            n = float(vm.group(1))
        except ValueError:
            continue
        if 40000 < n < 60000:
            ci = 0
            for ch in rm.group(1):
                ci = ci * 26 + (ord(ch) - 64)
            best = ci if best is None else min(best, ci)
    return best


def targets_from_mixed(mixed_xml, first_day):
    """货件区( max<first_day )中带 outlineLevel/hidden 的 col -> [(mn,mx,attrs)]"""
    m = COLS_BLK.search(mixed_xml)
    if not m:
        return []
    tg = []
    for mn, mx, _, a in parse_col_elems(m.group(1)):
        if mx >= first_day:
            continue
        keep = []
        for k in ("outlineLevel", "hidden", "collapsed"):
            v = attr(a, k)
            if v is not None:
                keep.append('%s="%s"' % (k, v))
        if keep:
            tg.append((mn, mx, " ".join(keep)))
    return tg


def merge_once(cur_cols_body, targets, first_day):
    """把 targets 合并进当前 <cols> body（仅货件区 <first_day），返回 (new_body, n_changed)。幂等。"""
    def strip_ol(a):
        a = re.sub(r'\b(outlineLevel|hidden|collapsed|min|max)="[^"]*"\s*', "", a)
        return re.sub(r"\s+", " ", a).strip()

    elems = parse_col_elems(cur_cols_body)
    # 覆盖表: 仅货件区列（< first_day）
    cover = {}
    for mn, mx, _, a in elems:
        base = strip_ol(a)
        for c in range(mn, min(mx, first_day - 1) + 1):
            cover[c] = base
    for mn, mx, add in targets:
        for c in range(mn, min(mx, first_day - 1) + 1):
            base = cover.get(c) or 'width="5.5" customWidth="1" style="3"'
            cover[c] = (add + " " + base).strip()
    # 重新按区间折叠输出（日列区 attrs 原样保留）
    new = []
    changed = 0
    for mn, mx, whole, a in elems:
        segs = []
        cur_attrs = None
        run = None
        for c in range(mn, mx + 1):
            if c >= first_day:
                at = re.sub(r'\b(min|max)="[^"]*"\s*', "", a)  # 日列区原样（含 outline/hidden）
                at = re.sub(r"\s+", " ", at).strip()
            else:
                at = cover.get(c, strip_ol(a))
            if at != cur_attrs:
                if run is not None:
                    segs.append((run, c - 1, cur_attrs))
                cur_attrs = at
                run = c
        segs.append((run, mx, cur_attrs))
        for s, e, at in segs:
            txt = '<col %s min="%d" max="%d"/>' % (at, s, e)
            if s == mn and e == mx and at == re.sub(r'\b(min|max)="[^"]*"\s*', "", a).strip():
                txt = whole  # 完全无变化保持原文
            if txt != whole or (s, e) != (mn, mx):
                changed += 1
            new.append(txt)
    # 目标区间里当前簿完全没覆盖的列 -> 新增
    have = set()
    for mn, mx, _, _ in elems:
        have.update(range(mn, mx + 1))
    for mn, mx, add in targets:
        if not any(c in have for c in range(mn, mx + 1)):
            new.append('<col %s min="%d" max="%d"/>' % (add, mn, mx))
            changed += 1
    body = "".join(new)
    return body, changed


def process(path, apply):
    name = os.path.basename(path)
    with zipfile.ZipFile(path) as z:
        payload = {}
        with zipfile.ZipFile(os.path.join(BAK_MIXED, name)) if os.path.exists(
                os.path.join(BAK_MIXED, name)) else _null() as zm:
            if zm is None:
                print("  !! mixed 备份缺簿:", name)
                return
            from core.record_extend import sheet_part_map
            smap = sheet_part_map(z)
            smap_m = sheet_part_map(zm)
            for sn, part in smap.items():
                if "每日库存分析表" not in sn:
                    continue
                cur = z.read(part).decode("utf-8")
                if sn not in smap_m:
                    print("  !! mixed 缺表:", name, sn)
                    continue
                old = zm.read(smap_m[sn]).decode("utf-8")
                fd = first_daycol_idx(cur)
                if not fd:
                    continue
                tg = targets_from_mixed(old, fd)
                if not tg:
                    continue
                m = COLS_BLK.search(cur)
                if not m:
                    continue
                new_body, nchg = merge_once(m.group(1), tg, fd)
                if nchg:
                    payload[part] = cur[:m.start()] + "<cols>" + new_body + "</cols>" + cur[m.end():]
                    print("  %s | %s: 目标区间=%s 改动段=%d" %
                          (name, sn, [(a, b) for a, b, _ in tg], nchg))
        if not payload:
            print("  ✅ 已就绪:", name)
            return
        if not apply:
            print("  (预演) 未写回:", name)
            return
        os.makedirs(BKDIR, exist_ok=True)
        bk = os.path.join(BKDIR, name)
        if not os.path.exists(bk):
            shutil.copy2(path, bk)
        tmp = path + ".ol.tmp"
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as out:
            for item in z.infolist():
                d = payload.get(item.filename)
                out.writestr(item, d.encode("utf-8") if d is not None else z.read(item.filename))
    # 沙箱安全替换：先挪旧再顶上
    try:
        bak_old = path + ".olold"
        if os.path.exists(bak_old):
            os.remove(bak_old)
        os.replace(path, bak_old)
        os.replace(tmp, path)
        os.remove(bak_old)
    except OSError:
        shutil.copyfile(tmp, path)
        os.remove(tmp)
    with zipfile.ZipFile(path) as z:
        print("  ✅ 写回:", name, "zip:", "OK" if z.testzip() is None else "损坏")


class _null:
    def __enter__(self):
        return None

    def __exit__(self, *a):
        return False


if __name__ == "__main__":
    apply = "--apply" in sys.argv
    sys.path.insert(0, os.path.join(BASE, "pipeline"))
    from core.record_extend import sheet_part_map  # noqa 提前失败
    files = sorted(glob.glob(os.path.join(WORK, "*库存销售表.xlsx")))
    print(("应用" if apply else "预演"), len(files), "本")
    for f in files:
        process(f, apply)
