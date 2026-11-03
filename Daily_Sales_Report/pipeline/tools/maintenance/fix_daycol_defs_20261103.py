# -*- coding: utf-8 -*-
"""
给 17 个分析表日列「逐列」重设 <col>：宽度沿用原定义，按日期设 outlineLevel/hidden。

背景：
  1) rebuild_month_day_cols_20261103 插入 11 月日列后 <cols> 未同步扩展 → 24 列无定义；
  2) 首版 fix_daycol_defs 只补「无定义」的列，已定义但落在 11-02 之后的列（74~78）
     仍不会被折叠 → 打开会看到 11/2~11/6 露在外面。

做法：把日列按「单列」重写（非日列区间保持原样），逐列判定：
      日期 >= 2026-11-02 → outlineLevel="1" hidden="1"
      否则              → 不带 outline/hidden（展开）
      宽度/style/customWidth 沿用原来覆盖该列的 <col>。
纯 XML 手术，幂等。
"""
import zipfile, re, shutil, os, sys, datetime, glob
sys.path.insert(0, r"C:/Users/dell/Desktop/work-main/sale_file_gen/Daily_Sales_Report/pipeline")
from core.record_extend import sheet_part_map

DSR = r"C:/Users/dell/Desktop/work-main/sale_file_gen/Daily_Sales_Report/1.每日销售数据表"
EPOCH = datetime.date(1899, 12, 30)
HIDE_FROM = datetime.date(2026, 11, 2)

COL_RE = re.compile(r"<col\b([^>]*?)/>")


def day_map(xml):
    m = re.search(r'<row[^>]*r="5"[^>]*>(.*?)</row>', xml, re.S)
    if not m:
        return {}
    out = {}
    for cm in re.finditer(r'<c\b([^>]*)>(.*?)</c>', m.group(1), re.S):
        rm = re.search(r'r="([A-Z]+)5"', cm.group(1))
        vm = re.search(r"<v>([^<]*)</v>", cm.group(2))
        if not (rm and vm):
            continue
        try:
            n = float(vm.group(1))
        except Exception:
            continue
        if 40000 < n < 60000:
            ci = 0
            for ch in rm.group(1):
                ci = ci * 26 + (ord(ch) - 64)
            out[ci] = EPOCH + datetime.timedelta(days=int(n))
    return out


def parse_cols(block):
    out = []
    for cm in COL_RE.finditer(block):
        a = cm.group(1)
        g = lambda k: (re.search(r'\b%s="([^"]*)"' % k, a) or [None, None])[1]
        try:
            mn, mx = int(g("min")), int(g("max"))
        except (TypeError, ValueError):
            continue
        out.append({"min": mn, "max": mx, "attrs": a})
    return out


def base_attrs(entries, c):
    """取覆盖 c 的原始 attrs（去掉 min/max/outline/hidden/collapsed）"""
    for e in entries:
        if e["min"] <= c <= e["max"]:
            a = e["attrs"]
            a = re.sub(r'\b(min|max|outlineLevel|hidden|collapsed)="[^"]*"', "", a)
            a = re.sub(r"\s+", " ", a).strip()
            if "width=" not in a:
                a = 'width="5.5" customWidth="1" ' + a
            return a
    return 'width="5.5" customWidth="1" style="3"'


def rebuild(xml, hide_from=None):
    m = re.search(r"<cols>.*?</cols>", xml, re.S)
    if not m:
        return xml, 0
    if hide_from is None:
        hide_from = HIDE_FROM
    block = m.group(0)
    entries = parse_cols(block)
    days = day_map(xml)
    if not days:
        return xml, 0
    dset = set(days)

    # 非日列区间：从原区间里剔掉日列
    def strip_mm(a):
        a = re.sub(r'\b(min|max)="[^"]*"', "", a)
        return re.sub(r"\s+", " ", a).strip()

    keep = []
    for e in entries:
        a, b = e["min"], e["max"]
        run = None
        for c in range(a, b + 1):
            if c in dset:
                if run is not None:
                    keep.append((run, c - 1, strip_mm(e["attrs"])))
                    run = None
            else:
                if run is None:
                    run = c
        if run is not None:
            keep.append((run, b, strip_mm(e["attrs"])))

    # 日列逐列
    percol = []
    for c in sorted(dset):
        d = days[c]
        a = base_attrs(entries, c)
        if d >= HIDE_FROM:
            a += ' outlineLevel="1" hidden="1"'
        percol.append((c, c, a))

    allc = [(a, b, at) for a, b, at in keep] + percol
    allc.sort(key=lambda x: x[0])
    body = "".join('<col min="%d" max="%d" %s/>' % (a, b, at.strip()) for a, b, at in allc)
    new = "<cols>%s</cols>" % body
    return xml.replace(block, new, 1), len(percol)


def process(path, apply=False, hide_from=None):
    tmp = path + ".tmp3"
    payload = {}
    with zipfile.ZipFile(path) as z:
        smap = sheet_part_map(z)
        for n, p in smap.items():
            if "每日库存分析表" not in n:
                continue
            x = z.read(p).decode("utf-8")
            x2, nd = rebuild(x, hide_from)
            if x2 != x:
                payload[p] = x2
        if not payload:
            print("  ✅ 已就绪:", os.path.basename(path))
            return
        print("  ⚠️ 重设:", os.path.basename(path), f"{len(payload)} 张表")
        if not apply:
            return
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as out:
            for item in z.infolist():
                d = payload.get(item.filename)
                out.writestr(item, d.encode("utf-8") if d is not None else z.read(item.filename))
    BK = os.path.join(os.path.dirname(DSR), ".workbuddy", "backup_coldefs_20261103")
    os.makedirs(BK, exist_ok=True)
    shutil.copy2(path, os.path.join(BK, os.path.basename(path)))
    try:
        os.replace(tmp, path)
    except OSError:
        shutil.copyfile(tmp, path)
        os.remove(tmp)
    with zipfile.ZipFile(path) as z:
        print("      zip:", "OK" if z.testzip() is None else "损坏")


if __name__ == "__main__":
    apply = "--apply" in sys.argv
    files = sorted(glob.glob(os.path.join(DSR, "*库存销售表.xlsx")))
    print(f"{'应用' if apply else '预演'} {len(files)} 本")
    for f in files:
        process(f, apply)
