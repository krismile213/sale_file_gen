# -*- coding: utf-8 -*-
"""方案A：年总销量列 `IF(SUM(月块)=0,X,SUM+X)` → `SUM(月块)`（XML 手术，不重存）

- 只替换公式文本，保留 <f> 原有属性（shared/array 都不动）
  · shared 主格改了，从属格按相对引用自动平移，语义仍正确
- 目标 <v> 缓存若为非空数值，一并清空，避免打开时读到旧值
- 备份 → .workbuddy/backup_year_total_20261102/
"""
import zipfile, re, os, glob, shutil, sys, datetime

BASE = r"C:\Users\dell\Desktop\work-main\sale_file_gen\Daily_Sales_Report\1.每日销售数据表"
BK = r"C:\Users\dell\Desktop\work-main\sale_file_gen\Daily_Sales_Report\.workbuddy\backup_year_total_20261102"
APPLY = "--apply" in sys.argv

# <f 属性>IF(SUM(C1:r)=0,Xr,SUM(C1:r)+Xr)</f>
F_RE = re.compile(
    r'<f([^>]*)>(IF\(SUM\(([A-Z]+)(\d+):([A-Z]+)\4\)=0,([A-Z]+)\4,SUM\(\3\4:\5\4\)\+\6\4\))</f>')
# 整个 <c ...> ... </c>，用来清 <v>
C_RE = re.compile(r'<c r="([A-Z]+)(\d+)"([^>]*)>(.*?)</c>', re.S)


def fix_xml(xml):
    """返回 (新xml, 改公式数, 清缓存数)"""
    # 先收集被改 shared 主格的 si（从属格要一起清缓存）
    si_set = set()
    for m in F_RE.finditer(xml):
        sm = re.search(r'si="(\d+)"', m.group(1))
        if sm:
            si_set.add(sm.group(1))

    changed = [0]

    def rep(m):
        attrs, c1, r1, c2 = m.group(1), m.group(3), m.group(4), m.group(5)
        changed[0] += 1
        return f"<f{attrs}>SUM({c1}{r1}:{c2}{r1})</f>"

    xml2 = F_RE.sub(rep, xml)

    cleared = [0]

    def clr(m):
        ref, row, attrs, body = m.group(1), m.group(2), m.group(3), m.group(4)
        hit = bool(re.search(r"<f[^>]*>SUM\([A-Z]+\d+:[A-Z]+\d+\)</f>", body))
        if not hit and si_set:
            sm = re.search(r'<f t="shared"[^>]*si="(\d+)"\s*/>', body)
            if sm and sm.group(1) in si_set:
                hit = True
        if not hit:
            return m.group(0)
        if re.search(r"<v>[^<]*</v>", body) and not re.search(r"<v\s*/>", body):
            cleared[0] += 1
            body2 = re.sub(r"<v>[^<]*</v>", "<v/>", body)
            return f'<c r="{ref}{row}"{attrs}>{body2}</c>'
        return m.group(0)

    xml2 = C_RE.sub(clr, xml2)
    return xml2, changed[0], cleared[0]


def main():
    files = sorted(glob.glob(os.path.join(BASE, "*库存销售表.xlsx")))
    print(f"共 {len(files)} 本 | 模式={'APPLY' if APPLY else 'DRY'}")
    if APPLY:
        os.makedirs(BK, exist_ok=True)

    grand_f = grand_c = 0
    for f in files:
        name = os.path.basename(f)
        if APPLY:
            shutil.copy2(f, os.path.join(BK, name))

        zin = zipfile.ZipFile(f, "r")
        items = [(i, zin.read(i.filename)) for i in zin.infolist()]
        zin.close()

        nf = nc = 0
        out = []
        for info, data in items:
            if re.match(r"xl/worksheets/sheet\d+\.xml$", info.filename):
                txt = data.decode("utf-8")
                if "IF(SUM(" in txt:
                    txt2, a, b = fix_xml(txt)
                    data = txt2.encode("utf-8")
                    nf += a
                    nc += b
            out.append((info, data))

        if nf == 0:
            print(f"  {name:<32} 无改动")
            continue
        print(f"  {name:<32} 改公式 {nf:>4} | 清缓存 {nc:>4}")
        grand_f += nf
        grand_c += nc

        if APPLY:
            tmp = f + ".tmp"
            with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zo:
                for info, data in out:
                    zo.writestr(info, data)
            os.replace(tmp, f)

    print(f"合计：改公式 {grand_f} 处，清缓存 {grand_c} 处")
    if APPLY:
        print("备份 →", BK)


if __name__ == "__main__":
    main()
