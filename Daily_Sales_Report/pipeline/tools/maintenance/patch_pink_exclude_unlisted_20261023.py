# -*- coding: utf-8 -*-
"""给「每日库存分析表」的粉色断货规则加 D 列排除：状态=「未上架」的行不标断货色。

背景（2026-10-23 用户拍板）：
    新品补行后，零库存的「未上架」SKU 会被粉色规则（可售≤0 且无在途 → 玫红）标成断货，
    但没上架不算断货，会造成误读。所以在行级基础条件里加一条 `$D{行}<>"未上架"`。

改什么：
    AND($A7&lt;&gt;"",N($F7)&lt;=0,...)  →  AND($A7&lt;&gt;"",$D7&lt;&gt;"未上架",N($F7)&lt;=0,...)
    · 只动含 LOOKUP 的 cfRule（粉色断货规则），其它条件格式不碰
    · 幂等：已含「未上架」的规则原样跳过
    · 状态列 D 已确认全 12 簿一致（A=SKU B=型号 C=分类 D=在售/未上架）

为什么走 XML 而不是 openpyxl：
    openpyxl 存盘会清空所有公式缓存值（日销列 SUMIFS），data_only 读出来全 None，
    预览/归档会看到空表。这里只改 XML 里的一小段公式字符串，其余条目原样搬运。

用法：
    python patch_pink_exclude_unlisted_20261023.py           # 干跑
    python patch_pink_exclude_unlisted_20261023.py --apply
"""
import os, re, sys, zipfile, traceback

HERE = os.path.dirname(os.path.abspath(__file__))
DSR = HERE
for _ in range(6):
    if os.path.isdir(os.path.join(DSR, "1.每日销售数据表")):
        break
    DSR = os.path.dirname(DSR)
BOOKS_DIR = os.path.join(DSR, "1.每日销售数据表")
SHEET_SUFFIX = "每日库存分析表"
APPLY = "--apply" in sys.argv

# AND($A7&lt;&gt;"",   —— 注意 XML 里 < > 是 &lt; &gt;，" 原样
PAT = re.compile(r'AND\(\$A(\d+)&lt;&gt;"",')


def sheet_files(z):
    """返回 [(sheetname, xml_path)]"""
    wbx = z.read("xl/workbook.xml").decode("utf-8")
    rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    rmap = {}
    for tag in re.findall(r"<Relationship\b[^>]*/?>", rels):
        i = re.search(r'\bId="([^"]+)"', tag)
        t = re.search(r'\bTarget="([^"]+)"', tag)
        if i and t:
            rmap[i.group(1)] = t.group(1)
    out = []
    for m in re.finditer(r"<sheet\b[^>]*/?>", wbx):
        tag = m.group(0)
        nm = re.search(r'\bname="([^"]+)"', tag)
        rid = re.search(r'\br:id="([^"]+)"', tag)
        if not (nm and rid):
            continue
        t = rmap.get(rid.group(1), "").lstrip("/")
        if not t.startswith("xl/"):
            t = "xl/" + t
        if t in z.namelist():
            out.append((nm.group(1), t))
    return out


def patch_xml(xml):
    """返回 (新xml, 替换数)"""
    n = 0

    def fix_rule(m):
        nonlocal n
        body = m.group(0)
        if "LOOKUP" not in body:
            return body
        if "未上架" in body:
            return body                     # 幂等

        def repl(mm):
            nonlocal n
            n += 1
            return f'AND($A{mm.group(1)}&lt;&gt;"",$D{mm.group(1)}&lt;&gt;"未上架",'

        return PAT.sub(repl, body)

    new = re.sub(r"<cfRule\b.*?</cfRule>", fix_rule, xml, flags=re.S)
    return new, n


def main():
    print(f"{'工作簿':<32}{'sheet':<26}{'改':<4}结果")
    print("-" * 78)
    tot_b = tot_s = tot_f = 0
    for fn in sorted(os.listdir(BOOKS_DIR)):
        if not fn.endswith("库存销售表.xlsx") or fn.startswith("~$"):
            continue
        cur = os.path.join(BOOKS_DIR, fn)
        try:
            # ⚠️ 读句柄必须先关闭，否则 Windows 下 os.replace 一定 PermissionError
            patches, rows = {}, []
            with zipfile.ZipFile(cur) as z:
                targets = [(s, p) for s, p in sheet_files(z)
                           if s.strip().endswith(SHEET_SUFFIX)]
                if not targets:
                    print(f"{fn[:30]:<32}{'-':<26}{0:<4}无分析表")
                    continue
                for s, p in targets:
                    xml = z.read(p).decode("utf-8")
                    new, n = patch_xml(xml)
                    rows.append((s, n))
                    if n:
                        patches[p] = new.encode("utf-8")

            for s, n in rows:
                print(f"{fn[:30]:<32}{s[:24]:<26}{n:<4}"
                      f"{'待改' if n else '已是最新'}")
            if not patches:
                continue
            tot_b += 1
            tot_s += len(patches)
            if not APPLY:
                continue

            tmp = cur + ".tmp"
            with zipfile.ZipFile(cur) as zin, \
                 zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    data = patches.get(item.filename) or zin.read(item.filename)
                    zi = zipfile.ZipInfo(item.filename, date_time=item.date_time)
                    zi.compress_type = item.compress_type
                    zi.external_attr = item.external_attr
                    zi.internal_attr = item.internal_attr
                    zi.create_system = item.create_system
                    zout.writestr(zi, data)
            with zipfile.ZipFile(tmp) as zz:
                bad = zz.testzip()
            if bad:
                raise RuntimeError(f"zip 损坏: {bad}")
            os.replace(tmp, cur)
            tot_f += 1
        except Exception as e:
            print(f"{fn[:30]:<32}{'!!':<26}{'':<4}{e!r}")
            traceback.print_exc()
    print("-" * 78)
    print(f"涉及 {tot_b} 本 / {tot_s} 张表；"
          f"{'已落盘 ' + str(tot_f) + ' 本' if APPLY else '干跑未落盘'}")


if __name__ == "__main__":
    main()
