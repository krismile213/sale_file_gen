# -*- coding: utf-8 -*-
"""给「每日库存分析表」日列粉色规则加上界 MIN(锚点,TODAY()+1)。

为什么不用 openpyxl 重写整个工作簿：
    openpyxl 存盘会丢掉所有公式缓存值（日销列是 SUMIFS），导致 data_only 读出来全是 None，
    归档副本/脚本预览会看到空表。所以这里**只改 sheet XML 里那一小段公式字符串**，
    zip 其余条目原样搬运 —— 缓存值、透视表缓存、富数据全部保留。

改什么：
    日列规则尾部 `,{L1}$5<{锚点}`  →  `,{L1}$5<MIN({锚点},TODAY()+1)`
    · 锚点正常（= 快照日）时 MIN 取锚点 → 与原公式完全等价，染色结果一格不变
    · 锚点被误设成未来 → MIN 封顶到今天，不会把还没到的日期标红
    · 打开历史归档时 TODAY() 变大 → MIN 仍取锚点 → 历史染色不受影响

用法：
    python patch_pink_ceiling_20261023.py            # 干跑
    python patch_pink_ceiling_20261023.py --apply
"""
import os, re, sys, shutil, zipfile, traceback

HERE = os.path.dirname(os.path.abspath(__file__))
DSR = HERE
for _ in range(6):
    if os.path.isdir(os.path.join(DSR, "1.每日销售数据表")):
        break
    DSR = os.path.dirname(DSR)
BOOKS_DIR = os.path.join(DSR, "1.每日销售数据表")
BK = os.path.join(DSR, ".workbuddy", "backup_pink_ceiling_20261023")
SHEET_SUFFIX = "每日库存分析表"
APPLY = "--apply" in sys.argv

# 日列上界： ,{L1}$5&lt;{锚点}   （XML 转义后 &lt;）
PAT = re.compile(r",([A-Z]{1,2}\$5)&lt;([^,&]+?)(?=[,)])")


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
        t = rmap.get(rid.group(1), "")
        t = t.lstrip("/")
        if not t.startswith("xl/"):
            t = "xl/" + t
        if t in z.namelist():
            out.append((nm.group(1), t))
    return out


def patch_xml(xml):
    """返回 (新xml, 替换数)"""
    n = 0

    def repl(m):
        nonlocal n
        col, anchor = m.group(1), m.group(2)
        if "MIN(" in anchor:
            return m.group(0)          # 幂等
        n += 1
        return f",{col}&lt;MIN({anchor},TODAY()+1)"

    # 只处理含 LOOKUP 的 cfRule 里的 formula
    def fix_rule(m):
        body = m.group(0)
        if "LOOKUP" not in body:
            return body
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
                targets = [(s, p) for s, p in sheet_files(z) if s.strip().endswith(SHEET_SUFFIX)]
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
                      f"{'待改' if n else '已是最新/无日列规则'}")
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
    print(f"涉及 {tot_b} 本 / {tot_s} 张表；{'已落盘 ' + str(tot_f) + ' 本' if APPLY else '干跑未落盘'}")


if __name__ == "__main__":
    main()
