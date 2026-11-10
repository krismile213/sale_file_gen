# -*- coding: utf-8 -*-
"""新品标黄规则的 C 列状态码 → D 列状态（2026-11-07）

背景
----
2026-11-07 「以商品资料为准」全量覆盖了分析表 C 列（商品分类），
本来 C 列里混着一批**状态码**：XB / XE / XC（与「待停售」同义 = 新品不做了）。
覆盖后这些码被换成资料的真实评级（B / E / C…），于是原来靠
    AND($F7=0,$G7=0,$D7<>"停售",$C7<>"待停售",$C7<>"XB",$C7<>"XE",$C7<>"XC",$XL7="")
排除掉的行，现在条件全部满足 → 会被误标成「新品黄」。

改法
----
C 列从今以后**只表示评级**，不再承载状态语义。新品判定改看 D 列状态：
    保留  $D7<>"停售"            （已停售不算新品）
    删除  $C7<>"待停售"          （改由 D 列兜，资料停售的 D 已经写「停售」）
    删除  $C7<>"XB"/"XE"/"XC"    （这三个码已不复存在）

新公式：
    AND($F7=0,$G7=0,$D7<>"停售",$X7="")

用法：
    python fix_new_product_rule_c2d_20261107.py            # 预览
    python fix_new_product_rule_c2d_20261107.py --apply    # 落盘（先备份）
"""
import argparse
import glob
import os
import re
import shutil
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
D = os.path.abspath(os.path.join(HERE, "..", "..", "..", "1.每日销售数据表"))

# 待删除的条件片段。
# ⚠️ 引号的转义形态**各簿不一致**：本批文件的 formula 里引号是**裸** `"XB"`，
#    不是 `&quot;XB&quot;`（openpyxl 存过的和 WPS 存过的写法不同）→ 两种都要试。
DROP = tuple(x for code in ("待停售", "XB", "XE", "XC")
             for x in (f'$C7&lt;&gt;"{code}",', f'$C7&lt;&gt;&quot;{code}&quot;,',
                       f'$C7<>"{code}",'))
# 只抓「新品规则」那种 formula：以 AND($F7=0,$G7=0 开头、到 </formula> 结束
RULE_RE = re.compile(
    r"(<formula>)(AND\(\$F7=0,\$G7=0,[^<]*?)(</formula>)")


def sheet_map(z):
    """sheet 名 → zip 内路径。

    ⚠️ rels 里 `Type/Target/Id` 三个属性的**顺序各簿不固定**（有的 Target 在前），
    不能按固定正则一次抓完，必须逐个 Relationship 元素解析属性。
    """
    rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    rid2t = {}
    for m in re.finditer(r"<Relationship\b([^>]*)/?>", rels):
        attrs = dict(re.findall(r'(\w+)="([^"]*)"', m.group(1)))
        rid, tgt = attrs.get("Id"), attrs.get("Target")
        if rid and tgt:
            rid2t[rid] = tgt.lstrip("/")
    wb = z.read("xl/workbook.xml").decode("utf-8")
    out = {}
    for m in re.finditer(r"<sheet\b([^>]*)/?>", wb):
        attrs = dict(re.findall(r'([\w:]+)="([^"]*)"', m.group(1)))
        name = attrs.get("name")
        rid = attrs.get("r:id")
        if name and rid and rid in rid2t:
            p = rid2t[rid]
            out[name] = p if p.startswith("xl/") else "xl/" + p
    return out


def patch_formula(f):
    """把一条新品规则里的 C 列状态码条件去掉。返回 (新公式, 是否改动)"""
    if "$C7" not in f:
        return f, False
    g = f
    for d in DROP:
        g = g.replace(d, "")
    # 清掉可能残留的连续逗号
    while ",," in g:
        g = g.replace(",,", ",")
    g = g.replace(",)", ")")
    return g, (g != f)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--apply", action="store_true")
    args = ap.parse_args()

    books = [p for p in sorted(glob.glob(os.path.join(D, "*.xlsx")))
             if not os.path.basename(p).startswith("~$")]
    total = 0
    for p in books:
        bn = os.path.basename(p)
        with zipfile.ZipFile(p) as z:
            names = z.namelist()
            smap = sheet_map(z)
            data = {n: z.read(n) for n in names}

        hit_sheets = []
        for sn, path in smap.items():
            if path not in data:
                continue
            x = data[path].decode("utf-8")
            if "$C7" not in x:
                continue
            nchg = 0
            def repl(m):
                nonlocal nchg
                new, chg = patch_formula(m.group(2))
                if chg:
                    nchg += 1
                    return m.group(1) + new + m.group(3)
                return m.group(0)
            y = RULE_RE.sub(repl, x)
            if nchg:
                data[path] = y.encode("utf-8")
                hit_sheets.append((sn, nchg))
                total += nchg

        if not hit_sheets:
            continue
        print(f"{bn}: " + ", ".join(f"{s}×{n}" for s, n in hit_sheets))
        if args.apply:
            bak = p + ".bak_cf_c2d"
            if not os.path.exists(bak):
                shutil.copy2(p, bak)
            # ⚠️ 不能用 os.replace(tmp, p)：沙箱/驱动的 rename 钩子会报 WinError 5
            #    （文件本身可写，实测 File.Open ReadWrite Share.None 能成功）。
            #    改为「原地重写」：以 wb 模式打开目标文件直接写 zip。
            with zipfile.ZipFile(p, "w", zipfile.ZIP_DEFLATED) as zo:
                for n in names:
                    zo.writestr(n, data[n])
            print(f"   已落盘（备份 {os.path.basename(bak)}）")

    print(f"\n共改动 {total} 条规则" + ("（已落盘）" if args.apply else "（预览，未落盘）"))


if __name__ == "__main__":
    main()
