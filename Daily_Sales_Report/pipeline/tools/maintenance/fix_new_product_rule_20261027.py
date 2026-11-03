# -*- coding: utf-8 -*-
"""
修正 patch_20261027_layout_cf.py 引入的两个问题（2026-10-27）：

1. 新品规则配错 dxfId —— find_yellow_dxf() 用
   `<cfRule[^>]*dxfId="(\\d+)"[^>]*>(.*?)</cfRule>` 抓「保底」规则，
   遇到自闭合 `<cfRule .../>` 时 `(.*?)` 会跨过后续规则，
   把别的规则的 dxfId 配上「保底」文本，结果 15/17 张表拿到了
   浅红 FFFFC7CE（SKU 重复色），不是正黄 FFFFFF00。
   → 本脚本按 dxfs 里 rgb=FFFFFF00 反查真正的黄 dxfId 再改回去。

2. 新品规则缺「非待停售」条件 —— 美国表 73 行符合四要素
   （F=0 / G=0 / D≠停售 / 初次销售为空），其中 64 行是人工标的待停售，
   按用户「只清 9 个新品」的决定，这些不该标黄。
   → 公式追加 `,$C7<>"待停售"`。

用法：
    python fix_new_product_rule_20261027.py            # 预览
    python fix_new_product_rule_20261027.py --apply    # 落盘（先备份）
"""
import argparse
import glob
import html
import os
import re
import shutil
import zipfile
from xml.etree import ElementTree as ET

D = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                 "..", "..", "..", "1.每日销售数据表")
D = os.path.abspath(D)
BACKUP = os.path.join(os.path.dirname(D), ".workbuddy", "backup_fix_20261027b")

YELLOW_RGB = "FFFFFF00"

# 已存在的新品规则（上一轮写的）：AND($F7=0,$G7=0,$D7<>"停售",$XX7="")
NEW_RULE_RE = re.compile(
    r'(<cfRule\b[^>]*?dxfId=")(\d+)("[^>]*>\s*<formula>)'
    r'AND\(\$F7=0,\$G7=0,\$D7&lt;&gt;&quot;停售&quot;,\$([A-Z]+)7=&quot;&quot;\)'
    r'(</formula>)'
)


def sheet_map(z):
    rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    rid2t = dict((m[0], m[1]) for m in
                 re.findall(r'Id="(rId\d+)"[^>]*Target="([^"]+)"', rels))
    wb = z.read("xl/workbook.xml").decode("utf-8")
    sh = re.findall(r'<sheet [^>]*name="([^"]+)"[^>]*r:id="(rId\d+)"', wb)
    return {n: ("xl/" + rid2t[r].lstrip("/")) for n, r in sh}


def dxf_list(st):
    """返回 [(index, 主色 rgb 或 None, 原始片段)]"""
    m = re.search(r"<dxfs.*?</dxfs>", st, re.S)
    if not m:
        return []
    out = []
    for i, it in enumerate(re.findall(r"<dxf>(.*?)</dxf>", m.group(0), re.S)):
        cols = re.findall(r'(?:bgColor|fgColor) rgb="([0-9A-Fa-f]{8})"', it)
        out.append((i, cols[0].upper() if cols else None, it))
    return out


def find_yellow(dl):
    """dxfs 里 rgb=FFFFFF00 的索引（正黄）。多个取第一个。"""
    for i, c, _ in dl:
        if c == YELLOW_RGB:
            return i
    return None


def fix_sheet(x, dl):
    """返回 (新 xml, 修改条数, 明细)"""
    yellow = find_yellow(dl)
    if yellow is None:
        return x, 0, "无正黄 dxf"
    old2new = {}
    n = [0]

    def rep(m):
        n[0] += 1
        old = int(m.group(2))
        fcol = m.group(4)
        old2new[old] = yellow
        f = ('AND($F7=0,$G7=0,$D7&lt;&gt;&quot;停售&quot;,'
             '$C7&lt;&gt;&quot;待停售&quot;,$%s7=&quot;&quot;)' % fcol)
        return "%s%d%s%s%s" % (m.group(1), yellow, m.group(3), f, m.group(5))

    x2 = NEW_RULE_RE.sub(rep, x)
    detail = "dxf %s->%s" % (sorted(old2new), yellow) if old2new else "未找到新品规则"
    return x2, n[0], detail


def process(path, apply=False):
    fn = os.path.basename(path)
    z = zipfile.ZipFile(path)
    sm = sheet_map(z)
    dl = dxf_list(z.read("xl/styles.xml").decode("utf-8"))
    targets = [(k, v) for k, v in sm.items() if "每日库存分析表" in k]
    parts, plan = {}, []
    for nm, sp in targets:
        x = z.read(sp).decode("utf-8", "replace")
        x2, n, detail = fix_sheet(x, dl)
        if n:
            ET.fromstring(x2)          # 良构校验
        parts[sp] = x2
        plan.append((nm, n, detail))
    for name in z.namelist():
        if name not in parts:
            parts[name] = z.read(name)
    z.close()

    if apply:
        os.makedirs(BACKUP, exist_ok=True)
        shutil.copy2(path, os.path.join(BACKUP, fn))
        tmp = path + ".tmp"
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zo:
            for name, data in parts.items():
                zo.writestr(name, data)
        try:
            os.replace(tmp, path)
        except PermissionError:
            with open(path, "wb") as f:
                f.write(open(tmp, "rb").read())
            os.remove(tmp)
    return fn, plan


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--apply", action="store_true")
    a = ap.parse_args()
    files = sorted(glob.glob(os.path.join(D, "*库存销售表.xlsx")))
    tot = 0
    for p in files:
        fn, plan = process(p, a.apply)
        print("==", fn)
        for nm, n, detail in plan:
            tot += n
            print("   %-24s 改动%d  %s" % (nm, n, detail))
    print("\n合计改动 %d 处  %s" % (tot, "已落盘" if a.apply else "（预览，未落盘）"))
    if a.apply:
        print("备份:", BACKUP)


if __name__ == "__main__":
    main()
