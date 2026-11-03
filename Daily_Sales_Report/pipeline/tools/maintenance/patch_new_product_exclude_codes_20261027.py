# -*- coding: utf-8 -*-
"""
新品标黄规则追加排除码 XB / XE / XC（2026-10-27，用户确认）。

背景：C 列是「商品分类 + 状态码」混合列，取值有 A~E、待停售、XB、XE、XC。
      XB/XE/XC 与「待停售」同义（不做了），但这些行 D 列仍是「在售」、
      可售/在途为 0、初次销售日期为空，会被新品规则顺带标黄
      （典型：美国表 G532，C='XB'）。

改法：把 17 张分析表里的新品规则
    AND($F7=0,$G7=0,$D7<>"停售",$C7<>"待停售",$X7="")
追加三个不等条件：
    AND($F7=0,$G7=0,$D7<>"停售",$C7<>"待停售",$C7<>"XB",$C7<>"XE",$C7<>"XC",$X7="")

用法：
    python patch_new_product_exclude_codes_20261027.py
    python patch_new_product_exclude_codes_20261027.py --apply
"""
import argparse
import glob
import os
import re
import shutil
import zipfile
from xml.etree import ElementTree as ET

HERE = os.path.dirname(os.path.abspath(__file__))
D = os.path.abspath(os.path.join(HERE, "..", "..", "..", "1.每日销售数据表"))
BACKUP = os.path.abspath(os.path.join(HERE, "..", "..", "..", ".workbuddy",
                                      "backup_exclude_codes_20261027"))

# 现有新品规则：AND($F7=0,$G7=0,$D7<>"停售",$C7<>"待停售",$XX7="")
RULE_RE = re.compile(
    r'(<formula>)AND\(\$F7=0,\$G7=0,\$D7&lt;&gt;&quot;停售&quot;,'
    r'\$C7&lt;&gt;&quot;待停售&quot;,\$([A-Z]+)7=&quot;&quot;\)(</formula>)'
)

EXTRA = ('$C7&lt;&gt;&quot;XB&quot;,$C7&lt;&gt;&quot;XE&quot;,$C7&lt;&gt;&quot;XC&quot;,')


def smap(z):
    rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    rid2t = dict((m[0], m[1]) for m in
                 re.findall(r'Id="(rId\d+)"[^>]*Target="([^"]+)"', rels))
    wb = z.read("xl/workbook.xml").decode("utf-8")
    return {n: ("xl/" + rid2t[r].lstrip("/"))
            for n, r in re.findall(r'<sheet [^>]*name="([^"]+)"[^>]*r:id="(rId\d+)"', wb)}


def patch(x):
    n = [0]

    def rep(m):
        n[0] += 1
        return ("%sAND($F7=0,$G7=0,$D7&lt;&gt;&quot;停售&quot;,"
                "$C7&lt;&gt;&quot;待停售&quot;,%s$%s7=&quot;&quot;)%s"
                % (m.group(1), EXTRA, m.group(2), m.group(3)))

    return RULE_RE.sub(rep, x), n[0]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--apply", action="store_true")
    a = ap.parse_args()
    tot = 0
    for p in sorted(glob.glob(os.path.join(D, "*库存销售表.xlsx"))):
        fn = os.path.basename(p)
        z = zipfile.ZipFile(p)
        parts, plan = {}, []
        for nm, sp in smap(z).items():
            if "每日库存分析表" not in nm:
                continue
            x = z.read(sp).decode("utf-8")
            x2, n = patch(x)
            if n:
                ET.fromstring(x2)
                parts[sp] = x2
                plan.append((nm, n))
                tot += n
        for name in z.namelist():
            if name not in parts:
                parts[name] = z.read(name)
        z.close()
        if not plan:
            continue
        print("==", fn)
        for nm, n in plan:
            print("   %-24s 改 %d 条" % (nm, n))
        if a.apply:
            os.makedirs(BACKUP, exist_ok=True)
            shutil.copy2(p, os.path.join(BACKUP, fn))
            tmp = p + ".tmp"
            with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zo:
                for name, data in parts.items():
                    zo.writestr(name, data)
            try:
                os.replace(tmp, p)
            except PermissionError:
                with open(p, "wb") as f:
                    f.write(open(tmp, "rb").read())
                os.remove(tmp)
    print("\n合计 %d 条  %s" % (tot, "已落盘" if a.apply else "（预览）"))


if __name__ == "__main__":
    main()
