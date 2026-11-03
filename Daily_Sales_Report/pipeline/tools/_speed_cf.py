# -*- coding: utf-8 -*-
"""按表头定位「实时速度」列，再判断该列有无红/蓝规则"""
import os
import re
import sys
import zipfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _cf_inventory as inv  # noqa: E402


def shared_strings(z):
    if "xl/sharedStrings.xml" not in z.namelist():
        return []
    sx = z.read("xl/sharedStrings.xml").decode("utf-8", "replace")
    return ["".join(re.findall(r"<t[^>]*>(.*?)</t>", si, re.S)) for si in re.findall(r"<si>(.*?)</si>", sx, re.S)]


def find_col(x, ss, want):
    """在 row 5/6 找表头等于 want 的列字母"""
    for row in ("5", "6"):
        m = re.search(r'<row r="%s"[^>]*>(.*?)</row>' % row, x, re.S)
        if not m:
            continue
        for cm in re.finditer(r'<c r="([A-Z]{1,3})%s"([^>]*)>(.*?)</c>' % row, m.group(1), re.S):
            attr, body = cm.group(2), cm.group(3)
            if 't="s"' in attr:
                i = re.search(r"<v>(\d+)</v>", body)
                v = ss[int(i.group(1))] if i and int(i.group(1)) < len(ss) else ""
            else:
                v = "".join(re.findall(r"<t[^>]*>(.*?)</t>", body, re.S))
            if v.strip() == want:
                return cm.group(1)
    return None


def main():
    names = sorted(f for f in os.listdir(inv.WORK) if f.endswith(".xlsx") and not f.startswith("~"))
    for fn in names:
        z = zipfile.ZipFile(os.path.join(inv.WORK, fn))
        ss = shared_strings(z)
        theme = inv.parse_theme(z)
        dxf = inv.parse_dxfs(z, theme)
        for nm, sp in inv.sheet_files(z):
            if "每日库存分析表" not in nm or sp not in z.namelist():
                continue
            x = z.read(sp).decode("utf-8", "replace")
            col = find_col(x, ss, "实时速度")
            red = blue = 0
            if col:
                for sq, body in re.findall(
                    r'<conditionalFormatting[^>]*sqref="([^"]*)"[^>]*>(.*?)</conditionalFormatting>', x, re.S
                ):
                    if not re.match(r"\$?%s\d" % col, sq):
                        continue
                    for r in re.findall(r"<cfRule[^>]*?/>|<cfRule[^>]*?>.*?</cfRule>", body, re.S):
                        fm = re.findall(r"<formula>(.*?)</formula>", r, re.S)
                        if not fm:
                            continue
                        ff = fm[0].replace("&lt;", "<").replace("&gt;", ">")
                        did = re.search(r'dxfId="(\d+)"', r)
                        c = ""
                        if did and int(did.group(1)) < len(dxf):
                            fg, bg = dxf[int(did.group(1))]
                            c = bg or fg
                        lab = inv.NAMED.get("FF" + c, c) if c else "(无色)"
                        if ">" in ff and "1.5" in ff:
                            red += 1
                        elif "<" in ff:
                            blue += 1
            print(f"{fn[:24]:<26}{nm[:16]:<18}实时速度列={str(col):<4}红={red} 蓝={blue}")
        z.close()


if __name__ == "__main__":
    main()
