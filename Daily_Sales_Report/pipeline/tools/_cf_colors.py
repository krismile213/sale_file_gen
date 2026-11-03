# -*- coding: utf-8 -*-
"""逐表汇总「实际出现的颜色 + 涉及列数」，用于和图例对账"""
import os
import re
import sys
import zipfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _cf_inventory as inv  # noqa: E402


def main():
    names = sorted(f for f in os.listdir(inv.WORK) if f.endswith(".xlsx") and not f.startswith("~"))
    for fn in names:
        z = zipfile.ZipFile(os.path.join(inv.WORK, fn))
        theme = inv.parse_theme(z)
        dxf = inv.parse_dxfs(z, theme)
        for nm, sp in inv.sheet_files(z):
            if "每日库存分析表" not in nm or sp not in z.namelist():
                continue
            x = z.read(sp).decode("utf-8", "replace")
            cols = {}
            for sq, body in re.findall(
                r'<conditionalFormatting[^>]*sqref="([^"]*)"[^>]*>(.*?)</conditionalFormatting>', x, re.S
            ):
                for rule in re.findall(r"<cfRule[^>]*?/>|<cfRule[^>]*?>.*?</cfRule>", body, re.S):
                    did = re.search(r'dxfId="(\d+)"', rule)
                    fg, bg = ("", "")
                    if did and int(did.group(1)) < len(dxf):
                        fg, bg = dxf[int(did.group(1))]
                    c = bg or fg
                    if not c:
                        continue
                    lab = inv.NAMED.get("FF" + c, c)
                    cols.setdefault(lab, set())
                    m = re.match(r"\$?([A-Z]{1,3})", sq)
                    if m:
                        cols[lab].add(m.group(1))
            print(f"{fn[:26]:<28} {nm[:18]:<20} " + " | ".join(f"{k}({len(v)})" for k, v in sorted(cols.items())))
        z.close()


if __name__ == "__main__":
    main()
