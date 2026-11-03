# -*- coding: utf-8 -*-
"""定点核查：duplicateValues 分布、黄色 sqref 实际跨度、G261 残留"""
import os
import re
import sys
import zipfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _cf_inventory as inv  # noqa: E402


def main():
    names = sorted(f for f in os.listdir(inv.WORK) if f.endswith(".xlsx") and not f.startswith("~"))
    print("%-28s %-20s %-10s %-22s %s" % ("工作簿", "表", "duplicate", "duplicate色", "黄色sqref"))
    print("-" * 110)
    for fn in names:
        z = zipfile.ZipFile(os.path.join(inv.WORK, fn))
        theme = inv.parse_theme(z)
        dxf = inv.parse_dxfs(z, theme)
        for nm, sp in inv.sheet_files(z):
            if "每日库存分析表" not in nm or sp not in z.namelist():
                continue
            x = z.read(sp).decode("utf-8", "replace")
            dup, dupc, yel = "-", "-", "-"
            for sq, body in re.findall(
                r'<conditionalFormatting[^>]*sqref="([^"]*)"[^>]*>(.*?)</conditionalFormatting>', x, re.S
            ):
                for rule in re.findall(r"<cfRule[^>]*?/>|<cfRule[^>]*?>.*?</cfRule>", body, re.S):
                    typ = re.search(r'type="([^"]+)"', rule)
                    did = re.search(r'dxfId="(\d+)"', rule)
                    fml = re.findall(r"<formula>(.*?)</formula>", rule, re.S)
                    fg, bg = ("", "")
                    if did and int(did.group(1)) < len(dxf):
                        fg, bg = dxf[int(did.group(1))]
                    c = bg or fg
                    lab = inv.NAMED.get("FF" + c, c) if c else "(无色)"
                    if typ and typ.group(1) == "duplicateValues":
                        dup = "有"
                        dupc = f"{lab} {sq[:22]}"
                    if lab == "黄" and fml and "保底" in fml[0]:
                        yel = sq
            print("%-28s %-20s %-10s %-22s %s" % (fn[:26], nm[:18], dup, dupc, yel))
        z.close()

    print("\n=== G261 残留规则 ===")
    for fn in names:
        z = zipfile.ZipFile(os.path.join(inv.WORK, fn))
        theme = inv.parse_theme(z)
        dxf = inv.parse_dxfs(z, theme)
        for nm, sp in inv.sheet_files(z):
            if "每日库存分析表" not in nm or sp not in z.namelist():
                continue
            x = z.read(sp).decode("utf-8", "replace")
            for sq, body in re.findall(
                r'<conditionalFormatting[^>]*sqref="([^"]*)"[^>]*>(.*?)</conditionalFormatting>', x, re.S
            ):
                if "G261" in body:
                    for rule in re.findall(r"<cfRule[^>]*?>.*?</cfRule>", body, re.S):
                        if "G261" in rule:
                            pri = re.search(r'priority="(\d+)"', rule)
                            f = re.findall(r"<formula>(.*?)</formula>", rule, re.S)
                            print(f"  {fn[:26]:<28} {nm[:18]:<20} p={pri.group(1) if pri else '?'} sqref={sq}")
                            print(f"       {(f[0] if f else '')[:110]}")
        z.close()


if __name__ == "__main__":
    main()
