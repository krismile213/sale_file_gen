# -*- coding: utf-8 -*-
"""把全量 CF 清单压成「工作簿 / 表 / 颜色 -> 列」矩阵，便于核对图例。"""
import os
import re
import zipfile

import _cf_inventory as inv

HERE = os.path.dirname(os.path.abspath(__file__))


def colof(sqref):
    """sqref 里所有区间取起始列字母，去重保序"""
    cols = []
    for part in sqref.split():
        m = re.match(r"\$?([A-Z]{1,3})", part)
        if m and m.group(1) not in cols:
            cols.append(m.group(1))
    return cols


def main():
    names = sorted(f for f in os.listdir(inv.WORK) if f.endswith(".xlsx") and not f.startswith("~"))
    for fn in names:
        p = os.path.join(inv.WORK, fn)
        try:
            z = zipfile.ZipFile(p)
        except Exception as e:
            print(fn, "OPEN FAIL", e)
            continue
        theme = inv.parse_theme(z)
        dxf = inv.parse_dxfs(z, theme)
        print("=" * 96)
        print("###", fn)
        for nm, sp in inv.sheet_files(z):
            if "每日库存分析表" not in nm or sp not in z.namelist():
                continue
            x = z.read(sp).decode("utf-8", "replace")
            print(f"  -- {nm}")
            agg = {}
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
                    col = bg or fg
                    label = inv.NAMED.get("FF" + col, col or "(无色)")
                    f0 = (fml[0] if fml else "").replace("&quot;", '"').replace("&gt;", ">").replace("&lt;", "<")
                    key = (label, typ.group(1) if typ else "?")
                    agg.setdefault(key, []).append((colof(sq), f0[:60]))
            for (label, typ), items in agg.items():
                cols = []
                for c, _ in items:
                    for c1 in c:
                        if c1 not in cols:
                            cols.append(c1)
                fs = [f for _, f in items if f]
                print(f"     {label:<14} {typ:<18} 列={','.join(cols)[:60]:<60}")
                if fs:
                    print(f"          f: {fs[0]}")
        z.close()


if __name__ == "__main__":
    main()
