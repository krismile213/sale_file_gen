# -*- coding: utf-8 -*-
"""严格校验：XML 良构性、CF 块 schema 位置、橙色规则唯一性、在途行覆盖"""
import os
import re
import sys
import xml.etree.ElementTree as ET
import zipfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _cf_inventory as inv  # noqa: E402

NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
ORDER = ["sheetPr", "dimension", "sheetViews", "sheetFormatPr", "cols", "sheetData",
         "sheetCalcPr", "sheetProtection", "protectedRanges", "scenarios", "autoFilter",
         "sortState", "dataConsolidate", "customSheetViews", "mergeCells", "phoneticPr",
         "conditionalFormatting", "dataValidations", "hyperlinks", "printOptions",
         "pageMargins", "pageSetup", "headerFooter", "rowBreaks", "colBreaks",
         "customProperties", "cellWatches", "ignoredErrors", "smartTags", "drawing",
         "legacyDrawing", "legacyDrawingHF", "picture", "oleObjects", "controls",
         "webPublishItems", "tableParts", "extLst"]


def shared_strings(z):
    if "xl/sharedStrings.xml" not in z.namelist():
        return []
    sx = z.read("xl/sharedStrings.xml").decode("utf-8", "replace")
    return ["".join(re.findall(r"<t[^>]*>(.*?)</t>", si, re.S)) for si in re.findall(r"<si>(.*?)</si>", sx, re.S)]


def d_values(x, ss):
    out = {}
    for rm in re.finditer(r'<row r="(\d+)"[^>]*>(.*?)</row>', x, re.S):
        r = int(rm.group(1))
        if r < 7:
            continue
        dm = re.search(r'<c r="D%d"([^>]*)>(.*?)</c>' % r, rm.group(2), re.S)
        if not dm:
            continue
        attr, body = dm.group(1), dm.group(2)
        if 't="s"' in attr:
            i = re.search(r"<v>(\d+)</v>", body)
            v = ss[int(i.group(1))] if i and int(i.group(1)) < len(ss) else ""
        elif 't="inlineStr"' in attr:
            v = "".join(re.findall(r"<t[^>]*>(.*?)</t>", body, re.S))
        else:
            vm = re.search(r"<v>(.*?)</v>", body)
            v = vm.group(1) if vm else ""
        out[r] = v.strip()
    return out


def main():
    names = sorted(f for f in os.listdir(inv.WORK) if f.endswith(".xlsx") and not f.startswith("~"))
    bad = []
    for fn in names:
        p = os.path.join(inv.WORK, fn)
        z = zipfile.ZipFile(p)
        ss = shared_strings(z)
        theme = inv.parse_theme(z)
        dxfs = inv.parse_dxfs(z, theme)

        for part in ("xl/styles.xml", "xl/sharedStrings.xml"):
            if part in z.namelist():
                try:
                    ET.fromstring(z.read(part))
                except Exception as e:
                    bad.append(f"{fn} {part} XML 解析失败: {e}")

        for nm, sp in inv.sheet_files(z):
            if "每日库存分析表" not in nm or sp not in z.namelist():
                continue
            raw = z.read(sp)
            try:
                root = ET.fromstring(raw)
            except Exception as e:
                bad.append(f"{fn} {nm} XML 解析失败: {e}")
                continue

            tags = [c.tag.replace(NS, "") for c in root]
            idx = [ORDER.index(t) for t in tags if t in ORDER]
            if idx != sorted(idx):
                bad.append(f"{fn} {nm} 子元素顺序异常: {[t for t in tags if t in ORDER][:14]}")

            x = raw.decode("utf-8", "replace")
            orange = []
            for sq, body in re.findall(
                r'<conditionalFormatting[^>]*sqref="([^"]*)"[^>]*>(.*?)</conditionalFormatting>', x, re.S
            ):
                for rule in re.findall(r"<cfRule[^>]*?/>|<cfRule[^>]*?>.*?</cfRule>", body, re.S):
                    did = re.search(r'dxfId="(\d+)"', rule)
                    if not did:
                        continue
                    i = int(did.group(1))
                    if i >= len(dxfs):
                        continue
                    fg, bg = dxfs[i]
                    if inv.NAMED.get("FF" + (bg or fg)) == "橙":
                        orange.append((sq, rule))

            dv = d_values(x, ss)
            in_transit = sorted(r for r, v in dv.items() if v == "在途")
            covered = "-"
            if orange:
                m = re.search(r"A7:A(\d+)", orange[0][0])
                if m:
                    hi = int(m.group(1))
                    if not in_transit:
                        covered = "无在途行"
                    else:
                        covered = "全含" if max(in_transit) <= hi else "有遗漏"
            print(f"{fn[:24]:<26}{nm[:18]:<20}橙规则={len(orange)} 在途行={len(in_transit):<4}覆盖={covered}")
            if len(orange) != 1:
                bad.append(f"{fn} {nm} 橙色规则数={len(orange)}（应为 1）")
            if covered == "有遗漏":
                bad.append(f"{fn} {nm} 在途行超出橙色 sqref 上界")
        z.close()

    print()
    if bad:
        print("!! 问题：")
        for b in bad:
            print("   ", b)
    else:
        print("严格校验全部通过（XML 良构 / schema 顺序 / 橙规则唯一 / 在途行全覆盖）")


if __name__ == "__main__":
    main()
