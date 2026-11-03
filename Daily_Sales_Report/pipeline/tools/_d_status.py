# -*- coding: utf-8 -*-
"""统计 17 张分析表 D 列（状态）的实际取值分布"""
import os
import re
import sys
import zipfile
from collections import Counter

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _cf_inventory as inv  # noqa: E402


def shared_strings(z):
    if "xl/sharedStrings.xml" not in z.namelist():
        return []
    sx = z.read("xl/sharedStrings.xml").decode("utf-8", "replace")
    return ["".join(re.findall(r"<t[^>]*>(.*?)</t>", si, re.S)) for si in re.findall(r"<si>(.*?)</si>", sx, re.S)]


def main():
    names = sorted(f for f in os.listdir(inv.WORK) if f.endswith(".xlsx") and not f.startswith("~"))
    allvals = Counter()
    for fn in names:
        z = zipfile.ZipFile(os.path.join(inv.WORK, fn))
        ss = shared_strings(z)
        for nm, sp in inv.sheet_files(z):
            if "每日库存分析表" not in nm or sp not in z.namelist():
                continue
            x = z.read(sp).decode("utf-8", "replace")
            cnt = Counter()
            for rm in re.finditer(r'<row r="(\d+)"[^>]*>(.*?)</row>', x, re.S):
                r = int(rm.group(1))
                if r < 7:
                    continue
                dm = re.search(r'<c r="D%d"([^>]*)>(.*?)</c>' % r, rm.group(2), re.S)
                if not dm:
                    cnt["(无D单元格)"] += 1
                    continue
                attr, body = dm.group(1), dm.group(2)
                if 't="s"' in attr:
                    i = re.search(r"<v>(\d+)</v>", body)
                    v = ss[int(i.group(1))] if i else ""
                elif 't="inlineStr"' in attr:
                    v = "".join(re.findall(r"<t[^>]*>(.*?)</t>", body, re.S))
                else:
                    vm = re.search(r"<v>(.*?)</v>", body)
                    v = vm.group(1) if vm else ""
                cnt[v.strip()] += 1
                allvals[v.strip()] += 1
            print(f"{fn[:26]:<28} {nm[:18]:<20} " + "  ".join(f"{k or '(空)'}={v}" for k, v in cnt.most_common(8)))
        z.close()
    print("\n=== 全表合并 ===")
    for k, v in allvals.most_common(20):
        print(f"   {k or '(空)'}: {v}")


if __name__ == "__main__":
    main()
