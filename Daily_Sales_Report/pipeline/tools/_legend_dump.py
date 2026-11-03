# -*- coding: utf-8 -*-
"""提取每张分析表 B1 图例原文（兼容 inlineStr / sharedString / 富文本）"""
import os
import re
import zipfile

import _cf_inventory as inv


def shared_strings(z):
    if "xl/sharedStrings.xml" not in z.namelist():
        return []
    sx = z.read("xl/sharedStrings.xml").decode("utf-8", "replace")
    out = []
    for si in re.findall(r"<si>(.*?)</si>", sx, re.S):
        out.append("".join(re.findall(r"<t[^>]*>(.*?)</t>", si, re.S)))
    return out


def main():
    names = sorted(f for f in os.listdir(inv.WORK) if f.endswith(".xlsx") and not f.startswith("~"))
    for fn in names:
        p = os.path.join(inv.WORK, fn)
        try:
            z = zipfile.ZipFile(p)
        except Exception as e:
            print(fn, "OPEN FAIL", e)
            continue
        ss = shared_strings(z)
        print("=" * 96)
        print("###", fn)
        for nm, sp in inv.sheet_files(z):
            if "每日库存分析表" not in nm or sp not in z.namelist():
                continue
            x = z.read(sp).decode("utf-8", "replace")
            m = re.search(r'<c r="B1"[^>]*?(?:/>|>(.*?)</c>)', x, re.S)
            if not m:
                print(f"  -- {nm}: B1 不存在")
                continue
            head = x[m.start():m.start() + 120]
            body = m.group(1) or ""
            if 't="s"' in head:
                idx = re.search(r"<v>(\d+)</v>", body)
                txt = ss[int(idx.group(1))] if idx and int(idx.group(1)) < len(ss) else "?"
            else:
                txt = "".join(re.findall(r"<t[^>]*>(.*?)</t>", body, re.S))
            txt = txt.replace("&quot;", '"').replace("&lt;", "<").replace("&gt;", ">").replace("&amp;", "&")
            print(f"  -- {nm}  (B1 len={len(txt)})")
            for seg in txt.split("\n"):
                if seg.strip():
                    print("       |", seg.strip())
        z.close()


if __name__ == "__main__":
    main()
