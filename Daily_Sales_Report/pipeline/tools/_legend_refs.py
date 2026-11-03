# -*- coding: utf-8 -*-
"""查欧盟英国簿里图例 sharedString 被哪些 sheet 的哪些单元格引用"""
import os
import re
import sys
import zipfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _cf_inventory as inv  # noqa: E402


def main():
    p = os.path.join(inv.WORK, "欧盟英国库存销售表.xlsx")
    z = zipfile.ZipFile(p)
    sx = z.read("xl/sharedStrings.xml").decode("utf-8", "replace")
    sis = re.findall(r"<si>(.*?)</si>", sx, re.S)
    texts = ["".join(re.findall(r"<t[^>]*>(.*?)</t>", si, re.S)) for si in sis]
    tgt = [i for i, t in enumerate(texts) if "橙色：新品未上架" in t]
    print("含图例的 si 索引:", tgt)
    for i in tgt:
        print(f"  si={i} len={len(texts[i])}")
        print(f"     {texts[i][:200]}")
        print()

    # 找引用
    for nm, sp in inv.sheet_files(z):
        if sp not in z.namelist():
            continue
        x = z.read(sp).decode("utf-8", "replace")
        for m in re.finditer(r'<c r="([A-Z]{1,3}\d+)"([^>]*t="s"[^>]*)>\s*<v>(\d+)</v>', x):
            idx = int(m.group(3))
            if idx in tgt:
                print(f"  引用: {nm[:22]:<24} {m.group(1):<8} -> si={idx}")
    z.close()


if __name__ == "__main__":
    main()
