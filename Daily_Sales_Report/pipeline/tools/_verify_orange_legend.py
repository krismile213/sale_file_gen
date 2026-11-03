# -*- coding: utf-8 -*-
"""复验橙色规则 + 图例补丁：CF块数、公式数、缓存值、图例文本、zip 完整性"""
import os
import re
import sys
import zipfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
import _cf_inventory as inv  # noqa: E402

HERE = os.path.dirname(os.path.abspath(__file__))
BAK = os.path.join(os.path.dirname(os.path.dirname(HERE)), ".workbuddy", "backup_orange_legend_20261026")


def shared_strings(z):
    if "xl/sharedStrings.xml" not in z.namelist():
        return []
    sx = z.read("xl/sharedStrings.xml").decode("utf-8", "replace")
    return ["".join(re.findall(r"<t[^>]*>(.*?)</t>", si, re.S)) for si in re.findall(r"<si>(.*?)</si>", sx, re.S)]


def stats(z, sp):
    x = z.read(sp).decode("utf-8", "replace")
    return {
        "cf": len(re.findall(r"<conditionalFormatting\b", x)),
        "f": len(re.findall(r"<f[ >/]", x)),
        "v": len(re.findall(r"<v>", x)),
    }


def b1_text(z, sp, ss):
    x = z.read(sp).decode("utf-8", "replace")
    m = re.search(r'<c r="B1"[^>]*?(?:/>|>(.*?)</c>)', x, re.S)
    if not m:
        return None
    head = x[m.start():m.start() + 120]
    body = m.group(1) or ""
    if 't="s"' in head:
        i = re.search(r"<v>(\d+)</v>", body)
        return ss[int(i.group(1))] if i and int(i.group(1)) < len(ss) else "?"
    return "".join(re.findall(r"<t[^>]*>(.*?)</t>", body, re.S))


def main():
    names = sorted(f for f in os.listdir(inv.WORK) if f.endswith(".xlsx") and not f.startswith("~"))
    ok_all = True
    print("=" * 118)
    print("一、橙色规则 + CF/公式/缓存值 对比（当前 vs 备份）")
    print("=" * 118)
    print(f"{'工作簿':<26}{'表':<20}{'CF现':>5}{'CF备':>5}{'公式现':>8}{'公式备':>8}{'缓存现':>8}{'缓存备':>8}  判定")
    for fn in names:
        p = os.path.join(inv.WORK, fn)
        z = zipfile.ZipFile(p)
        zp = os.path.join(BAK, fn)
        zb = zipfile.ZipFile(zp) if os.path.exists(zp) else None
        ss = shared_strings(z)
        for nm, sp in inv.sheet_files(z):
            if "每日库存分析表" not in nm or sp not in z.namelist():
                continue
            a = stats(z, sp)
            if zb and sp in zb.namelist():
                b = stats(zb, sp)
                bcf, bf, bv = b["cf"], b["f"], b["v"]
            else:
                bcf = bf = bv = "-"
            x = z.read(sp).decode("utf-8", "replace")
            has_orange = ('$D7=&quot;在途&quot;' in x) or ('$D7="在途"' in x)
            good = (
                has_orange
                and (bcf == "-" or a["cf"] == bcf + 1)
                and (bf == "-" or a["f"] == bf)
                and (bv == "-" or a["v"] == bv)
            )
            if not good:
                ok_all = False
            print(f"{fn[:24]:<26}{nm[:18]:<20}{a['cf']:>5}{str(bcf):>5}{a['f']:>8}{str(bf):>8}"
                  f"{a['v']:>8}{str(bv):>8}  {'OK' if good else '!! 检查'}  {'橙✓' if has_orange else '橙✗'}")
        z.close()
        if zb:
            zb.close()

    print()
    print("=" * 118)
    print("二、图例文本（应含 8 色关键词）")
    print("=" * 118)
    keys = ["粉红色：缺货（深粉", "橙色：新品未上架（状态=在途）", "黄色：保底品",
            "浅红：SKU重复", "淡红：该SKU出现在数据透视表", "浅绿：保底+机会产品", "红色：销量高于", "实时速度"]
    for fn in names:
        z = zipfile.ZipFile(os.path.join(inv.WORK, fn))
        ss = shared_strings(z)
        for nm, sp in inv.sheet_files(z):
            if "每日库存分析表" not in nm or sp not in z.namelist():
                continue
            t = b1_text(z, sp, ss) or ""
            miss = [k for k in keys if k not in t]
            # 新兴市场 6 表本来就没有蓝色分句
            flag = "OK" if not miss else "缺:" + ",".join(miss)
            if miss:
                ok_all = False
            print(f"  {fn[:24]:<26}{nm[:18]:<20}len={len(t):<4}{flag}")
        z.close()

    print()
    print("=" * 118)
    print("三、zip 完整性 + 欧洲表缓存值抽查（应 17.5/19.4/0.03/1.1）")
    print("=" * 118)
    for fn in names:
        p = os.path.join(inv.WORK, fn)
        with zipfile.ZipFile(p) as z:
            bad = z.testzip()
        print(f"  {'✓' if bad is None else '✗'} {fn[:40]:<42} testzip={'OK' if bad is None else bad}")
        if bad:
            ok_all = False

    z = zipfile.ZipFile(os.path.join(inv.WORK, "欧盟英国库存销售表.xlsx"))
    for nm, sp in inv.sheet_files(z):
        if nm != "欧洲-每日库存分析表":
            continue
        x = z.read(sp).decode("utf-8", "replace")
        for ref in ("AO7", "AN7", "AO452", "AN452"):
            m = re.search(r'<c r="%s"[^>]*>(.*?)</c>' % ref, x, re.S)
            v = re.search(r"<v>(.*?)</v>", m.group(1)) if m else None
            print(f"  {ref}: 缓存={v.group(1) if v else '无'}")
    z.close()

    print()
    print("总判定:", "全部通过" if ok_all else "有问题，见上")


if __name__ == "__main__":
    main()
