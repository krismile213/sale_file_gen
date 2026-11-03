# -*- coding: utf-8 -*-
"""修正 B1 图例文字（XML 手术，不动 CF/公式/缓存值）。

修两处：
  ① 蓝色分句笔误（11 张有蓝规则的表）
       原：蓝色：销量低于上月日均销量的2.5倍      ← 字面意思是 <2.5×，几乎全表都成立
       新：蓝色：销量低于上月日均的1/2.5（即日均÷2.5）   ← 实际公式 AND(AS7<$AO7/2.5,AS7>0)
       3 倍同理 → 1/3（日均÷3）
  ② 沙特主表缺分号（1 张）
       原：...高于0.5或上月日均≤0.5且日销≥2件...
       新：...高于0.5；或上月日均≤0.5且日销≥2件...
       （其余 5 张新兴市场表都带这个分号，只有沙特主表漏了）

不动的（需用户拍板）：
  · 橙色：新品未上架 —— 实际橙色是 A 列重复 SKU 高亮，且只有英国表有
  · 粉红色：缺货 —— 实际分深粉(无补给)/浅粉(补给在途)两档
  · 实时速度列的红/蓝（AN>AO*1.5 / AN<AO/2）图例未写

用法：
  python patch_legend_blue_20261026.py            # 干跑
  python patch_legend_blue_20261026.py --apply    # 落盘（先备份）
"""
import os, re, sys, shutil, zipfile, glob

HERE = os.path.dirname(os.path.abspath(__file__))
DSR = HERE
for _ in range(6):
    if os.path.isdir(os.path.join(DSR, "1.每日销售数据表")):
        break
    DSR = os.path.dirname(DSR)
WORK = os.path.join(DSR, "1.每日销售数据表")
BAK = os.path.join(DSR, ".workbuddy", "backup_legend_20261026")
APPLY = "--apply" in sys.argv

# (旧串, 新串)
FIXES = [
    ("蓝色：销量低于上月日均销量的2.5倍",
     "蓝色：销量低于上月日均的1/2.5（即日均÷2.5）"),
    ("蓝色：销量低于上月日均销量的3倍",
     "蓝色：销量低于上月日均的1/3（即日均÷3）"),
    ("且上月日均销量高于0.5或上月日均≤0.5",
     "且上月日均销量高于0.5；或上月日均≤0.5"),
]


def patch_bytes(data):
    """对 XML 字节做替换，返回 (新字节, 命中列表)"""
    txt = data.decode("utf-8")
    hits = []
    for old, new in FIXES:
        n = txt.count(old)
        if n:
            txt = txt.replace(old, new)
            hits.append((old[:24], n))
    return txt.encode("utf-8"), hits


def main():
    files = sorted(glob.glob(os.path.join(WORK, "*.xlsx")))
    total_files = 0
    total_hits = 0
    plan = []

    for p in files:
        name = os.path.basename(p)
        if name.startswith("~$") or name.endswith(".bak"):
            continue
        z = zipfile.ZipFile(p)
        changed = {}
        file_hits = []
        for info in z.infolist():
            if not info.filename.endswith(".xml"):
                continue
            data = z.read(info.filename)
            new, hits = patch_bytes(data)
            if hits:
                changed[info.filename] = new
                file_hits.extend(hits)
        z.close()
        if not changed:
            continue
        total_files += 1
        total_hits += sum(n for _, n in file_hits)
        print(f"{'[改]' if APPLY else '[将改]'} {name}")
        for fn, _ in changed.items():
            print(f"      {fn}")
        for old, n in file_hits:
            print(f"      ×{n}  {old}...")
        plan.append((p, changed, file_hits))

    print()
    print(f"合计：{total_files} 本工作簿、{total_hits} 处替换")
    if not APPLY:
        print("（干跑，未落盘。加 --apply 执行）")
        return
    if not plan:
        print("无需改动。")
        return

    os.makedirs(BAK, exist_ok=True)
    for p, changed, _ in plan:
        name = os.path.basename(p)
        bak = os.path.join(BAK, name)
        if not os.path.exists(bak):
            shutil.copy2(p, bak)

        tmp = p + ".tmp_legend"
        zin = zipfile.ZipFile(p)
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            for info in zin.infolist():
                data = changed.get(info.filename)
                if data is None:
                    data = zin.read(info.filename)
                zi = zipfile.ZipInfo(info.filename, date_time=info.date_time)
                zi.compress_type = zipfile.ZIP_DEFLATED
                zi.external_attr = info.external_attr
                zout.writestr(zi, data)
        zin.close()

        # 完整性校验
        with zipfile.ZipFile(tmp) as zt:
            bad = zt.testzip()
            if bad:
                os.remove(tmp)
                print(f"!! {name} testzip 失败于 {bad}，已放弃")
                continue
        try:
            os.replace(tmp, p)
        except PermissionError:
            # 被 Excel/预览占用时退化为原地写
            with open(tmp, "rb") as f:
                blob = f.read()
            with open(p, "wb") as f:
                f.write(blob)
            os.remove(tmp)
        print(f"已落盘 {name}  (备份 {bak})")

    print("\n=== 落盘后复验 ===")
    for p, _, _ in plan:
        name = os.path.basename(p)
        z = zipfile.ZipFile(p)
        found = []
        for fn in z.namelist():
            if not fn.endswith(".xml"):
                continue
            t = z.read(fn).decode("utf-8")
            for m in re.finditer(r"蓝色：([^<]{0,60})", t):
                found.append(m.group(1))
        z.close()
        print(f"  {name}")
        for f in sorted(set(found)):
            print(f"      蓝色：{f}")


if __name__ == "__main__":
    main()
