# -*- coding: utf-8 -*-
"""修正图例中「实时速度」列的表述（上一版写错了）。

错误表述：超过上月日均4倍（「实时速度」列同用此倍率与门槛，蓝同理）；
实际情况：实时速度列(T/AD/AN)是独立口径，与日列倍率无关——
          红 = $T7>$U7*1.5（高于上月日均1.5倍，无日均门槛）
          蓝 = AND($T7<$U7/2,$U7>2)（低于日均1/2 且 日均>2）
          新兴市场 6 表只有红、没有蓝。

正确表述：
  有蓝的表：超过上月日均4倍；实时速度列另按固定口径：红=高于上月日均1.5倍，蓝=低于日均1/2且日均>2；
  无蓝的表：超过上月日均4倍；实时速度列另按固定口径：红=高于上月日均1.5倍；

用法：
  python patch_legend_speed_fix_20261026.py            # 干跑
  python patch_legend_speed_fix_20261026.py --apply    # 落盘（先备份）
"""
import os
import re
import shutil
import sys
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
DSR = HERE
for _ in range(6):
    if os.path.isdir(os.path.join(DSR, "1.每日销售数据表")):
        break
    DSR = os.path.dirname(DSR)
WORK = os.path.join(DSR, "1.每日销售数据表")
BAK = os.path.join(DSR, ".workbuddy", "backup_legend_speed_20261026")
APPLY = "--apply" in sys.argv

OLD = "超过上月日均4倍（「实时速度」列同用此倍率与门槛，蓝同理）；"
NEW_WITH_BLUE = "超过上月日均4倍；实时速度列另按固定口径：红=高于上月日均1.5倍，蓝=低于日均1/2且日均>2；"
NEW_NO_BLUE = "超过上月日均4倍；实时速度列另按固定口径：红=高于上月日均1.5倍；"


def sheet_files(z):
    wb = z.read("xl/workbook.xml").decode("utf-8", "replace")
    rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8", "replace")
    rid2t = {}
    for tag in re.findall(r"<Relationship\b[^>]*/?>", rels):
        rid = re.search(r'\bId="([^"]+)"', tag)
        tgt = re.search(r'\bTarget="([^"]+)"', tag)
        if rid and tgt:
            rid2t[rid.group(1)] = tgt.group(1)
    out = []
    for tag in re.findall(r"<sheet\b[^>]*/?>", wb):
        nm = re.search(r'\bname="([^"]+)"', tag)
        rid = re.search(r'\br:id="([^"]+)"', tag)
        if not (nm and rid):
            continue
        t = rid2t.get(rid.group(1), "")
        if not t:
            continue
        sp = t if t.startswith("xl/") else "xl/" + t.lstrip("/")
        out.append((nm.group(1), sp.replace("xl/xl/", "xl/")))
    return out


def fix(txt):
    if OLD not in txt:
        return txt, 0
    new = NEW_WITH_BLUE if "蓝色：" in txt else NEW_NO_BLUE
    return txt.replace(OLD, new), txt.count(OLD)


def main():
    files = sorted(f for f in os.listdir(WORK) if f.endswith(".xlsx") and not f.startswith("~"))
    plan = []
    for fn in files:
        p = os.path.join(WORK, fn)
        z = zipfile.ZipFile(p)
        names = z.namelist()
        changed = {}
        notes = []
        sheets = [(nm, sp) for nm, sp in sheet_files(z) if "每日库存分析表" in nm and sp in names]
        if not sheets:
            z.close()
            continue

        # inlineStr 情况
        for nm, sp in sheets:
            x = z.read(sp).decode("utf-8")
            m = re.search(r'<c r="B1"[^>]*t="inlineStr"[^>]*>\s*<is>(.*?)</is>', x, re.S)
            if not m:
                continue
            plain = "".join(re.findall(r"<t[^>]*>(.*?)</t>", m.group(1), re.S))
            newtxt, n = fix(plain)
            if n:
                x = x[: m.start(1)] + "<t>" + newtxt + "</t>" + x[m.end(1):]
                changed[sp] = x.encode("utf-8")
                notes.append(f"{nm}: 修正 {n} 处（inlineStr，{'有蓝' if '蓝色：' in plain else '无蓝'}）")

        # sharedString 情况
        if "xl/sharedStrings.xml" in names:
            sx = z.read("xl/sharedStrings.xml").decode("utf-8")
            orig = sx
            sis = list(re.finditer(r"<si>(.*?)</si>", sx, re.S))
            wanted = set()
            for nm, sp in sheets:
                blob = changed.get(sp)
                x = blob.decode("utf-8") if blob is not None else z.read(sp).decode("utf-8")
                m = re.search(r'<c r="B1"[^>]*t="s"[^>]*>\s*<v>(\d+)</v>', x)
                if m:
                    wanted.add(int(m.group(1)))
            for idx in sorted(wanted):
                if idx >= len(sis):
                    continue
                plain = "".join(re.findall(r"<t[^>]*>(.*?)</t>", sis[idx].group(1), re.S))
                newtxt, n = fix(plain)
                if not n:
                    continue
                if sx.count(plain) != 1:
                    notes.append(f"si={idx}: 原文出现 {sx.count(plain)} 次，跳过")
                    continue
                sx = sx.replace(plain, newtxt, 1)
                notes.append(f"si={idx}: 修正 {n} 处（sharedString，{'有蓝' if '蓝色：' in plain else '无蓝'}）")
            if sx != orig:
                changed["xl/sharedStrings.xml"] = sx.encode("utf-8")

        z.close()
        if changed:
            plan.append((p, changed, notes))

    for p, changed, notes in plan:
        print(f"{'[改]' if APPLY else '[将改]'} {os.path.basename(p)}")
        for n in notes:
            print("      " + n)
    print()
    print(f"合计：{len(plan)} 本工作簿")
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
        tmp = p + ".tmp_speed"
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
        with zipfile.ZipFile(tmp) as zt:
            bad = zt.testzip()
            if bad:
                os.remove(tmp)
                print(f"!! {name} testzip 失败于 {bad}，已放弃")
                continue
        try:
            os.replace(tmp, p)
        except PermissionError:
            with open(tmp, "rb") as f:
                blob = f.read()
            with open(p, "wb") as f:
                f.write(blob)
            os.remove(tmp)
        print(f"已落盘 {name}")


if __name__ == "__main__":
    main()
