# -*- coding: utf-8 -*-
"""欧洲表红色口径改造（只动欧洲表，用户拍板「先只改欧洲表」）。

原口径（两档）：
  p=39   红·高销量  AND(AS7>$AO7*3, $AO7>1.5)
  p=3647 红·低基数  AND($AO7>0, $AO7<=1.5, AS7>=2, AS7>$AO7*4)

新口径（一档，绝对+相对双条件）：
  p=39   红  OR( AND($AO7>0, AS7>$AO7*2, AS7>=5) , AS7>=20 )

理由（2026-10-26 数据核查）：
  · 旧口径下 G520（日均17.5）卖 43 不红（43 < 52.5），而卖 2 件的小 SKU 却红；
    低基数档 223 格里有 166 格只卖 2~3 件 = 全部红格的 75%，噪声占主导。
  · 只降倍率仍抓不到大单：G520 卖 23~32 的倍数仅 1.3~1.8，相对口径永远不红。
  · 新口径 = 「绝对大单(≥20)」或「异常放量(>日均2倍 且 ≥5件)」，
    欧洲表红格 299→327，最低红销量 2→5，G520 的 20/23/24/26/32/43 全部变红。

用法：
  python patch_red_eu_20261026.py            # 干跑
  python patch_red_eu_20261026.py --apply    # 落盘（先备份）
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
BAK = os.path.join(DSR, ".workbuddy", "backup_red_eu_20261026")
APPLY = "--apply" in sys.argv

BOOK = "欧盟英国库存销售表.xlsx"
SHEET = "欧洲-每日库存分析表"

OLD_HI = ('<cfRule type="expression" dxfId="7" priority="39">'
          '<formula>AND(AS7&gt;$AO7*3,$AO7&gt;1.5)</formula></cfRule>')
NEW_HI = ('<cfRule type="expression" dxfId="7" priority="39">'
          '<formula>OR(AND($AO7&gt;0,AS7&gt;$AO7*2,AS7&gt;=5),AS7&gt;=20)</formula></cfRule>')

OLD_LO = ('<cfRule type="expression" dxfId="7" priority="3647">'
          '<formula>AND($AO7&gt;0,$AO7&lt;=1.5,AS7&gt;=2,AS7&gt;$AO7*4)</formula></cfRule>')

# 低基数档在块内是第一个 cfRule，删掉整段
OLD_LO_AT_HEAD = ('<conditionalFormatting sqref="AS7:CE1476">' + OLD_LO)
NEW_LO_AT_HEAD = '<conditionalFormatting sqref="AS7:CE1476">'


def sheet_path(z):
    wb = z.read("xl/workbook.xml").decode("utf-8", "replace")
    rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8", "replace")
    rid2t = {}
    for tag in re.findall(r"<Relationship\b[^>]*/?>", rels):
        rid = re.search(r'\bId="([^"]+)"', tag)
        tgt = re.search(r'\bTarget="([^"]+)"', tag)
        if rid and tgt:
            rid2t[rid.group(1)] = tgt.group(1)
    for tag in re.findall(r"<sheet\b[^>]*/?>", wb):
        nm = re.search(r'\bname="([^"]+)"', tag)
        rid = re.search(r'\br:id="([^"]+)"', tag)
        if nm and rid and nm.group(1) == SHEET:
            t = rid2t.get(rid.group(1), "")
            sp = t if t.startswith("xl/") else "xl/" + t.lstrip("/")
            return sp.replace("xl/xl/", "xl/")
    return None


def main():
    p = os.path.join(WORK, BOOK)
    z = zipfile.ZipFile(p)
    sp = sheet_path(z)
    x = z.read(sp).decode("utf-8")
    z.close()

    print(f"目标: {BOOK} / {SHEET}  ({sp})")
    print()
    print("替换检查:")
    for name, old in (("高销量档公式", OLD_HI), ("低基数档整条", OLD_LO)):
        print(f"   {name}: 出现 {x.count(old)} 次 (应为 1)")
    print(f"   块首结构: 出现 {x.count(OLD_LO_AT_HEAD)} 次 (应为 1)")

    if x.count(OLD_HI) != 1 or x.count(OLD_LO_AT_HEAD) != 1:
        print("\n!! 结构不符，放弃（可能有其他改动过）")
        return

    y = x.replace(OLD_LO_AT_HEAD, NEW_LO_AT_HEAD, 1)
    y = y.replace(OLD_HI, NEW_HI, 1)

    # 校验
    if y.count("OR(AND($AO7&gt;0,AS7&gt;$AO7*2,AS7&gt;=5),AS7&gt;=20)") != 1:
        print("!! 新公式未正确写入")
        return
    if "priority=\"3647\"" in y:
        print("!! 低基数档未删净")
        return
    n_before = x.count("<cfRule")
    n_after = y.count("<cfRule")
    print(f"\ncfRule 数: {n_before} → {n_after}（应 -1）")

    if not APPLY:
        print("\n（干跑，未落盘。加 --apply 执行）")
        return

    os.makedirs(BAK, exist_ok=True)
    bak = os.path.join(BAK, BOOK)
    if not os.path.exists(bak):
        shutil.copy2(p, bak)

    tmp = p + ".tmp_red"
    zin = zipfile.ZipFile(p)
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            data = y.encode("utf-8") if info.filename == sp else zin.read(info.filename)
            zi = zipfile.ZipInfo(info.filename, date_time=info.date_time)
            zi.compress_type = zipfile.ZIP_DEFLATED
            zi.external_attr = info.external_attr
            zout.writestr(zi, data)
    zin.close()

    with zipfile.ZipFile(tmp) as zt:
        bad = zt.testzip()
        if bad:
            os.remove(tmp)
            print(f"!! testzip 失败于 {bad}，已放弃")
            return
    try:
        os.replace(tmp, p)
    except PermissionError:
        with open(tmp, "rb") as f:
            blob = f.read()
        with open(p, "wb") as f:
            f.write(blob)
        os.remove(tmp)
    print(f"\n已落盘 {BOOK}  (备份 {bak})")


if __name__ == "__main__":
    main()
