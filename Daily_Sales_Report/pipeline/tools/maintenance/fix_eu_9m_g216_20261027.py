# -*- coding: utf-8 -*-
"""把欧洲分析表 9 月 G216 由 141 补正为 142（Amazon.ae 渠道 1 件），合计 24980→24981。

原因：欧洲分析表的「欧洲」口径是**派生地区列**（非英国、非北美的 Amazon 渠道全算欧洲），
包含 Amazon.ae / Amazon.sa；而 10 张站点表只覆盖 9 个欧盟站 + 土耳其。
2026-09 恰好有 1 件 G216 走 Amazon.ae → 分析表比站点表合计多 1 件，属正常口径差异。
8 月两边完全一致，不动。
"""
import os, re, shutil, zipfile
from xml.etree import ElementTree as ET

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.abspath(os.path.join(HERE, "..", "..", "..", ".."))
XLSX = os.path.join(ROOT, "Daily_Sales_Report", "1.每日销售数据表", "欧盟英国库存销售表.xlsx")
BACKUP = os.path.join(ROOT, "Daily_Sales_Report", ".workbuddy", "backup_g216_20261027")
SKU, COL = "G216", "CJ"          # 9月列
NEW_SKU_VAL, NEW_TOTAL = 142, 24981


def main():
    z = zipfile.ZipFile(XLSX)
    rels = z.read("xl/_rels/workbook.xml.rels").decode()
    rid2t = dict(re.findall(r'Id="(rId\d+)"[^>]*Target="([^"]+)"', rels))
    wbx = z.read("xl/workbook.xml").decode()
    sm = {n: "xl/" + rid2t[r].lstrip("/") for n, r in
          re.findall(r'<sheet [^>]*name="([^"]+)"[^>]*r:id="(rId\d+)"', wbx)}
    sst = []
    if "xl/sharedStrings.xml" in z.namelist():
        sx = z.read("xl/sharedStrings.xml").decode()
        sst = ["".join(re.findall(r"<t[^>]*>(.*?)</t>", si, re.S))
               for si in re.findall(r"<si>(.*?)</si>", sx, re.S)]
    sp = sm["欧洲-每日库存分析表"]
    x = z.read(sp).decode()
    parts = {n: z.read(n) for n in z.namelist()}
    z.close()

    # 找 SKU 行
    row = None
    for m in re.finditer(r'<row r="(\d+)"[^>]*>(.*?)</row>', x, re.S):
        if int(m.group(1)) < 7:
            continue
        c = re.search(r'<c r="A%d"(?![0-9])([^>]*?)(?:/>|>(.*?)</c>)' % int(m.group(1)),
                      m.group(2), re.S)
        if not c:
            continue
        at, b = c.group(1), c.group(2) or ""
        if 't="s"' in at:
            v = re.search(r"<v>(\d+)</v>", b)
            s = sst[int(v.group(1))] if v and int(v.group(1)) < len(sst) else ""
        elif 't="inlineStr"' in at:
            s = "".join(re.findall(r"<t[^>]*>(.*?)</t>", b, re.S))
        else:
            s = b
        if str(s).strip() == SKU:
            row = int(m.group(1))
            break
    if row is None:
        raise SystemExit("找不到 %s" % SKU)
    print("%s 在第 %d 行" % (SKU, row))

    ref = "%s%d" % (COL, row)
    m = re.search(r'<c r="%s"(?![0-9])([^>]*?)(?:/>|>(.*?)</c>)' % ref, x, re.S)
    if not m:
        raise SystemExit("找不到 %s" % ref)
    seg = m.group(0)
    old = re.search(r"<v>(-?[\d.]+)</v>", seg)
    print("  %s 旧值 = %s" % (ref, old.group(1) if old else "无"))
    if "<f" in seg:
        raise SystemExit("%s 有公式，拒绝改" % ref)
    new_seg = re.sub(r"<v>-?[\d.]+</v>", "<v>%d</v>" % NEW_SKU_VAL, seg, count=1)
    x = x[:m.start()] + new_seg + x[m.end():]

    # 合计行
    mt = re.search(r'<c r="%s6"(?![0-9])(?:[^>]*?)(?:/>|>(.*?)</c>)' % COL, x, re.S)
    if not mt:
        raise SystemExit("找不到 %s6" % COL)
    tseg = mt.group(0)
    oldt = re.search(r"<v>(-?[\d.]+)</v>", tseg)
    print("  %s6 旧值 = %s" % (COL, oldt.group(1) if oldt else "无"))
    if "<f" not in tseg:
        raise SystemExit("%s6 没有公式，拒绝改" % COL)
    new_t = re.sub(r"<v>-?[\d.]+</v>", "<v>%d</v>" % NEW_TOTAL, tseg, count=1)
    x = x[:mt.start()] + new_t + x[mt.end():]

    ET.fromstring(x)
    parts[sp] = x.encode("utf-8")
    os.makedirs(BACKUP, exist_ok=True)
    shutil.copy2(XLSX, os.path.join(BACKUP, os.path.basename(XLSX)))
    tmp = XLSX + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zo:
        for n, d in parts.items():
            zo.writestr(n, d)
    try:
        os.replace(tmp, XLSX)
    except PermissionError:
        with open(XLSX, "wb") as f:
            f.write(open(tmp, "rb").read())
        os.remove(tmp)
    print("已落盘，备份:", BACKUP)


if __name__ == "__main__":
    main()
