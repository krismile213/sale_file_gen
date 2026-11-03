# -*- coding: utf-8 -*-
"""10 张站点表的日列红色倍率 2.5 → 3.0（2026-10-27 用户确认）。

依据：欧洲分析表（组 D）日列红色是 AND(AS7>$AO7*3,$AO7>1.5)，
而 10 张站点表写的是 *2.5（当时抄了英国组的参数）。阈值 >1.5、蓝色 ÷3 都已一致，不用动。
实时速度列的 `$F7>$G7*1.5` / `AND($F7<$G7/2,$G7>2)` 是独立口径，**不动**。

安全性：先数 `$G7*2.5` / `$F7*2.5` 出现次数，不是恰好 1 次就报错退出，不做盲替换。
"""
import os, re, shutil, sys, zipfile
from xml.etree import ElementTree as ET

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.abspath(os.path.join(HERE, "..", "..", "..", ".."))
XLSX = os.path.join(ROOT, "Daily_Sales_Report", "1.每日销售数据表", "欧盟英国库存销售表.xlsx")
BACKUP = os.path.join(ROOT, "Daily_Sales_Report", ".workbuddy", "backup_site_red_20261027")
SITES = ["意大利", "法国", "西班牙", "德国", "土耳其", "荷兰", "瑞典", "波兰", "比利时", "爱尔兰"]


def main():
    z = zipfile.ZipFile(XLSX)
    rels = z.read("xl/_rels/workbook.xml.rels").decode()
    rid2t = dict(re.findall(r'Id="(rId\d+)"[^>]*Target="([^"]+)"', rels))
    wbx = z.read("xl/workbook.xml").decode()
    sm = {n: "xl/" + rid2t[r].lstrip("/") for n, r in
          re.findall(r'<sheet [^>]*name="([^"]+)"[^>]*r:id="(rId\d+)"', wbx)}
    parts = {n: z.read(n) for n in z.namelist()}
    z.close()

    for nm in SITES:
        sp = sm[nm]
        x = parts[sp].decode("utf-8")
        n_g = x.count("$G7*2.5")
        n_f = x.count("$F7*2.5")
        if n_g + n_f != 1:
            print("  ! %-5s 命中 $G7*2.5=%d / $F7*2.5=%d（期望合计 1），跳过" % (nm, n_g, n_f))
            continue
        x2 = x.replace("$G7*2.5", "$G7*3").replace("$F7*2.5", "$F7*3")
        ET.fromstring(x2)
        parts[sp] = x2.encode("utf-8")
        print("  %s: %s 改为 *3" % (nm, "$G7" if n_g else "$F7"))

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
