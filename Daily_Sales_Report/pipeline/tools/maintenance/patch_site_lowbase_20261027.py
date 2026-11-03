# -*- coding: utf-8 -*-
"""10 张站点表补「低基数」染色规则（2026-10-27 用户确认继续）。

口径 = 欧洲-每日库存分析表已有的低基数规则逐字对齐：
    AND($AO7>0,$AO7<=1.5,AS7>=2,AS7>$AO7*4)   （dxfId=7 红，sqref=AS7:CE1476）
站点表按各自布局推导（从日列红色规则反解，不靠记忆）：
    意/法/西  日列 P..AT，基准上月日均 $G → AND($G7>0,$G7<=1.5,P7>=2,P7>$G7*4)
    德系 7 张 少「库存」列，日列 O..AS，基准 $F → AND($F7>0,$F7<=1.5,O7>=2,O7>$F7*4)
与红色规则天然互斥（红要求基准>1.5，低基数要求 ≤1.5），dxfId 复用各表红色规则
的同一个（反解得出，10 表均为 7）。优先级取各表现有 max(priority)+1。

安全性：每表必须恰好命中 1 条日列红规则、且无既有低基数规则，否则跳过该表；
落盘前 ET 良构校验；备份到 .workbuddy/backup_site_lowbase_20261027。

用法：python patch_site_lowbase_20261027.py [--apply]
"""
import os, re, shutil, sys, zipfile
from xml.etree import ElementTree as ET

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.abspath(os.path.join(HERE, "..", "..", "..", ".."))
XLSX = os.path.join(ROOT, "Daily_Sales_Report", "1.每日销售数据表", "欧盟英国库存销售表.xlsx")
BACKUP = os.path.join(ROOT, "Daily_Sales_Report", ".workbuddy", "backup_site_lowbase_20261027")
SITES = ["意大利", "法国", "西班牙", "德国", "土耳其", "荷兰", "瑞典", "波兰", "比利时", "爱尔兰"]

RED_RE = re.compile(
    r'AND\(([A-Z]{1,2})7&gt;\$([FG])7\*3,\$([FG])7&gt;1\.5\)</formula></cfRule>')


def smap_of(z):
    rels = z.read("xl/_rels/workbook.xml.rels").decode()
    rid2t = {}
    for rel in re.findall(r"<Relationship\b[^>]*/?>", rels):
        i = re.search(r'Id="([^"]+)"', rel)
        t = re.search(r'Target="([^"]+)"', rel)
        if i and t:
            t = t.group(1).lstrip("/")
            rid2t[i.group(1)] = t if t.startswith("xl/") else "xl/" + t
    wbx = z.read("xl/workbook.xml").decode()
    return {n: rid2t[r] for n, r in
            re.findall(r'<sheet [^>]*?name="([^"]+)"[^>]*?r:id="(rId\d+)"', wbx)}


def main():
    apply = "--apply" in sys.argv
    zin = zipfile.ZipFile(XLSX)
    sm = smap_of(zin)
    parts = {n: zin.read(n) for n in zin.namelist()}
    zin.close()

    for nm in SITES:
        sp = sm[nm]
        x = parts[sp].decode("utf-8")
        if re.search(r'AND\(\$[FG]7&gt;0,\$[FG]7&lt;=1\.5,', x):
            print("  ! %-5s 已有低基数规则，跳过" % nm)
            continue
        # 找日列红规则所在 conditionalFormatting 块
        tgt = None
        for m in re.finditer(
                r'(<conditionalFormatting[^>]*?sqref="([^"]*)"[^>]*>)(.*?)(</conditionalFormatting>)',
                x, re.S):
            body = m.group(3)
            hits = list(RED_RE.finditer(body))
            if len(hits) == 1:
                if tgt:
                    raise SystemExit("  ! %s 命中多个块，中止" % nm)
                tgt = (m, hits[0], body)   # body 一并保存：循环会覆盖同名变量
        if not tgt:
            print("  ! %-5s 未找到日列红规则，跳过" % nm)
            continue
        m, red, body = tgt
        dcol, base = red.group(1), red.group(2)
        head = body.rfind("<cfRule", 0, red.start())   # 回溯到规则头拿 dxfId
        dxf = re.search(r'dxfId="(\d+)"', body[head:red.start()]).group(1)
        pris = [int(p) for p in re.findall(r'priority="(\d+)"', x)]
        pri = max(pris) + 1
        f = "AND($%s7&gt;0,$%s7&lt;=1.5,%s7&gt;=2,%s7&gt;$%s7*4)" % (base, base, dcol, dcol, base)
        rule = '<cfRule type="expression" priority="%d" dxfId="%s"><formula>%s</formula></cfRule>' % (pri, dxf, f)
        new_body = body + rule
        x2 = x[:m.start()] + m.group(1) + new_body + m.group(4) + x[m.end():]
        ET.fromstring(x2)                     # 良构校验（testzip 查不出语法错）
        parts[sp] = x2.encode("utf-8")
        print("  %s: sqref=%s 低基数=%s  dxfId=%s pri=%d" % (nm, m.group(2), f[:46], dxf, pri))

    if not apply:
        print("\n（预览，未落盘）")
        return
    os.makedirs(BACKUP, exist_ok=True)
    bak = os.path.join(BACKUP, os.path.basename(XLSX))
    if not os.path.exists(bak):
        shutil.copy2(XLSX, bak)
        print("备份:", bak)
    tmp = XLSX + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zo:
        for n, d in parts.items():
            zo.writestr(n, d)
    try:
        os.replace(tmp, XLSX)
    except PermissionError:
        with open(XLSX, "wb") as fo:
            fo.write(open(tmp, "rb").read())
        os.remove(tmp)
    print("已落盘:", XLSX)


if __name__ == "__main__":
    main()
