# -*- coding: utf-8 -*-
"""验证修补后的行：公式引用是否指向自身行（不是模板行 r7）"""
import zipfile, re

p = r"C:/Users/dell/Desktop/work-main/sale_file_gen/Daily_Sales_Report/1.每日销售数据表/欧盟英国库存销售表.xlsx"
z = zipfile.ZipFile(p)

CHECKS = {
    "xl/worksheets/sheet9.xml": ("英国", 7, [466, 470]),
    "xl/worksheets/sheet10.xml": ("欧洲", 7, [465, 468]),
}
for part, (label, tpl, targets) in CHECKS.items():
    x = z.read(part).decode()
    print("=" * 66)
    print(label)
    for rn in targets + [tpl]:
        m = re.search(r'<row r="%d"[^>]*>(.*?)</row>' % rn, x, re.S)
        cells = {}
        for mc in re.finditer(r'<c r="([A-Z]+)%d"([^>]*)(?:/>|>(.*?)</c>)' % rn, m.group(1), re.S):
            cells[mc.group(1)] = mc.group(3) or ""
        print("  --- r%d" % rn)
        for col in ("A", "B", "E", "F", "G", "H", "AC", "AD"):
            body = cells.get(col, "")
            f = re.search(r"<f[^>]*>(.*?)</f>", body, re.S)
            v = re.search(r"<v>(.*?)</v>", body, re.S)
            t = re.search(r"<t[^>]*>(.*?)</t>", body, re.S)
            if f:
                txt = f.group(1)
                # 检查是否含模板行号 7（错误信号）
                bad = re.search(r"(?<![$\d])%d(?![$\d])" % tpl, txt)
                flag = "  ⚠含r%d" % tpl if bad else ""
                print("    %-3s %s%s" % (col, txt[:95], flag))
            elif t:
                print("    %-3s %r" % (col, t.group(1)[:60]))
            elif v:
                print("    %-3s V=%s" % (col, v.group(1)[:20]))
            else:
                print("    %-3s (空)" % col)
