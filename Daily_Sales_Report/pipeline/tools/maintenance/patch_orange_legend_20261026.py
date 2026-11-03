# -*- coding: utf-8 -*-
"""橙色「新品未上架」落地 + B1 图例补全 8 种业务色（XML 手术）。

背景（2026-10-26 核查结论）：
  · 图例只写了 4 色，实际有 10 类；且「橙色：新品未上架」名不符实——
    实际橙色是 duplicateValues（A 列重复 SKU），且只有英国表是橙，其余 16 表是浅红。
  · D 列（状态）实际只有 在售/在途/停售/不销售 四种，**没有「未上架」**，
    所以原有粉色规则里的 `$D7<>"未上架"` 恒真、形同虚设。

本补丁做三件事：
  ① 17 张分析表各加一条橙色规则：AND($A7<>"",$D7="在途")，打在 A 列数据行
     priority = 该表 max+1（最低优先级，不抢红/蓝/粉的显示）
  ② 英国表的 duplicateValues 由橙(FFFF9900)改为浅红(FFFFC7CE)，与其余 16 表一致，
     让「橙色」只表示「新品未上架」
  ③ B1 图例补全为 8 色：粉红(两档)/橙/黄/浅红/淡红/浅绿/红/蓝 + 实时速度列说明

不动：G261 调试残留、无色 no-op 规则（用户拍板保留）。

用法：
  python patch_orange_legend_20261026.py            # 干跑
  python patch_orange_legend_20261026.py --apply    # 落盘（先备份）
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
BAK = os.path.join(DSR, ".workbuddy", "backup_orange_legend_20261026")
APPLY = "--apply" in sys.argv

ORANGE = "FFFF9900"
LIGHTRED = "FFFFC7CE"
PINK_DEEP = "FFFF8FA8"

# ---- 图例文本改写规则（按顺序执行，每条必须命中且只命中 1 次）----
LEGEND_FIXES = [
    ("粉红色：缺货；",
     "粉红色：缺货（深粉=无在途补给，浅粉=补给在途）；"),
    ("橙色：新品未上架；",
     "橙色：新品未上架（状态=在途）；"
     "黄色：保底品；"
     "浅红：SKU重复或某单日销量>20；"
     "淡红：该SKU出现在数据透视表；"
     "浅绿：保底+机会产品且当月单日最高>4；"),
    ("超过上月日均4倍；",
     "超过上月日均4倍（「实时速度」列同用此倍率与门槛，蓝同理）；"),
]

NEW_RULE = (
    '<conditionalFormatting sqref="A7:A{r1}">'
    '<cfRule type="expression" dxfId="{dxf}" priority="{pri}">'
    '<formula>AND($A7&lt;&gt;&quot;&quot;,$D7=&quot;在途&quot;)</formula>'
    "</cfRule></conditionalFormatting>"
)


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


def parse_dxfs(styles_xml):
    """返回 [(fg_rgb, bg_rgb), ...]，顺序即 dxfId"""
    m = re.search(r"<dxfs[^>]*>(.*?)</dxfs>", styles_xml, re.S)
    if not m:
        return [], None
    out = []
    for blk in re.findall(r"<dxf>(.*?)</dxf>", m.group(1), re.S):
        fill = re.search(r"<fill>(.*?)</fill>", blk, re.S)
        fg = bg = ""
        if fill:
            f = re.search(r'<fgColor[^>]*rgb="([0-9A-Fa-f]{8})"', fill.group(1))
            b = re.search(r'<bgColor[^>]*rgb="([0-9A-Fa-f]{8})"', fill.group(1))
            fg = f.group(1).upper() if f else ""
            bg = b.group(1).upper() if b else ""
        out.append((fg, bg))
    return out, m


def find_color_dxf(dxfs, rgb):
    for i, (fg, bg) in enumerate(dxfs):
        if rgb in (fg, bg):
            return i
    return -1


def append_dxf(styles_xml, rgb):
    """在 </dxfs> 前追加一个纯填充 dxf，并更新 count，返回 (新xml, 新索引)"""
    dxfs, m = parse_dxfs(styles_xml)
    idx = len(dxfs)
    new = (
        '<dxf><fill><patternFill patternType="solid">'
        f'<bgColor rgb="{rgb}"/></patternFill></fill></dxf>'
    )
    head = styles_xml[: m.end(1)]
    tail = styles_xml[m.end(1):]
    out = head + new + tail
    out = re.sub(r'(<dxfs count=")(\d+)(")', lambda mm: mm.group(1) + str(int(mm.group(2)) + 1) + mm.group(3), out, count=1)
    return out, idx


def last_data_row(sheet_xml, dxfs):
    """用深粉缺货规则的 sqref 末行作为数据行上界（该区间已被验证无缺口）"""
    for sq, body in re.findall(
        r'<conditionalFormatting[^>]*sqref="([^"]*)"[^>]*>(.*?)</conditionalFormatting>', sheet_xml, re.S
    ):
        for rule in re.findall(r"<cfRule[^>]*?/>|<cfRule[^>]*?>.*?</cfRule>", body, re.S):
            did = re.search(r'dxfId="(\d+)"', rule)
            if not did:
                continue
            i = int(did.group(1))
            if i < len(dxfs) and PINK_DEEP in dxfs[i]:
                m = re.search(r":\$?[A-Z]{1,3}\$?(\d+)$", sq)
                if m:
                    return int(m.group(1))
    return None


def patch_legend_text(txt):
    """对一段图例文本应用改写，返回 (新文本, 命中数)"""
    n = 0
    for old, new in LEGEND_FIXES:
        c = txt.count(old)
        n += c
        txt = txt.replace(old, new)
    return txt, n


def main():
    files = sorted(f for f in os.listdir(WORK) if f.endswith(".xlsx") and not f.startswith("~"))
    plan = []

    for fn in files:
        p = os.path.join(WORK, fn)
        z = zipfile.ZipFile(p)
        names = z.namelist()
        styles = z.read("xl/styles.xml").decode("utf-8") if "xl/styles.xml" in names else ""
        dxfs, _ = parse_dxfs(styles)
        sheets = [(nm, sp) for nm, sp in sheet_files(z) if "每日库存分析表" in nm and sp in names]
        if not sheets:
            z.close()
            continue

        orange_idx = find_color_dxf(dxfs, ORANGE)
        need_new_dxf = orange_idx < 0
        new_dxf_idx = len(dxfs) if need_new_dxf else orange_idx
        lightred_idx = find_color_dxf(dxfs, LIGHTRED)

        changed = {}
        notes = []
        for nm, sp in sheets:
            x = z.read(sp).decode("utf-8")
            orig = x

            # ---- ① 新增橙色规则 ----
            r1 = last_data_row(x, dxfs)
            pris = [int(v) for v in re.findall(r'priority="(\d+)"', x)]
            pri = (max(pris) + 1) if pris else 1
            if r1:
                if "$D7=&quot;在途&quot;" in x or '$D7="在途"' in x:
                    notes.append(f"{nm}: 已有在途规则，跳过")
                else:
                    rule = NEW_RULE.format(r1=r1, dxf=new_dxf_idx, pri=pri)
                    pos = x.rfind("</conditionalFormatting>")
                    if pos < 0:
                        notes.append(f"{nm}: 无 CF 块，跳过")
                    else:
                        pos += len("</conditionalFormatting>")
                        x = x[:pos] + rule + x[pos:]
                        notes.append(f"{nm}: +橙色规则 A7:A{r1} p={pri} dxf={new_dxf_idx}")
            else:
                notes.append(f"{nm}: 未找到深粉规则，无法定位数据行，跳过")

            # ---- ② 英国表 duplicateValues 改色 ----
            def recolor(mm):
                if int(mm.group(2)) == orange_idx and lightred_idx >= 0:
                    return mm.group(1) + str(lightred_idx) + mm.group(3)
                return mm.group(0)

            x2 = re.sub(r'(<cfRule[^>]*type="duplicateValues"[^>]*dxfId=")(\d+)(")', recolor, x)
            if x2 != x:
                notes.append(f"{nm}: duplicateValues 橙→浅红(dxf {lightred_idx})")
                x = x2

            # ---- ③ 图例（inlineStr 直接改单元格）----
            m = re.search(r'<c r="B1"[^>]*t="inlineStr"[^>]*>\s*<is>(.*?)</is>', x, re.S)
            if m:
                inner = m.group(1)
                plain = "".join(re.findall(r"<t[^>]*>(.*?)</t>", inner, re.S))
                newtxt, hits = patch_legend_text(plain)
                if hits:
                    x = x[: m.start(1)] + "<t>" + newtxt + "</t>" + x[m.end(1):]
                    notes.append(f"{nm}: 图例改写 {hits} 处（inlineStr）")

            if x != orig:
                changed[sp] = x.encode("utf-8")

        # ---- 图例在 sharedStrings 的情况（欧盟英国簿）----
        if "xl/sharedStrings.xml" in names:
            sx = z.read("xl/sharedStrings.xml").decode("utf-8")
            orig = sx
            sis = list(re.finditer(r"<si>(.*?)</si>", sx, re.S))
            # 只改「每日库存分析表」B1 实际引用的 si
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
                blk = sis[idx]
                plain = "".join(re.findall(r"<t[^>]*>(.*?)</t>", blk.group(1), re.S))
                newtxt, hits = patch_legend_text(plain)
                if not hits:
                    continue
                if sx.count(plain) != 1:
                    notes.append(f"si={idx}: 原文出现 {sx.count(plain)} 次，跳过")
                    continue
                sx = sx.replace(plain, newtxt, 1)
                notes.append(f"si={idx}: 图例改写 {hits} 处（sharedString）")
            if sx != orig:
                changed["xl/sharedStrings.xml"] = sx.encode("utf-8")

        if need_new_dxf and changed:
            newstyles, _ = append_dxf(styles, ORANGE)
            changed["xl/styles.xml"] = newstyles.encode("utf-8")
            notes.append(f"styles.xml: 追加橙色 dxf idx={new_dxf_idx}")

        z.close()
        if changed:
            plan.append((p, changed, notes))

    for p, changed, notes in plan:
        print(f"{'[改]' if APPLY else '[将改]'} {os.path.basename(p)}")
        for n in notes:
            print("      " + n)
        print("      涉及部件: " + ", ".join(sorted(changed)))

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

        tmp = p + ".tmp_orange"
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
        print(f"已落盘 {name}  (备份 {bak})")


if __name__ == "__main__":
    main()
