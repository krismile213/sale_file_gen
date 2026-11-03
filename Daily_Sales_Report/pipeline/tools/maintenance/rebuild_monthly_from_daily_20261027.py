# -*- coding: utf-8 -*-
"""
用「日销数据」重建欧盟簿的 8 月 / 9 月月度销量（2026-10-27，用户要求以日销为准）。

为什么换源：
  · 上一轮站点表 8/9 月取自 `月销售总表.xlsx`，SKU 写法与站点表不一致
    （总表写 G261，站点表写 G261IT），导致 9 个国家合计少算 217 件；
    而且总表比日销低 246 件（总表生成早于月末补单）。
  · 分析表 8 月列另有脏值：G598/G212/G272/G302/K001/K004/K005 七个 SKU
    的 1~8 月被写成同一份（299/276/251/222/242/288/250/280），
    而月销售总表里它们其实是 0~8 的小数 → 8 月合计虚高 7×299=2093。

数据源：
  8 月 ← `2026年每日数据/2026年8月/2026年8月-欧盟英国库存销售表.xlsx` 的「每日销售数据」
  9 月 ← `2026年每日数据/2026年9月/2026年9月-欧盟英国库存销售表.xlsx` 同上
  口径：有效性(AQ)=1、purchase-date 月份匹配、按 sales-channel × sku 汇总 quantity(O)
  已验证：8/9 月每个有销量的 SKU 在对应站点表里都已有行，缺失 0 个。

写入范围：
  · 分析表「8月」列 —— 英国按 Amazon.co.uk、欧洲按 9 国合计（修 299 脏值）
  · 10 张站点表「8月/9月」列
  · 第 6 行合计只改 `<v>` 缓存值，**不动共享公式**
  · 10 月列不动（已由当前「每日销售数据」填好并逐行校验过）

用法：
    python rebuild_monthly_from_daily_20261027.py
    python rebuild_monthly_from_daily_20261027.py --apply
"""
import argparse
import os
import re
import shutil
import zipfile
from collections import defaultdict
from xml.etree import ElementTree as ET

import openpyxl
from openpyxl.utils import get_column_letter as GL, column_index_from_string as CI

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.abspath(os.path.join(HERE, "..", "..", "..", ".."))
XLSX = os.path.join(ROOT, "Daily_Sales_Report", "1.每日销售数据表",
                    "欧盟英国库存销售表.xlsx")
ARCH = {
    8: os.path.join(ROOT, "Daily_Sales_Report", "1.每日销售数据表", "2026年每日数据",
                    "2026年8月", "2026年8月-欧盟英国库存销售表.xlsx"),
    9: os.path.join(ROOT, "Daily_Sales_Report", "1.每日销售数据表", "2026年每日数据",
                    "2026年9月", "2026年9月-欧盟英国库存销售表.xlsx"),
}
BACKUP = os.path.join(ROOT, "Daily_Sales_Report", ".workbuddy",
                      "backup_monthly_from_daily_20261027")

EU9 = ["Amazon.de", "Amazon.es", "Amazon.it", "Amazon.fr", "Amazon.ie",
       "Amazon.com.be", "Amazon.nl", "Amazon.se", "Amazon.pl"]
# 欧洲分析表 = 全部 10 张站点表之和（含土耳其；土耳其目前 8/9/10 月均为 0，
# 用 10 国口径是为了以后土耳其有量时不会漏）
EU_ALL = [ch for _, ch in [("意大利", "Amazon.it"), ("法国", "Amazon.fr"),
                           ("西班牙", "Amazon.es"), ("德国", "Amazon.de"),
                           ("土耳其", "Amazon.com.tr"), ("荷兰", "Amazon.nl"),
                           ("瑞典", "Amazon.se"), ("波兰", "Amazon.pl"),
                           ("比利时", "Amazon.com.be"), ("爱尔兰", "Amazon.ie")]]
UK = "Amazon.co.uk"
SITES = [("意大利", "Amazon.it"), ("法国", "Amazon.fr"), ("西班牙", "Amazon.es"),
         ("德国", "Amazon.de"), ("土耳其", "Amazon.com.tr"), ("荷兰", "Amazon.nl"),
         ("瑞典", "Amazon.se"), ("波兰", "Amazon.pl"), ("比利时", "Amazon.com.be"),
         ("爱尔兰", "Amazon.ie")]
DATA_FIRST_ROW = 7


# ---------------- 读日销 ----------------
# amzn.gr.<SKU>-<hash>[-LN|-VG] 是亚马逊全球开店跨站 listing 的别名 SKU，
# 实物就是主 SKU 卖的，必须归并回主 SKU，否则会被当成「表里没有的新 SKU」漏掉。
AMZN_GR = re.compile(r"^amzn\.gr\.", re.I)
SKU_HEAD = re.compile(r"^([A-Z]{1,3}\d{3,4}[A-Z]{0,3}(?:[-_ ]\d+(?=[-_ ]|$))?)")


def norm_sku(s):
    s = str(s).strip()
    if AMZN_GR.match(s):
        m = SKU_HEAD.match(s[8:])
        if m:
            return m.group(1)
    return s


def load_daily(path, month):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb["每日销售数据"]
    out = defaultdict(int)
    for r in ws.iter_rows(min_row=3, max_col=44, values_only=True):
        if not r[0] or r[42] != 1:
            continue
        if str(r[2])[:7] != "2026-%02d" % month:
            continue
        sku, ch, q = r[11], r[6], r[14]
        if not sku or not isinstance(q, (int, float)):
            continue
        out[(str(ch), norm_sku(sku))] += int(q)
    wb.close()
    return out


# ---------------- XML 工具 ----------------
def smap(z):
    rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    rid2t = dict((m[0], m[1]) for m in
                 re.findall(r'Id="(rId\d+)"[^>]*Target="([^"]+)"', rels))
    wb = z.read("xl/workbook.xml").decode("utf-8")
    return {n: ("xl/" + rid2t[r].lstrip("/"))
            for n, r in re.findall(r'<sheet [^>]*name="([^"]+)"[^>]*r:id="(rId\d+)"', wb)}


def sst_of(z):
    if "xl/sharedStrings.xml" not in z.namelist():
        return []
    sx = z.read("xl/sharedStrings.xml").decode("utf-8")
    return ["".join(re.findall(r"<t[^>]*>(.*?)</t>", si, re.S))
            for si in re.findall(r"<si>(.*?)</si>", sx, re.S)]


CELL_RE = r'<c r="([A-Z]+)%d"(?:[^>]*?)(?:/>|>.*?</c>)'


def patch_row(inner, row, targets, style_src):
    """targets = {列字母: 值}，就地改这一行。"""
    segs, gaps, pos = [], [], 0
    for m in re.finditer(CELL_RE % row, inner, re.S):
        gaps.append(inner[pos:m.start()])
        segs.append([ci_(m.group(1)), m.group(0)])
        pos = m.end()
    gaps.append(inner[pos:])
    for col, val in targets.items():
        idx = CI(col)
        ref = "%s%d" % (col, row)
        hit = next((k for k, (c, _) in enumerate(segs) if c == idx), None)
        if hit is not None:
            old = segs[hit][1]
            if "<f" in old:                       # 有公式的格子不动
                continue
            attrs = re.match(r'<c r="[A-Z]+\d+"([^>]*?)(?:/>|>)', old).group(1)
            attrs = re.sub(r'\s+t="[^"]*"', "", attrs)
            segs[hit][1] = '<c r="%s"%s><v>%d</v></c>' % (ref, attrs, val)
        elif val != 0:
            new = '<c r="%s"%s><v>%d</v></c>' % (ref, style_src, val)
            ins = len(segs)
            for k, (c, _) in enumerate(segs):
                if c > idx:
                    ins = k
                    break
            segs.insert(ins, [idx, new])
            gaps.insert(ins, "")
    return "".join(gaps[i] + segs[i][1] for i in range(len(segs))) + gaps[-1]


def ci_(s):
    n = 0
    for ch in s:
        n = n * 26 + (ord(ch) - 64)
    return n


def patch_sheet(x, plan):
    """plan = {行号: {列: 值}}，整表单遍重写。"""
    out, pos = [], 0
    for m in re.finditer(r'<row r="(\d+)"[^>]*>(.*?)</row>', x, re.S):
        r = int(m.group(1))
        out.append(x[pos:m.start()])
        if r in plan:
            out.append(x[m.start():m.start(2)])
            out.append(patch_row(m.group(2), r, plan[r], ""))
            out.append("</row>")
        else:
            out.append(m.group(0))
        pos = m.end()
    out.append(x[pos:])
    return "".join(out)


def header_map(x, sst, row=5):
    """{列字母: 表头文本}"""
    m = re.search(r'<row r="%d"[^>]*>(.*?)</row>' % row, x, re.S)
    out = {}
    if not m:
        return out
    for c in re.finditer(r'<c r="([A-Z]+)%d"([^>]*?)(?:/>|>(.*?)</c>)' % row,
                         m.group(1), re.S):
        a, b = c.group(2), c.group(3) or ""
        if 't="s"' in a:
            v = re.search(r"<v>(\d+)</v>", b)
            t = sst[int(v.group(1))] if v and int(v.group(1)) < len(sst) else ""
        elif 't="inlineStr"' in a:
            t = "".join(re.findall(r"<t[^>]*>(.*?)</t>", b, re.S))
        else:
            t = ""
        if t.strip():
            out[c.group(1)] = t.strip()
    return out


def month_cols(x, sst):
    """返回 2026 年月份的 {月: 列字母}（12月..1月 连续块，12月合计为 0 的那个）"""
    hd = header_map(x, sst)
    tot = re.search(r'<row r="6"[^>]*>(.*?)</row>', x, re.S)
    tot = tot.group(1) if tot else ""
    for col, txt in sorted(hd.items(), key=lambda kv: CI(kv[0])):
        if txt != "12月":
            continue
        c0 = CI(col)
        if not all(hd.get(GL(c0 + k)) == "%d月" % (12 - k) for k in range(12)):
            continue
        m = re.search(r'<c r="%s6"(?![0-9])[^>]*?(?:/>|>(.*?)</c>)' % col, tot, re.S)
        v = re.search(r"<v>(-?[\d.]+)</v>", m.group(1) or "") if m else None
        if v is None or abs(float(v.group(1))) < 0.5:
            return {12 - k: GL(c0 + k) for k in range(12)}
    return None


def sku_rows(x, sst):
    out = {}
    for m in re.finditer(r'<row r="(\d+)"', x):
        r = int(m.group(1))
        if r < DATA_FIRST_ROW:
            continue
        mm = re.search(r'<c r="A%d"(?![0-9])[^>]*?(?:/>|>(.*?)</c>)' % r, x, re.S)
        if not mm:
            continue
        b = mm.group(1) or ""
        a = re.search(r'<c r="A%d"([^>]*?)(?:/>|>)' % r, x)
        at = a.group(1) if a else ""
        if 't="s"' in at:
            v = re.search(r"<v>(\d+)</v>", b)
            sku = sst[int(v.group(1))] if v and int(v.group(1)) < len(sst) else None
        elif 't="inlineStr"' in at:
            sku = "".join(re.findall(r"<t[^>]*>(.*?)</t>", b, re.S))
        else:
            sku = b
        if sku and str(sku).strip():
            out[r] = str(sku).strip()
    return out


def get_cells(x, row, cols):
    """{列: 旧值}，只取数值缓存"""
    m = re.search(r'<row r="%d"[^>]*>(.*?)</row>' % row, x, re.S)
    out = {}
    if not m:
        return out
    for col in cols:
        c = re.search(r'<c r="%s%d"(?![0-9])[^>]*?(?:/>|>(.*?)</c>)' % (col, row),
                      m.group(1), re.S)
        if not c:
            out[col] = None
            continue
        v = re.search(r"<v>(-?[\d.]+)</v>", c.group(1) or "")
        out[col] = float(v.group(1)) if v else None
    return out


def set_total(x, col, val):
    """只改第 6 行合计的 <v>，保留 <f>"""
    m = re.search(r'<c r="%s6"(?![0-9])(?:[^>]*?)(?:/>|>.*?</c>)' % col, x, re.S)
    if not m:
        return x, False
    seg = m.group(0)
    if "<f" not in seg:
        return x, False
    if re.search(r"<v>-?[\d.]+</v>", seg):
        seg2 = re.sub(r"<v>-?[\d.]+</v>", "<v>%d</v>" % val, seg, count=1)
    else:
        seg2 = seg.replace("</c>", "<v>%d</v></c>" % val)
    return x[:m.start()] + seg2 + x[m.end():], True


# ---------------- 主流程 ----------------
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--apply", action="store_true")
    a = ap.parse_args()

    daily = {m: load_daily(p, m) for m, p in ARCH.items()}
    for m in (8, 9):
        print("日销 %d 月：%d 条 (渠道×SKU)" % (m, len(daily[m])))

    z = zipfile.ZipFile(XLSX)
    sm = smap(z)
    sst = sst_of(z)
    parts, report = {}, []

    for nm, sp in sm.items():
        sheet = nm.strip()
        is_ana = "每日库存分析表" in nm
        if not is_ana and sheet not in [s for s, _ in SITES]:
            continue
        x = z.read(sp).decode("utf-8")
        cols = month_cols(x, sst)
        if not cols:
            print("  ! %s 找不到 2026 月份块" % nm)
            continue
        rows = sku_rows(x, sst)

        if is_ana:
            uk = sheet.startswith("英国")
            months = [8, 9]
            def get(sku, m):
                if uk:
                    return daily[m].get((UK, sku), 0)
                return sum(daily[m].get((c, sku), 0) for c in EU_ALL)
        else:
            ch = dict(SITES)[sheet]
            months = [8, 9]
            def get(sku, m, ch=ch):
                return daily[m].get((ch, sku), 0)

        plan, tot = {}, {m: 0 for m in months}
        diffs = []
        for r, sku in rows.items():
            cell = {}
            old = get_cells(x, r, [cols[m] for m in months])
            for m in months:
                v = int(get(sku, m))
                cell[cols[m]] = v
                tot[m] += v
                ov = old.get(cols[m])
                if ov is None or abs(ov - v) > 0.5:
                    diffs.append((sku, m, ov, v))
            plan[r] = cell
        x = patch_sheet(x, plan)
        for m in months:
            x, ok = set_total(x, cols[m], tot[m])
            if not ok:
                print("  ! %s %d月 合计公式未找到" % (nm, m))
        ET.fromstring(x)
        parts[sp] = x
        report.append((nm, {m: tot[m] for m in months}, len(rows)))
        big = sorted(diffs, key=lambda t: -abs((t[2] or 0) - t[3]))[:5]
        print("  %-22s 改动 %d 格；最大 %s" % (
            nm, len(diffs),
            " / ".join("%s %d月 %s→%d" % (s, m, "-" if o is None else int(o), v)
                       for s, m, o, v in big)))

    for n in z.namelist():
        if n not in parts:
            parts[n] = z.read(n)
    z.close()

    for nm, tot, n in report:
        print("  %-24s SKU%-4d %s" % (nm, n,
              "  ".join("%d月=%-6d" % (m, v) for m, v in tot.items())))

    if a.apply:
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
        print("\n已落盘，备份:", BACKUP)
    else:
        print("\n（预览，未落盘）")


if __name__ == "__main__":
    main()
