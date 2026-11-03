# -*- coding: utf-8 -*-
"""2026-10-27 日销表四项修复（用户报的问题）

1. 列宽：日列（第5行是日期序列值的列）宽度统一 5.5 —— 解决「格子有大有小 / ### 号」
2. 表头被染：`cellIs greaterThan sqref=X$1:X$1048576` 这类从第 1 行开始的规则，
   sqref 起始行改成 7 —— 表头第 5 行是日期（序列值 ~46300 > 20），会被误染
3. 错位规则：`sqref=AW7:...` 配 `AX7=1` 这种列引用右移一列的残留规则，删掉
   （阿联酋 G555 的 10/18 因此被误标红）
4. 新品标黄：新增 CF `AND($F7=0,$G7=0,$D7<>"停售",<初次销售列>7="")` → 黄，
   sqref=A7:A{末行}。复用该表「保底」规则用的黄色 dxf

用法：
    python patch_20261027_layout_cf.py --dry-run     # 只看计划，不动文件
    python patch_20261027_layout_cf.py --apply       # 落盘（先自动备份）

备份：Daily_Sales_Report/.workbuddy/backup_fix_20261027/
"""
import argparse
import html
import os
import re
import shutil
import zipfile
from xml.etree import ElementTree as ET

from openpyxl.utils import column_index_from_string as CI, get_column_letter as L

HERE = os.path.dirname(os.path.abspath(__file__))
PIPE = os.path.dirname(os.path.dirname(HERE))          # .../pipeline
DSR = os.path.dirname(PIPE)                            # .../Daily_Sales_Report
WORK = os.path.join(DSR, "1.每日销售数据表")
BACKUP = os.path.join(DSR, ".workbuddy", "backup_fix_20261027")

NEW_WIDTH = 5.5
DATE_LO, DATE_HI = 45000, 48000      # Excel 日期序列值区间（2023~2031）
FIRST_ROW = 7


# ---------------- 基础 ----------------
def sheet_map(z):
    wb = z.read("xl/workbook.xml").decode("utf-8", "replace")
    rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8", "replace")
    rm = {m.group(1): m.group(2).lstrip("/")
          for m in re.finditer(r'<Relationship[^>]*Id="([^"]+)"[^>]*Target="([^"]+)"', rels)}
    out = {}
    for m in re.finditer(r'<sheet[^>]*name="([^"]*)"[^>]*r:id="([^"]+)"', wb):
        t = rm.get(m.group(2), "")
        out[m.group(1)] = t if t.startswith("xl/") else "xl/" + t
    return out


def sst_of(z):
    if "xl/sharedStrings.xml" not in z.namelist():
        return []
    s = z.read("xl/sharedStrings.xml").decode("utf-8", "replace")
    return [html.unescape(re.sub(r"<[^>]+>", "", m))
            for m in re.findall(r"<si>(.*?)</si>", s, re.S)]


def row_cells(x, rn):
    """{列字母: (t属性, 值)}"""
    m = re.search(r'<row r="%d"[^>]*>(.*?)</row>' % rn, x, re.S)
    if not m:
        return {}
    out = {}
    for cm in re.finditer(r'<c r="([A-Z]+)%d"([^>]*)>(.*?)</c>' % rn, m.group(1), re.S):
        col, attrs, inner = cm.group(1), cm.group(2), cm.group(3)
        t = re.search(r't="(\w+)"', attrs)
        v = re.search(r"<v>(.*?)</v>", inner, re.S)
        out[col] = (t.group(1) if t else "", v.group(1) if v else None)
    return out


def header_of(x, sst):
    """第5行 → {表头文本: 列字母}"""
    out = {}
    for col, (t, v) in row_cells(x, 5).items():
        if v is None:
            continue
        if t == "s":
            try:
                out[sst[int(v)].strip()] = col
            except (ValueError, IndexError):
                pass
        else:
            out[html.unescape(v).strip()] = col
    return out


def date_cols(x):
    """第5行里数值落在日期区间的列（按列序）"""
    cols = []
    for col, (t, v) in row_cells(x, 5).items():
        if t in ("", None) and v:
            try:
                f = float(v)
            except ValueError:
                continue
            if DATE_LO <= f <= DATE_HI:
                cols.append(col)
    cols.sort(key=CI)
    return cols


def max_row_of(x):
    rs = [int(m) for m in re.findall(r'<row r="(\d+)"', x)]
    return max(rs) if rs else 0


# ---------------- 各项修复 ----------------
def fix_widths(x, dcols):
    """日列宽度统一 NEW_WIDTH，保留 style/hidden/outlineLevel 等其它属性。"""
    m = re.search(r"<cols>(.*?)</cols>", x, re.S)
    if not m:
        return x, 0
    want = set(CI(c) for c in dcols)
    if not want:
        return x, 0
    lo, hi = min(want), max(want)
    new_cols, changed = [], 0
    for tag in re.findall(r"<col [^>]*/>", m.group(1)):
        mn = int(re.search(r'min="(\d+)"', tag).group(1))
        mx = int(re.search(r'max="(\d+)"', tag).group(1))
        wm = re.search(r'width="([\d.]+)"', tag)
        width = wm.group(1) if wm else None
        # 与日列区间有交集才处理
        a, b = max(mn, lo), min(mx, hi)
        if a > b:
            new_cols.append(tag)
            continue
        # 区间外的前后两段原样保留
        if mn < a:
            new_cols.append(re.sub(r'(min|max)="\d+"',
                                   lambda mm: '%s="%d"' % (mm.group(1), mn if mm.group(1) == "min" else a - 1),
                                   tag, count=2))
        seg = tag
        if a != mn or b != mx:
            seg = re.sub(r'min="\d+"', 'min="%d"' % a, seg)
            seg = re.sub(r'max="\d+"', 'max="%d"' % b, seg)
        if width is not None and abs(float(width) - NEW_WIDTH) > 1e-6:
            seg = re.sub(r'width="[\d.]+"', 'width="%g"' % NEW_WIDTH, seg)
            changed += 1
        elif width is None:
            seg = seg.replace("<col ", '<col width="%g" ' % NEW_WIDTH, 1)
            changed += 1
        new_cols.append(seg)
        if b < mx:
            new_cols.append(re.sub(r'(min|max)="\d+"',
                                   lambda mm: '%s="%d"' % (mm.group(1), b + 1 if mm.group(1) == "min" else mx),
                                   tag, count=2))
    x = x[:m.start(1)] + "".join(new_cols) + x[m.end(1):]
    return x, changed


def fix_header_cf(x):
    """`X$1:X$1048576` / `X1:X1048576` → 起始行改 7（只动整列规则）。"""
    n = 0

    def repl(m):
        nonlocal n
        head, sq, tail = m.group(1), m.group(2), m.group(3)
        parts = []
        hit = False
        for seg in sq.split():
            mm = re.match(r"^\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)$", seg)
            if mm and int(mm.group(2)) < FIRST_ROW and int(mm.group(4)) > 100000:
                seg = "%s%d:%s$%s" % (mm.group(1), FIRST_ROW, mm.group(3), mm.group(4))
                hit = True
            parts.append(seg)
        if hit:
            n += 1
        return head + 'sqref="' + " ".join(parts) + '"' + tail

    x = re.sub(r'(<conditionalFormatting[^>]*?)sqref="([^"]*)"([^>]*>)', repl, x)
    return x, n


def drop_misaligned(x):
    """删掉「sqref 在某列，公式却引用右邻列」的错位规则。"""
    n = 0
    out = []
    pos = 0
    for m in re.finditer(r'<conditionalFormatting[^>]*sqref="([^"]*)"[^>]*>(.*?)</conditionalFormatting>',
                         x, re.S):
        sq = m.group(1)
        body = m.group(2)
        fm = re.search(r"<formula>(.*?)</formula>", body, re.S)
        if not fm:
            continue
        f = html.unescape(fm.group(1))
        mm = re.match(r"^([A-Z]+)(\d+)=(\d+)$", f.strip())
        if not mm:
            continue
        # sqref 的单列 vs 公式列
        sm = re.match(r"^\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)$", sq.strip())
        if not sm or sm.group(1) != sm.group(3):
            continue
        if CI(mm.group(1)) == CI(sm.group(1)) + 1:
            out.append((m.start(), m.end()))
            n += 1
    for a, b in reversed(out):
        x = x[:a] + x[b:]
    return x, n


def add_new_product_rule(x, first_col, yellow_dxf):
    """新增「新品标黄」规则。first_col = 初次销售日期/时间列。"""
    if not yellow_dxf:
        return x, 0
    mrow = max_row_of(x)
    if mrow < FIRST_ROW:
        return x, 0
    pri = max([int(p) for p in re.findall(r'priority="(\d+)"', x)] or [0]) + 1
    f = ('AND($F%d=0,$G%d=0,$D%d&lt;&gt;&quot;停售&quot;,$%s%d=&quot;&quot;)'
         % (FIRST_ROW, FIRST_ROW, FIRST_ROW, first_col, FIRST_ROW))
    block = ('<conditionalFormatting sqref="A%d:A%d">'
             '<cfRule type="expression" dxfId="%s" priority="%d"><formula>%s</formula></cfRule>'
             '</conditionalFormatting>' % (FIRST_ROW, mrow, yellow_dxf, pri, f))
    # 插在最后一个 conditionalFormatting 之后（schema: CF 在 dataValidations 之前）
    last = None
    for mm in re.finditer(r"</conditionalFormatting>", x):
        last = mm
    if last is None:
        return x, 0
    x = x[:last.end()] + block + x[last.end():]
    return x, 1


def find_yellow_dxf(x):
    """找「保底」规则用的 dxfId（该表的黄）。"""
    for m in re.finditer(r'<cfRule[^>]*dxfId="(\d+)"[^>]*>(.*?)</cfRule>', x, re.S):
        if "保底" in html.unescape(m.group(2)):
            return m.group(1)
    return None


# ---------------- 主流程 ----------------
def process(path, apply=False):
    fn = os.path.basename(path)
    z = zipfile.ZipFile(path)
    sm = sheet_map(z)
    sst = sst_of(z)
    targets = [(k, v) for k, v in sm.items() if "每日库存分析表" in k]
    if not targets:
        z.close()
        return []
    plan = []
    parts = {}
    for nm, sp in targets:
        x = z.read(sp).decode("utf-8", "replace")
        hdr = header_of(x, sst)
        dcols = date_cols(x)
        fcol = hdr.get("初次销售日期") or hdr.get("初次销售时间")
        x2, nw = fix_widths(x, dcols)
        x2, nh = fix_header_cf(x2)
        x2, nd = drop_misaligned(x2)
        yellow = find_yellow_dxf(x2)
        x2, nn = add_new_product_rule(x2, fcol, yellow) if fcol else (x2, 0)
        ET.fromstring(x2)                    # 良构校验，不通过直接抛
        parts[sp] = x2
        plan.append(dict(sheet=nm, date_cols=len(dcols), width=nw, header_cf=nh,
                         dropped=nd, new_rule=nn, first_col=fcol, yellow=yellow))
    # 其它部件原样
    for n in z.namelist():
        if n not in parts:
            parts[n] = z.read(n)
    z.close()

    if apply:
        os.makedirs(BACKUP, exist_ok=True)
        shutil.copy2(path, os.path.join(BACKUP, fn))
        tmp = path + ".tmp"
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zo:
            for n, data in parts.items():
                zo.writestr(n, data)
        try:
            os.replace(tmp, path)
        except PermissionError:
            with open(path, "wb") as f:
                f.write(open(tmp, "rb").read())
            os.remove(tmp)
    return plan


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--apply", action="store_true")
    ap.add_argument("--dry-run", action="store_true")
    a = ap.parse_args()
    apply = a.apply and not a.dry_run
    print("模式：%s\n" % ("落盘" if apply else "dry-run（不动文件）"))
    for fn in sorted(os.listdir(WORK)):
        if not fn.endswith(".xlsx") or fn.startswith("~$"):
            continue
        p = os.path.join(WORK, fn)
        try:
            plan = process(p, apply=apply)
        except PermissionError as e:
            print("!! %-30s 文件被占用（Excel 开着？）%s" % (fn[:30], e))
            continue
        except Exception as e:
            print("!! %-30s 失败 %r" % (fn[:30], e))
            continue
        if not plan:
            print("-- %-30s 无分析表" % fn[:30])
            continue
        for pl in plan:
            print("%-30s %-16s 日列%2d 宽改%2d 表头CF%2d 删错位%d 新品规则%d 初次列=%s 黄dxf=%s"
                  % (fn[:30], pl["sheet"][:16], pl["date_cols"], pl["width"],
                     pl["header_cf"], pl["dropped"], pl["new_rule"],
                     pl["first_col"], pl["yellow"]))
    if apply:
        print("\n备份在：%s" % BACKUP)


if __name__ == "__main__":
    main()
