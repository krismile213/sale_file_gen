# -*- coding: utf-8 -*-
"""
两处第 6 行（合计行）公式修复（2026-10-27 晚）：

【1】我上一轮补齐站点表月度时，把第 6 行合计的共享公式写坏了。
     `fill_eu_site_monthly_20261027.py` 更新缓存值时整格重写，
     丢掉了 `<f t="shared" si="N"/>`，只剩 `<v>`。
     → 10 张站点表 × 3 列（8/9/10 月）= 30 格需要把共享公式补回去。
     共享 si 从同块主格（如 AV6 的 `ref="AV6:BH6" si="3"`）取。

【2】用户反馈：货件列第 6 行没有「统计全部数量」的公式。
     实测 17 张分析表里有 8 个货件列缺 `SUBTOTAL(9, <列>7:<列>末行)`：
       北美加拿大-美国 O｜英国 I｜欧洲 I,J｜沙特阿拉伯 I｜澳大利亚 I｜罗马尼亚 I｜美客多-墨西哥 I
     成因：新货件列插入后没把第 6 行公式一起带过来（左起 I 列最容易漏）。
     → 若该列落在某个共享公式的 ref 范围内，补 `<f t="shared" si="同组"/>`；
       否则写独立 `SUBTOTAL(9,列7:列{末行})`。缓存值同步算好。

用法：
    python fix_row6_formulas_20261027.py            # 预览
    python fix_row6_formulas_20261027.py --apply    # 落盘（先备份）
"""
import argparse
import glob
import os
import re
import shutil
import zipfile
from xml.etree import ElementTree as ET

from openpyxl.utils import get_column_letter as GL, column_index_from_string as CI

HERE = os.path.dirname(os.path.abspath(__file__))
D = os.path.abspath(os.path.join(HERE, "..", "..", "..", "1.每日销售数据表"))
BACKUP = os.path.abspath(os.path.join(HERE, "..", "..", "..", ".workbuddy",
                                      "backup_row6_20261027"))
SITES = ["意大利", "法国", "西班牙", "德国", "土耳其", "荷兰",
         "瑞典", "波兰", "比利时", "爱尔兰"]
TARGET_MONTHS = (8, 9, 10)


# ---------------- 基础 ----------------
def smap(z):
    rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    rid2t = dict((m[0], m[1]) for m in re.findall(r'Id="(rId\d+)"[^>]*Target="([^"]+)"', rels))
    wb = z.read("xl/workbook.xml").decode("utf-8")
    return {n: ("xl/" + rid2t[r].lstrip("/"))
            for n, r in re.findall(r'<sheet [^>]*name="([^"]+)"[^>]*r:id="(rId\d+)"', wb)}


def sst(z):
    if "xl/sharedStrings.xml" not in z.namelist():
        return []
    sx = z.read("xl/sharedStrings.xml").decode("utf-8")
    return ["".join(re.findall(r"<t[^>]*>(.*?)</t>", si, re.S))
            for si in re.findall(r"<si>(.*?)</si>", sx, re.S)]


def cells(x, r, S):
    out = {}
    for c in re.finditer(r'<c r="([A-Z]+)%d"([^>]*?)(?:/>|>(.*?)</c>)' % r, x, re.S):
        a, b = c.group(2), c.group(3) or ""
        if 't="s"' in a:
            v = re.search(r"<v>(\d+)</v>", b)
            t = S[int(v.group(1))] if v and int(v.group(1)) < len(S) else ""
        elif 't="inlineStr"' in a:
            t = "".join(re.findall(r"<t[^>]*>(.*?)</t>", b, re.S))
        else:
            t = re.sub(r"<[^>]+>", "", b)
        out[c.group(1)] = (a, b, t.strip())
    return out


def last_row(x):
    return max([int(m.group(1)) for m in re.finditer(r'<row r="(\d+)"', x)
                if int(m.group(1)) >= 7] or [0])


def col_sums(x, lo=7):
    """单遍扫全表，返回 {列字母: 该列 lo 行起的**数值**合计}。
    ⚠️ 货件列里有大量 `t="s"` 指向空串的单元格，
       不排除会把共享字符串索引当数字加进去（曾算出 1336 万）。
    单遍实现：按行切分后再在行内找单元格，避免逐格全串搜索导致 O(n^2)。"""
    sums = {}
    for rm in re.finditer(r'<row r="(\d+)"[^>]*>(.*?)</row>', x, re.S):
        if int(rm.group(1)) < lo:
            continue
        for c in re.finditer(r'<c r="([A-Z]+)\d+"([^>]*?)(?:/>|>(.*?)</c>)',
                             rm.group(2), re.S):
            attrs, b = c.group(2), c.group(3) or ""
            if "<f" in b:
                continue
            if re.search(r't="(s|str|d|inlineStr)"', attrs):
                continue
            v = re.search(r"<v>(-?[\d.]+)</v>", b)
            if not v:
                continue
            try:
                sums[c.group(1)] = sums.get(c.group(1), 0.0) + float(v.group(1))
            except ValueError:
                pass
    return {k: int(round(v)) for k, v in sums.items()}


def col_sum(x, col, lo, hi):
    """兼容旧调用：单列合计（内部走单遍）"""
    return col_sums(x, lo).get(col, 0)


def put_cell(x, ref, inner):
    """把 ref 单元格整体替换成 inner（inner 已含 <c ...> 标签）"""
    m = re.search(r'<c r="%s"(?![0-9])(?:[^>]*?)(?:/>|>.*?</c>)' % ref, x, re.S)
    if not m:
        return x, False
    return x[:m.start()] + inner + x[m.end():], True


# ---------------- 【1】站点表月度合计公式 ----------------
def find_month_block(x, S):
    r6 = cells(x, 6, S)
    r5 = cells(x, 5, S)
    for col in sorted(r5, key=CI):
        if r5[col][2] != "12月":
            continue
        c0 = CI(col)
        if all(r5.get(GL(c0 + k), ("", "", ""))[2] == "%d月" % (12 - k) for k in range(12)):
            v = r6.get(col, ("", "", ""))[1]
            if not re.search(r"<v>(-?[\d.]+)</v>", v) or float(
                    re.search(r"<v>(-?[\d.]+)</v>", v).group(1)) == 0:
                return c0
    return None


def fix_site(x, S):
    start = find_month_block(x, S)
    if not start:
        return x, 0, "找不到 2026 月份块"
    master = cells(x, 6, S).get(GL(start), ("", "", ""))[1]
    si = re.search(r'si="(\d+)"', master)
    if not si:
        return x, 0, "主格无共享 si"
    si = si.group(1)
    n = 0
    for m in TARGET_MONTHS:
        col = GL(start + (12 - m))
        cur = cells(x, 6, S).get(col, ("", "", ""))[1]
        if "<f" in cur:
            continue
        vm = re.search(r"<v>(-?[\d.]+)</v>", cur)
        val = vm.group(1) if vm else "0"
        attrs = cells(x, 6, S).get(col, (' s="23"', "", ""))[0]
        attrs = re.sub(r'\st="[^"]*"', "", attrs)
        x, ok = put_cell(x, "%s6" % col,
                         '<c r="%s6"%s><f t="shared" si="%s"/><v>%s</v></c>'
                         % (col, attrs, si, val))
        n += ok
    return x, n, "si=%s 补 %d 格" % (si, n)


# ---------------- 【2】货件列合计公式 ----------------
def fix_ship(x, S):
    r5 = cells(x, 5, S)
    diff = None
    for col in sorted(r5, key=CI):
        if "在途差额" in r5[col][2]:
            diff = CI(col)
            break
    if not diff:
        return x, 0, "无在途差额列"
    r1 = cells(x, 1, S)
    r6 = cells(x, 6, S)
    end = last_row(x)
    sums = col_sums(x, 7)
    # 收集现有共享公式组
    groups = []
    for col, (a, b, t) in r6.items():
        m = re.search(r'<f t="shared"(?: ref="([A-Z]+\d+:[A-Z]+\d+)")? si="(\d+)"', b)
        if m and m.group(1):
            p = m.group(1).split(":")
            groups.append((CI(re.sub(r"\d", "", p[0])), CI(re.sub(r"\d", "", p[1])),
                           m.group(2)))
    fixed = []
    for ci_ in range(9, diff):
        col = GL(ci_)
        if not r1.get(col, ("", "", ""))[2]:
            continue
        if "<f" in r6.get(col, ("", "", ""))[1]:
            continue
        attrs = r6.get(col, (' s="23"', "", ""))[0]
        attrs = re.sub(r'\st="[^"]*"', "", attrs)
        hit = next((g[2] for g in groups if g[0] <= ci_ <= g[1]), None)
        val = sums.get(col, 0)
        if hit:
            inner = ('<c r="%s6"%s><f t="shared" si="%s"/><v>%d</v></c>'
                     % (col, attrs, hit, val))
        else:
            inner = ('<c r="%s6"%s><f>SUBTOTAL(9,%s7:%s%d)</f><v>%d</v></c>'
                     % (col, attrs, col, col, end, val))
        x, ok = put_cell(x, "%s6" % col, inner)
        if ok:
            fixed.append(col)
    return x, len(fixed), "货件列 I..%s 补 %s" % (GL(diff - 1), fixed or "无")


# ---------------- 【3】重算货件列合计缓存值 ----------------
def recalc_ship(x, S):
    """把所有货件列第 6 行公式的缓存 <v> 按「只算数值」重算。
    修掉 col_sum 早期版本把 t="s" 空串格的共享索引当数字累加的错。"""
    r5 = cells(x, 5, S)
    diff = None
    for col in sorted(r5, key=CI):
        if "在途差额" in r5[col][2]:
            diff = CI(col)
            break
    if not diff:
        return x, 0, "无在途差额列"
    end = last_row(x)
    sums = col_sums(x, 7)
    n, changed = 0, []
    for ci_ in range(9, diff):
        col = GL(ci_)
        ref = "%s6" % col
        m = re.search(r'<c r="%s"(?![0-9])(?:[^>]*?)(?:/>|>.*?</c>)' % ref, x, re.S)
        if not m or "<f" not in m.group(0):
            continue
        val = sums.get(col, 0)
        old = re.search(r"<v>(-?[\d.]+)</v>", m.group(0))
        if old and abs(float(old.group(1)) - val) < 0.5:
            continue
        seg = m.group(0)
        if old:
            seg = seg[:old.start()] + "<v>%d</v>" % val + seg[old.end():]
        else:
            seg = seg.replace("</c>", "<v>%d</v></c>" % val)
        x = x[:m.start()] + seg + x[m.end():]
        n += 1
        changed.append("%s6 %s->%d" % (col, old.group(1) if old else "无", val))
    return x, n, "缓存值修正 %s" % (changed or "无")


# ---------------- 主流程 ----------------
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--apply", action="store_true")
    ap.add_argument("--recalc", action="store_true",
                    help="只重算货件列第 6 行缓存值")
    a = ap.parse_args()

    tot = 0
    for p in sorted(glob.glob(os.path.join(D, "*库存销售表.xlsx"))):
        fn = os.path.basename(p)
        z = zipfile.ZipFile(p)
        S = sst(z)
        parts, plan = {}, []
        for nm, sp in smap(z).items():
            if nm.strip() not in SITES and "每日库存分析表" not in nm:
                continue
            x = z.read(sp).decode("utf-8")
            n1, d1, n2, d2 = 0, "", 0, ""
            if a.recalc:
                if "每日库存分析表" in nm:
                    x, n2, d2 = recalc_ship(x, S)
            else:
                if nm.strip() in SITES:
                    x, n1, d1 = fix_site(x, S)
                if "每日库存分析表" in nm:
                    x, n2, d2 = fix_ship(x, S)
            if n1 or n2:
                ET.fromstring(x)
                parts[sp] = x
                plan.append((nm, n1, d1, n2, d2))
                tot += n1 + n2
        for name in z.namelist():
            if name not in parts:
                parts[name] = z.read(name)
        z.close()
        if not plan:
            continue
        print("==", fn)
        for nm, n1, d1, n2, d2 in plan:
            if n1:
                print("   月度合计 %-22s %s" % (nm, d1))
            if n2:
                print("   货件合计 %-22s %s" % (nm, d2))
        if a.apply:
            os.makedirs(BACKUP, exist_ok=True)
            shutil.copy2(p, os.path.join(BACKUP, fn))
            tmp = p + ".tmp"
            with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zo:
                for name, data in parts.items():
                    zo.writestr(name, data)
            try:
                os.replace(tmp, p)
            except PermissionError:
                with open(p, "wb") as f:
                    f.write(open(tmp, "rb").read())
                os.remove(tmp)
    print("\n合计修复 %d 处  %s" % (tot, "已落盘" if a.apply else "（预览，未落盘）"))
    if a.apply:
        print("备份:", BACKUP)


if __name__ == "__main__":
    main()
