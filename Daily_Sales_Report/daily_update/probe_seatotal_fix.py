#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
probe_seatotal_fix.py — 全量修复 SEA Total 月块列位 bug 损伤（2026-10-16，v2）

背景：
  southeast_asia_sale_file.py 旧版 SEA Total 写入用 block_start(7, 9, m)（base 误用
  年度区 G 列、且布局当 [国x8, 合计]），使当月国家值写到上月块、当月块国家列恒 0。
  模板真实结构（行1 表头实锤）：月块从 O(15) 起、块宽 9、**合计列在块首**
  （BZ1='8月东南亚销量'、CI1='9月东南亚销量' 均为月标签=合计列，中间夹 8 个国家列）。
  合计列位置 7+(m-1)*9+8 == 15+(m-1)*9 永远巧合撞对；国家列永远错位到上月块。
  每月跑污染上月块 + 模板继承 → 1~8 月块国家列全是污染值（滚动累积）：
    9月跑污染 8月块(CA~CH)、8月跑污染 7月块(BR~BY)……2月跑污染 1月块(P~W)、
    1月跑污染年度区 G~N。
  脚本已于 2026-10-16 修好（base=15、合计列块首），已生成文件需原地修复。

本探针（--trunk 修 trunk 202609 成品；否则修重算副本）：
  A. SEA Total 1~9 月块全部重写（513 SKU 行 + 行2）：
     合计列（块首）← 8 国和；8 国列 ← 各国 sheet 该月块合计列（块尾）同 SKU 值
     真值源自证：sum(各国 sheet 该月行2) == SEA Total 合计列行2 现值（模板值）
     （9月副本例外：== 现值+154，因菲律宾 sheet 已补 154 而 SEA Total 未补）
  B. 年度 sheet「2023-2026年销量情况 Sales」公式缓存全量同步
     （该工作簿 calcPr 无 fullCalcOnLoad，副本带缓存、trunk 为空缓存 <v/>）：
     行3~10（8 国行）AM~AU（1~9月公式列，范围起始列逐月+9）cache → 当月真值
     行2（SUM 公式行）AM2~AU2 cache → 当月 8 国和
     ⚠️ 模板既有错位（非本次引入，不改公式结构）：公式列比行1 日期标签早一个月
        （如 AT 列日期=9月但公式范围从 8月块 CA 起）；cache 按公式重算结果同步，
        保证公式与缓存一致。AL 列（年度区 G~N，陈旧区）不动。

用法：python probe_seatotal_fix.py [--apply] [--trunk]
"""
import os
import re
import sys
import zipfile
import xml.etree.ElementTree as ET

import openpyxl

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
ET.register_namespace("", NS)
REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
APPLY = "--apply" in sys.argv
TRUNK = "--trunk" in sys.argv

SEA_TOTAL = "SEA Total 东南亚总销量表"
YEAR_SHEET = "2023-2026年销量情况 Sales"
MONTHS = list(range(1, 10))          # 修复 1~9 月块
BLOCKS = {
    "Philippines 菲律宾销量表": 6, "Thailand 泰国销量表": 5, "Indonesia 印尼销量表": 4,
    "Singapore 新加坡销量表": 3, "Vietnam Lazada 越南销销量表": 1,
    "Malaysia Lazada 马来西亚销量表": 1, "Malaysia Shopee 虾皮跨境销量表": 9,
    "Group buying distributor团购经销商": 1,
}
COUNTRY_ORDER = ["菲律宾", "泰国", "印尼", "新加坡", "越南", "马来Lazada", "虾皮跨境", "团购"]
SHEET_BY_COUNTRY = dict(zip(COUNTRY_ORDER, BLOCKS.keys()))
YEAR_ROW_LABELS = ["菲律宾Philippines", "泰国Thailand", "印尼Indonesia", "新加坡Singapore",
                   "越南Vietnam", "马来西亚Malaysia Lazada", "虾皮跨境Malaysia Shopee",
                   "团购经销商Group buying distributor"]
YR_ROWS = {c: 3 + i for i, c in enumerate(COUNTRY_ORDER)}
YEAR_COL_BY_MONTH = {m: 38 + m for m in MONTHS}      # 1月=AM(39) … 9月=AU(47)

if TRUNK:
    TARGET = os.path.join(REPO, "trunk/result/东南亚销量表/202609",
                          "2026年东南亚销量表Southeast Asia Sales table.xlsx")
    EXP_9 = [3123, 3238, 4026, 174, 15, 11, 361, 0]  # trunk 口径（无 145/154）
else:
    TARGET = os.path.join(REPO, "菲律宾重算_202609",
                          "2026年东南亚销量表Southeast Asia Sales table.xlsx")
    EXP_9 = [3483, 3238, 4026, 174, 15, 11, 361, 0]  # 副本口径（145+154）


# ---------------------------------------------------------------- XML 工具
def split_ref(ref):
    m = re.match(r"([A-Z]+)(\d+)$", ref)
    col = 0
    for ch in m.group(1):
        col = col * 26 + (ord(ch) - 64)
    return col, int(m.group(2))


def fmt(num):
    f = float(num)
    return str(int(f)) if f.is_integer() else repr(f)


class SheetXml:
    """单 sheet XML 单元格级读写（save 前重读磁盘，防同文件多实例互相覆盖）"""

    def __init__(self, path, sheet_name):
        self.path = path
        with zipfile.ZipFile(path) as z:
            self.names = z.namelist()
            wbxml = z.read("xl/workbook.xml").decode("utf-8")
            rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
            relmap = {rel.get("Id"): rel.get("Target") for rel in ET.fromstring(rels)}
            target = None
            for nm, rid in re.findall(r'<sheet name="([^"]+)"[^>]*r:id="(rId\d+)"', wbxml):
                if nm == sheet_name:
                    target = relmap[rid]
                    break
            assert target, f"找不到 sheet {sheet_name!r}"
            t = target.lstrip("/")
            self.target = t if t.startswith("xl/") else "xl/" + t
        self.root = ET.fromstring(zipfile.ZipFile(path).read(self.target).decode("utf-8"))
        self.sd = self.root.find(f"{{{NS}}}sheetData")
        self.rows = {int(r.get("r")): r for r in self.sd.findall(f"{{{NS}}}row")}
        self.n_edit = 0

    def _row(self, rnum):
        if rnum not in self.rows:
            row = ET.Element(f"{{{NS}}}row", {"r": str(rnum)})
            done = False
            for r in self.sd.findall(f"{{{NS}}}row"):
                if int(r.get("r")) > rnum:
                    r.addprevious(row)
                    done = True
                    break
            if not done:
                self.sd.append(row)
            self.rows[rnum] = row
        return self.rows[rnum]

    def cell(self, ref, create=True):
        cnum, rnum = split_ref(ref)
        row = self._row(rnum)
        for c in row.findall(f"{{{NS}}}c"):
            if c.get("r") == ref:
                return c
        if not create:
            return None
        cell = ET.Element(f"{{{NS}}}c", {"r": ref})
        for c in list(row):
            cc, _ = split_ref(c.get("r"))
            if cc > cnum:
                c.addprevious(cell)
                return cell
        row.append(cell)
        return cell

    def num(self, ref):
        c = self.cell(ref, create=False)
        if c is None:
            return None
        if c.find(f"{{{NS}}}f") is not None:
            return None
        v = c.find(f"{{{NS}}}v")
        if v is None or v.text is None:
            return None
        return float(v.text)

    def set_num(self, ref, val):
        c = self.cell(ref)
        if c.get("t"):
            del c.attrib["t"]
        for f in c.findall(f"{{{NS}}}f"):
            c.remove(f)
        v = c.find(f"{{{NS}}}v")
        if v is None:
            v = ET.SubElement(c, f"{{{NS}}}v")
        v.text = fmt(val)
        self.n_edit += 1

    def cache(self, ref):
        c = self.cell(ref, create=False)
        if c is None:
            return None
        v = c.find(f"{{{NS}}}v")
        return v.text if (v is not None and v.text) else None

    def set_cache(self, ref, val):
        """只改公式格的缓存值 <v>，保留公式 <f> 不动"""
        c = self.cell(ref, create=False)
        assert c is not None, f"{ref} 不存在"
        f = c.find(f"{{{NS}}}f")
        assert f is not None, f"{ref} 不是公式格"
        v = c.find(f"{{{NS}}}v")
        if v is None:
            v = ET.SubElement(c, f"{{{NS}}}v")
        v.text = fmt(val)
        if c.get("t") in ("s", "str", "inlineStr"):
            del c.attrib["t"]
        self.n_edit += 1

    def save(self):
        new_xml = ET.tostring(self.root, encoding="UTF-8", xml_declaration=True)
        tmp = self.path + ".tmp"
        with zipfile.ZipFile(self.path) as z:
            fresh = {n: z.read(n) for n in self.names}
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zo:
            for n in self.names:
                zo.writestr(n, new_xml if n == self.target else fresh[n])
        os.replace(tmp, self.path)


def col_letter(n):
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


# ---------------------------------------------------------------- 0. 取各国真值
print("=" * 78)
print(f"[0] 目标：{'trunk 202609 成品' if TRUNK else '重算副本（菲律宾重算_202609）'}")
print(f"    {TARGET}")
wb = openpyxl.load_workbook(TARGET, data_only=True)
assert SEA_TOTAL in wb.sheetnames and YEAR_SHEET in wb.sheetnames

# per_month[m][country] = ({sku: 值}, 行2值)
per_month = {m: {} for m in MONTHS}
for c, sheet in SHEET_BY_COUNTRY.items():
    ws = wb[sheet]
    w = BLOCKS[sheet]
    for m in MONTHS:
        tcol = 3 + (m - 1) * w + w - 1          # 该国 sheet 该月合计列（块尾）
        vals = {}
        for r in range(3, ws.max_row + 1):
            v = ws.cell(r, 1).value
            if v is None or not str(v).strip():
                continue
            vals[str(v).strip()] = ws.cell(r, tcol).value or 0
        per_month[m][c] = (vals, ws.cell(2, tcol).value or 0)

for m in MONTHS:
    row2s = [per_month[m][c][1] for c in COUNTRY_ORDER]
    print(f"    {m:>2}月各国行2 = {row2s}  和 = {sum(row2s)}")

# ---------------------------------------------------------------- 1. 前置自证
sx = SheetXml(TARGET, SEA_TOTAL)
sea_rows = {}
for r in range(3, 516):
    v = wb[SEA_TOTAL].cell(r, 1).value
    if v is not None and str(v).strip():
        sea_rows[str(v).strip()] = r
assert len(sea_rows) == 513, f"SKU 行数 {len(sea_rows)} != 513"
all_skus = set()
for m in MONTHS:
    for c in COUNTRY_ORDER:
        all_skus |= set(per_month[m][c][0])
assert not (all_skus - set(sea_rows)), f"缺行: {sorted(all_skus - set(sea_rows))[:5]}"
print(f"    数据行自证：513 行，无缺失 ✓")

month_sum = {}      # m -> 8 国和（新值）
for m in MONTHS:
    blk = 15 + (m - 1) * 9                     # 合计列（块首）
    row2s = [per_month[m][c][1] for c in COUNTRY_ORDER]
    month_sum[m] = sum(row2s)
    cur_tot = sx.num(f"{col_letter(blk)}2")
    if m == 9 and not TRUNK:
        # 副本 9月：CI2 现值可能是 11154（v1 修复前）或 11308（v1 已修）
        assert month_sum[m] in ((cur_tot or 0), (cur_tot or 0) + 154), \
            f"9月和 {month_sum[m]} 与 CI2 {cur_tot} 不符（也不符 +154）"
        print(f"    9月自证：各国和 {month_sum[m]} == CI2 {cur_tot:.0f}"
              f"{'（v1 已修）' if month_sum[m] == cur_tot else '（+154 待修）'} ✓")
    else:
        assert month_sum[m] == cur_tot, \
            f"{m}月：各国 sheet 和 {month_sum[m]} != SEA Total 合计列现值 {cur_tot}"
        print(f"    {m}月自证：各国 sheet 行2 和 {month_sum[m]} == SEA Total "
              f"{col_letter(blk)}2 现值 {cur_tot:.0f} ✓")
assert [per_month[9][c][1] for c in COUNTRY_ORDER] == EXP_9, \
    f"9月各国值不符: {[per_month[9][c][1] for c in COUNTRY_ORDER]}"

# 数据行抽查：8月块国列行和 == BZ 合计列（若 v1 已修则为修复态自洽；
# 若未修则为污染态 CA~CH==CI。两种都说明数据行结构完好）
bad_fixed = [s for s, r in list(sea_rows.items())[:200]
             if sum(sx.num(f"{col_letter(79 + i)}{r}") or 0 for i in range(8))
             != (sx.num(f"BZ{r}") or 0)]
bad_polluted = [s for s, r in list(sea_rows.items())[:200]
                if sum(sx.num(f"{col_letter(79 + i)}{r}") or 0 for i in range(8))
                != (sx.num(f"CI{r}") or 0)]
assert not bad_fixed or not bad_polluted, \
    f"8月块数据行既非修复态也非污染态: fixed_bad={bad_fixed[:3]} polluted_bad={bad_polluted[:3]}"
state = "修复态（v1 已修）" if not bad_fixed else "污染态（v1 未修/trunk）"
print(f"    8月块数据行抽查（前200行）：{state}自洽 ✓")

# ---------------------------------------------------------------- 2. 写 SEA Total
print()
print("=" * 78)
print("[1] 重写 SEA Total 1~9 月块（513 行 × 9 列 × 9 月 + 行2）")
for s, r in sea_rows.items():
    for m in MONTHS:
        blk = 15 + (m - 1) * 9
        vals = [per_month[m][c][0].get(s, 0) for c in COUNTRY_ORDER]
        for i in range(8):
            sx.set_num(f"{col_letter(blk + 1 + i)}{r}", vals[i])
        sx.set_num(f"{col_letter(blk)}{r}", sum(vals))
for m in MONTHS:
    blk = 15 + (m - 1) * 9
    for i in range(8):
        sx.set_num(f"{col_letter(blk + 1 + i)}2", per_month[m][COUNTRY_ORDER[i]][1])
    sx.set_num(f"{col_letter(blk)}2", month_sum[m])
print(f"    拟改 {sx.n_edit} 格；各月合计列行2 → "
      f"{ {m: month_sum[m] for m in MONTHS} }")
if APPLY:
    sx.save()
    print(f"    [已写] SEA Total（{sx.n_edit} 格）")
else:
    print("    [干跑] 未写")

# ---------------------------------------------------------------- 3. 年度 sheet 缓存
print()
print("=" * 78)
print("[2] 年度 sheet 公式缓存全量同步（AM~AU = 1~9月公式列）")
sy = SheetXml(TARGET, YEAR_SHEET)
for i, c in enumerate(COUNTRY_ORDER):
    lbl = wb[YEAR_SHEET].cell(3 + i, 1).value
    assert lbl == YEAR_ROW_LABELS[i], f"年度 sheet 行{3+i} 标签异常: {lbl!r}"
print("    行3~10 国家标签 ✓")
before = {}
for m in MONTHS:
    col = col_letter(YEAR_COL_BY_MONTH[m])
    before[f"{col}3"] = sy.cache(f"{col}3")
print(f"    改前行3（菲律宾）缓存: {before}")
for m in MONTHS:
    col = col_letter(YEAR_COL_BY_MONTH[m])
    for i, c in enumerate(COUNTRY_ORDER):
        sy.set_cache(f"{col}{YR_ROWS[c]}", per_month[m][c][1])
    # 行2 SUM 缓存 = 当月 8 国和
    cur = sy.cache(f"{col}2")
    sy.set_cache(f"{col}2", month_sum[m])
    print(f"    {m}月列 {col}: 行3~10 → 各国当月值；行2 {cur} → {month_sum[m]}")
if APPLY:
    sy.save()
    print(f"    [已写] 年度 sheet（{sy.n_edit} 格）")
else:
    print(f"    [干跑] 拟改 {sy.n_edit} 格")
wb.close()

print()
print("=" * 78)
print("完成" + ("（已写盘）" if APPLY else "（干跑，未写盘）"))
