#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
probe_reseller154_apply.py — 把监控表 RESELLER 页的 154 件补入菲律宾「线下订单」列（2026-10-16）

背景（2026-10-16 用户点名的第二轮遗漏）：
  《DAILY MONITORING - SEPTEMBER 1-13(1).xlsx》的 **RESELLER 页**有 17 行 / 154 件、
  日期全为 9 SEP 2026（落在段A 窗口 9/1~9/12 内）。第一轮重算的段A只读了 SHP/LZD/Tiktok
  三页，RESELLER 页整页漏读 → 这 154 件新旧口径都没计过（旧三报告实测无对应行：
  同 SKU 同时段只有零星平台单且数量对不上），本次补入「线下订单」列。

  用户同时确认：10 月起菲律宾销量统一在一份新订单报告里，与 145 件手工单同口径——
  「订单类型」列除 Api订单 外的标签都按手工单计入线下渠道（已改进 month_sale_file.py：
  PH_MANUAL_ORDER_TYPES / PH_API_ORDER_TYPES，未知标签打 ⚠️ 警告）。RESELLER 页本身
  是 9 月过渡期产物，10 月起不再有，故本探针只把 154 计进 9 月副本，不进 trunk 主线。

落点（与印尼 YFDZ 200 件归「印尼线下订单」同口径）：
  A. 菲次表 Philippines_202609_sales.xlsx sheet2：G列(线下) += 154、H列(月合计) += 154、行2
  B. 重算副本主表 菲律宾 sheet：BC(55) += 154、BD(56) += 154、行2
  C. 重算副本月销售总表：2026年 sheet EM(143,菲9月)/HS(227,东南亚9月)/KA(287,总计9月)/KJ(296)
     17 个 SKU 行 += 对应件数；年度 sheet 2021-2026年销量情况 同步（G10/AA10 及派生格）
  不动：Shopee 列（154 从未计入任何平台列）、印尼/泰国次表、trunk 原件销量值。

用法：python probe_reseller154_apply.py [--apply]
"""
import os
import re
import sys
import zipfile
import datetime
import xml.etree.ElementTree as ET

import pandas as pd

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
ET.register_namespace("", NS)

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
APPLY = "--apply" in sys.argv

MONITOR = r"C:/Users/dell/Desktop/DAILY MONITORING - SEPTEMBER 1-13(1).xlsx"
SUB_TABLE = os.path.join(REPO, "东南亚次表分发_202609", "Philippines_202609_sales.xlsx")
MAIN_TABLE = os.path.join(REPO, "菲律宾重算_202609", "2026年东南亚销量表Southeast Asia Sales table.xlsx")
MONTH_TOTAL = os.path.join(REPO, "菲律宾重算_202609", "月销售总表.xlsx")

# 期望值（改前状态，来自 2026-10-16 实测；不一致就中止，防止改错文件）
EXP_BEFORE = {
    "ph_bc2": 145, "ph_bd2": 3329, "ph_ba2": 1809,
    "sub_g2": 145, "sub_h2": 3329, "sub_e2": 1809,
    "m_em": 3329, "m_hs": 11154, "m_ka": 63043,
    "y_g10": 63043, "y_g14": 595415, "y_j14": 4188161,
    "y_aa10": 11154, "y_aa14": 116863, "y_ad14": 832154,
}
EXP_AFTER = {
    "ph_bc2": 299, "ph_bd2": 3483,
    "sub_g2": 299, "sub_h2": 3483,
    "m_em": 3483, "m_hs": 11308, "m_ka": 63197,
    "y_g10": 63197, "y_g14": 595569, "y_j14": 4188315,
    "y_aa10": 11308, "y_aa14": 117017, "y_ad14": 832308,
}

# ---------------------------------------------------------------- 基础工具
def col_to_num(col):
    n = 0
    for ch in col:
        n = n * 26 + (ord(ch) - 64)
    return n


def num_to_col(n):
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def split_ref(ref):
    m = re.match(r"([A-Z]+)(\d+)$", ref)
    return col_to_num(m.group(1)), int(m.group(2))


def fmt(num):
    f = float(num)
    return str(int(f)) if f.is_integer() else repr(f)


class SheetXml:
    """单 sheet 的 XML 单元格级读写：只动目标格，公式与其它单元格字节不动"""

    def __init__(self, path, sheet_name):
        self.path = path
        with zipfile.ZipFile(path) as z:
            self.names = z.namelist()
            wbxml = z.read("xl/workbook.xml").decode("utf-8")
            rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
            relmap = {rel.get("Id"): rel.get("Target")
                      for rel in ET.fromstring(rels)}
            target = None
            for nm, rid in re.findall(r'<sheet name="([^"]+)"[^>]*r:id="(rId\d+)"', wbxml):
                if nm == sheet_name:
                    target = relmap[rid]
                    break
            assert target, f"找不到 sheet {sheet_name!r}"
            t = target.lstrip("/")
            self.target = t if t.startswith("xl/") else "xl/" + t
            self.blobs = {n: z.read(n) for n in self.names}
        self.root = ET.fromstring(self.blobs[self.target].decode("utf-8"))
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
        """读数值；公式格返回 None；文本格抛异常"""
        c = self.cell(ref, create=False)
        if c is None:
            return None
        if c.find(f"{{{NS}}}f") is not None:
            return None
        v = c.find(f"{{{NS}}}v")
        if v is None or v.text is None:
            return None
        if c.get("t") in ("s", "str", "inlineStr"):
            raise AssertionError(f"{ref} 是文本单元格（{c.get('t')}），不支持数值加法")
        return float(v.text)

    def add_num(self, ref, delta):
        """ref = ref + delta（原值 None 视作 0）"""
        old = self.num(ref)
        base = 0.0 if old is None else old
        c = self.cell(ref)
        if c.get("t"):
            del c.attrib["t"]
        for f in c.findall(f"{{{NS}}}f"):
            c.remove(f)
        v = c.find(f"{{{NS}}}v")
        if v is None:
            v = ET.SubElement(c, f"{{{NS}}}v")
        v.text = fmt(base + delta)
        self.n_edit += 1
        return base, base + delta

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

    def formula(self, ref):
        c = self.cell(ref, create=False)
        if c is None:
            return None
        f = c.find(f"{{{NS}}}f")
        return f.text if f is not None else None

    def save(self):
        # ⚠️ 必须重读磁盘最新内容做底：本实例构造时持有的 blobs 可能是旧的，
        # 若同一文件已被另一个 SheetXml 实例 save 过，用旧 blobs 重写整包会
        # 把对方的修改覆盖回旧值（2026-10-16 踩坑：月销总表 2026年 sheet 68 格
        # 被年度 sheet 的后一次 save 回滚）。
        new_xml = ET.tostring(self.root, encoding="UTF-8", xml_declaration=True)
        tmp = self.path + ".tmp"
        with zipfile.ZipFile(self.path) as z:
            fresh = {n: z.read(n) for n in self.names}
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zo:
            for n in self.names:
                zo.writestr(n, new_xml if n == self.target else fresh[n])
        os.replace(tmp, self.path)


def openpyxl_sku_rows(ws, col=1, start=3):
    m = {}
    for r in range(start, ws.max_row + 1):
        v = ws.cell(r, col).value
        if v is not None and str(v).strip():
            m[str(v).strip()] = r
    return m


# ---------------------------------------------------------------- 0. 取 154 明细（自证）
print("=" * 78)
print("[0] 从监控表 RESELLER 页重新提取 9 月明细（自证，防手抄错）")
df = pd.read_excel(MONITOR, sheet_name="RESELLER", header=None)
res = []
for i in range(len(df)):
    d = str(df.iloc[i, 8])
    if "SEP 2026" in d.upper():
        res.append((str(df.iloc[i, 11]).strip(), int(float(df.iloc[i, 12]))))
assert len(res) == 17, f"RESELLER 9 月行数应为 17，实为 {len(res)}"
assert sum(q for _, q in res) == 154, f"合计应为 154，实为 {sum(q for _, q in res)}"
assert len({s for s, _ in res}) == 17, "SKU 应互不重复"
RES = dict(res)
print(f"    17 行 / {sum(RES.values())} 件 ✓  {RES}")

# ---------------------------------------------------------------- A. 菲次表
print()
print("=" * 78)
print("[A] 菲次表 Philippines_202609_sales.xlsx（sheet2，openpyxl，0 公式）")
import openpyxl

wb = openpyxl.load_workbook(SUB_TABLE, data_only=True)
ws = wb.worksheets[1]
assert "Philippines" in ws.title, f"sheet2 名字异常: {ws.title!r}"
rows = openpyxl_sku_rows(ws)
missing = [s for s in RES if s not in rows]
assert not missing, f"次表找不到 SKU 行: {missing}"
# 表头自证
hdr = {c: str(ws.cell(1, c).value) for c in range(3, 9)}
assert "Offline" in hdr[7] and "线下" in hdr[7], f"G列不是线下列: {hdr[7]!r}"
assert "TOTAL" in hdr[8].upper(), f"H列不是月合计: {hdr[8]!r}"
assert "shope" in hdr[5].lower(), f"E列不是 Shopee: {hdr[5]!r}"
# 改前状态自证
assert ws.cell(2, 7).value == EXP_BEFORE["sub_g2"], f"次表 G2={ws.cell(2,7).value}"
assert ws.cell(2, 8).value == EXP_BEFORE["sub_h2"], f"次表 H2={ws.cell(2,8).value}"
assert ws.cell(2, 5).value == EXP_BEFORE["sub_e2"], f"次表 E2={ws.cell(2,5).value}"
for s, q in RES.items():
    r = rows[s]
    old_g = ws.cell(r, 7).value or 0
    old_h = ws.cell(r, 8).value or 0
    assert isinstance(old_g, (int, float)), f"{s} 次表 G{r} 非数值: {old_g!r}"
    ws.cell(r, 7).value = old_g + q
    ws.cell(r, 8).value = old_h + q
    print(f"    {s:<6} 行{r:<4} G: {old_g:>3} → {old_g+q:<4}   H: {old_h:>4} → {old_h+q:<4}")
ws.cell(2, 7).value = ws.cell(2, 7).value + 154
ws.cell(2, 8).value = ws.cell(2, 8).value + 154
print(f"    行2: G2 {EXP_BEFORE['sub_g2']} → {EXP_AFTER['sub_g2']}   "
      f"H2 {EXP_BEFORE['sub_h2']} → {EXP_AFTER['sub_h2']}")
if APPLY:
    wb.save(SUB_TABLE)
    print("    [已写] 次表")
else:
    print("    [干跑] 未写")

# ---------------------------------------------------------------- B. 主表菲律宾 sheet
print()
print("=" * 78)
print("[B] 重算副本主表 · Philippines 菲律宾销量表（XML 手术，保 441 个公式缓存）")
sx = SheetXml(MAIN_TABLE, "Philippines 菲律宾销量表")
import openpyxl as _oxl

_wb = _oxl.load_workbook(MAIN_TABLE, data_only=True)
ph_rows = openpyxl_sku_rows(_wb["Philippines 菲律宾销量表"])
_wb.close()
missing = [s for s in RES if s not in ph_rows]
assert not missing, f"主表找不到 SKU 行: {missing}"
assert sx.num("BC2") == EXP_BEFORE["ph_bc2"], f"主表 BC2={sx.num('BC2')}"
assert sx.num("BD2") == EXP_BEFORE["ph_bd2"], f"主表 BD2={sx.num('BD2')}"
assert sx.num("BA2") == EXP_BEFORE["ph_ba2"], f"主表 BA2={sx.num('BA2')}（Shopee 不应含 154）"
for s, q in RES.items():
    r = ph_rows[s]
    ob, nb = sx.add_num(f"BC{r}", q)
    od, nd = sx.add_num(f"BD{r}", q)
    print(f"    {s:<6} 行{r:<4} BC: {ob:>4.0f} → {nb:<4.0f}   BD: {od:>5.0f} → {nd:<5.0f}")
sx.add_num("BC2", 154)
sx.add_num("BD2", 154)
print(f"    行2: BC2 {EXP_BEFORE['ph_bc2']} → {EXP_AFTER['ph_bc2']}   "
      f"BD2 {EXP_BEFORE['ph_bd2']} → {EXP_AFTER['ph_bd2']}")
if APPLY:
    sx.save()
    print(f"    [已写] 主表菲律宾 sheet（{sx.n_edit} 格）")
else:
    print(f"    [干跑] 拟改 {sx.n_edit} 格")

# ---------------------------------------------------------------- C. 月销售总表
print()
print("=" * 78)
print("[C] 重算副本月销售总表（XML 手术，保 12210 个公式缓存）")
sxm = SheetXml(MONTH_TOTAL, "2026年")
_wb2 = _oxl.load_workbook(MONTH_TOTAL, data_only=True)
m_rows = openpyxl_sku_rows(_wb2["2026年"], start=4)
_wb2.close()
missing = [s for s in RES if s not in m_rows]
assert not missing, f"月销总表找不到 SKU 行: {missing}"
# 2026年 sheet：EM(143)菲 HS(227)东南亚 KA(287)总计 KJ(296)当年总销量
for s, q in RES.items():
    r = m_rows[s]
    oe, ne = sxm.add_num(f"EM{r}", q)
    oh, nh = sxm.add_num(f"HS{r}", q)
    ok, nk = sxm.add_num(f"KA{r}", q)
    oj, nj = sxm.add_num(f"KJ{r}", q)
    print(f"    {s:<6} 行{r:<4} EM: {oe:>4.0f}→{ne:<4.0f}  HS: {oh:>4.0f}→{nh:<4.0f}  "
          f"KA: {ok:>5.0f}→{nk:<5.0f}  KJ: {oj:>6.0f}→{nj:<6.0f}")

# 年度 sheet 2021-2026年销量情况（静态值）
sxy = SheetXml(MONTH_TOTAL, "2021-2026年销量情况")
g10 = sxy.num("G10"); g9 = sxy.num("G9"); f10 = sxy.num("F10")
g14 = sxy.num("G14"); f14 = sxy.num("F14"); j14 = sxy.num("J14")
aa10 = sxy.num("AA10"); aa9 = sxy.num("AA9"); z10 = sxy.num("Z10")
aa14 = sxy.num("AA14"); z14 = sxy.num("Z14"); ad14 = sxy.num("AD14")
assert (g10, g14, j14) == (EXP_BEFORE["y_g10"], EXP_BEFORE["y_g14"], EXP_BEFORE["y_j14"]), \
    f"年度 sheet 全球组现状异常: G10={g10} G14={g14} J14={j14}"
assert (aa10, aa14, ad14) == (EXP_BEFORE["y_aa10"], EXP_BEFORE["y_aa14"], EXP_BEFORE["y_ad14"]), \
    f"年度 sheet 东南亚组现状异常: AA10={aa10} AA14={aa14} AD14={ad14}"
new = {
    "G10": g10 + 154, "H10": (g10 + 154) / f10 - 1, "I10": (g10 + 154) / g9 - 1,
    "G14": g14 + 154, "I14": (g14 + 154) / f14 - 1, "J14": j14 + 154,
    "AA10": aa10 + 154, "AB10": (aa10 + 154) / z10 - 1, "AC10": (aa10 + 154) / aa9 - 1,
    "AA14": aa14 + 154, "AC14": (aa14 + 154) / z14 - 1, "AD14": ad14 + 154,
}
for ref, val in new.items():
    old = sxy.num(ref)
    sxy.set_num(ref, val)
    print(f"    {ref:<5}: {old!r:>22} → {val!r}")
assert new["G10"] == EXP_AFTER["y_g10"] and new["AA10"] == EXP_AFTER["y_aa10"]
if APPLY:
    sxm.save()
    print(f"    [已写] 月销总表 2026年 sheet（{sxm.n_edit} 格）")
    sxy.save()
    print(f"    [已写] 月销总表年度 sheet（{sxy.n_edit} 格）")
else:
    print(f"    [干跑] 拟改 2026年 {sxm.n_edit} 格 + 年度 {sxy.n_edit} 格")

print()
print("=" * 78)
print("完成" + ("（已写盘）" if APPLY else "（干跑，未写盘）"))
