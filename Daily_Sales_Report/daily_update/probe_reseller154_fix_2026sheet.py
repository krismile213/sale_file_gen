#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
probe_reseller154_fix_2026sheet.py — 补丁：重放月销总表「2026年」sheet 的 154 件修改

背景（2026-10-16）：
  probe_reseller154_apply.py 的 [C] 段有两个 bug：
  ① SheetXml.save() 用构造时持有的旧 blobs 重写整包 → sxy.save() 把 sxm 对
     「2026年」sheet 的 68 格修改回滚成旧值（年度 sheet 12 格是新的，未丢）。
  ② rels 正则漏配（Target 在 Id 前）→ KeyError rId3。
  两个 bug 均已在 probe_reseller154_apply.py 修好；本补丁用修好的 save() 逻辑
  重放「2026年」sheet 的 17 SKU × 4 列（EM/HS/KA/KJ）修改。

自证：改前 EM 列加总必须 = 3329（旧值，证明被回滚）；改后 = 3483。
"""
import os
import re
import sys
import zipfile
import xml.etree.ElementTree as ET

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
ET.register_namespace("", NS)
REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
MONTH_TOTAL = os.path.join(REPO, "菲律宾重算_202609", "月销售总表.xlsx")
APPLY = "--apply" in sys.argv

RES = {'G564': 20, 'G565': 50, 'G177': 7, 'G300': 3, 'G398': 8, 'G400': 3,
       'G461': 3, 'G343': 2, 'G510': 5, 'G484': 5, 'G426': 7, 'G356': 10,
       'G551': 11, 'G513': 1, 'G374': 1, 'G445': 8, 'G443': 10}
assert sum(RES.values()) == 154


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
        if c.get("t") in ("s", "str", "inlineStr"):
            raise AssertionError(f"{ref} 是文本单元格（{c.get('t')}）")
        return float(v.text)

    def add_num(self, ref, delta):
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

    def save(self):
        new_xml = ET.tostring(self.root, encoding="UTF-8", xml_declaration=True)
        tmp = self.path + ".tmp"
        with zipfile.ZipFile(self.path) as z:
            fresh = {n: z.read(n) for n in self.names}
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zo:
            for n in self.names:
                zo.writestr(n, new_xml if n == self.target else fresh[n])
        os.replace(tmp, self.path)


def openpyxl_sku_rows(ws, col=1, start=4):
    m = {}
    for r in range(start, ws.max_row + 1):
        v = ws.cell(r, col).value
        if v is not None and str(v).strip():
            m[str(v).strip()] = r
    return m


import openpyxl

# ---- 0. 前置自证：年度 sheet 应为改后值（sxy 最后保存，未被回滚）
sxy = SheetXml(MONTH_TOTAL, "2021-2026年销量情况")
g10, aa10 = sxy.num("G10"), sxy.num("AA10")
assert (g10, aa10) == (63197, 11308), f"年度 sheet 不是改后值: G10={g10} AA10={aa10}"
print(f"[0] 年度 sheet 已是改后值 ✓  G10={g10:.0f} AA10={aa10:.0f}")

# ---- 1. 取 SKU 行号
wb = openpyxl.load_workbook(MONTH_TOTAL, data_only=True)
m_rows = openpyxl_sku_rows(wb["2026年"], start=4)
wb.close()
missing = [s for s in RES if s not in m_rows]
assert not missing, f"找不到 SKU 行: {missing}"

# ---- 2. 前置自证：2026年 sheet EM/HS/KA 列当前应为旧值（被回滚的证据）
sxm = SheetXml(MONTH_TOTAL, "2026年")
for col, name, old in (("EM", "菲9月", 3329), ("HS", "东南亚9月", 11154), ("KA", "总计9月", 63043)):
    cur = sum(sxm.num(f"{col}{m_rows[s]}") or 0 for s in RES)
    tot = sum(sxm.num(f"{col}{r}") or 0 for r in m_rows.values())
    assert tot == old, f"{name} 列加总 {tot} != 回滚前旧值 {old}"
    print(f"[1] {name}({col}) 列加总 = {tot}（确认是被回滚的旧值）✓")

# ---- 3. 重放 17 SKU × 4 列
print()
print("[2] 重放 2026年 sheet 修改：")
for s, q in RES.items():
    r = m_rows[s]
    oe, ne = sxm.add_num(f"EM{r}", q)
    oh, nh = sxm.add_num(f"HS{r}", q)
    ok, nk = sxm.add_num(f"KA{r}", q)
    oj, nj = sxm.add_num(f"KJ{r}", q)
    print(f"    {s:<6} 行{r:<4} EM: {oe:>4.0f}→{ne:<4.0f}  HS: {oh:>4.0f}→{nh:<4.0f}  "
          f"KA: {ok:>5.0f}→{nk:<5.0f}  KJ: {oj:>6.0f}→{nj:<6.0f}")
if APPLY:
    sxm.save()
    print(f"    [已写] 2026年 sheet（{sxm.n_edit} 格）")
else:
    print(f"    [干跑] 拟改 {sxm.n_edit} 格")

# ---- 4. 复读验证
wb2 = openpyxl.load_workbook(MONTH_TOTAL, data_only=True)
ws2 = wb2["2026年"]
for col, name, exp in ((143, "EM菲9月", 3483), (227, "HS东南亚9月", 11308), (287, "KA总计9月", 63197)):
    tot = sum(ws2.cell(rr, col).value or 0 for rr in range(4, ws2.max_row + 1)
              if isinstance(ws2.cell(rr, col).value, (int, float)))
    flag = "✓" if (APPLY and tot == exp) or (not APPLY and tot != exp) else "✗"
    print(f"[3] {name} 数据行加总 = {tot}（期望 {exp}）{flag}")
wb2.close()
print()
print("完成" + ("（已写盘）" if APPLY else "（干跑，未写盘）"))
