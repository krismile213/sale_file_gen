#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
probe_fix_sub_sheet1.py — 同步菲律宾次表 Sheet1（英文表）到 145+154 新口径（2026-10-16）

背景：
  `东南亚次表分发_202609/Philippines_202609_sales.xlsx` 有两个 sheet：
    - Sheet1：纯英文交付表（SKU/Model/Type/Brand + 5 渠道列 + `Sep 2026 Total`），**行2 起即数据、无合计行**
    - sheet2「Philippines 菲律宾销量表 」：中文真身（月块结构），2026-10-16 两轮修订已改到 3,483
  但 Sheet1 自 2026-10-15 起就没跟上（还是 145 挪位前的旧值：Offline=0、Total=3,094），
  10-16 的 probe_ph_offline_reclass.py / probe_reseller154_apply.py 都只改了 sheet2 ⇒ 本次补同步。

做法：按 SKU 用 sheet2 的 9 月块渠道列重写 Sheet1 的 5 渠道 + Total 列；
sheet2 有而 Sheet1 没有的 SKU（G569）补行（Model/Type/Brand 从 sheet2 对应列取）。
自证：Sheet1 各列数据行加总 == sheet2 行2（168/558/1809/649/299/3483）。

用法：python probe_fix_sub_sheet1.py [--apply]
"""
import os
import sys

import openpyxl

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
SUB_TABLE = os.path.join(REPO, "东南亚次表分发_202609", "Philippines_202609_sales.xlsx")
APPLY = "--apply" in sys.argv

EXP = {"zhongshan": 168, "lazada": 558, "shopee": 1809, "tiktok": 649,
       "offline": 299, "total": 3483}
# sheet2 9月块列 → Sheet1 英文列（sheet2: C中山 D菲LZD E菲SHP F TikTok G Offline H Total）
MAP = [(3, 5, "zhongshan"), (4, 6, "lazada"), (5, 7, "shopee"),
       (6, 8, "tiktok"), (7, 9, "offline"), (8, 10, "total")]

wb = openpyxl.load_workbook(SUB_TABLE, data_only=True)
s1, s2 = wb["Sheet1"], wb["Philippines 菲律宾销量表 "]

# sheet2 SKU → 行号（行3起；只收 G/K 开头——与 trunk 脚本 L624 交付过滤一致，
# sheet2 还有 6 个 C 开头行不属于交付范围）
import re
s2rows = {}
for r in range(3, s2.max_row + 1):
    v = s2.cell(r, 1).value
    if v is not None and str(v).strip() and re.match(r"^[GK]\d", str(v).strip()):
        s2rows[str(v).strip()] = r
# Sheet1 SKU → 行号（行2 起即数据，无合计行）
s1rows = {}
for r in range(2, s1.max_row + 1):
    v = s1.cell(r, 1).value
    if v is not None and str(v).strip():
        s1rows[str(v).strip()] = r
missing = [s for s in s2rows if s not in s1rows]
print(f"[0] sheet2 G/K SKU {len(s2rows)} 行；Sheet1 SKU {len(s1rows)} 行；Sheet1 缺: {missing}")

# 改前状态自证（陈旧态 = 145 还混在 Shopee：SHP 1954 / Offline 0 / Total 3329）
old = {}
for c2, c1, name in MAP:
    old[name] = sum(s1.cell(r, c1).value or 0 for r in s1rows.values())
print(f"    改前 Sheet1 加总: {old}")
assert old["shopee"] == 1954 and old["offline"] == 0 and old["total"] == 3329, \
    f"改前状态异常（非预期陈旧态）: {old}"

# 重写（含补行）
n_edit = 0
appended = []
for s, r2 in s2rows.items():
    if s in s1rows:
        r1 = s1rows[s]
    else:
        r1 = s1.max_row + 1
        s1rows[s] = r1
        appended.append((s, r1))
        # Model/Type/Brand 从 sheet2 取（sheet2 B/C/D 列）
        s1.cell(r1, 1).value = s
        s1.cell(r1, 2).value = s2.cell(r2, 2).value
        s1.cell(r1, 3).value = s2.cell(r2, 3).value
        s1.cell(r1, 4).value = s2.cell(r2, 4).value
        n_edit += 4
        print(f"    补行 {s} → Sheet1 行{r1}（Model/Type/Brand 从 sheet2 同步）")
    for c2, c1, name in MAP:
        v = s2.cell(r2, c2).value or 0
        s1.cell(r1, c1).value = v
        n_edit += 1

print(f"\n[1] 拟改 {n_edit} 格（含 {len(appended)} 个补行）")
if APPLY:
    wb.save(SUB_TABLE)
    print("    [已写] Sheet1")
else:
    print("    [干跑] 未写")

# 复读验证
wb2 = openpyxl.load_workbook(SUB_TABLE, data_only=True)
w1 = wb2["Sheet1"]
print("\n[2] 复读验证（数据行加总 vs sheet2 行2）:")
ok = True
for c2, c1, name in MAP:
    tot = sum(w1.cell(r, c1).value or 0 for r in range(2, w1.max_row + 1)
              if isinstance(w1.cell(r, c1).value, (int, float)))
    exp = EXP[name]
    flag = "✓" if (APPLY and tot == exp) or (not APPLY and tot != exp) else "✗"
    if APPLY and tot != exp:
        ok = False
    print(f"    {name:<10} 列{c1}: 加总={tot}  期望={exp}  {flag}")
n_row = sum(1 for r in range(2, w1.max_row + 1) if w1.cell(r, 1).value)
print(f"    数据行数 = {n_row}（sheet2 = {len(s2rows)}）")
wb2.close()
wb.close()
print("\n完成" + ("（已写盘）" if APPLY else "（干跑，未写盘）") + ("" if ok or not APPLY else "，存在失败项！"))
