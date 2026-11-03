# -*- coding: utf-8 -*-
"""补修「北美加拿大库存销售表.xlsx」日列宽度残留 13（2026-10-14）。

背景：上一轮 `fix_col_widths_20261014.py` 修了 11 本，唯独本簿当时被占用而跳过；
本簿日列仍残留 5 段 width=13（openpyxl 默认列宽）的区间定义：
  36..42=13  45..54=13  62..63=13  67..72=13  73=13（三张分析表同构）
真值取自 pipeline 运行前的干净备份 `.workbuddy/backup_pre_run_20261014/北美加拿大库存销售表.xlsx`：
  36..42=4.75  45..54=5  62..63=4.375(美/墨略异)  67..72=5  73=5(加=3.375 墨=2.25)

做法：**纯 zip XML 手术**——只把分析表 `<cols>` 中 width="13" 的列条目宽度改为 pre-run 真值
（逐列查表；若某区间跨宽度边界则按需拆分）。折叠/隐藏/样式/条件格式/数据一律不动。

用法：
  python fix_na_ca_colwidth_20261014.py            # 干跑
  python fix_na_ca_colwidth_20261014.py --apply    # 落盘（先备份）
"""
import os
import re
import shutil
import zipfile
import sys

DSR = r"C:/Users/dell/Desktop/work-main/sale_file_gen/Daily_Sales_Report"
CUR = os.path.join(DSR, "1.每日销售数据表", "北美加拿大库存销售表.xlsx")
PRE = os.path.join(DSR, ".workbuddy", "backup_pre_run_20261014", "北美加拿大库存销售表.xlsx")
BAK = os.path.join(DSR, ".workbuddy", "backup_na_ca_colwidth_20261014")
APPLY = "--apply" in sys.argv


def attrs(t):
    return dict(re.findall(r'([\w:]+)="([^"]*)"', t))


def sheet_map(z):
    wb = z.read("xl/workbook.xml").decode("utf-8", "replace")
    rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8", "replace")
    rid2t = {attrs(m.group(0)).get("Id", ""): attrs(m.group(0)).get("Target", "")
             for m in re.finditer(r"<Relationship\b[^>]*/>", rels)}
    out = {}
    for m in re.finditer(r"<sheet\b[^>]*/>", wb):
        a = attrs(m.group(0))
        rid = a.get("r:id") or a.get("id") or ""
        t = rid2t.get(rid, "")
        if t.startswith("/"):
            t = t[1:]
        elif not t.startswith("xl/"):
            t = "xl/" + t
        out[a.get("name", "")] = t
    return out


def width_map(xml):
    """列索引 -> 宽度（展开区间）。"""
    m = {}
    for c in re.findall(r"<col [^>]*/>", xml):
        a = attrs(c)
        if a.get("width") is None:
            continue
        mn, mx = int(a["min"]), int(a["max"])
        w = a["width"]
        for i in range(mn, mx + 1):
            m[i] = w
    return m


def rebuild_cols(xml, pre_wm):
    """把 width='13' 的 col 条目按 pre_wm 修正；跨边界则拆分。返回 (新xml, 改动列表)。"""
    changes = []

    def repl(mt):
        raw = mt.group(0)
        a = attrs(raw)
        if a.get("width") != "13":
            return raw
        mn, mx = int(a["min"]), int(a["max"])
        segs = []
        for i in range(mn, mx + 1):
            w = pre_wm.get(i)
            if w is None or w == "13":
                segs = None
                break
            if segs and segs[-1][2] == w:
                segs[-1][1] = i
            else:
                segs.append([i, i, w])
        if not segs:
            return raw  # 无真值，保持原样
        # 用原条目的其它属性拼新条目
        other = {k: v for k, v in a.items() if k not in ("min", "max", "width")}
        parts = []
        for lo, hi, w in segs:
            kv = [f'min="{lo}"', f'max="{hi}"', f'width="{w}"']
            kv += [f'{k}="{v}"' for k, v in other.items()]
            parts.append("<col " + " ".join(kv) + "/>")
        changes.append((f"{mn}..{mx}", " ".join(f"{s[0]}..{s[1]}={s[2]}" for s in segs)))
        return "".join(parts)

    new_xml = re.sub(r"<col [^>]*/>", repl, xml)
    return new_xml, changes


with zipfile.ZipFile(PRE) as zp:
    pre_sheets = {sn: width_map(zp.read(p).decode("utf-8", "replace"))
                  for sn, p in sheet_map(zp).items()}

with zipfile.ZipFile(CUR) as z:
    cur_sheets = sheet_map(z)
    order = [i.filename for i in z.infolist()]
    parts = {i.filename: z.read(i.filename) for i in z.infolist()}

print(f"模式 {'落盘' if APPLY else '干跑'}")
total = 0
for sn, part in cur_sheets.items():
    if "分析" not in sn or part not in parts:
        continue
    pre_wm = pre_sheets.get(sn)
    if not pre_wm:
        print(f"  [{sn}] pre-run 无对应表，跳过")
        continue
    xml = parts[part].decode("utf-8", "replace")
    new_xml, changes = rebuild_cols(xml, pre_wm)
    if changes:
        print(f"  [{sn}] 修正 {len(changes)} 段：")
        for old, new in changes:
            print(f"       {old}=13  →  {new}")
        parts[part] = new_xml.encode("utf-8")
        total += len(changes)
    else:
        print(f"  [{sn}] 无 width=13 残留")

if APPLY and total:
    os.makedirs(BAK, exist_ok=True)
    dst = os.path.join(BAK, os.path.basename(CUR))
    if not os.path.exists(dst):
        shutil.copy2(CUR, dst)
    tmp = CUR + ".colfix.tmp.xlsx"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zo:
        for n in order:
            zo.writestr(n, parts[n])
    assert zipfile.ZipFile(tmp).testzip() is None, "zip 校验失败"
    try:
        os.replace(tmp, CUR)
    except OSError:
        shutil.copyfile(tmp, CUR)
        os.remove(tmp)
    print(f"→ 已落盘（{total} 段），备份 {BAK}")
else:
    print(f"合计待修正 {total} 段。")
    if not APPLY:
        print("（干跑，未改盘；加 --apply 落盘）")
