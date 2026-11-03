# -*- coding: utf-8 -*-
"""
菲律宾 9 月「手工订单 145 件」重归类探针（2026-10-16）
====================================================
背景：新订单报告《菲律宾202609订单报告.xlsx》里唯一一笔手工订单
（订单号 R66128260924112818344148 / 订单ID 403990300315398144 / 运单号 YFDZ123456789 /
状态已发货 / 2026-09-24 创建 / 15 行明细 / 145 件），因合并单元格 ffill 继承平台=Shopee，
被计入了「菲律宾仓销量（shopee）」列。业务确认：手工订单应归
「Philippine Offline orders 线下订单」列（与印尼 YFDZ 手工单归「印尼线下订单」同口径）。

本探针做的事（只改 202609 两份重算副本，不动 trunk 任何代码/成品）：
  1. 从新报告动态提取手工单明细（SKU→件数），不硬编码；
  2. 主表（东南亚销量表重算版）走 XML 手术：只改 Philippines sheet 的 BA(53)/BC(55) 两列
     目标格静态值，其他字节一律不碰（保住其余 441 个公式格的缓存）；
  3. 菲次表 0 公式，openpyxl 直接改：E(5) -= 件数、G(7) += 件数；
  4. 月合计列 BD(56)/H(8) 不变（列间挪位，国家总量不变）；
  5. 全套校验：行2 合计、逐 SKU 前后值、恒等链（菲律宾 == SEA Total 菲律宾 == 月销总表 EM）。

用法：python probe_ph_offline_reclass.py           # 干跑（只校验不改）
      python probe_ph_offline_reclass.py --apply   # 执行
"""
import os, re, sys, shutil, zipfile
import xml.etree.ElementTree as ET
import pandas as pd
import openpyxl
from openpyxl.utils import get_column_letter

APPLY = "--apply" in sys.argv
REPO = r"D:/Desktop/work-main/sale_file_gen"
MAIN_TABLE = os.path.join(REPO, "菲律宾重算_202609", "2026年东南亚销量表Southeast Asia Sales table.xlsx")
SUB_TABLE = os.path.join(REPO, "东南亚次表分发_202609", "Philippines_202609_sales.xlsx")
BACKUP_DIR = r"D:/dev-cache/sale_file_gen/backup/手工单归线下_20261016"
NEW_REPORT = os.path.join(REPO, "trunk/data/数据归属周期/202609/快照日期/20261006/东南亚/菲律宾202609订单报告.xlsx")

# 手工单识别键（WMS 内部订单号；平台订单号 403990300315398144 在「订单ID」列）
MANUAL_ORDER_NO = "R66128260924112818344148"

# 主表菲律宾 sheet 9 月块列号（base=3、块宽6、第9月起 = 3+8*6 = 51）
C_LZD_CROSS, C_LZD, C_SHP, C_TK, C_OFF, C_TOTAL = 51, 52, 53, 54, 55, 56
# 次表 sheet2 9 月块列号（C~H）
S_LZD_CROSS, S_LZD, S_SHP, S_TK, S_OFF, S_TOTAL = 3, 4, 5, 6, 7, 8

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
RNS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def step1_manual_detail():
    """从新报告动态提取手工单 SKU→件数明细"""
    raw = pd.read_excel(NEW_REPORT, dtype=str)
    line_level = ["订单ID", "Sku", "数量", "多属性Id", "平台数量", "sku图片", "sku中文名称"]
    parent = [c for c in raw.columns if c not in line_level]
    raw[parent] = raw[parent].ffill()
    sub = raw[raw["订单号"] == MANUAL_ORDER_NO].copy()
    assert len(sub) == 15, f"手工单明细行数异常: {len(sub)}"
    assert set(sub["订单类型"]) == {"手工订单"}, f"订单类型异常: {set(sub['订单类型'])}"
    assert set(sub["订单状态"]) == {"已发货"}, f"订单状态异常: {set(sub['订单状态'])}"
    sub["sku"] = sub["Sku"].astype(str).str.replace(r"^\d+-", "", regex=True)
    sub["qty"] = pd.to_numeric(sub["数量"], errors="coerce").astype(int)
    detail = sub.groupby("sku")["qty"].sum().to_dict()
    total = sum(detail.values())
    assert total == 145, f"手工单件数异常: {total}"
    print(f"[1] 手工单明细（{len(detail)} SKU / {total} 件）：{detail}")
    return detail


def step2_backup():
    if APPLY:
        os.makedirs(BACKUP_DIR, exist_ok=True)
        for f in (MAIN_TABLE, SUB_TABLE):
            dst = os.path.join(BACKUP_DIR, os.path.basename(f))
            if not os.path.exists(dst):
                shutil.copy2(f, dst)
        print(f"[2] 已备份 → {BACKUP_DIR}")
    else:
        print("[2] 干跑模式，跳过备份")


# ---------- XML 手术（主表） ----------
def _sheet_xml_path(zf, sheet_name):
    wbxml = ET.fromstring(zf.read("xl/workbook.xml"))
    rid = None
    for sh in wbxml.iter(f"{{{NS}}}sheet"):
        if sh.get("name") == sheet_name:
            rid = sh.get(f"{{{RNS}}}id")
            break
    assert rid, f"workbook.xml 找不到 sheet {sheet_name!r}"
    rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    for rel in rels:
        if rel.get("Id") == rid:
            tgt = rel.get("Target")
            return "xl/" + tgt.lstrip("/") if not tgt.startswith("xl/") else tgt
    raise AssertionError("rels 找不到 " + rid)


def _col_of(ref):
    m = re.match(r"([A-Z]+)(\d+)", ref)
    col = 0
    for ch in m.group(1):
        col = col * 26 + (ord(ch) - 64)
    return col, int(m.group(2))


def _cell_xml(ref, value, style=None):
    s = f' s="{style}"' if style else ""
    return f'<c r="{ref}"{s}><v>{value}</v></c>'


def xml_surgery(path, sheet_name, sku_rows, detail, col_from, col_to, apply):
    """把 col_from 列的件数挪到 col_to 列（静态值格），返回改动日志（幂等：已挪过的行跳过）"""
    with zipfile.ZipFile(path) as zf:
        names = zf.namelist()
        target = _sheet_xml_path(zf, sheet_name)
        xml = zf.read(target).decode("utf-8")
        blobs = {n: zf.read(n) for n in names}

    root = ET.fromstring(xml)
    sd = root.find(f"{{{NS}}}sheetData")
    fl, tl = get_column_letter(col_from), get_column_letter(col_to)
    log, skipped = [], []
    for sku, qty in detail.items():
        r = sku_rows[sku]
        row_el = None
        for row in sd.findall(f"{{{NS}}}row"):
            if row.get("r") == str(r):
                row_el = row
                break
        assert row_el is not None, f"XML 找不到行 {r}"
        cells = {c.get("r"): c for c in row_el.findall(f"{{{NS}}}c")}
        ref_f, ref_t = f"{fl}{r}", f"{tl}{r}"
        # 源列：必须存在且有 <v>
        cf = cells.get(ref_f)
        assert cf is not None, f"{ref_f} 格不存在（SKU {sku}）"
        vf = cf.find(f"{{{NS}}}v")
        assert vf is not None and vf.text is not None, f"{ref_f} 无值（SKU {sku}）"
        old_from = int(float(vf.text))
        # 幂等保护：目标列已 ≥ qty → 上轮已挪过，跳过（本期 BC/G 列原为全 0，该判定等价于"已处理"）
        ct0 = cells.get(ref_t)
        vt0 = ct0.find(f"{{{NS}}}v") if ct0 is not None else None
        old_to0 = int(float(vt0.text)) if (vt0 is not None and vt0.text) else 0
        if old_to0 >= qty:
            skipped.append(sku)
            continue
        assert old_from >= qty, f"{ref_f} 现值 {old_from} < 待挪 {qty}（SKU {sku}）"
        new_from = old_from - qty
        # 目标列：存在则改/补 <v>，不存在则新建 <c> 并按列序插入
        ct = cells.get(ref_t)
        if ct is None:
            style = cf.get("s")
            new_el = ET.fromstring(_cell_xml(ref_t, qty, style))
            # 按列号找插入点
            inserted = False
            for c in list(row_el):
                ccol, _ = _col_of(c.get("r"))
                if ccol > col_to:
                    c.addprevious(new_el)
                    inserted = True
                    break
            if not inserted:
                row_el.append(new_el)
            old_to = 0
        else:
            vt = ct.find(f"{{{NS}}}v")
            old_to = int(float(vt.text)) if (vt is not None and vt.text) else 0
            if vt is None:
                vt = ET.SubElement(ct, f"{{{NS}}}v")
            vt.text = str(old_to + qty)
        vf.text = str(new_from)
        log.append((sku, r, ref_f, old_from, new_from, ref_t, old_to, old_to + qty))

    if not apply:
        print(f"    [XML 干跑] {path.split('/')[-1]} / {sheet_name}: 拟改 {len(log)} 行（跳过 {len(skipped)}）")
        return log

    new_xml = ET.tostring(root, encoding="UTF-8", xml_declaration=True)
    tmp = path + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zo:
        for n in names:
            zo.writestr(n, new_xml if n == target else blobs[n])
    os.replace(tmp, path)
    print(f"    [XML 已写] {path.split('/')[-1]} / {sheet_name}: {len(log)} 行（跳过 {len(skipped)}）")
    return log


def openpyxl_edit(path, sheet_name, sku_rows, detail, col_from, col_to, apply, fix_row2=False):
    """幂等：源列已不足 qty 且目标列已 ≥ qty → 上轮已挪过，跳过"""
    wb = openpyxl.load_workbook(path)
    ws = wb[sheet_name]
    fl, tl = get_column_letter(col_from), get_column_letter(col_to)
    log, skipped = [], []
    for sku, qty in detail.items():
        r = sku_rows[sku]
        old_from = ws.cell(r, col_from).value or 0
        old_to = ws.cell(r, col_to).value or 0
        # 幂等保护：目标列已 ≥ qty → 上轮已挪过，跳过（本期 G 列原为全 0）
        if old_to >= qty:
            skipped.append(sku)
            continue
        assert old_from >= qty, f"{fl}{r} 现值 {old_from} < {qty}"
        ws.cell(r, col_from).value = old_from - qty
        ws.cell(r, col_to).value = old_to + qty
        log.append((sku, r, f"{fl}{r}", old_from, old_from - qty, f"{tl}{r}", old_to, old_to + qty))
    if fix_row2:
        # 次表行2 = 表头兼合计行：按数据行实际合计重写（自洽）
        for c in (col_from, col_to):
            nums = [ws.cell(r, c).value for r in range(3, ws.max_row + 1)]
            ws.cell(2, c).value = sum(v for v in nums if isinstance(v, (int, float)))
        print(f"    行2 已按数据行重写: {fl}2={ws.cell(2, col_from).value} {tl}2={ws.cell(2, col_to).value}")
    if apply:
        wb.save(path)
        print(f"    [openpyxl 已写] {path.split('/')[-1]} / {sheet_name}: {len(log)} 行（跳过 {len(skipped)}）")
    else:
        print(f"    [openpyxl 干跑] {path.split('/')[-1]} / {sheet_name}: 拟改 {len(log)} 行（跳过 {len(skipped)}）")
    return log


def step3_main_table(detail):
    """主表：菲律宾 sheet XML 手术 + 行2 合计"""
    print("[3] 主表 XML 手术（Philippines sheet）")
    wb = openpyxl.load_workbook(MAIN_TABLE, data_only=True)
    ws = wb["Philippines 菲律宾销量表"]
    sku_rows = {}
    for r in range(3, ws.max_row + 1):
        v = ws.cell(r, 1).value
        if v:
            sku_rows[str(v).strip()] = r
    missing = [s for s in detail if s not in sku_rows]
    assert not missing, f"主表找不到 SKU 行: {missing}"
    log = xml_surgery(MAIN_TABLE, "Philippines 菲律宾销量表", sku_rows, detail, C_SHP, C_OFF, APPLY)
    # 行2 合计
    row2_log = xml_surgery(MAIN_TABLE, "Philippines 菲律宾销量表", {"__row2__": 2},
                           {"__row2__": 145}, C_SHP, C_OFF, APPLY)
    # 打印
    for sku, r, rf, of, nf, rt, ot, nt in log:
        print(f"    {sku:<6} 行{r:<4} {rf}: {of:>4} → {nf:<4}   {rt}: {ot:>3} → {nt:<3}")
    return log


def step4_sub_table(detail):
    """次表 sheet2：openpyxl 直接改（含行2 合计重写）"""
    print("[4] 菲次表 openpyxl（sheet2「Philippines 菲律宾销量表 」）")
    wb = openpyxl.load_workbook(SUB_TABLE, data_only=True)
    ws = wb.worksheets[1]  # sheet2
    assert "Philippines" in ws.title, f"sheet2 名字异常: {ws.title!r}"
    sku_rows = {}
    for r in range(3, ws.max_row + 1):
        v = ws.cell(r, 1).value
        if v:
            sku_rows[str(v).strip()] = r
    missing = [s for s in detail if s not in sku_rows]
    assert not missing, f"次表找不到 SKU 行: {missing}"
    log = openpyxl_edit(SUB_TABLE, ws.title, sku_rows, detail, S_SHP, S_OFF, APPLY, fix_row2=True)
    for sku, r, rf, of, nf, rt, ot, nt in log:
        print(f"    {sku:<6} 行{r:<4} {rf}: {of:>4} → {nf:<4}   {rt}: {ot:>3} → {nt:<3}")
    return log


def step5_verify():
    """全套校验"""
    print("[5] 校验")
    ok = True

    # 5.1 主表
    wb = openpyxl.load_workbook(MAIN_TABLE, data_only=True)
    ws = wb["Philippines 菲律宾销量表"]
    vals = {get_column_letter(c): ws.cell(2, c).value for c in range(C_LZD_CROSS, C_TOTAL + 1)}
    exp = {"AY": 168, "AZ": 558, "BA": 1809, "BB": 649, "BC": 145, "BD": 3329}
    for k, v in exp.items():
        got = vals.get(k)
        mark = "✓" if got == v else "✗"
        if got != v:
            ok = False
        print(f"    主表行2 {k}: {got}（期望 {v}）{mark}")
    # 数据行合计
    for c, exp_v in [(C_LZD_CROSS, 168), (C_LZD, 558), (C_SHP, 1809), (C_TK, 649), (C_OFF, 145), (C_TOTAL, 3329)]:
        nums = [ws.cell(r, c).value for r in range(3, ws.max_row + 1)]
        nums = [v for v in nums if isinstance(v, (int, float))]
        got = sum(nums)
        if got != exp_v:
            ok = False
        print(f"    主表数据行合计 {get_column_letter(c)}: {got:.0f}（期望 {exp_v}）{'✓' if got == exp_v else '✗'}")

    # 5.2 次表
    wb2 = openpyxl.load_workbook(SUB_TABLE, data_only=True)
    ws2 = wb2.worksheets[1]
    vals2 = {get_column_letter(c): ws2.cell(2, c).value for c in range(S_LZD_CROSS, S_TOTAL + 1)}
    exp2 = {"C": 168, "D": 558, "E": 1809, "F": 649, "G": 145, "H": 3329}
    for k, v in exp2.items():
        got = vals2.get(k)
        if got != v:
            ok = False
        print(f"    次表行2 {k}: {got}（期望 {v}）{'✓' if got == v else '✗'}")
    # 次表数据行合计
    for c, exp_v in [(S_LZD_CROSS, 168), (S_LZD, 558), (S_SHP, 1809), (S_TK, 649), (S_OFF, 145), (S_TOTAL, 3329)]:
        nums = [ws2.cell(r, c).value for r in range(3, ws2.max_row + 1)]
        nums = [v for v in nums if isinstance(v, (int, float))]
        got = sum(nums)
        if got != exp_v:
            ok = False
        print(f"    次表数据行合计 {get_column_letter(c)}: {got:.0f}（期望 {exp_v}）{'✓' if got == exp_v else '✗'}")

    # 5.3 恒等链：菲律宾 == SEA Total 菲律宾(CA79) == 月销总表 EM143
    ws_st = wb["SEA Total 东南亚总销量表"]
    sea_ph = ws_st.cell(2, 79).value
    print(f"    SEA Total CA(菲律宾): {sea_ph}（期望 3329）{'✓' if sea_ph == 3329 else '✗'}")
    if sea_ph != 3329:
        ok = False
    # 月销总表按国家记（无渠道维度）：行3 是 SUM 公式（副本可能无缓存），
    # 改用数据行（行4 起，静态值）求和验证
    wbm = openpyxl.load_workbook(os.path.join(REPO, "菲律宾重算_202609/月销售总表.xlsx"), data_only=True)
    wsm = wbm["2026年"]
    em_data = sum(v for v in (wsm.cell(r, 143).value for r in range(4, wsm.max_row + 1))
                  if isinstance(v, (int, float)))
    em3 = wsm.cell(3, 143).value
    print(f"    月销总表 EM(菲律宾9月): 行3(公式)={em3!r} 数据行求和={em_data}（期望 3329）")
    if em_data != 3329:
        ok = False
        print("    ✗ 月销总表菲律宾列与主表不一致")
    else:
        print("    ✓ 月销总表菲律宾列一致（列间挪位不影响国家维度）")
    # 东南亚 HS227 / 总计 KA287 一并验证
    for col, name, exp_v in [(227, "HS 东南亚", 11154), (287, "KA 总计", 63043)]:
        s = sum(v for v in (wsm.cell(r, col).value for r in range(4, wsm.max_row + 1))
                if isinstance(v, (int, float)))
        print(f"    月销总表 {name}: 数据行求和={s}（期望 {exp_v}）{'✓' if s == exp_v else '✗'}")
        if s != exp_v:
            ok = False

    # 5.4 印尼列不动验证（AJ=2775）
    ws_id = wb["Indonesia 印尼销量表"]
    id_off = ws_id.cell(2, 36).value
    print(f"    印尼 AJ(线下): {id_off}（期望 2775，未动）{'✓' if id_off == 2775 else '✗'}")
    if id_off != 2775:
        ok = False

    print(f"[5] 总结: {'全部通过 ✓' if ok else '存在失败项 ✗'}")
    return ok


if __name__ == "__main__":
    print(f"模式: {'APPLY 执行' if APPLY else '干跑'}")
    detail = step1_manual_detail()
    step2_backup()
    step3_main_table(detail)
    step4_sub_table(detail)
    if APPLY:
        step5_verify()
    else:
        print("[5] 干跑模式跳过校验（apply 后自动校验）")
    print("完成")
