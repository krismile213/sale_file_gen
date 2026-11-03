# -*- coding: utf-8 -*-
"""欧洲表红色格子审计 + 导出可核对清单（纯 XML 解析，避开 13MB 簿 SIGTERM）

规则（欧洲 D 组，来自 sheet XML 实测）：
  red_hi prio=39   sqref=AS7:CE1476   AND(AS7>$AO7*3, $AO7>1.5)
  red_lo prio=3647 sqref=AS7:CE1476   AND($AO7>0,$AO7<=1.5,AS7>=2,AS7>$AO7*4)
  blue   prio=40   sqref=AS7:CE1476   AND(AS7<$AO7/3, AS7>0)
  T列红 prio=1606  sqref=AN7:AN1476   $AN7>$AO7*1.5     (可售列，无阈值)
  T列蓝 prio=1481  sqref=AN7:AN1476   AND($AN7<$AO7/2,$AO7>2)

输出：xlsx 到 断货分析/染色核对/ 下
"""
import os, re, zipfile, sys
from datetime import datetime, timedelta

HERE = os.path.dirname(os.path.abspath(__file__))
DSR = os.path.dirname(os.path.dirname(HERE))
WORK = os.path.join(DSR, "1.每日销售数据表")
OUTDIR = os.path.join(DSR, "断货分析", "染色核对")
BOOK = "欧盟英国库存销售表"
SHEET = "欧洲-每日库存分析表"

EPOCH = datetime(1899, 12, 30)
def serial2date(s):
    try:
        return (EPOCH + timedelta(days=float(s))).strftime("%Y-%m-%d")
    except Exception:
        return str(s)


def col2num(s):
    n = 0
    for ch in s:
        n = n * 26 + (ord(ch) - 64)
    return n


def num2col(n):
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def sheet_path(z, sheetname):
    wb = z.read("xl/workbook.xml").decode("utf-8")
    rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    rid2t = {}
    for m in re.finditer(r"<Relationship\b[^>]*/?>", rels):
        a = dict(re.findall(r'([\w:]+)="([^"]*)"', m.group(0)))
        rid2t[a.get("Id", "")] = a.get("Target", "")
    for m in re.finditer(r"<sheet\b[^>]*/?>", wb):
        a = dict(re.findall(r'([\w:]+)="([^"]*)"', m.group(0)))
        if a.get("name") == sheetname:
            t = rid2t.get(a.get("r:id", "") or a.get("id", ""), "")
            t = t.lstrip("/")
            if not t.startswith("xl/"):
                t = "xl/" + t
            return t
    return None


def shared_strings(z):
    try:
        sx = z.read("xl/sharedStrings.xml").decode("utf-8")
    except KeyError:
        return []
    out = []
    for si in re.finditer(r"<si>(.*?)</si>", sx, re.S):
        txt = "".join(re.findall(r"<t[^>]*>(.*?)</t>", si.group(1), re.S))
        txt = (txt.replace("&lt;", "<").replace("&gt;", ">")
               .replace("&quot;", '"').replace("&amp;", "&"))
        out.append(txt)
    return out


def parse_sheet(xml):
    """{row: {col: dict(v, t, f)}}"""
    out = {}
    for m in re.finditer(r'<row r="(\d+)"[^>]*>(.*?)</row>', xml, re.S):
        r = int(m.group(1))
        cells = {}
        for c in re.finditer(r'<c r="([A-Z]+)(\d+)"([^>]*?)(?:/>|>(.*?)</c>)', m.group(2), re.S):
            col, body, attrs = c.group(1), (c.group(4) or ""), c.group(3)
            t = re.search(r't="([^"]*)"', attrs)
            typ = t.group(1) if t else "n"
            f = re.search(r"<f[^>]*>(.*?)</f>", body, re.S)
            v = re.search(r"<v>(.*?)</v>", body, re.S)
            cells[col] = dict(t=typ, f=f.group(1) if f else None,
                              v=v.group(1) if v else None)
        out[r] = cells
    return out


def main():
    p = os.path.join(WORK, BOOK + ".xlsx")
    z = zipfile.ZipFile(p)
    sp = sheet_path(z, SHEET)
    xml = z.read(sp).decode("utf-8")
    ss = shared_strings(z)
    cells = parse_sheet(xml)

    def get(r, col):
        c = cells.get(r, {}).get(col)
        if not c:
            return None
        v, typ = c["v"], c["t"]
        if v is None:
            return None
        if typ == "s":
            try:
                i = int(v)
                return ss[i] if 0 <= i < len(ss) else None
            except ValueError:
                return v
        if typ in ("str", "inlineStr"):
            return v
        try:
            return float(v)
        except ValueError:
            return v

    def has_formula(r, col):
        c = cells.get(r, {}).get(col)
        return bool(c and c["f"])

    # 列位
    C_SKU, C_NAME, C_AVG = "A", "B", "AO"
    C_T = "AN"   # 可售库存列（T列规则用 AN）
    D0, D1 = col2num("AS"), col2num("CE")
    max_row = max(cells.keys())

    # 表头日期
    hdr = {}
    for n in range(D0, D1 + 1):
        hdr[num2col(n)] = get(5, num2col(n))

    # 锚点日：数据透视表!$B$11
    anchor = None
    try:
        piv = sheet_path(z, "数据透视表")
        if piv:
            pxml = z.read(piv).decode("utf-8")
            pc = parse_sheet(pxml)
            b11 = pc.get(11, {}).get("B")
            if b11 and b11["v"]:
                anchor = serial2date(b11["v"])
    except Exception:
        pass

    # 缓存值可信度抽查：日列是否 f+v 并存
    sample_ok = 0
    sample_f_only = 0
    for r in range(7, min(max_row, 60) + 1):
        for n in range(D0, D1 + 1):
            col = num2col(n)
            c = cells.get(r, {}).get(col)
            if not c:
                continue
            if c["f"] and c["v"] is not None:
                sample_ok += 1
            elif c["f"] and c["v"] is None:
                sample_f_only += 1

    print("=" * 96)
    print(f"{BOOK} / {SHEET}")
    print(f"sheet XML: {sp}")
    print(f"数据行 7~{max_row}   日列 AS..CE ({D1-D0+1} 列)  "
          f"{serial2date(hdr.get('AS'))} ~ {serial2date(hdr.get('CE'))}")
    print(f"锚点日(数据透视表!B11) = {anchor}")
    print(f"缓存值抽查(前60行日列): 公式+缓存值并存 {sample_ok} 格, 只有公式无缓存 {sample_f_only} 格")
    print("=" * 96)

    # 逐格判定
    hi, lo, bl, tred, tblue = [], [], [], [], []
    avg_num, avg_blank = 0, 0

    for r in range(7, max_row + 1):
        sku = get(r, C_SKU)
        if not sku:
            continue
        name = get(r, C_NAME)
        avg = get(r, C_AVG)
        if not isinstance(avg, (int, float)):
            avg_blank += 1
            avg = None
        else:
            avg_num += 1

        # T 列（可售库存）红/蓝
        t = get(r, C_T)
        if avg is not None and isinstance(t, (int, float)):
            if t > avg * 1.5:
                tred.append(dict(row=r, sku=sku, name=name, avg=avg, t=t, thr=avg * 1.5))
            elif t < avg / 2 and avg > 2:
                tblue.append(dict(row=r, sku=sku, name=name, avg=avg, t=t, thr=avg / 2))

        for n in range(D0, D1 + 1):
            col = num2col(n)
            v = get(r, col)
            if not isinstance(v, (int, float)):
                continue
            d = serial2date(hdr.get(col))
            base = dict(row=r, sku=sku, name=name, avg=avg, col=col, date=d, v=v)
            if avg is not None and avg > 1.5 and v > avg * 3:
                hi.append(dict(base, thr=avg * 3, mult=v / avg))
            if avg is not None and avg > 0 and avg <= 1.5 and v >= 2 and v > avg * 4:
                lo.append(dict(base, thr=avg * 4, mult=v / avg))
            if avg is not None and avg > 0 and v > 0 and v < avg / 3:
                bl.append(dict(base, thr=avg / 3, mult=v / avg))

    print(f"\n上月日均(AO): 数值 {avg_num} 行, 空/非数值 {avg_blank} 行")
    if avg_num:
        vals = sorted(x for x in (get(r, C_AVG) for r in range(7, max_row + 1))
                      if isinstance(x, (int, float)))
        print(f"  分布 min={vals[0]:.2f} p25={vals[len(vals)//4]:.2f} "
              f"中位={vals[len(vals)//2]:.2f} p75={vals[3*len(vals)//4]:.2f} max={vals[-1]:.2f}")

    print(f"\n命中统计：")
    print(f"  红·高销量  (日销>日均×3 且 日均>1.5)      : {len(hi):>5} 格")
    print(f"  红·低基数  (日均≤1.5 且 日销≥2 且 >日均×4): {len(lo):>5} 格")
    print(f"  蓝·低销量  (日销<日均÷3 且 日销>0)        : {len(bl):>5} 格")
    print(f"  可售列红   (可售>日均×1.5，无阈值)        : {len(tred):>5} 行")
    print(f"  可售列蓝   (可售<日均÷2 且 日均>2)        : {len(tblue):>5} 行")

    # 按日期分布
    from collections import Counter
    print(f"\n红色格按日期分布（red_hi）:")
    for d, n in sorted(Counter(x["date"] for x in hi).items()):
        print(f"  {d}  {n:>3} 格")

    # 导出 xlsx
    os.makedirs(OUTDIR, exist_ok=True)
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment
    wb = openpyxl.Workbook()

    RED = PatternFill("solid", fgColor="FFFF0000")
    WHITE_BOLD = Font(color="FFFFFFFF", bold=True)

    def style_header(ws, ncol):
        for c in range(1, ncol + 1):
            cell = ws.cell(1, c)
            cell.fill = RED
            cell.font = WHITE_BOLD
            cell.alignment = Alignment(horizontal="center")
        ws.freeze_panes = "A2"

    # Sheet1 口径说明
    ws0 = wb.active
    ws0.title = "计算口径"
    rows0 = [
        ("表", f"{BOOK} / {SHEET}"),
        ("快照锚点日", anchor or "?"),
        ("日销列范围", f"AS..CE（{D1-D0+1} 列）  {serial2date(hdr.get('AS'))} ~ {serial2date(hdr.get('CE'))}"),
        ("上月日均列", "AO（$AO，绝对列引用）"),
        ("可售库存列", "AN"),
        ("", ""),
        ("红色·高销量 red_hi", "AND(AS7 > $AO7*3, $AO7>1.5)"),
        ("  → 含义", "当天日销 > 上月日均×3，且上月日均本身>1.5（低日均不参与，防小基数噪声）"),
        ("  → 优先级", "prio=39"),
        ("", ""),
        ("红色·低基数补偿 red_lo", "AND($AO7>0, $AO7<=1.5, AS7>=2, AS7>$AO7*4)"),
        ("  → 含义", "上月日均≤1.5（小基数），当天日销≥2 且 >日均×4"),
        ("  → 优先级", "prio=3647"),
        ("", ""),
        ("蓝色·低销量 blue", "AND(AS7 < $AO7/3, AS7>0)"),
        ("  → 含义", "当天日销 < 上月日均÷3 且 >0"),
        ("  → 优先级", "prio=40"),
        ("", ""),
        ("可售列红（AN）", "$AN7 > $AO7*1.5"),
        ("  → 注意", "无日均阈值保护：日均=0 时任何正可售都标红"),
        ("可售列蓝（AN）", "AND($AN7 < $AO7/2, $AO7>2)"),
        ("", ""),
        ("注", "数值取自 sheet XML 缓存值（公式重算前的最后一次结果）"),
    ]
    ws0.append(["项目", "内容"])
    for a, b in rows0:
        ws0.append([a, b])
    style_header(ws0, 2)
    ws0.column_dimensions["A"].width = 26
    ws0.column_dimensions["B"].width = 80

    # Sheet2 红·高销量
    ws1 = wb.create_sheet("红-高销量")
    ws1.append(["行号", "SKU", "型号", "上月日均(AO)", "日列", "日期", "当日日销",
                "判定阈值(日均×3)", "实际倍数", "公式", "是否真红"])
    for x in sorted(hi, key=lambda y: (y["date"], -y["v"])):
        ws1.append([x["row"], x["sku"], x["name"], round(x["avg"], 2), x["col"], x["date"],
                    x["v"], round(x["thr"], 2), round(x["mult"], 2),
                    f"AND({x['col']}{x['row']}>$AO{x['row']}*3,$AO{x['row']}>1.5)", ""])
    style_header(ws1, 11)
    for col, w in zip("ABCDEFGHIJK", [7, 12, 30, 14, 7, 12, 10, 18, 10, 46, 10]):
        ws1.column_dimensions[col].width = w

    # Sheet3 红·低基数
    ws2 = wb.create_sheet("红-低基数补偿")
    ws2.append(["行号", "SKU", "型号", "上月日均(AO)", "日列", "日期", "当日日销",
                "判定阈值(日均×4)", "实际倍数", "公式", "是否真红"])
    for x in sorted(lo, key=lambda y: (y["date"], -y["v"])):
        ws2.append([x["row"], x["sku"], x["name"], round(x["avg"], 2), x["col"], x["date"],
                    x["v"], round(x["thr"], 2), round(x["mult"], 2),
                    f"AND($AO{x['row']}>0,$AO{x['row']}<=1.5,{x['col']}{x['row']}>=2,"
                    f"{x['col']}{x['row']}>$AO{x['row']}*4)", ""])
    style_header(ws2, 11)
    for col, w in zip("ABCDEFGHIJK", [7, 12, 30, 14, 7, 12, 10, 18, 10, 60, 10]):
        ws2.column_dimensions[col].width = w

    # Sheet4 可售列红
    ws3 = wb.create_sheet("可售列红")
    ws3.append(["行号", "SKU", "型号", "上月日均(AO)", "可售库存(AN)",
                "阈值(日均×1.5)", "实际倍数", "日均=0?", "公式", "是否真红"])
    for x in sorted(tred, key=lambda y: -y["t"]):
        ws3.append([x["row"], x["sku"], x["name"], round(x["avg"], 2), x["t"],
                    round(x["thr"], 2), round(x["t"] / x["avg"], 2) if x["avg"] else "∞",
                    "是" if x["avg"] == 0 else "",
                    f"$AN{x['row']}>$AO{x['row']}*1.5", ""])
    style_header(ws3, 10)
    for col, w in zip("ABCDEFGHIJ", [7, 12, 30, 14, 14, 16, 10, 9, 26, 10]):
        ws3.column_dimensions[col].width = w

    out = os.path.join(OUTDIR, "欧洲表红色核对_20261026.xlsx")
    wb.save(out)
    print(f"\n已导出: {out}")

    # 控制台预览 top 20
    print(f"\n【红·高销量 Top 20】")
    print(f"{'行':>4} {'SKU':<8} {'型号':<24} {'日均':>6} {'日期':<11} {'日销':>5} {'阈值':>6} {'倍数':>5}")
    for x in sorted(hi, key=lambda y: -y["v"])[:20]:
        print(f"{x['row']:>4} {str(x['sku']):<8} {(x['name'] or '')[:22]:<24} "
              f"{x['avg']:>6.2f} {x['date']:<11} {x['v']:>5.0f} {x['thr']:>6.1f} {x['mult']:>4.1f}x")

    print(f"\n【红·低基数补偿 全部 {len(lo)} 格】")
    print(f"{'行':>4} {'SKU':<8} {'型号':<24} {'日均':>6} {'日期':<11} {'日销':>5} {'阈值':>6} {'倍数':>5}")
    for x in sorted(lo, key=lambda y: (y["avg"], -y["v"]))[:40]:
        print(f"{x['row']:>4} {str(x['sku']):<8} {(x['name'] or '')[:22]:<24} "
              f"{x['avg']:>6.2f} {x['date']:<11} {x['v']:>5.0f} {x['thr']:>6.1f} {x['mult']:>4.1f}x")

    print(f"\n【可售列红 日均=0 的行（这些最可能被质疑）】")
    zero_avg = [x for x in tred if x["avg"] == 0]
    print(f"  共 {len(zero_avg)} 行")
    for x in zero_avg[:25]:
        print(f"  行{x['row']:>4} {str(x['sku']):<8} {(x['name'] or '')[:24]:<26} 日均=0 可售={x['t']:.0f}")


if __name__ == "__main__":
    main()
