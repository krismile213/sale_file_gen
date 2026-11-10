# -*- coding: utf-8 -*-
"""refresh_product_master.py —— 按规则自动重算商品资料的评级/类型/ASIN

数据源（全部来自已有报表，无需人工）：
    月销售总表  trunk/result/数据归属周期/{期}/月销售总表.xlsx  → 评级
    库存报告    trunk/data/数据归属周期/{期}/快照日期/{最新}/*管理库存报告*.csv → ASIN
    中文品名    资料自身 → 类型

用法：
    python pipeline/tools/refresh_product_master.py                    # 只读演练
    python pipeline/tools/refresh_product_master.py --period 202610    # 指定基准月
    python pipeline/tools/refresh_product_master.py --apply            # 真正写回（自动备份）
    python pipeline/tools/refresh_product_master.py --out r.txt --xlsx r.xlsx

⚠️ 设计边界（踩过的坑，别改）：
    1. 评级体系不自动切换：现值 A~E 走常规规则，现值 X 系列走 X 规则，空值（已停售）不动。
    2. 状态列永不自动改 —— 停售是人工决策（实测 204 个停售里 32 个当月还在卖）。
    3. 型号/英文品名只在为空时补，不回刷老品（老品命中率只有 35~50%）。
    4. 评级必须用「上一个完整月」的月销表，月中跑当月表会偏低。
"""
import argparse
import calendar
import csv
import glob
import io
import os
import re
import shutil
import sys
from collections import Counter, defaultdict
from datetime import datetime

_HERE = os.path.dirname(os.path.abspath(__file__))              # pipeline/tools
DSR = os.path.dirname(os.path.dirname(_HERE))                   # Daily_Sales_Report
SALE = os.path.dirname(DSR)                                     # sale_file_gen
sys.path.insert(0, os.path.join(DSR, "pipeline", "core"))
import product_master as PM                                     # noqa: E402

# ----------------------------------------------------------------- 评级规则

# 常规档阈值（用户口径，已用 367/367 实测验证）
T_A_M1, T_B_M3, T_C_M6, T_D_M6 = 500, 500, 250, 200

# 新品档日均阈值
X_STEPS = ((16.7, "XA"), (5.5, "XB"), (1.38, "XC"), (1.11, "XD"))

REGULAR = ("A", "B", "C", "D", "E")


def calc_regular(m1, m3, m6):
    """1 个月 / 3 个月 / 6 个月销量 → A~E"""
    if m1 >= T_A_M1:
        return "A"
    if m3 >= T_B_M3:
        return "B"
    if m6 >= T_C_M6:
        return "C"
    if m6 >= T_D_M6:
        return "D"
    return "E"


def calc_x(month_qty, days):
    """当月销量 / 当月天数 → XA~XE；0 销量 → X"""
    if month_qty <= 0:
        return "X"
    d = month_qty / float(days)
    for thr, code in X_STEPS:
        if d >= thr:
            return code
    return "XE"


# ----------------------------------------------------------------- 类型解析

_PAD_RE = re.compile(r"(Pad|Tab\b|Tablet|Inches)", re.I)


def ptype(cn):
    """中文品名 → 类型（实测 652/652 复现）"""
    s = cn or ""
    if "手机壳" in s:
        return "手机壳"
    if "水凝膜" in s or "软膜" in s:
        return "水凝膜"
    if "3D热弯" in s:
        return "3D热弯"
    if "平板" in s or (_PAD_RE.search(s) and "钢化膜" in s):
        return "平板钢化膜"
    if "曲面" in s and "钢化膜" in s:
        return "3D曲面"
    if "镜头膜" in s and "钢化膜" not in s:
        return "镜头膜"
    if "钢化膜" in s:
        return "钢化膜"
    return None


_SPEC_RE = re.compile(
    r"(高清|钢化膜|水凝膜|手机壳|软膜|镜头膜|3D|全屏|背膜"
    r"|标准版|透明|丝印|2\.5D|黑色|白色|蓝色|红色|粉色|紫色|绿色|灰色"
    r"|陶拉思|战斗骑士|幻影|磨砂|液态硅胶)")

_BRAND_ALIAS = {
    "三星": "SAMSUNG", "samsung": "SAMSUNG", "小米": "XIAOMI", "xiaomi": "XIAOMI",
    "红米": "REDMI", "redmi": "REDMI", "一加": "ONEPLUS", "oneplus": "ONEPLUS",
    "华为": "HUAWEI", "huawei": "HUAWEI", "荣耀": "HONOR", "honor": "HONOR",
    "苹果": "APPLE", "apple": "APPLE", "iphone": "APPLE", "oppo": "OPPO",
    "vivo": "VIVO", "realme": "REALME", "google": "GOOGLE", "pixel": "GOOGLE",
    "sony": "SONY", "索尼": "SONY", "motorola": "MOTOROLA", "nokia": "NOKIA",
    "infinix": "INFINIX", "tecno": "TECNO", "nothing": "NOTHING", "poco": "POCO",
    "zte": "ZTE",
}


def parse_model(cn):
    """中文品名 → 型号（截到第一个规格词）。老品命中率低，只给新品用。"""
    m = _SPEC_RE.search(cn or "")
    if not m:
        return None
    head = (cn or "")[:m.start()].strip()
    return head or None


def brand_of_model(model):
    """型号 → 品牌代码"""
    s = re.sub(r"\s+", "", model or "").lower()
    for k in sorted(_BRAND_ALIAS, key=len, reverse=True):
        if s.startswith(k):
            return _BRAND_ALIAS[k]
    return None


# ----------------------------------------------------------------- 数据源

def find_msx(period):
    p = os.path.join(SALE, "trunk", "result", "数据归属周期", period, "月销售总表.xlsx")
    return p if os.path.isfile(p) else None


def latest_period():
    root = os.path.join(SALE, "trunk", "result", "数据归属周期")
    ds = sorted(d for d in os.listdir(root)
                if re.fullmatch(r"\d{6}", d) and os.path.isdir(os.path.join(root, d)))
    return ds[-1] if ds else None


def _cellnum(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return 0.0


def _read_year_sheet(ws):
    """读月销表某年 sheet → {sku: {(年,月): 销量}}，总计区块动态定位"""
    out = defaultdict(dict)
    it = ws.iter_rows(values_only=True)
    try:
        h1 = next(it)
    except StopIteration:
        return out
    tcol = None
    for j, v in enumerate(h1):
        if isinstance(v, str) and v.strip() == "总计":
            tcol = j
            break
    if tcol is None:
        return out
    # 月份在 tcol..tcol+11
    for row in it:
        if not row or tcol >= len(row):
            continue
        sku = row[0]
        if not isinstance(sku, str) or not sku.strip():
            continue
        if sku.strip().upper() in ("SKU", "产品型号"):
            continue        # 表头行（r3），不是数据
        vals = [_cellnum(row[tcol + k]) if tcol + k < len(row) else 0.0 for k in range(12)]
        out[sku.strip()] = vals
    return out


def load_sales(period):
    """→ ({sku: [(y,m,val)...]}, 年份)，含跨年补齐"""
    p = find_msx(period)
    if not p:
        return {}, None
    y = int(period[:4])
    wb = load_workbook_ro(p)
    data = {}
    try:
        for sheet, yy in ((f"{y}年", y), (f"{y - 1}年", y - 1)):
            if sheet not in wb.sheetnames:
                continue
            raw = _read_year_sheet(wb[sheet])
            for sku, vals in raw.items():
                slot = data.setdefault(sku, {})
                for i, v in enumerate(vals):
                    slot[(yy, i + 1)] = v
    finally:
        wb.close()
    return data, y


def load_workbook_ro(path):
    from openpyxl import load_workbook
    return load_workbook(path, read_only=True, data_only=True)


def latest_snapshot(period):
    root = os.path.join(SALE, "trunk", "data", "数据归属周期", period, "快照日期")
    if not os.path.isdir(root):
        return None
    ds = sorted(d for d in os.listdir(root) if re.fullmatch(r"\d{8}", d))
    return os.path.join(root, ds[-1]) if ds else None


_ENC = ("utf-8-sig", "utf-8", "shift_jis", "cp932", "gbk", "utf-16")


def _pick_col(hdr, want):
    """在表头里挑列索引：want='sku' / 'asin'

    ⚠️ 日本站报告的列名是日文（「出品者 SKU」「商品 SKU」「ASIN」「商品名」），
    不能硬找 'sku'。SKU 列优先取「出品者 SKU」（= 卖家自己的 G 编号），
    排除「商品 SKU」（那是亚马逊内部编号）。
    """
    low = [(i, (h or "").strip()) for i, h in enumerate(hdr)]
    if want == "asin":
        for i, h in low:
            if h.lower() == "asin":
                return i
        return None
    cand = [(i, h) for i, h in low if "sku" in h.lower()]
    if not cand:
        return None
    for i, h in cand:
        if "商品" not in h and "product" not in h.lower():
            return i
    return cand[0][0]


def load_asins(snap):
    """从库存报告取 {sku: asin}（多编码兼容，日/澳是 shift_jis）

    ⚠️ shift_jis 文件里常有个别非法字节，严格解码会整份失败 → 先严格试，全失败再用
    errors='replace' 兜底重试。宁可丢几行文本，也不要整站 ASIN 缺失。
    """
    res, bad = {}, []
    if not snap:
        return res, bad
    for f in glob.glob(os.path.join(snap, "*管理库存报告*.csv")):
        raw = open(f, "rb").read()
        txt = None
        for enc in _ENC:
            try:
                txt = raw.decode(enc)
                break
            except UnicodeDecodeError:
                continue
        if txt is None:
            for enc in _ENC:
                try:
                    txt = raw.decode(enc, errors="replace")
                    break
                except (UnicodeDecodeError, LookupError):
                    continue
        if txt is None:
            bad.append(os.path.basename(f))
            continue
        rows = list(csv.reader(io.StringIO(txt)))
        if not rows:
            bad.append(os.path.basename(f))
            continue
        ki, ai = _pick_col(rows[0], "sku"), _pick_col(rows[0], "asin")
        if ki is None or ai is None:
            bad.append(os.path.basename(f))
            continue
        for r in rows[1:]:
            if len(r) <= max(ki, ai):
                continue
            sku, a = r[ki].strip(), r[ai].strip()
            if not sku or not a or "N/A" in a.upper() or sku.upper() == "SKU":
                continue
            res.setdefault(sku, a)
    return res, bad


# ----------------------------------------------------------------- 主逻辑

def shift_ym(y, m, back):
    """从 (y,m) 往前推 back 个月"""
    t = y * 12 + (m - 1) - back
    return t // 12, t % 12 + 1


def window(slot, y, m, n):
    tot = 0.0
    for i in range(n):
        tot += slot.get(shift_ym(y, m, i), 0.0)
    return tot


def analyse(period, master, sales, y, asins):
    mo = int(period[4:6])
    days = calendar.monthrange(y, mo)[1]
    rows = []
    stat = Counter()
    for sku, a in master.items():
        real = str(a.get("score") or "").strip().upper()
        slot = sales.get(sku)
        cn = str(a.get("name_cn") or "").strip()
        # --- 评级
        if real in REGULAR:
            if slot:
                calc = calc_regular(window(slot, y, mo, 1),
                                    window(slot, y, mo, 3),
                                    window(slot, y, mo, 6))
                stat["reg_total"] += 1
                if calc == real:
                    stat["reg_same"] += 1
                else:
                    stat["reg_diff"] += 1
                    rows.append(dict(kind="评档", sku=sku, col="score", old=real,
                                     new=calc, note=f"{mo}月{window(slot,y,mo,1):.0f} "
                                                    f"/3月{window(slot,y,mo,3):.0f} "
                                                    f"/6月{window(slot,y,mo,6):.0f}"))
            else:
                stat["reg_nosale"] += 1
                rows.append(dict(kind="查无销量", sku=sku, col=None, old=real, new="",
                                 note="月销表无此 SKU，可能是更早停了"))
        elif real.startswith("X"):
            if slot:
                calc = calc_x(window(slot, y, mo, 1), days)
                stat["x_total"] += 1
                if calc == real:
                    stat["x_same"] += 1
                else:
                    stat["x_diff"] += 1
                    rows.append(dict(kind="评档", sku=sku, col="score", old=real,
                                     new=calc, note=f"{mo}月{window(slot,y,mo,1):.0f} "
                                                    f"日均{window(slot,y,mo,1)/days:.2f}"))
            else:
                stat["x_nosale"] += 1
        else:
            stat["blank"] += 1          # 空评级 = 已停售，不动

        # --- 类型
        t_real = str(a.get("type") or "").strip()
        t_calc = ptype(cn)
        if not t_real and t_calc:
            rows.append(dict(kind="补类型", sku=sku, col="type", old="", new=t_calc, note=cn[:40]))
            stat["type_fill"] += 1
        elif t_real and t_calc and t_real != t_calc:
            rows.append(dict(kind="类型异", sku=sku, col=None, old=t_real, new=t_calc,
                             note="人工核对品名"))
            stat["type_diff"] += 1
        elif t_real == t_calc:
            stat["type_same"] += 1

        # --- ASIN
        a_real = str(a.get("asin") or "").strip()
        if not a_real or "N/A" in a_real.upper():
            new = asins.get(sku)
            if new:
                rows.append(dict(kind="补ASIN", sku=sku, col="asin", old=a_real, new=new,
                                 note="来自当月库存报告"))
                stat["asin_fill"] += 1
            else:
                stat["asin_miss"] += 1

        # --- 型号（只补空）
        if not str(a.get("model") or "").strip():
            new = parse_model(cn)
            if new:
                rows.append(dict(kind="补型号", sku=sku, col="model", old="", new=new, note=cn[:40]))
                stat["model_fill"] += 1

    # 月销表有、资料没有 = 新品待建档
    for sku in sales:
        if sku not in master:
            stat["new_sku"] += 1
            rows.append(dict(kind="待建档", sku=sku, col=None, old="", new="",
                             note="月销表有、资料无 → 新品需手工建档"))
    return rows, stat


# ----------------------------------------------------------------- 写入

def apply_changes(rows, period, rename=True):
    p = PM.find_path()
    bak = p + f".bak_rating_{datetime.now():%m%d_%H%M}"
    shutil.copy2(p, bak)
    from openpyxl import load_workbook
    wb = load_workbook(p)          # 不用 read_only，要写
    ws = wb["商品资料"]
    # ⚠️ 表头有两个「SKU」：A 列是主键，O 列是公式列（K009ZH 这种带语言后缀的）。
    # 用 dict 构建会被后者覆盖 → 主键列必须写死 A 列，并断言表头真的是 SKU。
    if str(ws.cell(2, 1).value or "").strip() != "SKU":
        wb.close()
        raise RuntimeError(f"表头异常：A2 = {ws.cell(2, 1).value!r}，应为 'SKU'")
    hdr = {}
    for c in ws[2]:
        if c.value:
            k = str(c.value).strip()
            if k not in hdr:            # first-wins，保留左侧主键列
                hdr[k] = c.column
    colmap = {"score": None, "type": None, "asin": None, "model": None, "sku": 1}
    for k, c in hdr.items():
        if "销量得分" in k:
            colmap["score"] = c
        elif k == "类型":
            colmap["type"] = c
        elif k == "ASIN":
            colmap["asin"] = c
        elif k == "型号":
            colmap["model"] = c
    if not colmap["score"]:
        wb.close()
        raise RuntimeError("表头定位失败：找不到「N月销量得分」列")

    row_of = {}
    for r in range(3, ws.max_row + 1):
        v = ws.cell(r, colmap["sku"]).value
        if isinstance(v, str) and v.strip():
            row_of[v.strip()] = r

    n = 0
    per = Counter()
    for d in rows:
        if not d.get("col"):
            continue
        r = row_of.get(d["sku"])
        c = colmap.get(d["col"])
        if not r or not c:
            continue
        ws.cell(r, c).value = d["new"]
        n += 1
        per[d["kind"]] += 1

    # 同步列名月份（幂等：无论原来是几月，都改成基准月）
    if rename:
        cell = ws.cell(2, colmap["score"])
        old = str(cell.value or "")
        new = re.sub(r"\d+月销量得分", f"{int(period[4:6])}月销量得分", old)
        if new != old:
            cell.value = new
            per["列名"] += 1

    wb.save(p)
    wb.close()
    return n, dict(per), bak


# ----------------------------------------------------------------- 输出

def dump(rows, stat, period, path=None):
    out = []
    out.append(f"商品资料自动重算  period={period}  {datetime.now():%Y-%m-%d %H:%M}")
    out.append("")
    out.append("== 统计 ==")
    labels = [
        ("reg_same", "常规档一致"), ("reg_diff", "常规档需改档"), ("reg_nosale", "常规档但月销表无记录"),
        ("x_same", "X 类一致"), ("x_diff", "X 类需改档"), ("x_nosale", "X 类但月销表无记录"),
        ("blank", "空评级(已停售，不动)"),
        ("type_same", "类型一致"), ("type_fill", "类型可补"), ("type_diff", "类型异常需核"),
        ("asin_fill", "ASIN 可补"), ("asin_miss", "ASIN 仍缺"), ("model_fill", "型号可补"),
        ("new_sku", "新品待建档"),
    ]
    for k, lab in labels:
        out.append(f"  {lab:<24s}{stat.get(k, 0):6d}")
    g = lambda k: stat.get(k, 0)
    tot = g("reg_total") + g("x_total")
    same = g("reg_same") + g("x_same")
    if tot:
        out.append(f"  {'评级总体一致率':<24s}{same / tot * 100:5.1f}%  ({same}/{tot})")
    out.append("")
    order = ["评档", "补ASIN", "补类型", "类型异", "补型号", "查无销量", "待建档"]
    by = defaultdict(list)
    for d in rows:
        by[d["kind"]].append(d)
    for k in order:
        if k not in by:
            continue
        out.append(f"-- {k}（{len(by[k])}）")
        for d in by[k][:400]:
            tag = "*" if d.get("col") else " "
            out.append(f" {tag} {d['sku']:<12s} {str(d['old']):>8s} -> {str(d['new']):<8s} {d['note']}")
        if len(by[k]) > 400:
            out.append(f"   …… 另有 {len(by[k]) - 400} 条")
        out.append("")
    txt = "\n".join(out)
    if path:
        with open(path, "w", encoding="utf-8") as f:
            f.write(txt)
    return txt


def dump_xlsx(rows, path):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    wb = Workbook()
    ws = wb.active
    ws.title = "重算差异"
    hdr = ["簿", "类型", "SKU", "现资料值", "按规则算出", "备注", "能否自动写"]
    ws.append(hdr)
    for c in ws[1]:
        c.font = Font(bold=True)
        c.fill = PatternFill("solid", fgColor="DDEBF7")
    COLOR = {"评档": "FCE4E4", "补ASIN": "D9F2D9", "补类型": "D9F2D9",
             "补型号": "D9F2D9", "类型异": "FFF2CC", "待建档": "F8CBAD", "查无销量": "EDEDED"}
    for d in rows:
        ws.append(["商品资料", d["kind"], d["sku"], d["old"], d["new"], d["note"],
                   "可" if d.get("col") else "否"])
        f = COLOR.get(d["kind"])
        if f:
            for c in ws[ws.max_row]:
                c.fill = PatternFill("solid", fgColor=f)
    for col, w in zip("ABCDEFG", (12, 12, 14, 14, 16, 40, 12)):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    wb.save(path)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--period", default=None, help="基准月，如 202610（默认取最新周期）")
    ap.add_argument("--apply", action="store_true", help="真正写回商品资料.xlsx")
    ap.add_argument("--no-rename", action="store_true", help="不改销量得分列名的月份前缀")
    ap.add_argument("--out", default=None)
    ap.add_argument("--xlsx", default=None)
    args = ap.parse_args()

    period = args.period or latest_period()
    if not period:
        print("[!] 找不到任何月销表")
        return 2
    master = PM.load(force=True)
    if not master:
        print("[!] 读不到商品资料")
        return 2
    sales, y = load_sales(period)
    if not sales:
        print(f"[!] 月销表为空或找不到：{period}")
        return 2
    snap = latest_snapshot(period)
    asins, bad = load_asins(snap)

    print(f"period = {period}   基准月 = {int(period[4:6])} 月")
    print(f"商品资料 = {PM.find_path()}  ({len(master)} SKU)")
    print(f"月销表   = {find_msx(period)}  ({len(sales)} SKU)")
    print(f"快照     = {snap}  ({len(asins)} 个 ASIN)")
    if bad:
        print(f"[!] {len(bad)} 份库存报告读不开：{bad[:3]}")
    print()

    rows, stat = analyse(period, master, sales, y, asins)
    txt = dump(rows, stat, period, args.out)
    print(txt)
    if args.xlsx:
        dump_xlsx(rows, args.xlsx)
        print(f"已导出 {args.xlsx}")

    if args.apply:
        n, per, bak = apply_changes(rows, period, rename=not args.no_rename)
        print(f"\n[写入] {n} 格  {per}")
        print(f"       备份 {os.path.basename(bak)}")
    else:
        print("\n[演练] 未写入。加 --apply 执行。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
