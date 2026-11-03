# -*- coding: utf-8 -*-
"""
日销 vs 月销 对撞（只读）
========================
目的：验证「日销表（1.每日销售数据表）的 9 月销量」与「月销总表（trunk/result）9 月列」是否对得上。

口径（两侧同一把尺子，独立重算、不看公式缓存）：
  亚马逊：order-status != Cancelled
          AND sales-channel 在 channel_country_map 内（= 有效亚马逊站点渠道；排除 Non-Amazon）
          AND sku 不以 Amazon.Found./amzn.gr. 开头（月销侧过滤）
  归属日：两种都算，用来看差异到底出在哪 ——
          A) 原始日期 = purchase-date 前 10 位（日销表 LEFT(C,10) 口径）
          B) 本地时区日期 = purchase-date 转站点时区后取日期（月销管线 tz_convert 口径）
  值：quantity 求和

用法：python compare_daily_vs_monthly.py [YYYY-MM]
"""
# --- pipeline 路径引导（2026-10-14 目录收敛自动注入）---
import os as _os, sys as _sys
_bp = _os.path.dirname(_os.path.abspath(__file__))
for _i in range(6):
    if _os.path.isdir(_os.path.join(_bp, "1.每日销售数据表")):
        break
    _bp = _os.path.dirname(_bp)
DSR = _bp
_bcore = _os.path.join(DSR, "pipeline", "core")
if _bcore not in _sys.path:
    _sys.path.insert(0, _bcore)
# --- 引导结束 ---
import os, sys, re, datetime
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
# [normalize] DSR 由上方引导块赋值
ROOT = os.path.dirname(DSR)
ARCH = os.path.join(DSR, "1.每日销售数据表", "2026年每日数据", "2026年9月")
MONTHLY = os.path.join(ROOT, "trunk", "result", "数据归属周期", "202609", "月销售总表.xlsx")

YM = sys.argv[1] if len(sys.argv) > 1 else "2026-09"
D0, D1 = f"{YM}-01", f"{YM}-30"

CH2SITE = {
    "Amazon.co.uk": "英国", "Amazon.fr": "法国", "Amazon.de": "德国", "Amazon.it": "意大利",
    "Amazon.es": "西班牙", "Amazon.com.tr": "土耳其", "Amazon.com.be": "比利时", "Amazon.se": "瑞典",
    "Amazon.nl": "荷兰", "Amazon.pl": "波兰", "Amazon.ie": "爱尔兰", "Amazon.com.au": "澳大利亚",
    "Amazon.co.jp": "日本", "Amazon.com": "美国", "Amazon.com.mx": "墨西哥", "Amazon.ca": "加拿大",
    "Amazon.sa": "沙特阿拉伯", "Amazon.ae": "阿拉伯联合酋长国", "Amazon.in": "印度",
}
SITE2COL = {
    "英国": "英国", "法国": "欧洲", "德国": "欧洲", "意大利": "欧洲", "西班牙": "欧洲",
    "土耳其": "欧洲", "比利时": "欧洲", "瑞典": "欧洲", "荷兰": "欧洲", "波兰": "欧洲", "爱尔兰": "欧洲",
    "日本": "日本", "澳大利亚": "澳洲", "美国": "美国", "加拿大": "加拿大", "墨西哥": "墨西哥",
    "沙特阿拉伯": "沙特", "阿拉伯联合酋长国": "阿联酋", "印度": "印度",
}
TZ = {"英国": "Europe/London", "欧洲": "Europe/London", "日本": "Asia/Tokyo",
      "澳大利亚": "Australia/Sydney", "阿拉伯联合酋长国": "Asia/Dubai", "沙特阿拉伯": "Asia/Riyadh",
      "美国": "America/Los_Angeles", "加拿大": "America/Los_Angeles", "墨西哥": "America/Los_Angeles",
      "印度": "Asia/Kolkata"}

BOOKS = [
    ("欧盟英国库存销售表.xlsx", "每日销售数据", "amazon"),
    ("北美加拿大库存销售表.xlsx", "每日销售数据", "amazon"),
    ("Interflowd北美加拿大库存销售表.xlsx", "每日销售数据", "amazon"),
    ("日本库存销售表.xlsx", "每日销售数据", "amazon"),
    ("澳大利亚库存销售表.xlsx", "每日销售数据", "amazon"),
    ("Interflowd澳大利亚库存销售表.xlsx", "每日销售数据", "amazon"),
    ("印度库存销售表.xlsx", "每日销售数据", "amazon"),
    ("沙特库存销售表.xlsx", "每日销售数据沙特", "amazon"),
    ("阿联酋库存销售表.xlsx", "每日销售数据阿联酋", "amazon"),
    ("Interflowd阿联酋库存销售表.xlsx", "每日销售数据阿联酋", "amazon"),
    ("罗马尼亚库存销售表.xlsx", "订单数据", "emag"),
    ("美客多库存销售表.xlsx", "订单数据", "ml"),
]
MESES = {"enero": 1, "febrero": 2, "marzo": 3, "abril": 4, "mayo": 5, "junio": 6, "julio": 7,
         "agosto": 8, "septiembre": 9, "setiembre": 9, "octubre": 10, "noviembre": 11, "diciembre": 12}
ML_BAD = ("退款", "退货", "取消", "reembols", "devoluci", "cancelad", "anulaci")


def qint(v):
    try:
        return int(float(v))
    except (TypeError, ValueError):
        return 0


def date_of(v):
    if isinstance(v, datetime.datetime):
        return v.strftime("%Y-%m-%d")
    s = str(v).strip()
    return s[:10] if len(s) >= 10 else ""


def tz_date_of(v, tzname):
    if tzname is None:
        return date_of(v)
    import pandas as pd
    try:
        ts = pd.Timestamp(v)
        if ts.tzinfo is None:
            return ts.strftime("%Y-%m-%d")
        return ts.tz_convert(tzname).strftime("%Y-%m-%d")
    except Exception:
        return date_of(v)


def ml_date(s):
    s = str(s).strip()
    m = re.match(r"(\d{1,2})\s+de\s+([a-záéíóú]+)\s+de\s+(\d{4})", s, re.I)
    if m:
        day, mon, yr = int(m.group(1)), m.group(2).lower(), int(m.group(3))
        if mon in MESES:
            return f"{yr:04d}-{MESES[mon]:02d}-{day:02d}"
    return s[:10]


def read_amazon_book(fp, sheet_name):
    wb = openpyxl.load_workbook(fp, read_only=True, data_only=True)
    ws = wb[sheet_name]
    hdr, idx = None, {}
    for row in ws.iter_rows(min_row=1, max_row=8, max_col=60):
        vals = {str(c.value).strip(): c.column for c in row if c.value is not None}
        if any(k.startswith("order-item-id") for k in vals):
            hdr, idx = row[0].row, vals
            break
    if hdr is None:
        wb.close()
        return None
    def col(k):
        if k in idx:
            return idx[k]
        for kk, cc in idx.items():
            if kk.startswith(k):
                return cc
        return None
    c_pd, c_st, c_ch, c_sk, c_qt = (col("purchase-date"), col("order-status"),
                                    col("sales-channel"), col("sku"), col("quantity"))
    out = []
    for row in ws.iter_rows(min_row=hdr + 1, max_col=60):
        if c_pd is None:
            break
        pv = row[c_pd - 1].value
        if pv is None:
            continue
        raw = date_of(pv)
        if not (f"{YM[:4]}-08-20" <= raw <= f"{YM[:4]}-10-05"):
            continue  # 只看窗口附近，加速
        st = str(row[c_st - 1].value or "") if c_st else ""
        ch = str(row[c_ch - 1].value or "") if c_ch else ""
        sk = str(row[c_sk - 1].value or "") if c_sk else ""
        qt = qint(row[c_qt - 1].value) if c_qt else 0
        out.append((raw, st.strip(), ch.strip(), sk.strip(), qt, pv))
    wb.close()
    return out


def read_emag_book(fp, sheet_name):
    wb = openpyxl.load_workbook(fp, read_only=True, data_only=True)
    ws = wb[sheet_name]
    rows = [(str(r[1] or "").strip()[:10], str(r[12] or "").strip(), qint(r[7]), str(r[4] or "").strip())
            for r in ws.iter_rows(min_row=2, values_only=True)]
    wb.close()
    return rows


def read_ml_book(fp, sheet_name):
    wb = openpyxl.load_workbook(fp, read_only=True, data_only=True)
    ws = wb[sheet_name]
    rows = [(ml_date(r[1]), str(r[2] or "").strip(), qint(r[6]), str(r[18] or "").strip())
            for r in ws.iter_rows(min_row=2, values_only=True)]
    wb.close()
    return rows


def read_monthly():
    wb = openpyxl.load_workbook(MONTHLY, read_only=True, data_only=True)
    ws = wb["2026年"]
    r1 = next(ws.iter_rows(min_row=1, max_row=1, values_only=True))
    r2 = next(ws.iter_rows(min_row=2, max_row=2, values_only=True))
    groups, cur = {}, None
    for j, v in enumerate(r1, 1):
        if v:
            cur = str(v).strip()
            groups[cur] = j
    want = {g: c0 + 8 for g, c0 in groups.items()
            if c0 + 8 <= len(r2) and str(r2[c0 + 8 - 1] or "").strip() == "9月"}
    sums = {g: 0 for g in want}
    for row in ws.iter_rows(min_row=4, max_row=ws.max_row, values_only=True):
        for g, c in want.items():
            if c - 1 < len(row):
                sums[g] += qint(row[c - 1])
    wb.close()
    return sums


def main():
    monthly = read_monthly()
    print("=" * 104)
    print(f"日销 vs 月销 对撞 | 期间 {YM}（月销表 9 月列，数据行求和）")
    print("=" * 104)
    for g in ["欧洲", "英国", "美国", "加拿大", "日本", "澳洲", "阿联酋", "沙特", "墨西哥",
              "印度", "亚马逊", "eMAG", "TEMU平台", "美客多", "总计"]:
        if g in monthly:
            print(f"    {g:<8} = {monthly[g]:>10,}")
    print()

    per_book = []
    for fn, sn, kind in BOOKS:
        fp = os.path.join(ARCH, f"2026年9月-{fn}")
        name = fn.replace("库存销售表.xlsx", "")
        if not os.path.exists(fp):
            print(f"[缺文件] {name}")
            continue
        if kind == "amazon":
            rows = read_amazon_book(fp, sn)
            if rows is None:
                print(f"[跳过] {name} 找不到表头")
                continue
            agg = {}
            for raw, st, ch, sk, qt, pv in rows:
                if st == "Cancelled" or ch not in CH2SITE:
                    continue
                if sk.startswith("Amazon.Found.") or sk.startswith("amzn.gr."):
                    continue
                col = SITE2COL[CH2SITE[ch]]
                in_raw = D0 <= raw <= D1
                tzd = tz_date_of(pv, TZ.get(col))
                in_tz = D0 <= tzd <= D1
                if not (in_raw or in_tz):
                    continue
                a = agg.setdefault(col, [0, 0, 0])
                a[0] += qt if in_raw else 0
                a[1] += qt if in_tz else 0
                a[2] += 1
            per_book.append((name, agg))
        elif kind == "emag":
            rows = read_emag_book(fp, sn)
            tot = sum(qt for d, st, qt, sk in rows
                      if D0 <= d <= D1 and st not in ("Storno", "Canceled by customer"))
            n = sum(1 for d, st, qt, sk in rows if D0 <= d <= D1)
            per_book.append((name, {"eMAG": [tot, tot, n]}))
        elif kind == "ml":
            rows = read_ml_book(fp, sn)
            tot = n = 0
            bad = 0
            for d, est, uni, sk in rows:
                if not (D0 <= d <= D1):
                    continue
                if any(k in est.lower() for k in ML_BAD):
                    bad += 1
                    continue
                tot += uni
                n += 1
            per_book.append((name, {"美客多": [tot, tot, n]}))
            print(f"  [ML] 9月窗口内 有效行 {n}｜剔除(取消/退货/退款) {bad}｜件 {tot}")

    print("-" * 104)
    print("逐簿 9 月销量（非Cancelled + 有效亚马逊渠道 + 剔 Amazon.Found/amzn.gr）")
    print("-" * 104)
    for name, agg in per_book:
        for col, (a, b, n) in sorted(agg.items()):
            print(f"  {name:<20} → {col:<8} 原始日期 {a:>8,} 件 | 时区日期 {b:>8,} 件 | 命中行 {n}")

    side = {}
    for name, agg in per_book:
        for col, (a, b, n) in agg.items():
            s = side.setdefault(col, [0, 0])
            s[0] += a
            s[1] += b

    print()
    print("=" * 104)
    print("对撞表（单位：件）")
    print("=" * 104)
    print(f"{'月销列':<10}{'月销表':>10}{'日销[原始]':>12}{'差':>8}{'日销[时区]':>12}{'差':>8}")
    tm = t1 = t2 = 0
    for col in ["欧洲", "英国", "美国", "加拿大", "日本", "澳洲", "阿联酋", "沙特", "墨西哥",
                "印度", "eMAG", "美客多"]:
        m = monthly.get(col, 0)
        d1, d2 = side.get(col, [0, 0])
        tm += m; t1 += d1; t2 += d2
        print(f"{col:<10}{m:>10,}{d1:>12,}{d1 - m:>+8,}{d2:>12,}{d2 - m:>+8,}")
    print("-" * 104)
    print(f"{'合计':<10}{tm:>10,}{t1:>12,}{t1 - tm:>+8,}{t2:>12,}{t2 - tm:>+8,}")
    print()
    print("参考：月销表『亚马逊』列 =", f"{monthly.get('亚马逊', 0):,}",
          "|『总计』列 =", f"{monthly.get('总计', 0):,}")
    # 亚马逊列 vs 十日亚马逊列之和
    amz_cols = ["欧洲", "英国", "美国", "加拿大", "日本", "澳洲", "阿联酋", "沙特", "墨西哥", "印度"]
    print("      十国亚马逊列之和 =", f"{sum(monthly.get(c, 0) for c in amz_cols):,}",
          "（应等于『亚马逊』列）")


if __name__ == "__main__":
    main()
