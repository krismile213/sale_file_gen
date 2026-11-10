# -*- coding: utf-8 -*-
"""fix_aux_cols_20261107.py —— 把『每日销售数据』的辅助列写成静态值

背景（2026-11-07 实测）：
  openpyxl 每存一次盘就会清空全部公式缓存。三本簿（Interflowd 北美 / 日本 /
  Interflowd 澳洲）自被管道写过之后从未在 Excel 重算过，导致
    purchase date / 有效性 / 地区  三列缓存值恒为 None
  → 分析表日销公式 SUMIFS(... $AK:$AK, 日期, $AO:$AO, 地区, $AM:$AM, 有效性)
    永远匹配不到行 → AH(当月销量)=0 → AD(实时速度) 恒 0 → 断货判断全错。

修法（用户 2026-11-07 拍板）：辅助列不再依赖公式，直接按下列口径写静态值：
    purchase date = LEFT(purchase-date, 10)
    有效性        = (order-status != "Cancelled" 且 MID(sales-channel,1,7)="Amazon.")
                    北美簿额外放行 Non-Amazon / Non-Amazon CA（TEMU）
    地区          = 按 sales-channel 映射到站点名

用法：
    python pipeline/tools/fix_aux_cols_20261107.py --dry-run
    python pipeline/tools/fix_aux_cols_20261107.py --apply
"""
import os
import sys
import shutil
import argparse
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter, column_index_from_string as ci

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))
from pipeline.core import config as C

SHEET = "每日销售数据"

# sales-channel → 地区名
CH2SITE = {
    "Amazon.com": "美国",
    "Amazon.ca": "加拿大",
    "Amazon.com.mx": "墨西哥",
    "Amazon.com.au": "澳大利亚",
    "Amazon.co.jp": "日本",
    "Amazon.in": "印度",
    "Amazon.co.uk": "英国",
    "Amazon.de": "欧洲",
    "Amazon.fr": "欧洲",
    "Amazon.it": "欧洲",
    "Amazon.es": "欧洲",
    "Amazon.nl": "欧洲",
    "Amazon.se": "欧洲",
    "Amazon.pl": "欧洲",
    "Amazon.com.be": "欧洲",
    "Amazon.ae": "阿联酋",
    "Amazon.sa": "沙特阿拉伯",
}

AUX_NAMES = ("purchase date", "purchase-date", "有效性", "型号", "地区")


def find_raw_cols(ws, hdr_row, max_scan=60):
    """找原始报告列 purchase-date / order-status / sales-channel 的列号"""
    raw = {}
    for c in range(1, max_scan + 1):
        v = ws.cell(row=hdr_row, column=c).value
        if not isinstance(v, str):
            continue
        h = v.strip().lower()
        if h == "purchase-date":
            raw.setdefault("date", c)
        elif h == "order-status":
            raw.setdefault("status", c)
        elif h == "sales-channel":
            raw.setdefault("channel", c)
    return raw


def find_aux_cols(ws, max_scan=60):
    """在前 5 行里找**辅助列**（报告原生列之后额外加的）→ {'date'|'valid'|'site'|'model': 列号}

    ⚠️ 必须排除报告原生列：原始报告本身就有 purchase-date(C)、order-status(E)、
    sales-channel(G)。辅助列里的「purchase date」（中间是空格）才是要写的那一列。
    """
    out = {}
    for r in range(1, 6):
        for c in range(1, max_scan + 1):
            v = ws.cell(row=r, column=c).value
            if not isinstance(v, str):
                continue
            s = v.strip()
            # 注意：辅助列名是「purchase date」（空格），原生列是「purchase-date」（连字符）
            if s.lower() == "purchase date":
                out.setdefault("date", c)
            elif s == "有效性":
                out.setdefault("valid", c)
            elif s == "地区":
                out.setdefault("site", c)
            elif s == "型号":
                out.setdefault("model", c)
        if out:
            break
    return out


def find_data_start(ws):
    """首个数据行：A 列是 amazon-order-id（不是表头文字）"""
    for r in range(1, 8):
        v = ws.cell(row=r, column=1).value
        if isinstance(v, str) and v.strip() and v.strip().lower() != "amazon-order-id":
            return r
    return 3


def compute_row(ws, r, col_status, col_channel, col_pdate, site, allow_temu=False):
    st = ws.cell(row=r, column=col_status).value
    ch = ws.cell(row=r, column=col_channel).value
    pd = ws.cell(row=r, column=col_pdate).value
    st = ("" if st is None else str(st)).strip()
    ch = ("" if ch is None else str(ch)).strip()
    # purchase date
    if isinstance(pd, datetime):
        d = pd.strftime("%Y-%m-%d")
    elif pd is None:
        d = ""
    else:
        d = str(pd).strip()[:10]
    # 有效性
    ok = (st != "Cancelled") and ch[:7] == "Amazon."
    if not ok and allow_temu and ch.startswith("Non-Amazon"):
        ok = True
    return d, (1 if ok else 0)


def resolve_site(ws, r, col_channel):
    ch = ws.cell(row=r, column=col_channel).value
    ch = ("" if ch is None else str(ch)).strip()
    return CH2SITE.get(ch, ch or "")


def process_book(path, allow_temu, dry_run):
    fn = os.path.basename(path)
    wb = load_workbook(path)
    if SHEET not in wb.sheetnames:
        print(f"  {fn}: 无『{SHEET}』，跳过")
        wb.close()
        return None
    ws = wb[SHEET]
    aux = find_aux_cols(ws)
    if not aux:
        print(f"  {fn}: 找不到辅助列表头，跳过")
        wb.close()
        return None
    r0 = find_data_start(ws)
    hdr_row = r0 - 1
    raw = find_raw_cols(ws, hdr_row)
    if len(raw) < 3:
        print(f"  {fn}: 原始列不全 {raw}，跳过")
        wb.close()
        return None

    last = ws.max_row
    # 数据末行（A 列非空）
    while last >= r0 and ws.cell(row=last, column=1).value in (None, ""):
        last -= 1

    n_date = n_valid = n_site = 0
    samples = []
    for r in range(r0, last + 1):
        if ws.cell(row=r, column=1).value in (None, ""):
            continue
        d, v = compute_row(ws, r, raw["status"], raw["channel"], raw["date"], None, allow_temu)
        site = resolve_site(ws, r, raw["channel"])
        for kind, c in aux.items():
            if kind == "date":
                if ws.cell(row=r, column=c).value != d:
                    n_date += 1
                if not dry_run:
                    ws.cell(row=r, column=c).value = d
            elif kind == "valid":
                if ws.cell(row=r, column=c).value != v:
                    n_valid += 1
                if not dry_run:
                    ws.cell(row=r, column=c).value = v
            elif kind == "site":
                if ws.cell(row=r, column=c).value != site:
                    n_site += 1
                if not dry_run:
                    ws.cell(row=r, column=c).value = site
        if len(samples) < 3:
            samples.append((r, d, v, site))
    print(f"  {fn}: 数据行 r{r0}~r{last}，辅助列="
          + str({k: get_column_letter(c) for k, c in aux.items()}))
    print(f"      待写：date {n_date} / valid {n_valid} / site {n_site}   样例={samples}")
    if not dry_run:
        wb.save(path)
    wb.close()
    return n_date + n_valid + n_site


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--apply", action="store_true")
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--book", default=None, help="只处理指定簿（文件名）")
    a = ap.parse_args()
    dry = not a.apply

    base = C.WORK_DIR if hasattr(C, "WORK_DIR") else None
    if not base or not os.path.isdir(base):
        print("!! 找不到工作目录")
        return 1
    files = sorted(f for f in os.listdir(base)
                   if f.endswith(".xlsx") and not f.startswith("~$"))
    if a.book:
        files = [f for f in files if f == a.book]
    print(f"{'[DRY]' if dry else '[APPLY]'} 目标 {len(files)} 本簿，目录 {base}")
    total = 0
    for f in files:
        p = os.path.join(base, f)
        # TEMU 放行：北美系簿（美国/加拿大/墨西哥）
        allow_temu = any(k in f for k in ("北美", "美国", "加拿大"))
        if not dry:
            bak = p + ".bak_auxcols"
            if not os.path.exists(bak):
                shutil.copy2(p, bak)
        n = process_book(p, allow_temu, dry)
        total += (n or 0)
    print(f"合计变更 {total} 格")
    return 0


if __name__ == "__main__":
    sys.exit(main())
