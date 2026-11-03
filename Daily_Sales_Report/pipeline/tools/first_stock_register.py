# -*- coding: utf-8 -*-
"""型号首库存登记：扫全部快照日期的库存报告，登记每款型号「真正有库存」的日期。

定义（与用户口径对齐）：
- 首次出现   = 第一次在任何站点的库存报告里看到该 SKU（= listing 建好）
- 首可售日   = 第一次 afn-fulfillable-quantity > 0（真正能卖）
- 首在途日   = 第一次 在途(working+shipped+receiving) > 0（开始发货）

用法：
    python tools/first_stock_register.py                # 干跑，只打印统计
    python tools/first_stock_register.py --apply        # 生成/覆盖 xlsx
    python tools/first_stock_register.py --apply --sku G598,G599   # 只看指定 SKU

输出：Daily_Sales_Report/断货分析/型号首库存登记.xlsx
      sheet「首库存登记」：一行一个 SKU，跨站点汇总
      sheet「站点明细」  ：一行一个 (SKU, 站点)，看具体哪天在哪个站点有货
"""
import os, sys, csv, glob, re, argparse
from collections import defaultdict

DSR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
PERIODS = ["202608", "202609", "202610"]
OUT_XLSX = os.path.join(DSR, "断货分析", "型号首库存登记.xlsx")

# 与 core/analysis._SKU_RE 保持一致（2026-10-24 放宽字母后缀 G136R/G590J）
SKU_RE = re.compile(r"^[A-Z]{1,3}\d{3,4}[A-Z]{0,3}([-_ ].*)?$")

# 库存报告文件名 → 站点名（顺序敏感，先匹配长的）
SITE_RULES = [
    ("亚马逊阿拉伯联合酋长国", "阿联酋-亚马逊"),
    ("亚马逊沙特阿拉伯", "沙特-亚马逊"),
    ("亚马逊澳大利亚", "澳大利亚-亚马逊"),
    ("亚马逊加拿大", "加拿大-亚马逊"),
    ("亚马逊墨西哥", "墨西哥-亚马逊"),
    ("亚马逊印度", "印度-亚马逊"),
    ("亚马逊日本", "日本-亚马逊"),
    ("亚马逊欧盟", "欧洲-亚马逊"),
    ("亚马逊英国", "英国-亚马逊"),
    ("亚马逊美国", "美国-亚马逊"),
    ("亚马逊北美", "北美-亚马逊"),
    ("interflow阿拉伯联合酋长国", "阿联酋-interflow"),
    ("interflow沙特阿拉伯", "沙特-interflow"),
    ("interflow澳大利亚", "澳大利亚-interflow"),
    ("interflow加拿大", "加拿大-interflow"),
    ("interflow美国", "美国-interflow"),
    ("interflow北美", "北美-interflow"),
    ("美客多墨西哥", "墨西哥-美客多"),
    ("EMAG罗马尼亚", "罗马尼亚-EMAG"),
    ("多国库存", "多国库存"),
]


def site_of(fname):
    for key, name in SITE_RULES:
        if key in fname:
            return name
    return None


def read_table(path):
    """返回 (header, rows)。

    库存报告有三种形态，统一在这里吃掉：
    · 亚马逊英文 CSV   —— 表头第 1 行，标准列名
    · 亚马逊中文 CSV   —— **GBK 且表头已含不可逆 '?' 替换**，列名认不出，只能靠 "SKU" 字样定位列 0
    · EMAG / 美客多 xlsx —— 前几行是说明页，真实表头在第 N 行、甚至不在第一个 sheet
    """
    if path.lower().endswith((".xlsx", ".xlsm")):
        try:
            import openpyxl
        except ImportError:
            return None, []
        try:
            # 不用 read_only：EMAG/美客多的 sheet dimension 元数据是假的（max_row=1/max_col=1），
            # read_only 模式会只吐 1 列。这两个文件都只有几百行，全读代价可忽略。
            wb = openpyxl.load_workbook(path, data_only=True)
        except Exception:
            return None, []
        # 逐个 sheet 找「含 SKU 字样的表头行」
        for sn in wb.sheetnames:
            ws = wb[sn]
            hdr, hdr_i = None, 0
            for i, r in enumerate(ws.iter_rows(min_row=1, max_row=min(60, ws.max_row or 60),
                                               max_col=min(60, ws.max_column or 60),
                                               values_only=True), 1):
                cells = [str(x).strip().lower() for x in r if x is not None]
                if any(c in ("sku", "vendor_ext_id") or "sku" in c for c in cells):
                    hdr, hdr_i = list(r), i
                    break
            if hdr is None:
                continue
            rest = [list(r) for r in ws.iter_rows(min_row=hdr_i + 1,
                                                  max_col=len(hdr),
                                                  values_only=True)]
            wb.close()
            return hdr, rest
        wb.close()
        return None, []

    for enc in ("utf-8-sig", "gb18030", "utf-8"):
        for err in ("strict", "replace"):
            try:
                with open(path, "r", encoding=enc, errors=err, newline="") as f:
                    sample = f.read(16384)
                    f.seek(0)
                    delim = "\t" if sample.count("\t") > sample.count(",") else ","
                    rows = list(csv.reader(f, delimiter=delim))
                if rows:
                    return rows[0], rows[1:]
            except Exception:
                continue
    return None, []


def find_cols(header):
    """定位 sku / 可售 / 在途 列号。

    返回 dict(sku=, avail=, inb=[...], degraded=bool)。
    degraded=True 表示表头已损坏、只拿到 SKU 列，数值列不可用。
    """
    if not header:
        return None
    hl = [str(h or "").strip().strip('"').lower() for h in header]

    def idx(*names):
        for n in names:
            if n in hl:
                return hl.index(n)
        return None

    sku = idx("sku", "msku", "seller-sku", "seller sku", "vendor_ext_id")
    if sku is None:
        # 表头乱码兜底：中文版报告的列 0 是「店铺 SKU」，仍带 "SKU" 字样
        for i, h in enumerate(hl):
            if "sku" in h:
                sku = i
                break
    if sku is None:
        return None

    avail = idx("afn-fulfillable-quantity", "可售数量", "可售库存",
                "proposed_quantity", "在库商品可售数量")
    w = idx("afn-inbound-working-quantity")
    s = idx("afn-inbound-shipped-quantity")
    r = idx("afn-inbound-receiving-quantity")
    inb = [x for x in (w, s, r) if x is not None]
    degraded = avail is None and not inb
    return dict(sku=sku, avail=avail, inb=inb, degraded=degraded)


def num(v):
    try:
        return float(str(v).replace(",", "").strip() or 0)
    except Exception:
        return 0.0


def scan(only_skus=None):
    """返回 {sku: {"site": {site: {...}}, ...}}"""
    data = defaultdict(lambda: defaultdict(dict))
    files = []
    for p in PERIODS:
        base = os.path.join(DSR, "..", "trunk", "data", "数据归属周期", p, "快照日期")
        base = os.path.normpath(base)
        if not os.path.isdir(base):
            continue
        for day in sorted(os.listdir(base)):
            d = os.path.join(base, day)
            if not os.path.isdir(d):
                continue
            for fp in sorted(glob.glob(os.path.join(d, "*.csv")) +
                             glob.glob(os.path.join(d, "*.txt")) +
                             glob.glob(os.path.join(d, "*.xlsx"))):
                fn = os.path.basename(fp)
                if "库存" not in fn:
                    continue
                site = site_of(fn)
                if site is None:
                    continue
                files.append((day, site, fp))

    parsed = {}
    for day, site, fp in files:
        header, rows = read_table(fp)
        cols = find_cols(header)
        if not cols:
            parsed[(day, site)] = "解析失败"
            continue
        hl = [str(h or "").strip().strip('"').lower() for h in (header or [])]
        name_idx = None
        for cand in ("product-name", "product_name", "商品名称", "产品"):
            if cand in hl:
                name_idx = hl.index(cand)
                break
        n = 0
        for r in rows:
            i = cols["sku"]
            if i >= len(r):
                continue
            sku = str(r[i] or "").strip().strip('"')
            if not SKU_RE.match(sku):
                continue
            if only_skus and sku not in only_skus:
                continue
            avail = num(r[cols["avail"]]) if cols["avail"] is not None and cols["avail"] < len(r) else None
            inb = None
            if cols["inb"]:
                inb = sum(num(r[j]) for j in cols["inb"] if j < len(r))
            rec = data[sku][site]
            rec.setdefault("days", set()).add(day)
            if avail is not None and avail > 0:
                rec.setdefault("avail_days", set()).add(day)
            if inb is not None and inb > 0:
                rec.setdefault("inb_days", set()).add(day)
            # 当前值（取最新一天）
            if day >= rec.get("last_day", ""):
                rec["last_day"] = day
                rec["cur_avail"] = avail
                rec["cur_inb"] = inb
                if name_idx is not None and name_idx < len(r) and not rec.get("name"):
                    rec["name"] = str(r[name_idx] or "")[:120]
            n += 1
        parsed[(day, site)] = ("降级·仅记出现" if cols["degraded"] else n)
    return data, files, parsed


def agg(data):
    rows = []
    for sku in sorted(data):
        sites = data[sku]
        all_days = set()
        av_days = set()
        in_days = set()
        cur_av = cur_inb = 0.0
        first_av_site = first_in_site = ""
        name = ""
        for site, rec in sites.items():
            all_days |= rec.get("days", set())
            if rec.get("avail_days"):
                av_days |= rec["avail_days"]
                if not first_av_site:
                    first_av_site = site
            if rec.get("inb_days"):
                in_days |= rec["inb_days"]
                if not first_in_site:
                    first_in_site = site
            cur_av += rec.get("cur_avail") or 0
            cur_inb += rec.get("cur_inb") or 0
            if not name and rec.get("name"):
                name = rec["name"]
        rows.append(dict(
            sku=sku, name=name,
            first_seen=min(all_days) if all_days else "",
            first_avail=min(av_days) if av_days else "",
            first_inbound=min(in_days) if in_days else "",
            first_avail_site=first_av_site,
            first_inbound_site=first_in_site,
            cur_avail=cur_av, cur_inb=cur_inb,
            n_sites=len(sites),
            state=("有货" if cur_av > 0 else
                   "在途" if cur_inb > 0 else
                   "零库存" if (av_days or in_days or all_days) else "无记录"),
        ))
    return rows


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--apply", action="store_true")
    ap.add_argument("--sku", default="")
    a = ap.parse_args()
    only = {x.strip() for x in a.sku.split(",") if x.strip()} or None

    data, files, parsed = scan(only)
    rows = agg(data)

    print(f"扫描库存报告 {len(files)} 份，涉及 {len({k[0] for k in parsed})} 个快照日")
    bad = [k for k, v in parsed.items() if v == "解析失败"]
    if bad:
        print(f"  !! 解析失败 {len(bad)} 份：{bad[:5]}")
    print(f"SKU 共 {len(rows)} 个\n")

    print(f"{'SKU':<8}{'首次出现':<12}{'首可售':<12}{'首在途':<12}{'当前可售':>9}{'当前在途':>9}  状态")
    print("-" * 78)
    for r in rows:
        print(f"{r['sku']:<8}{r['first_seen'] or '-':<12}{r['first_avail'] or '-':<12}"
              f"{r['first_inbound'] or '-':<12}{r['cur_avail']:>9.0f}{r['cur_inb']:>9.0f}  {r['state']}")

    if not a.apply:
        print("\n[干跑] 加 --apply 生成 xlsx")
        return

    import openpyxl
    from openpyxl.styles import Font, Alignment, PatternFill
    from openpyxl.utils import get_column_letter

    # ── 保留历史「起售日期」（2026-10-24 修）───────────────────────────────
    # 这个脚本每次跑批前都会重建整个 xlsx，会把 runner 写进来的「起售日期」列清掉。
    # 起售日期是**历史事实**（首次有可售库存那天），不能被后续跑批改写或抹掉，
    # 所以重建前先把旧值捞出来，写新表时原样带回去。
    _keep_start = {}
    if os.path.exists(OUT_XLSX):
        try:
            _old = openpyxl.load_workbook(OUT_XLSX, data_only=True, read_only=True)
            _ows = _old["首库存登记"]
            _oh = next(_ows.iter_rows(min_row=1, max_row=1, values_only=True), ())
            if "起售日期" in _oh:
                _ci = _oh.index("起售日期")
                for _r in _ows.iter_rows(min_row=2, values_only=True):
                    if _r and _r[0] and len(_r) > _ci and _r[_ci]:
                        _keep_start[str(_r[0]).strip()] = str(_r[_ci]).strip()
            _old.close()
        except Exception as _e:
            print(f"  !! 读旧起售日期失败（将重建为空）：{_e}")

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "首库存登记"
    hdr = ["SKU", "型号", "首次出现", "首可售日", "首在途日",
           "首可售站点", "首在途站点", "当前可售", "当前在途", "涉及站点数", "状态",
           "起售日期"]
    ws.append(hdr)
    for r in rows:
        ws.append([r["sku"], r["name"], r["first_seen"], r["first_avail"], r["first_inbound"],
                   r["first_avail_site"], r["first_inbound_site"],
                   r["cur_avail"], r["cur_inb"], r["n_sites"], r["state"],
                   # 起售日期必须来自「首可售日」；没有首可售日（只有过在途）
                   # 一律留空 —— 用户口径：在途不算起售。
                   (_keep_start.get(str(r["sku"]).strip(), "") if r["first_avail"] else "")])
    if _keep_start:
        _n = sum(1 for r in rows if str(r["sku"]).strip() in _keep_start)
        print(f"  保留历史起售日期 {_n} 条（共 {len(_keep_start)} 条在旧表里）")

    ws2 = wb.create_sheet("站点明细")
    ws2.append(["SKU", "站点", "首次出现", "首可售日", "首在途日",
                "最新日期", "最新可售", "最新在途"])
    for sku in sorted(data):
        for site in sorted(data[sku]):
            rec = data[sku][site]
            ws2.append([
                sku, site,
                min(rec.get("days", set())) if rec.get("days") else "",
                min(rec["avail_days"]) if rec.get("avail_days") else "",
                min(rec["inb_days"]) if rec.get("inb_days") else "",
                rec.get("last_day", ""), rec.get("cur_avail"), rec.get("cur_inb"),
            ])

    for sh, widths in ((ws, [9, 52, 11, 11, 11, 18, 18, 10, 10, 10, 9, 11]),
                       (ws2, [9, 20, 11, 11, 11, 11, 10, 10])):
        for i, w in enumerate(widths, 1):
            sh.column_dimensions[get_column_letter(i)].width = w
        for c in sh[1]:
            c.font = Font(bold=True)
            c.fill = PatternFill("solid", fgColor="DDEBF7")
            c.alignment = Alignment(horizontal="center")
        sh.freeze_panes = "A2"

    os.makedirs(os.path.dirname(OUT_XLSX), exist_ok=True)
    wb.save(OUT_XLSX)
    print(f"\n已写入 {OUT_XLSX}  （{len(rows)} 行 + {sum(len(v) for v in data.values())} 条站点明细）")


if __name__ == "__main__":
    main()
