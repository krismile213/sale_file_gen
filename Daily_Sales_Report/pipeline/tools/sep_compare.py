# -*- coding: utf-8 -*-
"""9月销量对比：人工(backup每日销售数据) vs 程序(快照订单报告独立重算)

口径（与日销表一致）：
  有效性 = order-status != Cancelled 且 sales-channel 以 "Amazon." 开头
           （北美额外含 merchant-order-id 带 TEMU 的单）
  归属   = purchase-date 左10位（下单日，报告原文时区）
  值     = quantity 求和

人工侧 = backup_20261008 的每日销售数据（跑批前原貌，人工粘贴到下载日 9-18）
程序侧 = 快照 202609 订单报告全量重算（含完整 9 月）

对比窗口 = 人工侧 9 月最大下单日（自动取人工贴到的最后一天）
输出 = 9月销量对比-程序vs人工.xlsx
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
import csv, os, re, sys
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import config as C

BACKUP = C.BACKUP_DIR
# ⚠️ 本工具是「9月 vs 10月」一次性对撞，固定读 **9 月快照**；不能跟 C.SNAP 走
# （后者自 2026-10-14 起=「最新快照」，会漂到 10 月，导致 202609 订单报告找不到）。
SNAP = os.path.join(C.SNAP_BASE, "202609", "快照日期", "20261006")
OUT_XLSX = os.path.join(C.DSR, "9月销量对比-程序vs人工.xlsx")

ITEMS = [
    ("欧盟英国",      "欧盟英国库存销售表.xlsx",              "每日销售数据",      "亚马逊欧洲站点202609订单报告.txt",              False),
    ("北美(美加墨)",  "北美加拿大库存销售表.xlsx",            "每日销售数据",      "亚马逊北美站点202609订单报告.txt",              True),
    ("日本",          "日本库存销售表.xlsx",                  "每日销售数据",      "亚马逊日本站点202609订单报告.txt",              False),
    ("沙特",          "沙特库存销售表.xlsx",                  "每日销售数据沙特",  "亚马逊沙特阿拉伯站点202609订单报告.txt",        False),
    ("澳大利亚",      "澳大利亚库存销售表.xlsx",              "每日销售数据",      "亚马逊澳大利亚站点202609订单报告.txt",          False),
    ("印度",          "印度库存销售表.xlsx",                  "每日销售数据",      "亚马逊印度站点202609订单报告.txt",              False),
    ("阿联酋",        "阿联酋库存销售表.xlsx",                "每日销售数据阿联酋","亚马逊阿拉伯联合酋长国站点202609订单报告.txt",  False),
    ("Interflowd北美","Interflowd北美加拿大库存销售表.xlsx",  "每日销售数据",      "interflow北美站点202609订单报告.txt",           True),
    ("Interflowd澳洲","Interflowd澳大利亚库存销售表.xlsx",    "每日销售数据",      "interflow澳大利亚站点202609订单报告.txt",       False),
    ("Interflowd阿联酋","Interflowd阿联酋库存销售表.xlsx",    "每日销售数据阿联酋","interflow阿拉伯联合酋长国站点202609订单报告.txt", False),
]

SEP_MONTH = ("2026-09-01", "2026-09-30")


def pdate(v):
    if isinstance(v, datetime):
        return v.strftime("%Y-%m-%d")
    s = str(v).strip()
    return s[:10] if len(s) >= 10 else s


def qint(v):
    try:
        return int(float(v))
    except (TypeError, ValueError):
        return 0


def is_valid(status, channel, moid, temu):
    if (status or "").strip() == "Cancelled":
        return False
    ch = (channel or "").strip()
    if ch.startswith("Amazon."):
        return True
    if temu and "TEMU" in (moid or "").upper():
        return True
    return False


def new_agg():
    return dict(day={}, chan={}, sku={}, total=0, nrows=0)


def add(m, day, chan, sku, qty):
    m["day"][day] = m["day"].get(day, 0) + qty
    m["chan"][chan] = m["chan"].get(chan, 0) + qty
    m["sku"][sku] = m["sku"].get(sku, 0) + qty
    m["total"] += qty
    m["nrows"] += 1


def read_backup_sales(fp, sheet_name):
    """人工侧：backup 工作簿每日销售数据 → 9月有效行"""
    wb = openpyxl.load_workbook(fp, read_only=True, data_only=True)
    ws = wb[sheet_name]
    hdr, idx = None, {}
    for row in ws.iter_rows(min_row=1, max_row=6, max_col=50):
        vals = {str(c.value).strip(): c.column for c in row if c.value is not None}
        if any(k.startswith("order-item-id") for k in vals):
            hdr = row[0].row
            idx = vals
            break
    if hdr is None:
        wb.close()
        raise RuntimeError(f"{fp} 找不到表头行")
    need = ["purchase-date", "order-status", "sales-channel", "quantity", "sku", "merchant-order-id"]
    col = {}
    for k in need:
        c = idx.get(k)
        if c is None:
            for kk, cc in idx.items():
                if kk.startswith(k):
                    c = cc
                    break
        col[k] = c
    out = []
    for row in ws.iter_rows(min_row=hdr + 1, max_col=50):
        pd_ = row[col["purchase-date"] - 1].value if col["purchase-date"] else None
        d = pdate(pd_) if pd_ is not None else ""
        if not (SEP_MONTH[0] <= d <= SEP_MONTH[1]):
            continue
        st = row[col["order-status"] - 1].value if col["order-status"] else None
        ch = row[col["sales-channel"] - 1].value if col["sales-channel"] else None
        qt = row[col["quantity"] - 1].value if col["quantity"] else None
        sk = row[col["sku"] - 1].value if col["sku"] else None
        mo = row[col["merchant-order-id"] - 1].value if col["merchant-order-id"] else None
        out.append((d, (st or "").strip() if st else "", (ch or "").strip() if ch else "",
                    qint(qt), (str(sk).strip() if sk else ""), (str(mo).strip() if mo else "")))
    wb.close()
    return out


def read_report_orders(fp, sep="\t"):
    for enc in C.ENC_CANDIDATES:
        try:
            with open(fp, encoding=enc, newline="") as f:
                rows = list(csv.reader(f, delimiter=sep))
            break
        except UnicodeDecodeError:
            continue
    hdr = [h.strip() for h in rows[0]]
    ix = {h: i for i, h in enumerate(hdr)}
    out = []
    for r in rows[1:]:
        if len(r) <= max(ix.get("purchase-date", 0), ix.get("quantity", 0)):
            continue
        d = (r[ix["purchase-date"]].strip())[:10]
        if not (SEP_MONTH[0] <= d <= SEP_MONTH[1]):
            continue
        out.append((d, r[ix["order-status"]].strip(), r[ix["sales-channel"]].strip(),
                    qint(r[ix["quantity"]]), r[ix["sku"]].strip(),
                    r[ix["merchant-order-id"]].strip()))
    return out


def agg(rows, temu, d0=None, d1=None):
    m = new_agg()
    for d, st, ch, qt, sk, mo in rows:
        if d0 and not (d0 <= d <= d1):
            continue
        if not is_valid(st, ch, mo, temu):
            continue
        add(m, d, ch or "(空)", sk or "(空)", qt)
    return m


# ---------------- 透视图锚点（人工透视缓存值） ----------------
def read_pivot_cache():
    """从 9 月工作副本抓欧盟/北美透视图的渠道汇总缓存值"""
    res = {}
    wb = openpyxl.load_workbook(os.path.join(C.WORK_DIR, "欧盟英国库存销售表.xlsx"),
                                read_only=True, data_only=True)
    ws = wb["数据透视表"]
    for row in ws.iter_rows(min_row=5, max_row=40, min_col=10, max_col=11):
        lab, val = row[0].value, row[1].value
        if lab and isinstance(val, (int, float)):
            res[str(lab).strip()] = int(val)
    wb.close()
    na = {}
    wb = openpyxl.load_workbook(os.path.join(C.WORK_DIR, "北美加拿大库存销售表.xlsx"),
                                read_only=True, data_only=True)
    ws = wb["数据透视图"]
    for row in ws.iter_rows(min_row=5, max_row=12, min_col=10, max_col=11):
        lab, val = row[0].value, row[1].value
        if lab and isinstance(val, (int, float)):
            na[str(lab).strip()] = int(val)
    wb.close()
    return {"欧盟英国": res, "北美(美加墨)": na}


# ---------------- 输出 Excel ----------------
THIN = Border(*[Side(style="thin", color="D0D0D0")] * 4)
H_FILL = PatternFill("solid", fgColor="1F4E79")
H_FONT = Font(color="FFFFFF", bold=True, size=10)
DIFF_FILL = PatternFill("solid", fgColor="FFF2A8")
TOT_FONT = Font(bold=True)
SEC_FONT = Font(bold=True, size=11, color="1F4E79")


def sec_title(ws, r, text):
    c = ws.cell(row=r, column=1, value=text)
    c.font = SEC_FONT
    return r + 1


def write_table(ws, r0, headers, rows, diff_cols=(), total_row=True):
    """写小表，返回下一可用行。diff_cols: 差异列号(1-based within headers)高亮非零"""
    for j, h in enumerate(headers, 1):
        c = ws.cell(row=r0, column=j, value=h)
        c.fill, c.font, c.border = H_FILL, H_FONT, THIN
        c.alignment = Alignment(horizontal="center")
    r = r0
    for row in rows:
        r += 1
        for j, v in enumerate(row, 1):
            c = ws.cell(row=r, column=j, value=v)
            c.border = THIN
            if j in diff_cols and isinstance(v, (int, float)) and v != 0:
                c.fill = DIFF_FILL
        if total_row and row is rows[-1] and str(row[0]).startswith(("合计", "总计")):
            for j in range(1, len(headers) + 1):
                ws.cell(row=r, column=j).font = TOT_FONT
    return r + 2


def main():
    results = {}   # name -> dict(manual=..., prog_full=..., prog_win=..., win=(d0,d1))
    for name, fn, sn, rp, temu in ITEMS:
        print(f"== {name}")
        bp = os.path.join(BACKUP, fn)
        man_rows = read_backup_sales(bp, sn)
        man_full = agg(man_rows, temu)
        man_days = [d for d in man_full["day"]]
        if man_days:
            win_end = max(man_days)
        else:
            win_end = "2026-09-17"
        rep_rows = read_report_orders(os.path.join(SNAP, rp))
        prog_full = agg(rep_rows, temu)
        prog_win = agg(rep_rows, temu, "2026-09-01", win_end)
        results[name] = dict(man=man_full, prog=prog_full, prog_win=prog_win,
                             win=("2026-09-01", win_end), nrep=len(rep_rows))
        print(f"   人工(9月): {man_full['total']}件/{man_full['nrows']}行 窗口末={win_end} | "
              f"程序(窗口): {prog_win['total']}件 | 程序(全月): {prog_full['total']}件")

    piv = read_pivot_cache()

    wb = openpyxl.Workbook()
    # ---------- 总览 ----------
    ws = wb.active
    ws.title = "总览"
    ws.sheet_view.showGridLines = False
    ws.cell(row=1, column=1, value="9月销量对比总览（单位：件 quantity 求和）").font = Font(bold=True, size=13)
    ws.cell(row=2, column=1, value="口径：非Cancelled 且 sales-channel 以 Amazon. 开头（北美额外含 merchant-order-id 带 TEMU 的单）；按 purchase-date 前10位（下单日）归属").font = Font(size=9, color="808080")
    ws.cell(row=3, column=1, value="数据源：人工=backup_20261008（=根目录12份9月工作副本在10-08跑批前的原样备份，人工做到9-17，Interflowd三份停在9-14）；程序=202609快照订单报告独立重算").font = Font(size=9, color="808080")
    CAUSE = {
        "欧盟英国": "取消时差(净-3)",
        "北美(美加墨)": "9-01报告缺头(-116) + 下载日时点差(+168) + 取消时差",
        "日本": "取消时差",
        "沙特": "取消时差(微)",
        "澳大利亚": "9-01报告头部截断 + 取消时差",
        "印度": "取消时差(微)",
        "阿联酋": "完全一致",
        "Interflowd北美": "9-01报告头部截断(10:03起)",
        "Interflowd澳洲": "完全一致",
        "Interflowd阿联酋": "完全一致",
    }
    hdr = ["国家", "对比窗口", "人工件数", "程序件数(同窗口)", "差异(程序-人工)",
           "人工订单行", "程序订单行(同窗口)", "程序9月全月件数", "差异主因"]
    rows = []
    for name, *_ in ITEMS:
        r = results[name]
        d0, d1 = r["win"]
        rows.append([name, f"{d0}~{d1}", r["man"]["total"], r["prog_win"]["total"],
                     r["prog_win"]["total"] - r["man"]["total"],
                     r["man"]["nrows"], r["prog_win"]["nrows"], r["prog"]["total"],
                     CAUSE.get(name, "")])
    tot = ["合计", "", sum(x[2] for x in rows), sum(x[3] for x in rows),
           sum(x[4] for x in rows), sum(x[5] for x in rows), sum(x[6] for x in rows),
           sum(x[7] for x in rows), ""]
    rows.append(tot)
    nxt = write_table(ws, 4, hdr, rows, diff_cols=(5,))
    ws.cell(row=nxt, column=1, value="注1：美客多/罗马尼亚快照目录无订单报告，无法独立重算，未纳入本对比。").font = Font(size=9, color="808080")
    ws.cell(row=nxt + 1, column=1, value="注2：差异主因详见「差异根因」sheet——核心发现：快照订单报告丢掉了 UTC 9-01 07:00 之前的北美订单（导出date range按美西时间设置）。").font = Font(size=9, color="C00000")
    for col, w in zip("ABCDEFGHI", [16, 24, 10, 14, 14, 10, 16, 14, 44]):
        ws.column_dimensions[col].width = w

    # ---------- 透视图锚点 ----------
    ws2 = wb.create_sheet("透视图锚点对比")
    ws2.sheet_view.showGridLines = False
    ws2.cell(row=1, column=1, value="人工透视图缓存值 vs 程序同窗口重算").font = Font(bold=True, size=13)
    ws2.cell(row=2, column=1, value="人工缓存=9月工作副本透视图sheet静态值（下载日期9-18，即只含9-01~人工最后粘贴日的单）；程序=同窗口订单报告重算").font = Font(size=9, color="808080")
    r = 4
    for name, cache in piv.items():
        if not cache:
            continue
        r = sec_title(ws2, r, f"◆ {name}")
        r0 = results[name]
        d0, d1 = r0["win"]
        rows = []
        tot_m = tot_p = 0
        for chan, mv in cache.items():
            if chan in ("总计", "求和项:quantity", "行标签"):
                continue
            pv = r0["prog_win"]["chan"].get(chan, 0)
            rows.append([chan, mv, pv, pv - mv])
            tot_m += mv
            tot_p += pv
        pv_total = cache.get("总计")
        rows.append(["合计(渠道加总)", tot_m, tot_p, tot_p - tot_m])
        if pv_total is not None:
            rows.append(["透视表内总计", pv_total, tot_p, tot_p - pv_total])
        r = write_table(ws2, r, ["渠道", "人工缓存值", f"程序({d0}~{d1})", "差异"], rows, diff_cols=(4,))
    for col, w in zip("ABCD", [18, 12, 16, 10]):
        ws2.column_dimensions[col].width = w

    # ---------- 差异根因 ----------
    ws4 = wb.create_sheet("差异根因")
    ws4.sheet_view.showGridLines = False
    ws4.cell(row=1, column=1, value="差异根因（逐条实测验证）").font = Font(bold=True, size=13)
    causes = [
        ["#", "根因", "影响国家", "量级", "证据"],
        ["1", "快照订单报告导出 date range 按美西时间 9-01 00:00 设定，\n丢失了 UTC 9-01 07:00 之前的全部订单",
         "北美(+116件人工多)\ninterflow北美 / interflow澳洲 / 阿联酋 9-01 头部疑似同样被截",
         "北美 9-01：人工 308 行(00:00~23:59 全天) vs 报告 189 行(07:04 起)；\n报告 8-31=0 行而欧洲报告 8-31 有 66 行(欧洲未截)；\n9-02 起两边完全一致(271行、时间范围一致)"],
        ["2", "下载日时点差：人工在 9-18 下载报告只贴到下载时刻，\n快照报告(10-06导出)含 9-18 全天",
         "全部国家（窗口最后一天）",
         "北美 9-18：人工 74 件 vs 报告 242 件；\n这是数据时点性质，不算错误。窗口末日对比请人工补齐后再看"],
        ["3", "取消时差：人工按天贴单，贴完之后才取消的单\n人工表里状态仍是 Shipped；快照报告是最终状态(Cancelled)",
         "全部国家（每天 0~5 件）",
         "北美 9-01 后每天差 1~4 件；日本累计 -29 件、澳洲 -9 件属于此模式"],
        ["4", "结论：差异 ≠ 程序算错，而是「报告源不完整 + 时点差」。\n人工的每日销售数据（每天贴）比一次性导出的快照报告更完整",
         "修正方法：重导订单报告时 date range 取 8-25 ~ 10-06\n（跨月边界多留几天），再重跑 sep_compare 即可归零",
         "除上述 1/2/3 外，窗口内无任何口径差异：\n渠道分布、SKU 分布均能逐行对上"],
    ]
    r = 3
    for i, row in enumerate(causes):
        for j, v in enumerate(row, 1):
            c = ws4.cell(row=r, column=j, value=v)
            c.border = THIN
            c.alignment = Alignment(wrap_text=True, vertical="top")
            if i == 0:
                c.fill, c.font = H_FILL, H_FONT
            elif row[0] == "4":
                c.font = Font(bold=True, color="C00000")
    ws4.column_dimensions["A"].width = 4
    ws4.column_dimensions["B"].width = 46
    ws4.column_dimensions["C"].width = 34
    ws4.column_dimensions["D"].width = 20
    ws4.column_dimensions["E"].width = 60
    for i in range(4, 3 + len(causes)):
        ws4.row_dimensions[i].height = 66

    # ---------- 每国明细 ----------
    for name, fn, sn, rp, temu in ITEMS:
        r0 = results[name]
        d0, d1 = r0["win"]
        ws3 = wb.create_sheet(name[:28])
        ws3.sheet_view.showGridLines = False
        ws3.cell(row=1, column=1, value=f"{name} 9月销量对比（人工 backup vs 程序 报告重算）").font = Font(bold=True, size=12)
        # 按日
        days = sorted(set(r0["man"]["day"]) | set(r0["prog"]["day"]))
        r = sec_title(ws3, 3, f"◆ 按日（窗口 {d0}~{d1} 高亮区，全月为程序值）")
        rows = []
        tm = tp = tw = tf = 0
        for d in days:
            m = r0["man"]["day"].get(d, 0)
            f = r0["prog"]["day"].get(d, 0)
            w = r0["prog_win"]["day"].get(d, 0) if d <= d1 else None
            rows.append([d, m if d <= d1 else "", w if w is not None else "", (w - m) if d <= d1 else "", f])
            if d <= d1:
                tm += m
                tp += w if w is not None else 0
            tf += f
        rows.append(["合计", tm, tp, tp - tm, tf])
        r = write_table(ws3, r, ["日期", "人工(窗口内)", "程序(窗口内)", "差异", "程序(9月全月)"],
                        rows, diff_cols=(4,))
        # 按渠道
        chans = sorted(set(r0["man"]["chan"]) | set(r0["prog_win"]["chan"]),
                       key=lambda c: -(r0["prog_win"]["chan"].get(c, 0) + r0["man"]["chan"].get(c, 0)))
        r = sec_title(ws3, r, f"◆ 按渠道（窗口 {d0}~{d1}）")
        rows = []
        tm = tp = 0
        for ch in chans:
            m = r0["man"]["chan"].get(ch, 0)
            p = r0["prog_win"]["chan"].get(ch, 0)
            rows.append([ch, m, p, p - m, r0["prog"]["chan"].get(ch, 0)])
            tm += m
            tp += p
        rows.append(["合计", tm, tp, tp - tm, r0["prog"]["total"]])
        r = write_table(ws3, r, ["渠道", "人工(窗口内)", "程序(窗口内)", "差异", "程序(9月全月)"],
                        rows, diff_cols=(4,))
        # SKU 差异 Top
        skus = set(r0["man"]["sku"]) | set(r0["prog_win"]["sku"])
        diffs = [(s, r0["man"]["sku"].get(s, 0), r0["prog_win"]["sku"].get(s, 0)) for s in skus]
        diffs = [x for x in diffs if x[2] - x[1] != 0]
        diffs.sort(key=lambda x: -abs(x[2] - x[1]))
        r = sec_title(ws3, r, f"◆ SKU差异Top20（窗口 {d0}~{d1}，共{len(diffs)}个SKU有差异）")
        rows = [[s, m, p, p - m] for s, m, p in diffs[:20]]
        if rows:
            r = write_table(ws3, r, ["SKU", "人工(窗口内)", "程序(窗口内)", "差异"], rows, diff_cols=(4,), total_row=False)
        else:
            ws3.cell(row=r, column=1, value="（窗口内无 SKU 级差异）").font = Font(size=9, color="808080")
        for col, w in zip("ABCDE", [22, 13, 13, 10, 14]):
            ws3.column_dimensions[col].width = w

    wb.save(OUT_XLSX)
    print(f"\n输出 → {OUT_XLSX}")

    # 控制台摘要
    print("\n===== 摘要 =====")
    for name, *_ in ITEMS:
        r = results[name]
        d = r["prog_win"]["total"] - r["man"]["total"]
        print(f"{name}: 人工{r['man']['total']} 程序{r['prog_win']['total']} 差异{d:+d} | 程序全月 {r['prog']['total']}")


if __name__ == "__main__":
    main()
