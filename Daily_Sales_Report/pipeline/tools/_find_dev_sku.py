# -*- coding: utf-8 -*-
"""查 G598/G599/G600/G601（开发中）在库存报告 / 订单报告 / 12 本工作簿里的出现情况。

要点：
- 前缀匹配（实际 SKU 常带后缀，如 G598-2PACK），不是精确等值
- 全部结果写文件，不走 stdout（之前 stdout 被吞过）
- 库存报告统一看 sku 列；订单报告看首行表头里的 sku 列
"""
import os, sys, csv, glob, io, json

DSR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
WORK = os.path.join(DSR, "1.每日销售数据表")
SNAP = os.path.normpath(os.path.join(DSR, "..", "trunk", "data", "数据归属周期",
                                     "202610", "快照日期", "20261023"))
TARGETS = ["G598", "G599", "G600", "G601"]
OUT = os.path.join(DSR, "pipeline", "tools", "_find_dev_sku_out.txt")

lines = []
def p(s=""):
    lines.append(str(s))


def pref(v):
    """返回命中的前缀；不命中返回 None"""
    s = str(v).strip().strip('"') if v is not None else ""
    if not s:
        return None
    for t in TARGETS:
        if s == t or s.startswith(t + "-") or s.startswith(t + "_") \
           or s.startswith(t + " ") or s.startswith(t + "."):
            return t
        # 也容忍 G598A / G5980 这类紧贴后缀（同系列编号）
        if s.startswith(t) and len(s) > len(t) and not s[len(t)].isdigit():
            return t
    return None


def read_csv_rows(path):
    """返回 (header, rows)。自动嗅探编码/分隔符。

    注意：澳大利亚/日本/澳interflow 三份库存报告是 GBK 且表头已含不可逆的 '?'
    替换，任何严格解码都会失败 —— 用 errors="replace" 兜底，表头乱码不影响
    SKU 列（值是 ASCII），调用方会退化为按列位取。
    """
    for enc in ("utf-8-sig", "gb18030", "utf-8"):
        for err in ("strict", "replace"):
            try:
                with open(path, "r", encoding=enc, errors=err, newline="") as f:
                    sample = f.read(8192)
                    f.seek(0)
                    delim = "\t" if sample.count("\t") > sample.count(",") else ","
                    rows = list(csv.reader(f, delimiter=delim))
                if rows:
                    return rows[0], rows[1:]
            except Exception:
                continue
    return None, []


# ---------------------------------------------------------------- ① 库存报告
p("=" * 88)
p("① 库存报告（20261023 快照目录）—— 按 sku 列前缀匹配")
p("=" * 88)
inv_hits = {t: [] for t in TARGETS}

files = sorted(glob.glob(os.path.join(SNAP, "*")))
files = [f for f in files if f.lower().endswith((".csv", ".txt"))]
for fp in files:
    base = os.path.basename(fp)
    if "订单报告" in base:
        continue                       # 订单报告单独在第 ② 段
    header, rows = read_csv_rows(fp)
    if not header:
        p(f"  [跳过] {base} 无法解析")
        continue
    hlow = [h.strip().strip('"').lower() for h in header]
    sku_idx = None
    for cand in ("sku", "msku", "seller-sku", "seller sku", "商品sku", "sku编码"):
        if cand in hlow:
            sku_idx = hlow.index(cand)
            break
    # 表头乱码兜底：库存报告 SKU 恒在前两列（店铺SKU / 卖家SKU）
    if sku_idx is None:
        p(f"  [表头乱码，按列位0/1兜底] {base[:40]}")
        cands = [0, 1]
    else:
        cands = [sku_idx]
    n = 0
    for r in rows:
        t = None
        for ci in cands:
            if ci < len(r):
                t = pref(r[ci])
                if t:
                    break
        if t:
            n += 1
            rec = {}
            for i, h in enumerate(hlow):
                if i < len(r) and i < 40:
                    rec[h] = r[i]
            inv_hits[t].append((base, rec))
    p(f"  {base[:52]:<54} sku列={hlow[sku_idx] if sku_idx is not None else '0/1':<10} 命中 {n}")

p()
for t in TARGETS:
    p(f"--- {t} : 库存命中 {len(inv_hits[t])} 条 ---")
    for base, rec in inv_hits[t]:
        keep = {k: v for k, v in rec.items()
                if k in ("sku", "asin", "product-name", "afn-fulfillable-quantity",
                         "afn-total-quantity", "afn-inbound-working-quantity",
                         "afn-inbound-shipped-quantity", "afn-inbound-receiving-quantity",
                         "afn-reserved-quantity", "store", "fnsku")}
        p(f"   [{base[:34]}] {json.dumps(keep, ensure_ascii=False)}")
    p()

# ---------------------------------------------------------------- ② 订单报告
p("=" * 88)
p("② 订单报告（20261023）—— 按 sku 列前缀匹配，统计销量")
p("=" * 88)
for fp in files:
    base = os.path.basename(fp)
    if "订单报告" not in base:
        continue
    header, rows = read_csv_rows(fp)
    if not header:
        p(f"  [跳过] {base} 无法解析")
        continue
    hlow = [h.strip().strip('"').lower() for h in header]
    sku_idx = None
    for cand in ("sku", "msku", "seller-sku", "seller sku"):
        if cand in hlow:
            sku_idx = hlow.index(cand)
            break
    if sku_idx is None:
        p(f"  [表头乱码，跳过] {base}  header={hlow[:6]}")
        continue
    cnt = {t: 0 for t in TARGETS}
    for r in rows:
        if sku_idx < len(r):
            t = pref(r[sku_idx])
            if t:
                cnt[t] += 1
    nz = {k: v for k, v in cnt.items() if v}
    p(f"  {base[:52]:<54} {'命中 ' + str(nz) if nz else '无命中'}")

# ---------------------------------------------------------------- ③ 12 本工作簿
p()
p("=" * 88)
p("③ 12 本工作簿 A 列前缀匹配（含所有 sheet）")
p("=" * 88)
try:
    import openpyxl
except ImportError:
    p("  openpyxl 未安装，跳过")
else:
    wb_hits = {t: [] for t in TARGETS}
    for fn in sorted(os.listdir(WORK)):
        if not fn.endswith(".xlsx") or fn.startswith("~$"):
            continue
        try:
            wb = openpyxl.load_workbook(os.path.join(WORK, fn), data_only=True,
                                        read_only=True)
        except Exception as e:
            p(f"  !! {fn}: {e}")
            continue
        for sn in wb.sheetnames:
            ws = wb[sn]
            ri = 0
            for row in ws.iter_rows(min_col=1, max_col=1, values_only=True):
                ri += 1
                if ri > 4000:
                    break
                t = pref(row[0] if row else None)
                if t:
                    wb_hits[t].append((fn, sn, ri, row[0]))
        wb.close()
    p()
    for t in TARGETS:
        p(f"--- {t} : 工作簿命中 {len(wb_hits[t])} 处 ---")
        for fn, sn, r, v in wb_hits[t]:
            p(f"   {fn[:30]:<32}{sn[:24]:<26} 行{r:<6} {v}")
        if not wb_hits[t]:
            p("   （无）")
        p()

with open(OUT, "w", encoding="utf-8") as f:
    f.write("\n".join(lines))
print("WROTE", OUT, len(lines), "lines")
