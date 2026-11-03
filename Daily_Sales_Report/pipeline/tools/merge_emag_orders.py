# -*- coding: utf-8 -*-
"""把拆成多份的 EMAG 全站点订单报告合并成一份（表头统一为中文）。

背景（2026-10-27）：用户把 eMAG 全站点订单报告按日期段拆成两份导出：
  EMAG全站点202610订单报告1.xlsx  10-08 ~ 10-26，100 单，中文表头
  EMAG全站点202610订单报告2.xlsx  09-30 ~ 10-08， 25 单，英文表头
两份订单号零交集、日期互补。pipeline 只认 `EMAG全站点{m}订单报告.xlsx`，
所以要合并成一份。英文表头按列序映射回中文（两份 32 列同序，已实测）。

用法：
  python merge_emag_orders.py 20261027            # 干跑，只报数
  python merge_emag_orders.py 20261027 --apply    # 生成合并文件（不删原件）
"""
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
DSR = HERE
for _ in range(6):
    if os.path.isdir(os.path.join(DSR, "1.每日销售数据表")):
        break
    DSR = os.path.dirname(DSR)
PROJECT = os.path.dirname(DSR)

# 英文表头 → 中文（与中文版报告同序，这里显式列出，不依赖列序）
EN2CN = {
    "Order number": "订单编号",
    "Order date": "订单日期",
    "AWB number": "空运提单号码",
    "Product name": "产品名称",
    "Product code": "产品代码",
    "PNK": "PNK",
    "Serial numbers": "Serial numbers",
    "Quantity": "数量",
    "Price without VAT/Unit": "不含增值税价格/件",
    "Price with vat": "含增值税价格",
    "Currency": "货币",
    "VAT": "增值税",
    "Order status": "订单状态",
    "Payment method": "付款方式",
    "Shipping method": "发货方式",
    "External delivery point ID": "派送地点外部ID",
    "Delivery point name": "派送地点名称",
    "Payment status": "付款状态",
    "Maximum finalization date": "最晚完成日期",
    "Maximum pickup date": "Pickup",
    "Customer name": "客户姓名",
    "Juridical person": "是法人",
    "VAT number": "增值税号",
    "Phone number": "电话号码",
    "Shipping name": "交货方名称",
    "Shipping phone": "交货电话",
    "Shipping address": "送货地址",
    "Shipping postal code": "送货邮编",
    "Billing name": "发票持有人",
    "Billing address": "开票地址",
    "Billing postal code": "发票邮编",
    "Observations": "评述",
}

KEY = "订单编号"
DATE = "订单日期"


def read_xlsx(path):
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]
    it = ws.iter_rows(values_only=True)
    hdr = [str(x).strip() if x is not None else "" for x in next(it)]
    rows = [list(r) for r in it if r and r[0] is not None]
    wb.close()
    return hdr, rows


def find_parts(snap, period):
    """返回 [(序号, 路径)]。

    兼容 eMAG/浏览器常见的分段命名（月底一页 100 条，订单多时会导出多份）：
      EMAG全站点202610订单报告1.xlsx      ← eMAG 手工加序号
      EMAG全站点202610订单报告(1).xlsx    ← 浏览器重名自动加 (1)
      EMAG全站点202610订单报告（1）.xlsx  ← 中文括号
      EMAG全站点202610订单报告-2.xlsx / _2 / " 2"
    不匹配主文件 `EMAG全站点{period}订单报告.xlsx`（无数字）。
    """
    out = []
    if not os.path.isdir(snap):
        return out
    pat = re.compile(r"^EMAG全站点%s订单报告[\s\-_（(]*(\d+)[\s）)]*\.xlsx$" % period)
    for f in sorted(os.listdir(snap)):
        mm = pat.match(f)
        if mm:
            out.append((int(mm.group(1)), os.path.join(snap, f)))
    out.sort()
    return out


def _key_idx(hdr):
    """表头里「订单编号」的下标（英文表头先映射）。"""
    m = [EN2CN.get(h, h) for h in hdr]
    return m.index(KEY) if KEY in m else None


def order_ids(path):
    """读一份报告的订单号集合；读失败返回 None。"""
    try:
        hdr, rows = read_xlsx(path)
    except Exception:
        return None
    i = _key_idx(hdr)
    if i is None:
        return None
    return {str(r[i]).strip() for r in rows if i < len(r) and str(r[i]).strip()}


def _gaps(days):
    """日期断档：[(前, 后, 缺失天数)]。"""
    from datetime import date
    out = []
    for a, b in zip(days, days[1:]):
        try:
            d = (date.fromisoformat(b) - date.fromisoformat(a)).days
        except ValueError:
            continue
        if d > 1:
            out.append((a, b, d - 1))
    return out


def _need_merge(target, parts):
    """要不要（重新）合并？返回 (bool, 原因)。

    · 主文件不存在 → 合并
    · 主文件存在 → 比对订单号集合：分段里出现主文件没有的单，说明主文件不完整
      （典型场景：浏览器重复下载产生 `报告(1).xlsx`，主文件只是第一页）→ 重新合并
    · 分段全被主文件覆盖 → 不动（幂等，不重写文件）
    """
    if not os.path.exists(target):
        return True, "主文件缺失"
    base = order_ids(target)
    if base is None:
        return True, "主文件读不出订单号，按不完整处理"
    extra = set()
    for _, p in parts:
        s = order_ids(p)
        if s:
            extra |= (s - base)
    if extra:
        return True, "主文件缺 %d 单（分段里有主文件没有的订单）" % len(extra)
    return False, "主文件已包含全部分段数据，不动"


def merge_day(snap, period, apply=True, quiet=False):
    """把分段 EMAG 报告合并成 `EMAG全站点{period}订单报告.xlsx`。

    返回 dict(ok, reason, target, rows, parts, stats, gaps)。
    幂等：分段数据已被主文件覆盖时不重写；确需覆盖时先备份旧主文件。
    """
    def say(*a):
        if not quiet:
            print(*a)

    target = os.path.join(snap, "EMAG全站点%s订单报告.xlsx" % period)
    parts = find_parts(snap, period)

    if not parts:
        return dict(ok=False, reason="无分段文件", target=target,
                    rows=0, parts=[], stats=[], gaps=[])

    need, why = _need_merge(target, parts)
    if not need:
        return dict(ok=False, reason=why, target=target,
                    rows=0, parts=parts, stats=[], gaps=[])
    if len(parts) < 2 and os.path.exists(target):
        # 主文件 + 单份分段：只有当分段带来新订单时才走（_need_merge 已判过）
        say("  [EMAG] 主文件之外发现 1 份补充分段（%s）" % os.path.basename(parts[0][1]))
    elif len(parts) < 2:
        return dict(ok=False, reason="只有 1 份分段，不自动改名", target=target,
                    rows=0, parts=parts, stats=[], gaps=[])

    say("  [EMAG] %s → 合并 %d 份分段：" % (why, len(parts)))
    for n, p in parts:
        say("     %d) %s" % (n, os.path.basename(p)))

    base_hdr, _ = read_xlsx(parts[0][1])
    if KEY not in base_hdr:
        base_hdr = [EN2CN.get(h, h) for h in base_hdr]

    merged, stats = {}, []
    for n, p in parts:
        hdr, rows = read_xlsx(p)
        mapped = [EN2CN.get(h, h) for h in hdr]
        unknown = [h for h in mapped if h not in base_hdr]
        idx = {}
        for i, h in enumerate(mapped):
            if h in base_hdr and h not in idx:
                idx[h] = i
        added = dup = 0
        for r in rows:
            rec = [""] * len(base_hdr)
            for h, i in idx.items():
                if i < len(r):
                    rec[base_hdr.index(h)] = r[i]
            k = str(rec[base_hdr.index(KEY)]).strip()
            if not k:
                continue
            if k in merged:
                dup += 1
                continue
            merged[k] = rec
            added += 1
        days = sorted({str(r[base_hdr.index(DATE)])[:10]
                       for r in rows if r and str(r[base_hdr.index(DATE)])[:4] == "2026"}
                      ) if rows else []
        stats.append((os.path.basename(p), len(rows), added, dup, unknown,
                      (days[0], days[-1]) if days else ("", "")))
        say("     %-40s 行=%-4d 新增=%-4d 重叠=%-3d %s"
            % (os.path.basename(p), len(rows), added, dup,
               ("%s~%s" % (days[0], days[-1])) if days else ""))

    rows = sorted(merged.values(), key=lambda r: str(r[base_hdr.index(DATE)]))
    if not rows:
        return dict(ok=False, reason="合并后 0 行，放弃", target=target,
                    rows=0, parts=parts, stats=stats, gaps=[])

    day_cnt = {}
    for r in rows:
        d = str(r[base_hdr.index(DATE)])[:10]
        day_cnt[d] = day_cnt.get(d, 0) + 1
    days = sorted(day_cnt)
    gaps = _gaps(days)
    say("     → 合并 %d 单，%s ~ %s（%d 天）" % (len(rows), days[0], days[-1], len(days)))
    if gaps:
        say("     ⚠ 日期不连续 %d 处（可能是漏导，也可能该时段确实无单，请核对）：" % len(gaps))
        for a, b, n in gaps[:6]:
            say("        %s 之后缺 %d 天 → %s" % (a, n, b))
        if len(gaps) > 6:
            say("        …共 %d 处" % len(gaps))

    if apply:
        import openpyxl
        if os.path.exists(target):
            bakdir = os.path.join(DSR, ".workbuddy", "backup_emag_merge")
            os.makedirs(bakdir, exist_ok=True)
            import shutil
            from datetime import datetime
            bak = os.path.join(bakdir, "%s-%s.xlsx" % (
                os.path.splitext(os.path.basename(target))[0],
                datetime.now().strftime("%Y%m%d%H%M%S")))
            shutil.copy2(target, bak)
            say("     → 旧主文件已备份 %s" % os.path.basename(bak))
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "纸张 1"
        ws.append(base_hdr)
        for r in rows:
            ws.append(r)
        wb.save(target)
        say("     → 已生成 %s（原件保留）" % os.path.basename(target))
    return dict(ok=True, reason=why, target=target, rows=len(rows),
                parts=parts, stats=stats, gaps=gaps)


def main():
    day = sys.argv[1] if len(sys.argv) > 1 else "20261027"
    apply = "--apply" in sys.argv
    m = day[:6]
    snap = os.path.join(PROJECT, "trunk", "data", "数据归属周期", m,
                        "快照日期", day)
    print("快照 %s" % snap)
    res = merge_day(snap, m, apply=apply, quiet=False)
    if not res["ok"]:
        print("未合并：%s" % res["reason"])
        return 0                       # 「不用合并」不是错误，正常退出
    if not apply:
        print("\n（干跑，未写文件。加 --apply 生成 %s）"
              % os.path.basename(res["target"]))
    elif res.get("gaps"):
        print("\n⚠ 合并完成，但日期有 %d 处断档（见上），请核对是否漏导了某一段。"
              % len(res["gaps"]))
    return 0


if __name__ == "__main__":
    sys.exit(main())

