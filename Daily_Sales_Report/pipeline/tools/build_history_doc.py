# -*- coding: utf-8 -*-
"""
从 history_extract_2025_2026.jsonl 生成《人工做表重要信息总档》markdown 文档。
只读 JSONL 与生成 md，不动工作簿。换期后改 extract 脚本的 MONTHS 再复跑即可。
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
import json
import os
import re
import sys
from collections import OrderedDict

sys.stdout.reconfigure(encoding="utf-8")

HERE = os.path.dirname(os.path.abspath(__file__))
JSONL = os.path.join(HERE, "history_extract_2025_2026.jsonl")
OUT = os.path.join(DSR, "人工做表重要信息总档-2025年1月至2026年8月.md")

recs = [json.loads(l) for l in open(JSONL, encoding="utf-8")]
months = list(OrderedDict.fromkeys(r["month"] for r in recs))
books = OrderedDict()
for r in recs:
    books.setdefault(r["month"], []).append(r)

# 书种：从文件名提取「{月份}-」后的部分
def species(fn):
    return re.sub(r"^\d{4}年\d{1,2}月-", "", fn).replace(".xlsx", "").replace("（不运营）", "")

species_first = OrderedDict()   # 书种 -> 首次出现月份
for m in months:
    for r in books[m]:
        sp = species(r["file"])
        species_first.setdefault(sp, m)

L = []
w = L.append

w("# 人工做表重要信息总档（2025年1月 ~ 2026年8月）")
w("")
w("> 生成方式：`daily_update/extract_history_2025_2026.py` 只读扫描全部历史工作簿提取，")
w("> 机器可读明细在 `daily_update/history_extract_2025_2026.jsonl`（每本一行）。")
w("> 用途：后续任何核对，先来这里查「当时人工做出来的数」和「当时的口径」。")
w("")
n_books = len(recs)
w(f"- **范围**：{months[0]} ~ {months[-1]}，共 {len(months)} 个月、**{n_books} 本工作簿**")
w("- **存档位置**：`1.每日销售数据表\\2025年每日数据\\{月}\\` 与 `\\2026年每日数据\\{月}\\`，文件名 `{月}-{书种}.xlsx`")
w("")
w("## 一、表种进退时间线")
w("")
w("| 书种 | 首月 | 末月(统计范围内) | 备注 |")
w("|---|---|---|---|")
for sp, first in species_first.items():
    last = max(m for m in months for r in books[m] if species(r["file"]) == sp)
    note = ""
    if "新加坡" in sp:
        note = "2025年6月标「不运营」，7月起退场"
    elif "Interflowd" in sp:
        note = "Interflow 平台分账簿（加 crawler 独立导出）"
    elif "罗马尼亚" in sp:
        note = "eMAG；库存明细源=优惠更新模板里的「库存」sheet"
    elif "美客多" in sp:
        note = "墨西哥；2025年7月为仓别口径过渡版（见大事记②）"
    w(f"| {sp} | {first} | {last} | {note} |")
w("")
w("## 二、口径演变大事记（核对前必读）")
w("")
w("1. **日窗恒 38 列**：起点通常为上月 25 日（30 天月的上个月 24 日、2 月前的 1 月 22 日），")
w("   终点 = 当月末；前 7~10 天是上月末，用于「实时速度/上月日均」等滚动指标。")
w("   ⚠️ **例外（2025年2、3月两期）**：全部书的日窗为 1/25~3/3——2 月期随 38 列模板溢出到 3 月初，")
w("   3 月期未滚动窗口。但这不影响「月总销量」：3 月簿「3月总销量」为全月量级（如日本 3,252），")
w("   **月总列的值与日窗终点脱钩**，核对时不要用日窗终点推断数据覆盖范围。")
w("2. **2025-04 起分析表整体左移一列**：预留库存 F→E、可售 G→F、在途 H→G、调拨合计 I→H、")
w("   在途差额 T→S、实时速度 U→T、上月日均 V→U（1-3 月为 F/G/H/I/T/U/V，4 月起 E/F/G/H/S/T/U）。")
w("3. **美客多 2025年6~9月为拆仓过渡期**（口径与前后月不可直接纵向比）：")
w("   6 月起新增「转入FULL仓明细 / 海外仓明细 / FULL仓明细」；7 月为仓别口径版")
w("   （分析表表头=物流仓库存/FULL仓总库存/在途库存/调拨在途合计，库存 sheet 拆 4 个）；")
w("   8 月回归「库存明细」但残留海外仓调拨台账/物流仓明细，且多出空白 Sheet1；")
w("   9 月起完全回归标准结构（Sheet1 残留至 2026-02 前后）。")
w("4. **罗马尼亚（2025-12 起）无「在途」表头**：eMAG 无在途概念，分析表 G 列缺失；")
w("   库存明细 A1 为快照锚点，明细列为 PNK(c4)/产品代码(c5)/当前库存(c27)/stock_type(c28)。")
w("5. **多国分区（大簿）**：北美加拿大簿内分 美国/墨西哥/加拿大 三个库存区；")
w("   欧盟英国簿内分 英国/欧洲（+「多国库存」）等区——大簿的分析表货件列更宽，")
w("   差额/速度列位比小簿靠右（1-3月 AD/AE，4 月起 AC/AD）。")
w("6. **月总销量列**：主列 X（4 月起；1-3 月为 Y），年累计列在 BX 附近（随月份/书种浮动），")
w("   更早年累计再往右（如 2025-01 有 2024年总/2023年总）。")
w("7. **2026-05 起 Interflowd 三本独立成簿**（北美加拿大/澳大利亚，07 起加阿联酋）。")
w("")
w("## 三、分月台账（合计行数值 = 人工当月做表结果）")
w("")
w("列说明：SKU=分析表数据行数；预留/可售/在途/调拨/差额=分析表合计行对应列；")
w("月总=「N月总销量」合计；做表日=文件最后修改时间（人工收尾归档日）。单位：件。")
w("")
for m in months:
    w(f"### {m}")
    w("")
    w("| 书种 | 做表日 | SKU行 | 预留 | 可售 | 在途 | 调拨合计 | 差额 | 月总销量 | 年累计 | 日窗 |")
    w("|---|---|---|---|---|---|---|---|---|---|---|")
    for r in books[m]:
        an = r.get("analysis", {})
        tot = an.get("totals", {})
        mon_cols = an.get("month_total_cols", {})
        # 月总列 = 名字含「N月总销量」的；年累计 = 名字含「年总销量」的第一个
        mon_v = ""
        for name, col in mon_cols.items():
            if re.search(r"^\d{1,2}月总销量$", name):
                mon_v = tot.get(name, "")
                mon_col = col
        yr_v = ""
        for name, col in sorted(mon_cols.items()):
            if "年总销量" in name:
                yr_v = tot.get(name, "")
                break
        win = an.get("window", {})
        win_s = f"{win.get('first','')}~{win.get('last','')}" if win else ""
        def g(k):
            v = tot.get(k, "")
            if isinstance(v, float) and v.is_integer():
                return int(v)
            return v

        def gi(v):
            return int(v) if isinstance(v, float) and v.is_integer() else v
        w(f"| {species(r['file'])} | {r['mtime']} | {an.get('sku_rows','')} "
          f"| {g('预留库存')} | {g('可售库存')} | {g('在途')} | {g('调拨在途合计')} | {g('在途差额')} "
          f"| {gi(mon_v) if mon_v != '' else ''} | {gi(yr_v) if yr_v != '' else ''} | {win_s} |")
    w("")

w("## 四、各书种 sheet 结构（按月，差异才列）")
w("")
for sp in species_first:
    w(f"### {sp}")
    w("")
    w("| 月份 | sheet 名单（按工作簿内顺序） |")
    w("|---|---|")
    prev = None
    for m in months:
        rs = [r for r in books[m] if species(r["file"]) == sp]
        if not rs:
            continue
        r = rs[0]
        names = " / ".join(s["name"] for s in r["sheets"])
        mark = "" if prev is None or names == prev else " ←**变化**"
        w(f"| {m} | {names}{mark} |")
        prev = names
    w("")

w("## 五、逐本明细附录（锚点 / 订单行数 / 全 sheet 尺寸）")
w("")
w("锚点=库存快照类 sheet 的 A1（日期=快照日，数字=SKU 数）；订单行=订单/每日销售数据 sheet 的行数（含表头）。")
w("")
for m in months:
    w(f"### {m}")
    w("")
    for r in books[m]:
        sp = species(r["file"])
        anchors = r.get("stock_anchors", {})
        sales = r.get("sales_rows", {})
        a_s = "；".join(f"{k}={v}" for k, v in anchors.items()) if anchors else "-"
        s_s = "；".join(f"{k}={v}行" for k, v in sales.items()) if sales else "-"
        sheets = "；".join(f"{s['name']}({s['rows']}×{s['cols']})" for s in r["sheets"])
        w(f"- **{sp}**（{r['size_kb']}KB，做表 {r['mtime']}）")
        w(f"  - sheet：{sheets}")
        w(f"  - 锚点：{a_s}｜{s_s}")
    w("")

with open(OUT, "w", encoding="utf-8") as f:
    f.write("\n".join(L))
print(f"生成完成: {OUT}")
print(f"行数: {len(L)}")
