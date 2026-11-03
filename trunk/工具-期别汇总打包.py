# -*- coding: utf-8 -*-
"""
期别汇总打包（2026-11-02 新增；任务⑲，独立工具、只读 trunk/ 下已有产物）
=====================================================================

做什么
------
把**某一期**的成品（result）+ 源数据（data）+ 核对凭证收进一个交付文件夹，
结构对齐人工归档习惯（`trunk/result/result-{期}期/`），并自动生成：

  · **数据速览-{期}.md** —— 当期 vs 上期各列组数字（直接从月销总表累加，不写死）
  · **README-{期}期汇总说明.md** —— **只在不存在时**生成骨架，已有则一字不动
    （口径变化、异常说明这类叙述性内容机器写不出来，留给人/助手补）
  · **文件清单.csv** —— 全包相对路径 / 字节 / MD5（清单本身不登记）

用法（在项目根 `sale_file_gen/` 下）
------------------------------------
    set PYTHONPATH=C:\\Users\\dell\\AppData\\Roaming\\Python\\Python314\\site-packages
    C:\\Python314\\python.exe trunk\\工具-期别汇总打包.py [归属月=202610] [快照日=20261102] [--no-src]

    --no-src   不收源数据（只打成品，包小很多；默认收本期实际读取的那一个快照）

幂等：重复跑只会补齐新增文件并刷新清单/速览，不删任何东西；同名同 MD5 的文件跳过不重拷。

⚠️ 三条约定（都是踩过的）
------------------------
  1. **只收「本期实际读取」的那一个快照**，不收整月 14 个历史快照（202610 全月 337 MB）。
  2. **库龄总表只有当期一份**（`data/库龄报表/` 是滚动文件，没有按期分目录），别按上期路径找。
  3. `~$*.xlsx` 这类 Office 锁文件一律不收（收进去就是垃圾）。
"""
import hashlib
import io
import os
import re
import shutil
import sys

import openpyxl
from openpyxl.utils import get_column_letter as gl

HERE = os.path.dirname(os.path.abspath(__file__))          # = trunk/
RESULT = os.path.join(HERE, "result")
DATA = os.path.join(HERE, "data")
OUT_ROOT = RESULT


def n0(v):
    if v is None:
        return 0
    if isinstance(v, (int, float)):
        return v
    s = str(v).strip()
    if s in ("", "00", "-"):
        return 0
    try:
        return float(s)
    except ValueError:
        return 0


def md5_of(path):
    h = hashlib.md5()
    with open(path, "rb") as fh:
        for chunk in iter(lambda: fh.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def copy_idem(src, dst):
    """同 MD5 跳过（幂等）；返回 (是否新拷, 字节)"""
    os.makedirs(os.path.dirname(dst), exist_ok=True)
    if os.path.isfile(dst) and os.path.getsize(dst) == os.path.getsize(src) \
            and md5_of(src) == md5_of(dst):
        return False, os.path.getsize(dst)
    shutil.copy2(src, dst)
    return True, os.path.getsize(dst)


def collect(dst, period, snapshot, with_src=True):
    """把成品/源数据/中间产物/核对收进 dst；返回 (新拷数, 跳过数)"""
    new = skip = 0
    # 1) 成品：本期 + 上期（上期只为对比）
    prev = f"{int(period[:4]) - (1 if period[4:] == '01' else 0):04d}" + \
           (f"{int(period[4:]) - 1:02d}" if period[4:] != "01" else "12")
    for tag in ("数据归属周期", "东南亚销量表", "欧洲销量表", "北美销量表"):
        for per in (prev, period):
            d = os.path.join(RESULT, tag, per)
            if not os.path.isdir(d):
                continue
            for f in sorted(os.listdir(d)):
                if not f.lower().endswith(".xlsx") or f.startswith("~$"):
                    continue
                n, _ = copy_idem(os.path.join(d, f), os.path.join(dst, tag, per, f))
                new += n
                skip += (not n)
    # 2) 库龄（滚动文件，只有当期一份）
    age = os.path.join(DATA, "库龄报表", "亚马逊EMAG库龄情况总表.xlsx")
    if os.path.isfile(age):
        n, _ = copy_idem(age, os.path.join(dst, "库龄", os.path.basename(age)))
        new += n
        skip += (not n)
    # 3) 源数据：本期实际读取的那一个快照 + 库龄 + 货件
    if with_src:
        base = os.path.join(DATA, "数据归属周期", period)
        for rel in (os.path.join("快照日期", snapshot), os.path.join("库龄", snapshot), "货件"):
            s_dir = os.path.join(base, rel)
            if not os.path.isdir(s_dir):
                print(f"  ⚠️ 源数据缺 {rel}（跳过）")
                continue
            for root, _, fs in os.walk(s_dir):
                for f in fs:
                    if f.startswith("~$"):
                        continue
                    sp = os.path.join(root, f)
                    dp = os.path.join(dst, "源数据", period, rel, os.path.relpath(sp, s_dir))
                    n, _ = copy_idem(sp, dp)
                    new += n
                    skip += (not n)
    # 4) 中间产物
    for f in ("se_a_sku_sale_df.csv", "total_sku_sale_df.csv",
              "processed_month_sale_df.csv", "total_last30d_se_a_order_df.csv"):
        sp = os.path.join(HERE, "sale_file_tmp", f)
        if os.path.isfile(sp):
            n, _ = copy_idem(sp, os.path.join(dst, "中间产物", f))
            new += n
            skip += (not n)
    # 5) 核对凭证（工具-跨境订单状态核对.py 的产物）
    for per in (prev, period):
        p = os.path.join(RESULT, "核对", f"跨境订单状态核对-{per}.xlsx")
        if os.path.isfile(p):
            n, _ = copy_idem(p, os.path.join(dst, "核对", os.path.basename(p)))
            new += n
            skip += (not n)
    return new, skip


def col_sums(path, month):
    """按【2026年】sheet 行1 的分组标签，累加第 month 月列（数据行 4 起）"""
    ws = openpyxl.load_workbook(path, data_only=True)["2026年"]
    groups = [(str(c.value).strip(), c.column) for c in ws[1]
              if c.value and str(c.value).strip() not in ("", "2026年")]
    out = {}
    for name, c0 in groups:
        col = c0 + month - 1
        if col > ws.max_column:
            continue
        out[name] = int(round(sum(n0(ws.cell(r, col).value) for r in range(4, ws.max_row + 1))))
    rows = sum(1 for r in range(4, ws.max_row + 1) if ws.cell(r, 297).value not in (None, ""))
    return out, rows


def write_brief(dst, period, snapshot):
    """生成 数据速览-{期}.md（纯数字，全部现算）"""
    cur_p = os.path.join(RESULT, "数据归属周期", period, "月销售总表.xlsx")
    if not os.path.isfile(cur_p):
        print("  ⚠️ 没有本期月销总表，跳过数据速览")
        return None
    month = int(period[4:])
    prev_month_of_year = month - 1
    cur, rows = col_sums(cur_p, month)
    same_year_prev = col_sums(cur_p, prev_month_of_year)[0] if month > 1 else {}
    lines = [f"# 数据速览 · {period}（快照 {snapshot}）", "",
             "> 本文件由 `trunk/工具-期别汇总打包.py` 现算生成，改完成品重跑即可刷新。"
             "列组顺序取自月销总表行1。", "",
             f"数据行 **{rows}** 个 SKU；当月为 **{month} 月**"
             + (f"，上期列取同表 {prev_month_of_year} 月" if same_year_prev else "（1 月无同年上期）"), "",
             "| 列组 | 当月 | 同年上月 | 增减 |", "|---|---:|---:|---:|"]
    for name, v in cur.items():
        pv = same_year_prev.get(name)
        if pv is None:
            lines.append(f"| {name} | {v:,} | — | — |")
        else:
            lines.append(f"| {name} | {v:,} | {pv:,} | {v - pv:+,} |")
    # 库龄
    age = os.path.join(DATA, "库龄报表", "亚马逊EMAG库龄情况总表.xlsx")
    if os.path.isfile(age):
        wb = openpyxl.load_workbook(age, data_only=True)
        if "本期明细" in wb.sheetnames:
            ws = wb["本期明细"]
            # ⚠️ 表头在行2；**末列不是件数**（末列是「180+占比」）。件数列的名字是「合计」，
            #    180 天以上是「180天以上」—— 2026-11-02 第一版按 len(hdr) 取末列，
            #    把占比列当件数累加，算出「1,536 件」这种一眼假的值。按名字取，别按位置。
            hdr = [c.value for c in ws[2]]
            col = {str(h).strip(): i + 1 for i, h in enumerate(hdr) if h}
            site_c = col.get("站点", 2)
            qty_c = col.get("合计") or col.get("总库存") or col.get("库存数量")
            old_c = col.get("180天以上")
            if qty_c is None:
                print("  ⚠️ 库龄明细里找不到件数列（合计），跳过库龄数字")
            else:
                skus, qty, old = 0, 0, 0
                for r in range(3, ws.max_row + 1):
                    if ws.cell(r, site_c).value:
                        skus += 1
                        qty += n0(ws.cell(r, qty_c).value)
                        if old_c:
                            old += n0(ws.cell(r, old_c).value)
                txt = f"**{skus:,} SKU / {int(round(qty)):,} 件"
                if old_c:
                    txt += f" / 180 天以上 {int(round(old)):,} 件（{old / qty * 100:.1f}%）" if qty else ""
                lines += ["", f"库龄（`data/库龄报表/亚马逊EMAG库龄情况总表.xlsx`）：{txt}**"]
    p = os.path.join(dst, f"数据速览-{period}.md")
    io.open(p, "w", encoding="utf-8").write("\n".join(lines) + "\n")
    return p


def write_readme_if_absent(dst, period, snapshot):
    p = os.path.join(dst, f"README-{period}期汇总说明.md")
    if os.path.isfile(p):
        return False
    body = f"""# {period} 期 · 汇总交付包

> 归属周期 **{period}**｜快照日 **{snapshot}**｜由 `trunk/工具-期别汇总打包.py` 生成

## 目录

| 目录 | 内容 |
|---|---|
| `数据归属周期/{period}/月销售总表.xlsx` | 月销总表（主成品） |
| `东南亚销量表/{period}/` | 东南亚总表 + 菲/印/泰三国次表 |
| `欧洲销量表/{period}/` | 欧洲总表 + 10 国次表 |
| `北美销量表/{period}/` | 北美总表 + 美/加/墨次表 |
| `库龄/` | 库龄总表（4 sheet：本期明细/站点汇总/历史汇总/取数清单） |
| `核对/` | 跨境订单状态核对（订单号 + 状态逐单） |
| `源数据/{period}/` | 本期实际读取的快照、库龄、货件 |
| `中间产物/` | 主脚本往下游传的 CSV |
| `数据速览-{period}.md` | 各列组当月/上月数字（现算） |
| `文件清单.csv` | 全包 MD5 |

## 待补（人工/助手填写）

- [ ] 本期口径变化与原因
- [ ] 异常与已知事项
- [ ] 校验结论（恒等链 / 验收工具 / 逐单核对）
"""
    io.open(p, "w", encoding="utf-8").write(body)
    return True


def write_manifest(dst):
    rows = []
    for root, _, fs in os.walk(dst):
        for f in fs:
            p = os.path.join(root, f)
            rel = os.path.relpath(p, dst).replace("\\", "/")
            if rel == "文件清单.csv":
                continue
            rows.append((rel, os.path.getsize(p), md5_of(p)))
    rows.sort()
    with io.open(os.path.join(dst, "文件清单.csv"), "w", encoding="utf-8-sig", newline="") as fh:
        fh.write("相对路径,字节,MD5\n")
        for r, s, h in rows:
            fh.write(f"{r},{s},{h}\n")
    return len(rows), sum(r[1] for r in rows)


def main(period="202610", snapshot="20261102", with_src=True):
    if not re.fullmatch(r"\d{6}", period) or not re.fullmatch(r"\d{8}", snapshot):
        print("❌ 参数格式：归属月 YYYYMM、快照日 YYYYMMDD")
        return 2
    dst = os.path.join(OUT_ROOT, f"result-{period}期")
    print("=" * 92)
    print(f"期别汇总打包 | 归属月={period} 快照={snapshot} 含源数据={'是' if with_src else '否'}")
    print(f"目标：{dst}")
    print("=" * 92)
    os.makedirs(dst, exist_ok=True)
    # 先写文字件，再登记清单（否则清单漏掉它们）
    made = write_readme_if_absent(dst, period, snapshot)
    if made:
        print("  + README 骨架已生成（已有则不覆盖）")
    b = write_brief(dst, period, snapshot)
    if b:
        print(f"  + {os.path.basename(b)}")
    new, skip = collect(dst, period, snapshot, with_src)
    print(f"  成品/源数据：新拷 {new} 个，MD5 相同跳过 {skip} 个")
    n, size = write_manifest(dst)
    print(f"  文件清单：登记 {n} 个 / {size / 1048576:.1f} MB")
    print(f"\n✅ 打包完成 → {dst}")
    return 0


if __name__ == "__main__":
    os.chdir(HERE)
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    sys.exit(main(args[0] if len(args) > 0 else "202610",
                  args[1] if len(args) > 1 else "20261102",
                  with_src="--no-src" not in sys.argv))
