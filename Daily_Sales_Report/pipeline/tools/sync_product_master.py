# -*- coding: utf-8 -*-
"""sync_product_master.py —— 商品资料 ↔ 日销表分析表 属性列对照 / 回填

用法：
    python pipeline/tools/sync_product_master.py                 # 只对照，打印差异报告
    python pipeline/tools/sync_product_master.py --out r.txt     # 报告落文件
    python pipeline/tools/sync_product_master.py --apply --dry   # 演练：打印将写入的内容，不落盘
    python pipeline/tools/sync_product_master.py --apply         # 实际回填（自动备份 .bak_pm）
    python pipeline/tools/sync_product_master.py --book 日本      # 只处理某一本

对照的四个属性列（各簿列位不同，按行5/行6 标签自动探测，不写死列号）：
    C 列「商品分类」(行6 标签)          ← 商品资料「X月销量得分」
    膜种类列（膜的种类/膜的类型/种类）   ← 商品资料「类型」
    品牌列                              ← 商品资料「品牌」
    判 列                               ← 只看不改（常规/新品/保底，运营人工标记）

判定分级（关键：不无脑覆盖）：
    A 空 → 资料有值          可自动补    --apply 写入
    B 有值 且 两侧语义等价    视为一致    （如 索尼/Sony、iPhone/苹果）
    C 有值 且 两侧实打实不同  列为差异    **默认不动**，报告里给人工裁决
    加 --overwrite-score / --overwrite-membrane / --overwrite-brand 才强制覆盖

⚠️ 已知口径差（别当成 bug）：
    · 商品资料的评级列叫「10月销量得分」= 单月口径；分析表的商品分类是历史累计口径
      → 大量 A→B、B→C 属于「月度重评」，不是填错。默认不覆盖就是为此。
    · 资料的「状态」若为停售/不销售，但分析表里还有该 SKU → 列 [停] 风险项（只报不改）。
"""
import os
import sys
import glob
import shutil
import argparse
from collections import defaultdict

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "core"))

from openpyxl import load_workbook

import product_master as PM

HERE = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
DSR_DIR = os.path.join(HERE, "1.每日销售数据表")

DATA_FIRST_ROW = 7
MEMBRANE_LABELS = ("膜的种类", "膜的类型", "种类", "膜类型")
BRAND_LABELS = ("品牌",)
ASIN_LABELS = ("ASIN",)
JUDGE_LABELS = ("判",)

# 品牌中英同义（只用于「是否同一品牌」的判定，不拿来改写单元格）
BRAND_ALIAS = {
    "sony": "索尼", "索尼": "索尼",
    "apple": "苹果", "iphone": "苹果", "苹果": "苹果",
    "samsung": "三星", "三星": "三星",
    "xiaomi": "小米", "redmi": "小米", "小米": "小米", "红米": "小米",
    "honor": "荣耀", "荣耀": "荣耀",
    "huawei": "华为", "华为": "华为",
    "oneplus": "一加", "一加": "一加",
    "oppo": "OPPO", "vivo": "vivo", "realme": "Realme", "poco": "Poco",
    "google": "Google", "motorola": "Motorola", "nokia": "Nokia",
    "nothing": "Nothing", "zte": "努比亚", "nubia": "努比亚", "努比亚": "努比亚",
    "asus": "华硕", "华硕": "华硕", "blackshark": "Blackshark", "infinix": "Infinix",
    "tecno": "Tecno", "fcnt": "FCNT",
}


def _norm(v):
    if v is None:
        return ""
    s = str(v).strip()
    if s.endswith(".0") and s[:-2].isdigit():
        s = s[:-2]
    return s


def norm_score(v):
    """销量得分归一：'b' → 'B'；数字 0 视为空"""
    s = _norm(v).upper()
    return "" if s in ("0", "NONE", "NAN", "-") else s


def score_of(a):
    """由商品资料记录推出「商品分类」应写什么。

    ⚠️ 关键发现（2026-11-07 实测）：资料里 207 个「评级为空」的 SKU，
    **状态全部是 停售/不销售**。也就是说「评级空」的真实含义是「已停售、不再评级」，
    不是「没数据」。日销表侧对应的档叫 `待停售`（资料里没有这一档，是表侧标记）。
    → 直接把空值清成空白会丢掉 158 条有效的停售标记，必须写成 `待停售`。
    """
    sc = norm_score(a.get("score"))
    if sc:
        return sc
    if _norm(a.get("status")) in ("停售", "不销售"):
        return "待停售"
    return ""


def canon_brand(v):
    s = _norm(v)
    if not s:
        return ""
    return BRAND_ALIAS.get(s.lower(), s)


def same_type(a, b):
    x, y = _norm(a), _norm(b)
    if x == y:
        return True
    if not x or not y:
        return False
    return x in y or y in x


def scan_sheet_header(ws):
    r5 = r6 = None
    n = 0
    for row in ws.iter_rows(min_row=1, max_row=8, max_col=200, values_only=True):
        n += 1
        if n == 5:
            r5 = row
        elif n == 6:
            r6 = row
        if n >= 8:
            break
    return r5 or (), r6 or ()


def locate_cols(hdr5, hdr6):
    out = {"score": None, "membrane": None, "brand": None, "asin": None, "judge": None}
    for j, v in enumerate(hdr6):
        if isinstance(v, str) and "商品分类" in v.strip():
            out["score"] = j + 1
            break
    if out["score"] is None:
        out["score"] = 3
    for j, v in enumerate(hdr5):
        if not isinstance(v, str):
            continue
        s = v.strip()
        if out["membrane"] is None and s in MEMBRANE_LABELS:
            out["membrane"] = j + 1
        elif out["brand"] is None and s in BRAND_LABELS:
            out["brand"] = j + 1
        elif out["asin"] is None and s in ASIN_LABELS:
            out["asin"] = j + 1
        elif out["judge"] is None and s in JUDGE_LABELS:
            out["judge"] = j + 1
    return out


def iter_rows_of(ws, cols):
    maxc = max([c for c in cols.values() if c] + [1])
    r = DATA_FIRST_ROW
    for row in ws.iter_rows(min_row=DATA_FIRST_ROW, max_row=ws.max_row,
                            min_col=1, max_col=maxc, values_only=True):
        sku = row[0] if row else None
        if isinstance(sku, str) and sku.strip():
            def g(k):
                c = cols.get(k)
                if not c:
                    return None
                i = c - 1
                return row[i] if i < len(row) else None
            yield r, sku.strip(), g("score"), g("membrane"), g("brand"), g("asin"), g("judge")
        r += 1


def analyse_book(path, master, want):
    """只读扫描一本 → (stat, plan, diffs)"""
    stat = defaultdict(int)
    plan = []
    diffs = []
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        for sn in wb.sheetnames:
            if "分析表" not in sn:
                continue
            ws = wb[sn]
            h5, h6 = scan_sheet_header(ws)
            cols = locate_cols(h5, h6)
            for r, sku, sc, mem, brd, asin, jud in iter_rows_of(ws, cols):
                stat["rows"] += 1
                a = PM.attrs(sku, master)
                if not a:
                    stat["miss"] += 1
                    diffs.append(f"  [!] {sn} r{r} {sku:18s} 商品资料未收录")
                    continue
                mt = (a.get("_matched") or "").strip()
                tag = f"(资料→{mt})" if mt and mt != sku else ""

                # ---------- 商品分类 ← 销量得分（停售无评级 → 待停售）
                if cols["score"]:
                    new = score_of(a)
                    old = norm_score(sc)
                    if not new:
                        # 既无评级又非停售：资料确实没结论，默认不动（--blank-clear 才清）
                        if old:
                            stat["score_nodata"] += 1
                            if want["blank"]:
                                diffs.append(f"  [空评] {sn} r{r} {sku:18s}{tag:10s} 表={old:6s} 资料=(空)")
                                plan.append((sn, r, cols["score"], sku, "分类(清空)", old, None))
                    elif not old:
                        stat["score_fill"] += 1
                        plan.append((sn, r, cols["score"], sku, "分类", old, new))
                    elif old == new:
                        stat["score_same"] += 1
                    else:
                        stat["score_diff"] += 1
                        diffs.append(f"  [≠类] {sn} r{r} {sku:18s}{tag:10s} 表={old:6s} 资料={new}")
                        if want["score"]:
                            plan.append((sn, r, cols["score"], sku, "分类", old, new))

                # ---------- 膜种类 ← 类型
                if cols["membrane"]:
                    new = _norm(a.get("type"))
                    old = _norm(mem)
                    if not new:
                        if old and want["blank"]:
                            diffs.append(f"  [空膜] {sn} r{r} {sku:18s}{tag:10s} 表={old:8s} 资料=(空)")
                            plan.append((sn, r, cols["membrane"], sku, "膜种类(清空)", old, None))
                    elif not old:
                        stat["mem_fill"] += 1
                        plan.append((sn, r, cols["membrane"], sku, "膜种类", old, new))
                    elif same_type(old, new):
                        stat["mem_same"] += 1
                        # 以资料为准：别名也要归一到资料的写法（钢化膜(2.5D) → 钢化膜）
                        if old != new and want["membrane"]:
                            stat["mem_alias"] += 1
                            plan.append((sn, r, cols["membrane"], sku, "膜种类(归一)", old, new))
                    else:
                        stat["mem_diff"] += 1
                        diffs.append(f"  [≠膜] {sn} r{r} {sku:18s}{tag:10s} 表={old:8s} 资料={new}")
                        if want["membrane"]:
                            plan.append((sn, r, cols["membrane"], sku, "膜种类", old, new))

                # ---------- 品牌 ← 品牌
                if cols["brand"]:
                    new = _norm(a.get("brand"))
                    old = _norm(brd)
                    if not new:
                        if old and want["blank"]:
                            diffs.append(f"  [空牌] {sn} r{r} {sku:18s}{tag:10s} 表={old:10s} 资料=(空)")
                            plan.append((sn, r, cols["brand"], sku, "品牌(清空)", old, None))
                    elif not old:
                        stat["brand_fill"] += 1
                        plan.append((sn, r, cols["brand"], sku, "品牌", old, new))
                    elif canon_brand(old) == canon_brand(new):
                        stat["brand_same"] += 1
                        if old != new:
                            stat["brand_alias"] += 1
                            # 以资料为准：译名也统一成资料写法（Sony 而不是 索尼）
                            if want["brand_alias"]:
                                plan.append((sn, r, cols["brand"], sku, "品牌(译名归一)", old, new))
                    else:
                        stat["brand_diff"] += 1
                        diffs.append(f"  [≠牌] {sn} r{r} {sku:18s}{tag:10s} 表={old:10s} 资料={new}")
                        if want["brand"]:
                            plan.append((sn, r, cols["brand"], sku, "品牌", old, new))

                # ---------- 停售风险（只报不改）
                st = _norm(a.get("status"))
                if st in ("停售", "不销售"):
                    stat["stop"] += 1
                    diffs.append(f"  [停]  {sn} r{r} {sku:18s}{tag:10s} 资料状态={st}")
    finally:
        wb.close()
    return dict(stat), plan, diffs


def apply_book(path, plan, bak_suffix=None):
    if not plan:
        return 0, {}
    # 备份名带时间戳：同一本可能被回填多轮，.bak_pm 只留第一份会把中间态覆盖掉
    tag = bak_suffix or ("bak_pm_" + __import__("time").strftime("%m%d_%H%M"))
    bak = f"{path}.{tag}"
    if not os.path.exists(bak):
        shutil.copy2(path, bak)
    wb = load_workbook(path)
    try:
        per = defaultdict(int)
        for sn, r, c, sku, field, old, new in plan:
            wb[sn].cell(r, c).value = new
            per[sn] += 1
        wb.save(path)
    finally:
        wb.close()
    return len(plan), dict(per)


def hall_clone_brand(path, master, thresh=4):
    """识别「补行克隆残留」的品牌批：连续 ≥thresh 行的品牌值完全相同，
    而商品资料给这些 SKU 的品牌却不全同（说明这串不是真品牌，是模板行抄下来的）。

    → {row: (sheet, old_brand, correct_brand, sku)}
    这是 2026-11-07 在美国表 r394~446（全 iPhone）、罗马尼亚 r301~313（全 Realme）、
    英国 r466~472（全 一加）发现的真实错误批：补行克隆模板行时品牌没被覆盖。
    """
    found = {}
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        for sn in wb.sheetnames:
            if "分析表" not in sn:
                continue
            ws = wb[sn]
            h5, h6 = scan_sheet_header(ws)
            cols = locate_cols(h5, h6)
            if not cols["brand"]:
                continue
            seq = []          # [(row, sku, brand, correct)]
            for r, sku, sc, mem, brd, asin, jud in iter_rows_of(ws, cols):
                a = PM.attrs(sku, master) or {}
                correct = _norm(a.get("brand"))
                seq.append((r, sku, _norm(brd), correct))
            i, n = 0, len(seq)
            while i < n:
                j = i
                while j + 1 < n and seq[j + 1][2] == seq[i][2] and seq[j + 1][2]:
                    j += 1
                run = seq[i:j + 1]
                if len(run) >= thresh:
                    corr = {c for _, _, _, c in run if c}
                    olds = {o for _, _, o, _ in run}
                    if len(corr) > 1:      # 资料上本来就是多个品牌 → 这串是残留
                        for r, sku, old, correct in run:
                            if correct and canon_brand(correct) != canon_brand(old):
                                found[r] = (sn, old, correct, sku)
                i = j + 1
    finally:
        wb.close()
    return found


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--apply", action="store_true")
    ap.add_argument("--dry", action="store_true")
    ap.add_argument("--book", default=None)
    ap.add_argument("--overwrite-score", action="store_true")
    ap.add_argument("--overwrite-membrane", action="store_true")
    ap.add_argument("--overwrite-brand", action="store_true",
                    help="品牌不同也覆盖（默认：只覆盖「克隆残留」批，见 HALL_*）")
    ap.add_argument("--no-brand-hall", action="store_true",
                    help="连克隆残留品牌也不覆盖（纯只读）")
    ap.add_argument("--out", default=None)
    ap.add_argument("--show", type=int, default=0, help="每本最多打印多少条明细（0=不限）")
    ap.add_argument("--xlsx", default=None, help="差异清单导出为 xlsx")
    ap.add_argument("--force-master", action="store_true",
                    help="【以商品资料为准】一刀切：分类/膜/品牌全部按资料改写，"
                         "译名归一，资料为空也清空。等价于同时打开下列 5 个开关。")
    ap.add_argument("--blank-clear", action="store_true",
                    help="资料为空时把表里已有值清空（配合 --force-master 自动开启）")
    args = ap.parse_args()

    if args.force_master:
        args.overwrite_score = True
        args.overwrite_membrane = True
        args.overwrite_brand = True
        args.blank_clear = True

    want = {"score": args.overwrite_score,
            "membrane": args.overwrite_membrane,
            "brand": args.overwrite_brand,
            "blank": args.blank_clear,
            "brand_alias": args.force_master}

    master = PM.load()
    # 只认带「库存销售表」的簿：目录里还放着对照报告 xlsx，别把它当数据表扫
    books = [p for p in sorted(glob.glob(os.path.join(DSR_DIR, "*库存销售表.xlsx")))
             if not os.path.basename(p).startswith("~$")]
    if args.book:
        books = [p for p in books if args.book in os.path.basename(p)]

    out = []
    out.append("=" * 100)
    out.append("商品资料 ↔ 日销表分析表  属性对照报告")
    out.append("=" * 100)
    out.append(f"资料文件：{PM.find_path()}   （SKU {len(master)} 个）")
    out.append("")
    out.append(f"{'簿':<20s}{'行':>5s}{'未收录':>6s}{'分类同':>6s}{'分类补':>6s}{'分类≠':>6s}"
               f"{'无评级':>6s}{'膜同':>5s}{'膜补':>5s}{'膜≠':>5s}"
               f"{'牌同':>5s}{'牌补':>5s}{'牌≠':>5s}{'停售':>5s}")
    out.append("-" * 100)

    T = defaultdict(int)
    details = []
    total_plan = 0
    xlsx_rows = []          # 导出 xlsx 用：每条差异一行
    score_pairs = defaultdict(int)
    for p in books:
        bn = os.path.basename(p).replace("库存销售表.xlsx", "")
        st, plan, diffs = analyse_book(p, master, want)
        # ---- 补行克隆残留品牌：认出来就直接纳入 plan（这是错值，不是口径差）
        hall = {} if args.no_brand_hall else hall_clone_brand(p, master)
        if hall:
            have = {(x[0], x[1], x[2]) for x in plan}
            cols_by_sheet = {}
            wbt = load_workbook(p, read_only=True, data_only=True)
            for sn in sorted(set(v[0] for v in hall.values())):
                h5, h6 = scan_sheet_header(wbt[sn])
                cols_by_sheet[sn] = locate_cols(h5, h6)
            wbt.close()
            for r, (sn, old, correct, sku) in hall.items():
                c = cols_by_sheet[sn]["brand"]
                if (sn, r, c) in have:
                    continue
                st["brand_hall"] = st.get("brand_hall", 0) + 1
                plan.append((sn, r, c, sku, "品牌(克隆残留)", old, correct))
                diffs.append(f"  [残牌] {sn} r{r} {sku:18s} 表={old:10s} 资料={correct}"
                             f"   ← 补行克隆模板行留下的错值")
        for k, v in st.items():
            T[k] += v
        total_plan += len(plan)
        out.append(f"{bn:<20s}{st.get('rows',0):5d}{st.get('miss',0):6d}"
                   f"{st.get('score_same',0):6d}{st.get('score_fill',0):6d}{st.get('score_diff',0):6d}"
                   f"{st.get('score_nodata',0):6d}{st.get('mem_same',0):5d}{st.get('mem_fill',0):5d}"
                   f"{st.get('mem_diff',0):5d}{st.get('brand_same',0):5d}{st.get('brand_fill',0):5d}"
                   f"{st.get('brand_diff',0):5d}{st.get('stop',0):5d}")
        # 差异清单（供 xlsx 导出）
        for ln in diffs:
            s = ln.strip()
            if s.startswith("[≠类]"):
                kind = "分类不同"
            elif s.startswith("[≠膜]"):
                kind = "膜种类不同"
            elif s.startswith("[≠牌]"):
                kind = "品牌不同"
            elif s.startswith("[残牌]"):
                kind = "品牌克隆残留(可修)"
            elif s.startswith("[停]"):
                kind = "资料已停售"
            elif s.startswith("[空评]"):
                kind = "资料无评级(清空)"
            elif s.startswith("[空膜]"):
                kind = "资料无类型(清空)"
            elif s.startswith("[空牌]"):
                kind = "资料无品牌(清空)"
            else:
                kind = "资料未收录"
            xlsx_rows.append((bn, kind, s))
        # 分类档位漂移方向统计
        for ln in diffs:
            if "[≠类]" in ln:
                try:
                    seg = ln.split("表=")[1]
                    a_, b_ = seg.split("资料=")
                    score_pairs[(a_.strip(), b_.strip())] += 1
                except Exception:
                    pass
        if diffs:
            details.append((bn, len(plan), diffs))
        if args.apply and args.dry and plan:
            print(f"[演练] {bn}: 将写 {len(plan)} 格")
            for sn, r, c, sku, field, old, new in plan[:8]:
                print(f"   {sn} r{r} {sku} {field}: '{old}' → '{new}'")
        elif args.apply and not args.dry:
            n, per = apply_book(p, plan)
            if n:
                print(f"[写入] {bn}: {n} 格   备份 {os.path.basename(p)}.bak_pm")
                for k, v in per.items():
                    print(f"        {k}: {v}")

    out.append("-" * 100)
    out.append(f"{'合计':<20s}{T.get('rows',0):5d}{T.get('miss',0):6d}"
               f"{T.get('score_same',0):6d}{T.get('score_fill',0):6d}{T.get('score_diff',0):6d}"
               f"{T.get('score_nodata',0):6d}{T.get('mem_same',0):5d}{T.get('mem_fill',0):5d}"
               f"{T.get('mem_diff',0):5d}{T.get('brand_same',0):5d}{T.get('brand_fill',0):5d}"
               f"{T.get('brand_diff',0):5d}{T.get('stop',0):5d}")
    out.append("")
    out.append(f"可自动补格数：{total_plan}  （分类 {T.get('score_fill',0)} / "
               f"膜种类 {T.get('mem_fill',0)} / 品牌 {T.get('brand_fill',0)}）")
    out.append(f"品牌译名不同但同一品牌、已判为一致的：{T.get('brand_alias',0)} 格")
    out.append("")
    out.append("差异明细（[!] 未收录 / [≠类] 分类不同 / [≠膜] 膜种类不同 / [≠牌] 品牌不同 / [停] 资料停售）：")
    for bn, nplan, diffs in details:
        out.append("")
        out.append(f"── {bn}   （可自动补 {nplan} 格，差异 {len(diffs)} 条）")
        show = diffs if not args.show else diffs[:args.show]
        out.extend(show)
        if args.show and len(diffs) > args.show:
            out.append(f"   …… 另 {len(diffs)-args.show} 条未显示")

    txt = "\n".join(out)

    # ---- 分类档位漂移方向（只统计，不改任何东西）
    if score_pairs:
        print()
        print("分类档位漂移 TOP20（表 → 资料）：")
        tot_diff = sum(score_pairs.values())
        down = up = 0
        ORDER = ["A", "B", "C", "D", "E"]
        for (a_, b_), n in sorted(score_pairs.items(), key=lambda x: -x[1]):
            tag = ""
            if a_ in ORDER and b_ in ORDER:
                if ORDER.index(b_) > ORDER.index(a_):
                    tag = "↓降档"
                    down += n
                else:
                    tag = "↑升档"
                    up += n
            elif a_.startswith("X") and not b_.startswith("X"):
                tag = "转常规"
            elif a_ == "待停售":
                tag = "待停售→他档"
            print(f"  {a_:>6s} → {b_:<6s} {n:5d}  {tag}")
        print(f"  合计 {tot_diff} 条：降档 {down} / 升档 {up}")
        print("  ⇒ 若降档占绝对多数，说明是资料按单月重评、表的分类是历史累计口径，属口径差不是填错。")

    # ---- xlsx 导出
    if args.xlsx and xlsx_rows:
        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill, Alignment
        wb = Workbook()
        ws = wb.active
        ws.title = "差异清单"
        ws.append(["簿", "类型", "差异明细"])
        for c in ("A1", "B1", "C1"):
            ws[c].font = Font(bold=True, color="FFFFFF")
            ws[c].fill = PatternFill("solid", fgColor="4F6BED")
        COLOR = {"资料未收录": "FFF2CC", "资料已停售": "F8CBAD",
                 "品牌克隆残留(可修)": "D9F2D9", "分类不同": "FCE4E4",
                 "资料无评级(清空)": "EDEDED", "资料无类型(清空)": "EDEDED",
                 "资料无品牌(清空)": "EDEDED"}
        for bn, kind, s in xlsx_rows:
            ws.append([bn, kind, s])
            if kind in COLOR:
                fill = PatternFill("solid", fgColor=COLOR[kind])
                for j in (1, 2, 3):
                    ws.cell(ws.max_row, j).fill = fill
        ws.column_dimensions["A"].width = 20
        ws.column_dimensions["B"].width = 18
        ws.column_dimensions["C"].width = 100
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = f"A1:C{ws.max_row}"
        wb.save(args.xlsx)
        print(f"\n差异清单（{len(xlsx_rows)} 行）：{args.xlsx}")

    if args.out:
        with open(args.out, "w", encoding="utf-8") as f:
            f.write(txt)
        print(txt.split("差异明细")[0])
        print(f"\n完整报告：{args.out}")
    else:
        print(txt)


if __name__ == "__main__":
    main()
