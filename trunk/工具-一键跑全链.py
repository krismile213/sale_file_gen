# -*- coding: utf-8 -*-
"""
工具-一键跑全链.py（任务㉕ / 2026-11-10）
====================================================================
一条命令跑完「月销表 → 东南亚 → 欧洲 → 北美」，并在**每一步之后做断言**，
把过去靠人工翻日志、开 Excel 对数才能发现的问题变成"跑完就红"。

为什么要它
    · 固定的 4 脚本顺序 + 欧盟英国合并，全靠人记；漏一步就是静默错数
      （2026-11-10 那次：interflow 北美报告导成 11 月 → 少 809 件，日志里只写了"剔除上月尾单"）。
    · 恒等链（月销表 == 东南亚表 == 欧洲表 == 北美表）过去要手工开 4 个簿对；
      这里全部自动化，0 差才算过。
    · 定案数字落盘到 `定案数字.json`，下一期准备模板时可以反查"这份成品是不是最新定案版"。

用法
    python 工具-一键跑全链.py 202610 20261107                 # 完整跑（含合并欧盟英国报告）
    python 工具-一键跑全链.py 202610 20261107 --verify-only    # 不出数，只验现有成品
    python 工具-一键跑全链.py 202610 20261107 --from 3         # 从第 3 步（东南亚）开始跑
    python 工具-一键跑全链.py 202610 20261107 --skip-merge     # 跳过欧盟/英国报告合并
    python 工具-一键跑全链.py 202610 20261107 --allow-low-cover # 报告区间体检命中严重项也放行
    python 工具-一键跑全链.py 202610 20261107 --dry-run        # 只打印计划

步骤编号
    1  合并欧盟/英国亚马逊订单报告（幂等；只补「英国版有、欧盟版没有、归属本月」的行）
    2  month_sale_file.py           → 月销售总表 + sale_file_tmp 中间产物
    3  southeast_asia_sale_file.py  → 东南亚销量表（含国家订单明细 sheet）
    4  europe_sale_file.py          → 欧洲销量表（10 国 + 10 份英文交付件）
    5  north_america_sale_file.py   → 北美销量表（3 国 + 3 份英文交付件）
    6  交叉验收（恒等链 / 分组对撞 / 明细件数 / 定案数字落盘）

退出码
    0 = 全链路 + 全部断言通过      1 = 验收断言失败（有成品但不一致）
    2 = 某个生成脚本失败              3 = 前置条件不满足（缺模板/文件被占用）
====================================================================
"""
import argparse
import datetime
import json
import os
import re
import subprocess
import sys

TRUNK_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, TRUNK_DIR)

from mst_read import (group_month_totals, group_map_warnings, sheet_xml,   # noqa: E402
                      col_sum)
from openpyxl.utils import get_column_letter                               # noqa: E402

LOG_DIR = os.path.join(TRUNK_DIR, "sale_file_tmp")
FINAL_JSON = os.path.join(TRUNK_DIR, "定案数字.json")

# ---- 区域表结构（与三个生成脚本里的 BASE / COUNTRIES 保持一致） ----
# 东南亚各国 sheet 有两种版式（2026-11-10 实测）：
#   · 多渠道路径：每月块 = [渠道列…][月合计列]，块宽 = 渠道数 + 1
#   · 单列路径（马来/越南 Lazada）：每月就一列，**没有合计列**，块宽 = 1
# 元组 = (月销表分组, sheet 名, 块宽, 月值相对块首的偏移)
SEA_ROWS = [("菲律宾", "Philippines 菲律宾销量表", 6, 5),
            ("泰国", "Thailand 泰国销量表", 5, 4),
            ("印尼", "Indonesia 印尼销量表", 4, 3),
            ("新加坡", "Singapore 新加坡销量表", 3, 2),
            ("马来", "Malaysia Lazada 马来西亚销量表", 1, 0),
            ("越南", "Vietnam Lazada 越南销销量表", 1, 0),   # ★ 模板里就是这个「销销量」错字，别改
            ("虾皮境内", "Malaysia Shopee 虾皮跨境销量表", 9, 8)]
SEA_TOTAL_SHEET = "SEA Total 东南亚总销量表"
SEA_TOTAL_BASE = 15          # 第一个月块首列（1-based）；每块 = 合计 + 8 国 = 9 列
MST_ALIAS = {"虾皮境内": "虾皮"}      # 本表口径名 → 月销总表分组名

EU_ROWS = [("比利时", "Belgium 比利时销量表", 1), ("德国", "Germany 德国销量表", 1),
           ("西班牙", "Spain 西班牙销量表", 1), ("法国", "France 法国销量表", 1),
           ("意大利", "Italy 意大利销量表", 1), ("荷兰", "Netherlands 荷兰销量表", 1),
           ("波兰", "Poland 波兰销量表", 1), ("瑞典", "Sweden 瑞典销量表", 1),
           ("爱尔兰", "Ireland 爱尔兰销量表", 1), ("英国", "United Kingdom 英国销量表", 1)]
EU_TOTAL_SHEET = "EU Total 欧洲总销量表"
EU_TOTAL_BASE = 6

NA_ROWS = [("美国", "United States 美国销量表", 2),
           ("加拿大", "Canada 加拿大销量表", 1),
           ("墨西哥", "Mexico 墨西哥销量表", 2)]
NA_TOTAL_SHEET = "NA Total 北美总销量表"
NA_TOTAL_BASE = 6

# 日志里值得实时显示的片段（其余只进日志文件，避免刷屏）
SHOW_KEYS = ("【", "✅", "⚠️", "❌", "🚨", "★", "合计", "校验", "补充订单", "明细", "任务")


# ================================================================== 小工具
def _cell(x, ref):
    m = re.search(r'<c r="%s"[^>]*?(?:/>|><v>([^<]*)</v></c>)' % ref, x)
    try:
        return float(m.group(1)) if m and m.group(1) else 0.0
    except ValueError:
        return 0.0


def _pct(a, b):
    return f"{a - b:+.0f}" if abs(a - b) > 0.5 else "0"


def _sheetnames(path):
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True)
    try:
        return list(wb.sheetnames)
    finally:
        wb.close()


def _qty_col(path, sheet_name):
    """在明细 sheet 的第 2 行表头里找「数量」列的列字母（openpyxl 读，兼容共享/内联字符串）"""
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        row2 = next(wb[sheet_name].iter_rows(min_row=2, max_row=2, max_col=20, values_only=True))
    finally:
        wb.close()
    for i, v in enumerate(row2 or []):
        if v is not None and "数量" in str(v):
            return get_column_letter(i + 1)
    return None


def _is_locked(path):
    """文件是否被 Excel/WPS 独占（能读到打不开写 = 被占用）"""
    if not os.path.isfile(path):
        return False
    try:
        with open(path, "r+b"):
            return False
    except OSError:
        return True


def _sep(title=""):
    print("\n" + "=" * 74)
    if title:
        print(title)
        print("=" * 74)


# ================================================================== 路径
def paths(period, snapshot):
    snap = os.path.join(TRUNK_DIR, "data", "数据归属周期", period, "快照日期", snapshot)
    res = os.path.join(TRUNK_DIR, "result")
    year = period[:4]
    return {
        "snap": snap,
        "mst_tmpl": os.path.join(snap, "月销售总表.xlsx"),
        "mst_out": os.path.join(res, "数据归属周期", period, "月销售总表.xlsx"),
        "sea_tmpl": os.path.join(snap, "东南亚", f"{year}年东南亚销量表Southeast Asia Sales table.xlsx"),
        "sea_out": os.path.join(res, "东南亚销量表", period,
                                f"{year}年东南亚销量表Southeast Asia Sales table.xlsx"),
        "eu_out": os.path.join(res, "欧洲销量表", period,
                               f"{year}年欧洲销量表Europe Sales table.xlsx"),
        "na_out": os.path.join(res, "北美销量表", period,
                               f"{year}年北美销量表North America Sales table.xlsx"),
    }


# ================================================================== 跑一步
def run_script(script, period, snapshot, log_path, allow_low_cover=False,
               merge=False, extra_args=None):
    """跑一个 python 脚本；返回 (ok, 耗时秒)。日志全文写 log_path，重要行实时打印。"""
    env = os.environ.copy()
    env["MSF_PERIOD"] = period
    env["MSF_SNAPSHOT"] = snapshot
    env["PYTHONIOENCODING"] = "utf-8"
    env.pop("MSF_ALLOW_LOW_COVER", None)
    if allow_low_cover:
        env["MSF_ALLOW_LOW_COVER"] = "1"

    cmd = [sys.executable, os.path.join(TRUNK_DIR, script)]
    if extra_args:
        cmd += extra_args

    t0 = datetime.datetime.now()
    shown = total = 0
    lines = []
    with open(log_path, "w", encoding="utf-8", errors="replace") as log:
        log.write(f"$ {' '.join(cmd)}\n[cwd] {TRUNK_DIR}\n"
                  f"[env] MSF_PERIOD={period} MSF_SNAPSHOT={snapshot}"
                  f"{' MSF_ALLOW_LOW_COVER=1' if allow_low_cover else ''}\n")
        log.write("=" * 74 + "\n")
        p = subprocess.Popen(cmd, cwd=TRUNK_DIR, env=env,
                             stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                             text=True, encoding="utf-8", errors="replace", bufsize=1)
        for line in p.stdout:
            lines.append(line)
            log.write(line)
            total += 1
            if any(k in line for k in SHOW_KEYS):
                s = line.rstrip()
                if s:
                    print("   " + (s if len(s) < 160 else s[:157] + "…"))
                    shown += 1
        p.wait()
    secs = (datetime.datetime.now() - t0).total_seconds()
    ok = p.returncode == 0
    if not ok:
        print(f"\n   ---- 失败，日志尾部（全文：{log_path}）----")
        for line in lines[-40:]:
            print("   " + line.rstrip())
    else:
        print(f"   （日志 {total} 行，已显示 {shown} 行；全文 {log_path}）")
    return ok, secs


# ================================================================== 验收
class Checker:
    def __init__(self):
        self.items = []          # (ok, 分类, label, detail)

    def add(self, ok, cat, label, detail=""):
        self.items.append((ok, cat, label, detail))
        print(f"   {'✅' if ok else '❌'} [{cat}] {label}" + (f"  {detail}" if detail else ""))
        return ok

    def eq(self, cat, label, got, exp, tol=0.5):
        return self.add(abs(got - exp) <= tol, cat, label, f"{got:.0f} vs {exp:.0f} ({_pct(got, exp)})")

    @property
    def failed(self):
        return [i for i in self.items if not i[0]]


def verify(period, snapshot, verbose=True):
    """恒等链全量验收。返回 (Checker, 定案数字 dict 或 None)"""
    m = int(period[4:])
    P = paths(period, snapshot)
    ck = Checker()

    for k in ("mst_out", "sea_out", "eu_out", "na_out"):
        if not os.path.isfile(P[k]):
            ck.add(False, "文件", f"缺成品 {os.path.basename(P[k])}", P[k])
            return ck, None

    mst = group_month_totals(P["mst_out"], m, year=period[:4])
    _sep(f"① 月销总表 {period} / {m} 月 —— 分组结构与恒等式")
    for g in ("亚马逊", "东南亚", "总计", "欧洲", "英国", "美国", "加拿大", "墨西哥", "美客多", "印尼", "菲律宾", "泰国", "京东", "eMAG", "TEMU"):
        print(f"   · {g:<6} {mst[g]:>9.0f}")
    for w in group_map_warnings(mst):
        ck.add(False, "月销表", "恒等式不符", w)
    if not group_map_warnings(mst):
        ck.add(True, "月销表", "亚马逊=十国和 / 东南亚=八国和 / 总计=各块和 全部成立")
    ck.add(mst["总计"] > 0, "月销表", "总计 > 0", f"{mst['总计']:.0f}")

    # ---------------- ② 东南亚表 ----------------
    _sep(f"② 东南亚销量表 —— 逐国对撞 / SEA Total 块 / 订单明细")
    sx = {}
    sea_gets = {}
    for cn, sheet, width, delta in SEA_ROWS:
        x = sx.setdefault(sheet, sheet_xml(P["sea_out"], sheet))
        col = get_column_letter(3 + (m - 1) * width + delta)     # 当月「月合计」列
        v = _cell(x, f"{col}2")
        sea_gets[cn] = v
        ck.eq("东南亚表", f"{cn} = 月销表{MST_ALIAS.get(cn, cn)}列", v, mst[MST_ALIAS.get(cn, cn)])
    xt = sheet_xml(P["sea_out"], SEA_TOTAL_SHEET)
    c0 = SEA_TOTAL_BASE + (m - 1) * 9
    head = _cell(xt, f"{get_column_letter(c0)}2")
    cnts = [_cell(xt, f"{get_column_letter(c0 + i)}2") for i in range(1, 9)]
    ck.eq("东南亚表", "SEA Total 块首 = 8 国之和", head, sum(cnts))
    ck.eq("东南亚表", "SEA Total 块首 = 月销表东南亚列", head, mst["东南亚"])
    bad = []
    for r in range(3, 900):
        a = _cell(xt, f"{get_column_letter(c0)}{r}")
        b = sum(_cell(xt, f"{get_column_letter(c0 + i)}{r}") for i in range(1, 9))
        if (a or b) and abs(a - b) > 0.5:
            bad.append(r)
    ck.add(not bad, "东南亚表", "SEA Total 逐 SKU 行：合计列 = 8 国列之和",
           "0 行不一致" if not bad else f"{len(bad)} 行不一致，首行 {bad[:5]}")
    # 订单明细 sheet：名称形如「Philippines 菲律宾订单明细 10月」，件数列按表头「数量 Qty」定位
    for cn, en in (("菲律宾", "Philippines"), ("泰国", "Thailand"), ("印尼", "Indonesia")):
        nm = next((s for s in _sheetnames(P["sea_out"]) if s.startswith(en) and "订单明细" in s), None)
        if not nm:
            ck.add(False, "明细", f"{cn} 订单明细 sheet 缺失")
            continue
        xd = sheet_xml(P["sea_out"], nm)
        qcol = _qty_col(P["sea_out"], nm)
        if not qcol:
            ck.add(False, "明细", f"{cn} 明细里找不到「数量」列")
            continue
        ck.eq("明细", f"{cn} 明细件数 = 国家月合计",
              col_sum(xd, qcol, min_row=3), sea_gets.get(cn, -1))

    # ---------------- ③ 欧洲表 ----------------
    _sep("③ 欧洲销量表 —— 逐国 vs 月销表")
    eu9 = 0.0
    eu_uk = 0.0
    for cn, sheet, nch in EU_ROWS:
        width = nch + 1
        x = sheet_xml(P["eu_out"], sheet)
        col = get_column_letter(3 + (m - 1) * width + nch)
        v = _cell(x, f"{col}2")
        if cn == "英国":
            eu_uk = v
            ck.eq("欧洲表", "英国 = 月销表英国列", v, mst["英国"])
        else:
            eu9 += v
        print(f"   · {cn:<4} {v:>8.0f}")
    ck.eq("欧洲表", "九国之和 = 月销表欧洲列", eu9, mst["欧洲"])
    n_eu = len(EU_ROWS)
    xe = sheet_xml(P["eu_out"], EU_TOTAL_SHEET)
    tc = EU_TOTAL_BASE + n_eu + 1 + (m - 1) * (n_eu + 1) + n_eu
    ck.eq("欧洲表", "EU Total 当月合计 = 十国之和",
          _cell(xe, f"{get_column_letter(tc)}2"), eu9 + eu_uk)
    ck.eq("欧洲表", "欧洲表总合计 = 月销表(欧洲+英国)", eu9 + eu_uk, mst["欧洲"] + mst["英国"])

    # ---------------- ④ 北美表 ----------------
    _sep("④ 北美销量表 —— 逐国 vs 月销表")
    na = {}
    for cn, sheet, nch in NA_ROWS:
        width = nch + 1
        x = sheet_xml(P["na_out"], sheet)
        col = get_column_letter(3 + (m - 1) * width + nch)
        na[cn] = _cell(x, f"{col}2")
    ck.eq("北美表", "美国 = 月销表美国+TEMU", na["美国"], mst["美国"] + mst["TEMU"])
    ck.eq("北美表", "加拿大 = 月销表加拿大", na["加拿大"], mst["加拿大"])
    ck.eq("北美表", "墨西哥 = 月销表墨西哥+美客多", na["墨西哥"], mst["墨西哥"] + mst["美客多"])
    n_na = len(NA_ROWS)
    xn = sheet_xml(P["na_out"], NA_TOTAL_SHEET)
    tc = NA_TOTAL_BASE + n_na + 1 + (m - 1) * (n_na + 1) + n_na
    ck.eq("北美表", "NA Total 当月合计 = 三国之和", _cell(xn, f"{get_column_letter(tc)}2"), sum(na.values()))

    # ---------------- ⑤ 汇总 ----------------
    _sep("⑤ 汇总")
    print(f"   月销表总计 {mst['总计']:.0f} = 亚马逊 {mst['亚马逊']:.0f} + 东南亚 {mst['东南亚']:.0f}"
          f" + 京东 {mst['京东']:.0f} + eMAG {mst['eMAG']:.0f} + TEMU {mst['TEMU']:.0f}"
          f" + 美客多 {mst['美客多']:.0f}")
    print(f"   欧洲表 {eu9 + mst['英国']:.0f} / 北美表 {sum(na.values()):.0f} / 东南亚表 {head:.0f}")
    fails = ck.failed
    ck.add(not fails, "总体验收", f"{len(ck.items)} 项检查，失败 {len(fails)} 项")

    numbers = {
        "snapshot": snapshot,
        "verified_at": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "总计": int(round(mst["总计"])), "亚马逊": int(round(mst["亚马逊"])),
        "东南亚": int(round(mst["东南亚"])), "京东": int(round(mst["京东"])),
        "eMAG": int(round(mst["eMAG"])), "TEMU": int(round(mst["TEMU"])),
        "美客多": int(round(mst["美客多"])),
        "欧洲": int(round(mst["欧洲"])), "英国": int(round(mst["英国"])),
        "美国": int(round(mst["美国"])), "加拿大": int(round(mst["加拿大"])),
        "墨西哥": int(round(mst["墨西哥"])),
    }
    for cn, _sheet, _w, _d in SEA_ROWS:
        numbers[MST_ALIAS.get(cn, cn)] = int(round(mst[MST_ALIAS.get(cn, cn)]))
    return ck, numbers


# ================================================================== 定案数字
def save_final_numbers(period, numbers):
    data = {}
    if os.path.isfile(FINAL_JSON):
        try:
            with open(FINAL_JSON, encoding="utf-8") as f:
                data = json.load(f)
        except Exception:                                     # noqa: BLE001
            data = {}
    data[period] = numbers
    with open(FINAL_JSON, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2, sort_keys=True)
    print(f"   ✅ 定案数字已写入 {FINAL_JSON}")


def load_final_numbers(period):
    if not os.path.isfile(FINAL_JSON):
        return None
    try:
        with open(FINAL_JSON, encoding="utf-8") as f:
            return json.load(f).get(period)
    except Exception:                                         # noqa: BLE001
        return None


# ================================================================== 主流程
STEPS = {
    1: ("合并欧盟/英国亚马逊订单报告", "工具-合并欧盟英国报告.py"),
    2: ("月销售总表", "month_sale_file.py"),
    3: ("东南亚销量表", "southeast_asia_sale_file.py"),
    4: ("欧洲销量表", "europe_sale_file.py"),
    5: ("北美销量表", "north_america_sale_file.py"),
}


def main():
    ap = argparse.ArgumentParser(description="一键跑全链（月销表→东南亚→欧洲→北美）+ 断言验收")
    ap.add_argument("period", help="数据归属周期 YYYYMM，如 202610")
    ap.add_argument("snapshot", help="快照日期 YYYYMMDD，如 20261107")
    ap.add_argument("--from", dest="from_step", type=int, default=1,
                    help="从第几步开始（1=合并 2=月销 3=东南亚 4=欧洲 5=北美）")
    ap.add_argument("--skip-merge", action="store_true", help="跳过第 1 步（欧盟/英国报告合并）")
    ap.add_argument("--verify-only", action="store_true", help="不跑脚本，只验收现有成品")
    ap.add_argument("--allow-low-cover", action="store_true",
                    help="报告区间体检命中严重项也继续（透传 MSF_ALLOW_LOW_COVER=1）")
    ap.add_argument("--dry-run", action="store_true", help="只打印计划")
    args = ap.parse_args()

    period, snapshot = args.period, args.snapshot
    if not re.fullmatch(r"\d{6}", period):
        sys.exit("❌ 期别应为 YYYYMM，例如 202610")
    if not re.fullmatch(r"\d{8}", snapshot):
        sys.exit("❌ 快照日期应为 YYYYMMDD，例如 20261107")

    P = paths(period, snapshot)
    os.makedirs(LOG_DIR, exist_ok=True)
    stamp = datetime.datetime.now().strftime("%H%M%S")
    print(f"[期别] {period}    [快照] {snapshot}")
    print(f"[工作目录] {TRUNK_DIR}")
    print(f"[日志] {LOG_DIR}\\runall_{period}_{snapshot}_{stamp}_*.log")

    # ---------------- 前置检查 ----------------
    _sep("0/5 前置检查")
    problems = []
    if not os.path.isdir(P["snap"]):
        problems.append(f"快照目录不存在：{P['snap']}")
    for key, label in (("mst_tmpl", "月销总表模板"), ("sea_tmpl", "东南亚表模板")):
        if not os.path.isfile(P[key]):
            problems.append(f"{label}不存在：{P[key]}（先跑 工具-准备期模板.py）")
    for key, label in (("eu_out", "欧洲表成品"), ("na_out", "北美表成品")):
        if _is_locked(P[key]):
            problems.append(f"{label} 正被 Excel/WPS 打开，先关掉：{P[key]}")
    if _is_locked(P["mst_out"]):
        problems.append(f"月销售总表正被打开，先关掉：{P['mst_out']}")
    if _is_locked(P["sea_out"]):
        problems.append(f"东南亚销量表正被打开，先关掉：{P['sea_out']}")
    plan = [f"{i}. {STEPS[i][0]}" for i in sorted(STEPS)
            if i >= args.from_step and not (i == 1 and args.skip_merge)]
    if args.verify_only:
        plan = ["（--verify-only：只跑第 6 步验收）"]
    print("   计划：" + " → ".join(plan))
    if problems:
        for p in problems:
            print("   ❌ " + p)
        return 3
    print("   ✅ 前置条件满足")
    if args.dry_run:
        print("\n[dry-run] 未执行任何脚本")
        return 0

    # ---------------- 跑步骤 ----------------
    t_all = datetime.datetime.now()
    for i in sorted(STEPS):
        if args.verify_only or i < args.from_step or (i == 1 and args.skip_merge):
            continue
        name, script = STEPS[i]
        _sep(f"{i}/5 {name}  ({script})")
        log_path = os.path.join(LOG_DIR, f"runall_{period}_{snapshot}_{stamp}_{i}_{script.replace('.py', '')}.log")
        if i == 1:
            ok, secs = run_script(script, period, snapshot, log_path,
                                  extra_args=[period, snapshot])
        else:
            ok, secs = run_script(script, period, snapshot, log_path,
                                  allow_low_cover=args.allow_low_cover)
        print(f"   {'✅ 完成' if ok else '❌ 失败'}  用时 {secs:.0f}s")
        if not ok:
            print(f"\n❌ 第 {i} 步（{name}）失败，后续步骤已中止。日志：{log_path}")
            return 2

    # ---------------- 验收 ----------------
    _sep("6/6 交叉验收（恒等链 / 分组对撞 / 明细件数）")
    ck, numbers = verify(period, snapshot)
    if numbers is None:
        print("\n❌ 成品不全，验收无法进行")
        return 1
    if ck.failed:
        print(f"\n❌ 验收未通过：{len(ck.failed)} 项失败（明细见上）")
        return 1
    save_final_numbers(period, numbers)
    print(f"\n✅ 全链跑通，用时 {(datetime.datetime.now() - t_all).total_seconds():.0f}s")
    print(f"   定案：总计 {numbers['总计']:,} = 亚马逊 {numbers['亚马逊']:,} + 东南亚 {numbers['东南亚']:,}"
          f" + 京东 {numbers['京东']:,} + eMAG {numbers['eMAG']:,} + 美客多 {numbers['美客多']:,}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
