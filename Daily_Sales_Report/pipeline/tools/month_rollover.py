# -*- coding: utf-8 -*-
"""月度换月程序：9月归档 + 10月开新

窗口口径（用户 2026-10-09 确定）：当月表保留【上月7天 + 当月整月 + 下月1天】；
促销期可扩窗口 —— 备案项，暂不实现（见《每日导数清单-2026年10月.md》）。

换月动作（对照 2026年8月 人工归档件逐格核实，probe5/probe6）：
  ① 每日销售数据：
     归档版 = 保留 [宽松下界(现状起点), 2026-10-01]（右界=下月1天，同人工8月件"止于9-01"；
              人工8月件起点 7-23/7-24 比严格7天宽 → 归档左界不裁，宁多勿少）
     10月版 = 保留 [2026-09-24, ∞)（10月的上月7天 = 9-24~9-30，9月整块删除）
  ② 库存记录 / 每日库存分析表：删 日期行 日期 > 2026-10-01 的尾列
     （当前根目录记录列到 10-06 → 归档删 10-02..10-06；10月版原样保留9月列——
       左侧删列会破坏跨表公式引用，设计规则：10月版不动记录/分析表列）
  ③ 透视图定格（仅模拟透视件；欧盟/北美/Interflowd北美=真透视表件零改动）：
     - 锚点格（"下载日期"格，如 沙特阿拉伯可售库存!B1）值 → 2026-10-01（下月1日，8月件全为9-1）
     - 透视图公式里 DAY(<本sheet的=TODAY()格引用>) → DAY('<锚点sheet>'!$B$1)（人工同款改写）
     - =TODAY() 公式本身保留不动（人工8月件如此，仅显示格会漂移）
  10月版：③全部不做。引用TODAY的件（沙特/澳/印/阿联酋/Int阿联酋）自动跟随系统日期；
     日本/Interflowd澳 引用静态锚点格，由 runner 日常跑时滚动锚点值（见 runner 扩展任务）。

透视图锚点格（probe5 实测）：
  日本='可售+在途库存'!C1  沙特='沙特阿拉伯可售库存'!B1  澳/印='可售库存'!B1
  阿联酋='阿联酋可售库存'!B1  Interflowd澳='在途库存'!C1  Interflowd阿联酋='阿联酋可售库存'!B1

用法：
  python month_rollover.py --dry-run          # 打印每份文件的变更计划
  python month_rollover.py --apply            # 实跑（混合态备份→归档→10月版→校验）
  python month_rollover.py --apply --phase archive|october|verify
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
import os, sys, shutil, zipfile, re, argparse, gc, calendar
from copy import copy
from datetime import datetime, date, timedelta
from openpyxl import load_workbook
from openpyxl.formula.translate import Translator
from openpyxl.utils import get_column_letter

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
# 任务㉓：日列补建/列宽修复两个维护模块在 maintenance/ 下（ensure_day_cols 懒加载用）
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "maintenance"))
import config as C
from runner import _header_row, restore_parts  # noqa: F401
import status_sync                              # 方案B：封存前状态校正（2026-11-02）

WORK = C.WORK_DIR
ARCHIVE_DIR = os.path.join(WORK, "2026年每日数据", "2026年9月")
BACKUP_MIXED = os.path.join(C.DSR, ".workbuddy", "backup_mixed_20261009")

ARCH_PREFIX = "2026年9月-"

# 窗口（默认值 = 2026-09 → 2026-10，即首次换月的留档值；
# 日常换月请用 --from 202610 --to 202611，下面 _apply_window() 会整组覆盖）
FROM_PERIOD, TO_PERIOD = "202609", "202610"
ARCH_LO, ARCH_HI = "2026-08-01", "2026-10-01"   # 归档：左界宽松不裁（现状起点8-23/8-24），右界=下月1天
OCT_LO = "2026-09-24"                            # 10月版：上月7天起点（9-24~9-30承接+10月）
ANCHOR_D = date(2026, 10, 1)                     # 归档锚点值 = 下月1日（人工8月件=9-1 同款）
NEW_LO = OCT_LO                                  # 新版本的窗口左界（= 旧月最后 7 天起点）
OLD_MONTH_SPAN = ("2026-09-01", "2026-09-30")    # 校验时统计「旧月销量」的区间


def _derive(from_period, to_period):
    """由 旧月/新月 的 YYYYMM 推出整套窗口常量。

    口径（沿用 2026-10-09 用户确定的那套，9→10 已人工核实）：
      · 归档版：左界不裁（宁多勿少），右界 = 新月 1 号
      · 新月版：保留 [旧月最后 7 天起点, ∞)
      · 锚点格定格到 新月 1 号
    """
    fy, fm = int(from_period[:4]), int(from_period[4:6])
    ty, tm = int(to_period[:4]), int(to_period[4:6])
    last = date(fy, fm, calendar.monthrange(fy, fm)[1])
    return dict(
        archive_dir=os.path.join(WORK, "2026年每日数据", f"{fy}年{fm}月"),
        arch_prefix=f"{fy}年{fm}月-",
        arch_lo="1900-01-01",                  # 等同「不裁左界」
        arch_hi=date(ty, tm, 1).isoformat(),
        new_lo=(last - timedelta(days=6)).isoformat(),
        anchor_d=date(ty, tm, 1),
        old_span=(date(fy, fm, 1).isoformat(), last.isoformat()),
        from_period=from_period, to_period=to_period,
    )


def _apply_window(from_period, to_period):
    """把 --from/--to 推导出的窗口写到模块全局，供各阶段共用。"""
    global ARCHIVE_DIR, ARCH_PREFIX, ARCH_LO, ARCH_HI, OCT_LO, ANCHOR_D, NEW_LO
    global OLD_MONTH_SPAN, FROM_PERIOD, TO_PERIOD, BACKUP_MIXED
    d = _derive(from_period, to_period)
    # 备份目录必须按目标周期区分：backup_mixed() 有「已存在就不覆盖」的保护，
    # 若沿用固定名（原 20261009），第二次换月会因目录已存在而**一份都不备份**。
    BACKUP_MIXED = os.path.join(C.DSR, ".workbuddy", f"backup_mixed_{to_period}")
    ARCHIVE_DIR, ARCH_PREFIX = d["archive_dir"], d["arch_prefix"]
    ARCH_LO, ARCH_HI = d["arch_lo"], d["arch_hi"]
    OCT_LO = NEW_LO = d["new_lo"]
    ANCHOR_D = d["anchor_d"]
    OLD_MONTH_SPAN = d["old_span"]
    FROM_PERIOD, TO_PERIOD = from_period, to_period
    print(f"[窗口] {from_period} → {to_period}"
          f" | 归档 → {os.path.basename(ARCHIVE_DIR)}"
          f" | 归档销售 {ARCH_LO}~{ARCH_HI}"
          f" | 新版销售 ≥{NEW_LO} | 锚点 {ANCHOR_D}")

# 模拟透视件锚点格 → config.ANCHORS（与 runner 共用；真透视表件不在表内 → 归档零公式改动）
ANCHORS = C.ANCHORS

# 9月程序值（sep_compare 输出，校验对照）
SEP_PROG = {"欧盟英国": 31782, "北美": 7484, "日本": 5070, "沙特": 1424,
            "澳大利亚": 1598, "印度": 1107, "阿联酋": 965,
            "Interflowd北美": 567, "Interflowd澳": 236, "Interflowd阿联酋": 27}
VERIFY_KEY = {
    "欧盟英国库存销售表.xlsx": "欧盟英国", "北美加拿大库存销售表.xlsx": "北美",
    "日本库存销售表.xlsx": "日本", "沙特库存销售表.xlsx": "沙特",
    "澳大利亚库存销售表.xlsx": "澳大利亚", "印度库存销售表.xlsx": "印度",
    "阿联酋库存销售表.xlsx": "阿联酋",
    "Interflowd北美加拿大库存销售表.xlsx": "Interflowd北美",
    "Interflowd澳大利亚库存销售表.xlsx": "Interflowd澳",
    "Interflowd阿联酋库存销售表.xlsx": "Interflowd阿联酋",
}
def _sales_cfg(wb_cfg):
    """取 sales 配置项。2026-11-02：sales 可能是列表（美客多 9月补缺口 + 当月
    两份报告都写同一个 sheet），旧代码按 dict 取值会 TypeError。取最后一条即可
    ——它们指向同一个 sheet，裁剪窗口不受影响。"""
    sc = wb_cfg.get("sales")
    if isinstance(sc, (list, tuple)):
        sc = sc[-1] if sc else None
    return sc if isinstance(sc, dict) else None


SALES_SHEET = {w["file"]: (_sales_cfg(w) or {}).get("sheet")
               for w in C.WORKBOOKS if _sales_cfg(w)}
RECORD_SHEETS = {w["file"]: w.get("records") or [] for w in C.WORKBOOKS}


def norm_day(v):
    """pc 值 → 'YYYY-MM-DD'；非日期返回 None"""
    if v is None:
        return None
    if hasattr(v, "strftime"):
        return v.strftime("%Y-%m-%d")
    s = str(v).strip()
    if re.match(r"^\d{4}-\d{2}-\d{2}", s):
        return s[:10]
    return None


# ---------------- 销售窗口删行 ----------------

def prune_sales_window(ws, lo, hi=None, dry=False):
    """每日销售数据窗口删行（逐行判定，行序乱必须如此）。
    返回 (删除行数, 连续段数)。段数少时按段 delete_rows（probe5 实测 ≤11 段）。"""
    hdr = _header_row(ws)
    headers = {str(c.value).strip(): c.column for c in ws[hdr] if c.value is not None}
    pc = headers.get("purchase-date")
    if not pc:
        print(f"    [!] {ws.title}: 无 purchase-date 列，跳过删行")
        return (0, 0)
    last = ws.max_row
    del_rows, last_data = [], hdr
    for row in ws.iter_rows(min_row=hdr + 1, max_row=last, min_col=pc, max_col=pc):
        v = row[0].value
        if v is None:
            continue
        r = row[0].row
        last_data = max(last_data, r)
        d = norm_day(v)
        if d is None:
            continue  # 非日期值不动
        if d < lo or (hi and d > hi):
            del_rows.append(r)
    if not del_rows:
        return (0, 0)
    # 合并连续段
    runs = []
    s = e = del_rows[0]
    for r in del_rows[1:]:
        if r == e + 1:
            e = r
        else:
            runs.append((s, e)); s = e = r
    runs.append((s, e))
    if dry:
        return (len(del_rows), len(runs))
    if len(runs) > 300:  # 保险丝：正常 ≤11 段，异常散乱时人工介入
        raise RuntimeError(f"{ws.title}: 待删行散成 {len(runs)} 段，超出预期，停止（需人工复核）")
    # 删行前：幸存行 新行号→旧行号 映射（delete_rows 不会改写公式引用，须人工平移公式）
    del_set = set(del_rows)
    old_of = {}
    shift = 0
    for r in range(hdr + 1, last_data + 1):
        if r in del_set:
            shift += 1
        else:
            old_of[r - shift] = r
    # 公式列（首数据行的公式格；全表列结构一致，probe7 已核实）
    fcols = [c for c in range(1, ws.max_column + 1)
             if isinstance(ws.cell(row=hdr + 1, column=c).value, str)
             and ws.cell(row=hdr + 1, column=c).value.startswith("=")]
    # 删行（从后往前按段）
    for s, e in reversed(runs):
        ws.delete_rows(s, e - s + 1)
    # 公式引用平移：幸存行内公式从旧行号译到新行号
    fixed = 0
    for new_r in sorted(old_of):
        old_r = old_of[new_r]
        if new_r == old_r:
            continue
        for c in fcols:
            cell = ws.cell(row=new_r, column=c)
            v = cell.value
            if isinstance(v, str) and v.startswith("="):
                L = get_column_letter(c)
                try:
                    cell.value = Translator(v, origin=f"{L}{old_r}").translate_formula(f"{L}{new_r}")
                    fixed += 1
                except Exception:
                    pass
    # autoFilter 收缩
    af = ws.auto_filter.ref
    if af:
        m = re.match(r"([A-Z]+)(\d+):([A-Z]+)(\d+)", af)
        if m:
            ws.auto_filter.ref = f"{m.group(1)}{m.group(2)}:{m.group(3)}{last_data - len(del_rows)}"
    print(f"    [公式] {ws.title}: 平移{fixed}格公式（{len(fcols)}列 × 幸存行）")
    return (len(del_rows), len(runs))


# ---------------- 日期横排 sheet 尾列裁剪 ----------------

def _date_row(ws, probe_rows=(1, 10)):
    """找日期横排所在行（datetime 值最多的行）。返回 (行号, [(列,日期)]) 或 None"""
    best = None
    for row in ws.iter_rows(min_row=probe_rows[0], max_row=probe_rows[1], max_col=ws.max_column):
        dd = [(c.column, c.value.date()) for c in row if hasattr(c.value, "strftime")]
        if dd and (best is None or len(dd) > len(best[1])):
            best = (row[0].row, dd)
    return best


def prune_date_cols(ws, hi, dry=False, label=""):
    """删 日期 > hi 的【尾部连续】列（从末列往左收，遇 ≤hi 或非日期格即停）。
    行内其他位置的孤立日期格（备注/到货日等）一律不动。返回删除列数。"""
    dr = _date_row(ws)
    if not dr:
        return 0
    _, dates = dr
    if not dates:
        return 0
    by_col = dict(dates)
    maxc = ws.max_column
    del_cols = []
    c = maxc
    while c in by_col and by_col[c] > hi:
        del_cols.append(c)
        c -= 1
    if not del_cols:
        return 0
    del_cols.reverse()  # 升序；连续性由收集方式天然保证
    if dry:
        return len(del_cols)
    ws.delete_cols(del_cols[0], len(del_cols))
    return len(del_cols)


# ---------------- 透视图定格（模拟透视件专用） ----------------

def find_todays(ws):
    return [c.coordinate for row in ws.iter_rows() for c in row
            if isinstance(c.value, str) and c.value.replace(" ", "").upper().startswith("=TODAY")]


def retarget_today_to_anchor(wb, anchor_sheet, anchor_coord):
    """透视图公式里 DAY(<本sheet的TODAY格引用>[±算术尾巴]) → DAY('<锚点sheet>'!$X$Y[尾巴])。
    尾巴=引用后面的运算（如 DAY(透视图!$F$1-1) 的 '-1' 在括号内，澳/印写法，probe5）。
    返回改写公式数。人工8月归档同款动作。"""
    col, row = re.match(r"([A-Z]+)(\d+)", anchor_coord).groups()
    new_inner = f"'{anchor_sheet}'!${col}${row}"
    n = 0
    for sn in pivot_sheets(wb):
        ws = wb[sn]
        todays = find_todays(ws)
        if not todays:
            continue
        tpat = re.compile(r"([A-Z]+)(\d+)")
        today_pairs = {tpat.match(t).groups() for t in todays if tpat.match(t)}
        for rows in ws.iter_rows():
            for c in rows:
                v = c.value
                if not (isinstance(v, str) and v.startswith("=") and "DAY(" in v.upper()):
                    continue

                def sub(mo):
                    arg = mo.group(1)
                    bare = arg.replace("$", "").strip()
                    if "!" in bare:
                        bare = bare.rpartition("!")[2]
                    bare = bare.strip("'").strip()
                    m3 = re.match(r"^([A-Z]+\d+)(.*)$", bare)  # 坐标 + 可选算术尾巴
                    if m3:
                        coord_part, tail = m3.group(1), m3.group(2)
                        m2 = tpat.fullmatch(coord_part)
                        if m2 and m2.groups() in today_pairs:
                            return "DAY(" + new_inner + tail + ")"
                    return mo.group(0)

                nv = re.sub(r"DAY\(([^()]*)\)", sub, v, flags=re.I)
                if nv != v:
                    c.value = nv
                    n += 1
    return n


def set_anchors(wb, anchor_list, dry=False):
    """锚点格值 → ANCHOR_D。返回改动格数。"""
    n = 0
    for sn, coord in anchor_list:
        if sn not in wb.sheetnames:
            print(f"    [!] 锚点sheet不存在: {sn!r}")
            continue
        c = wb[sn][coord]
        if dry:
            n += 1
        else:
            old = c.value
            c.value = datetime(ANCHOR_D.year, ANCHOR_D.month, ANCHOR_D.day)
            n += 1
            print(f"    [锚点] {sn}!{coord}: {old} → {ANCHOR_D}")
    return n


def pivot_sheets(wb):
    return [n for n in wb.sheetnames if "透视" in n]


def sales_sheet(wb_cfg, wb):
    sc = _sales_cfg(wb_cfg)
    if not sc:
        return None
    sn = sc.get("sheet")
    return wb[sn] if sn and sn in wb.sheetnames else None


# ---------------- 阶段1：9月归档 ----------------

def build_archive(dry, fix_status=True):
    os.makedirs(ARCHIVE_DIR, exist_ok=True)
    print(f"\n########## 阶段1：{FROM_PERIOD} 归档 → {ARCHIVE_DIR}")
    # 方案B（2026-11-02 用户拍板）：封存前用当期最新订单报告的**最终状态**校正明细。
    # 不校正的话，`append_sales` 只增不删 → 取消单（首次抓到时是 Pending）会被永久
    # 计入历史，10 月实测 517 件 / 0.95%。见 core/status_sync.py 顶部说明。
    rep, old_ym = None, "%s-%s" % (FROM_PERIOD[:4], FROM_PERIOD[4:6])
    if fix_status:
        try:
            rep, snap, nf = status_sync.load_final_status(FROM_PERIOD)
            print(f"  状态校正源：快照 {snap} / {nf} 份订单报告 / 合并 {len(rep)} 单")
        except SystemExit as e:
            print(f"  [!] 状态校正跳过：{e}")
            print(f"  [!] 归档簿将保留未回退的取消单（日销表历史会偏高约 1%）")
    else:
        print("  [状态校正] 已禁用（--no-status-fix）")
    for wb_cfg in C.WORKBOOKS:
        f = wb_cfg["file"]
        src = os.path.join(WORK, f)
        dst = os.path.join(ARCHIVE_DIR, ARCH_PREFIX + f)
        if wb_cfg.get("skip"):
            if not dry:
                shutil.copy2(src, dst)
            print(f"  {f}: 美客多/eMAG 原样复制")
            continue
        wb = load_workbook(src, data_only=False)
        acts = []
        ws = sales_sheet(wb_cfg, wb)
        # a0) 封存前状态校正（方案B）：取消单的状态改 Cancelled、quantity 改 0
        if ws is not None and rep is not None:
            stt = status_sync.fix_sales_sheet(ws, rep, dry=dry)
            if not stt.get("ok"):
                print(f"    [!] {f}: 状态校正跳过（{stt.get('err')}）")
            elif stt["fixed"]:
                acts.append(f"状态校正{stt['fixed']}行/{stt['qty']:.0f}件")
        # a) 每日销售数据窗口（右界裁剪）
        if ws is not None:
            n, runs = prune_sales_window(ws, ARCH_LO, ARCH_HI, dry=True)
            if n:
                acts.append(f"销售删{n}行/{runs}段")
                if not dry:
                    prune_sales_window(ws, ARCH_LO, ARCH_HI, dry=False)
        # b) 记录/分析表删 >10-01 尾列
        for rn in wb_cfg.get("records", []):
            target = rn if rn in wb.sheetnames else next(
                (x for x in wb.sheetnames if x.strip() == rn.strip()), None)
            if target:
                n = prune_date_cols(wb[target], ANCHOR_D, dry=True)
                if n:
                    acts.append(f"{target.strip()}删列{n}")
                    if not dry:
                        prune_date_cols(wb[target], ANCHOR_D, dry=False)
        for sn in wb.sheetnames:
            if "每日库存分析表" in sn:
                n = prune_date_cols(wb[sn], ANCHOR_D, dry=True)
                if n:
                    acts.append(f"{sn}删列{n}")
                    if not dry:
                        prune_date_cols(wb[sn], ANCHOR_D, dry=False)
        # c) 锚点+公式重定向（仅模拟透视件）
        if f in ANCHORS:
            if dry:
                acts.append(f"锚点→{ANCHOR_D.isoformat()} + DAY公式重定向")
            else:
                na = set_anchors(wb, ANCHORS[f])
                nr = retarget_today_to_anchor(wb, ANCHORS[f][0][0], ANCHORS[f][0][1])
                acts.append(f"锚点×{na} 公式重定向×{nr}")
        # d) 月汇总重算并冻结旧月列（明细已校正 → 静态值同步降下来）
        if (not dry) and rep is not None and ws is not None:
            for x in status_sync.freeze_month_cols(wb, old_ym):
                if x.get("err"):
                    print(f"    [!] 月汇总 {x['sheet']}: {x['err']}")
                    continue
                dd = x["after"] - x["before"]
                if abs(dd) >= 1:
                    acts.append(f"{x['sheet']}月汇总{dd:+.0f}")
        if dry:
            print(f"  {f}: {'; '.join(acts) if acts else '(零改动→纯复制)'}")
        else:
            if acts:
                tmp = dst + ".new.xlsx"
                wb.save(tmp)
                wb.close()
                restore_parts(src, tmp)
                _safe_replace(tmp, dst)
            else:
                wb.close()
                shutil.copy2(src, dst)
            print(f"  {f}: {'; '.join(acts) if acts else '纯复制'} → {os.path.basename(dst)}")
        del wb
        gc.collect()


# ---------------- 阶段2：10月新版 ----------------

def build_october(dry):
    print(f"\n########## 阶段2：{TO_PERIOD} 新版 ← {WORK}")
    for wb_cfg in C.WORKBOOKS:
        f = wb_cfg["file"]
        src = os.path.join(WORK, f)
        if wb_cfg.get("skip"):
            print(f"  {f}: 美客多/eMAG 原样保留")
            continue
        wb = load_workbook(src, data_only=False)
        acts = []
        ws = sales_sheet(wb_cfg, wb)
        if ws is not None:
            n, runs = prune_sales_window(ws, OCT_LO, None, dry=True)
            if n:
                acts.append(f"销售删旧月块{n}行/{runs}段")
                if not dry:
                    prune_sales_window(ws, OCT_LO, None, dry=False)
        if dry:
            print(f"  {f}: {'; '.join(acts) if acts else '(零改动)'}")
        else:
            if acts:
                tmp = src + ".new.xlsx"
                wb.save(tmp)
                wb.close()
                restore_parts(src, tmp)
                os.replace(tmp, src)
            else:
                wb.close()
            print(f"  {f}: {'; '.join(acts) if acts else '不变'}")
        del wb
        gc.collect()


# ---------------- 阶段2.5：新月日列补建（任务㉓，2026-11-03） ----------------

def ensure_day_cols(dry):
    """当前簿分析表日销列补建至 TO_PERIOD 月末 + 逐列重设 <col>（宽度/折叠）。

    复用 maintenance/ 的两个已验证 XML 手术模块（11-03 首次实战 17 张表）：
      · rebuild_month_day_cols.process(path, dry, month_end)  右移+插列+公式/CF/DV/cols 同步
      · fix_daycol_defs.process(path, apply, hide_from)       日列逐列宽度 5.5 + 未来日折叠
    幂等：日列已到 month_end 的表自动跳过。折叠的最终口径由日常跑批
    analysis.sync_day_col_groups 按快照日接管，这里只是先把列补齐、给宽度。
    """
    ty, tm = int(TO_PERIOD[:4]), int(TO_PERIOD[4:6])
    mend = date(ty, tm, calendar.monthrange(ty, tm)[1])
    print(f"\n########## 阶段2.5：{TO_PERIOD} 日列补建（至 {mend.isoformat()}）")
    import rebuild_month_day_cols_20261103 as RDC
    import fix_daycol_defs_20261103 as FDD
    os.makedirs(BACKUP_MIXED, exist_ok=True)
    total = 0
    for wb_cfg in C.WORKBOOKS:
        f = wb_cfg["file"]
        p = os.path.join(WORK, f)
        if not os.path.exists(p):
            print(f"  {f}: 缺文件!")
            continue
        bp = os.path.join(BACKUP_MIXED, f)
        if not dry and not os.path.exists(bp):
            shutil.copy2(p, bp)          # 单独 --phase daycols 跑时也有备份
        try:
            n = RDC.process(p, dry=dry, month_end=mend)
        except Exception as e:
            print(f"  {f}: !! 日列补建失败 {e!r}")
            continue
        total += n
        if n and not dry:
            FDD.process(p, apply=True, hide_from=date(ty, tm, 1))
    print(f"  合计新增列: {total}")


# ---------------- 阶段3：校验 ----------------

def _scan_sales_file(p, sheet_name):
    """read_only 扫销售表：返回 (hdr, 范围min~max, 9月有效quantity求和) 或 None
    口径与 SEP_PROG 对齐：非Cancelled + Amazon.*渠道(或北美TEMU单) + quantity求和"""
    wb = load_workbook(p, read_only=True, data_only=True)
    if sheet_name not in wb.sheetnames:
        wb.close()
        return None
    ws = wb[sheet_name]
    hdr = pc = st = ch = mo = qt = None
    for ri, row in enumerate(ws.iter_rows(min_row=1, max_row=6, max_col=80), 1):
        vals = {}
        for ci, c in enumerate(row, 1):
            if c.value is not None:
                vals[str(c.value).strip()] = ci
        if any(k.startswith("order-item-id") for k in vals):
            hdr, pc = ri, vals.get("purchase-date")
            st = vals.get("order-status")
            ch = vals.get("sales-channel")
            mo = vals.get("merchant-order-id")
            qt = vals.get("quantity")
            break
    if not pc:
        wb.close()
        return None
    q9, dmin, dmax = 0, None, None
    for row in ws.iter_rows(min_row=hdr + 1, max_col=80):
        d = norm_day(row[pc - 1].value)
        if not d:
            continue
        dmin = d if dmin is None or d < dmin else dmin
        dmax = d if dmax is None or d > dmax else dmax
        if OLD_MONTH_SPAN[0] <= d <= OLD_MONTH_SPAN[1]:
            stat = (row[st - 1].value or "") if st else ""
            chan = (row[ch - 1].value or "") if ch else ""
            merch = (row[mo - 1].value or "") if mo else ""
            if str(stat).strip() != "Cancelled" and (
                    str(chan).startswith("Amazon.") or "TEMU" in str(merch).upper()):
                q = row[qt - 1].value if qt else 1
                try:
                    q9 += int(float(q))
                except (TypeError, ValueError):
                    q9 += 1
    wb.close()
    return (hdr, dmin, dmax, q9)


def _last_record_date(p, rec_names):
    wb = load_workbook(p, read_only=True, data_only=False)
    out = []
    for rn in rec_names:
        tgt = rn if rn in wb.sheetnames else next(
            (x for x in wb.sheetnames if x.strip() == rn.strip()), None)
        if not tgt:
            out.append((rn, "缺sheet"))
            continue
        dr = _date_row(wb[tgt], probe_rows=(1, 6))
        out.append((tgt.strip(), dr[1][-1][1].strftime("%m-%d") if dr else "无日期"))
    wb.close()
    return out


# 换月前基线（--apply 时自动采集）：归档版 9月销量 必须与基线一致
BASELINE = {}


def capture_baseline():
    for f in VERIFY_KEY:
        p = os.path.join(WORK, f)
        if os.path.exists(p):
            r = _scan_sales_file(p, SALES_SHEET.get(f, "每日销售数据"))
            if r:
                BASELINE[f] = r[3]
    print(f"换月前基线（表内旧月销量 {OLD_MONTH_SPAN[0]}~{OLD_MONTH_SPAN[1]}, quantity口径）: {BASELINE}")


def verify(dry_run_hint=False):
    print(f"\n########## 阶段3：校验")
    print("-- 归档版（旧月销量 vs 换月前基线 / 范围 / 记录列尾 / 锚点）--")
    for f, key in VERIFY_KEY.items():
        p = os.path.join(ARCHIVE_DIR, ARCH_PREFIX + f)
        if not os.path.exists(p):
            print(f"  [归档] {f}: 缺文件!")
            continue
        r = _scan_sales_file(p, SALES_SHEET.get(f, "每日销售数据"))
        if r:
            _, dmin, dmax, q9 = r
            base = BASELINE.get(f)
            if base is not None:
                flag = " 与换月前一致 ✓" if q9 == base else f"  ←←← 换月前={base}，不一致!"
                print(f"  [归档] {f}: 旧月有效销量={q9}{flag} 范围={dmin}~{dmax}")
            else:
                prog = SEP_PROG.get(key)
                dd = q9 - prog if prog is not None else None
                dds = f"{dd:+d}" if dd is not None else "?"
                print(f"  [归档] {f}: 旧月有效销量={q9} (报告终态程序值={prog}, Δ={dds};"
                      f" 表内含人工粘贴时点状态, 差异属口径内) 范围={dmin}~{dmax}")
        rec = _last_record_date(p, RECORD_SHEETS.get(f, []))
        bad = [f"{n}:{d}" for n, d in rec if d not in ("无日期",) and "2026-" + str(d) > ANCHOR_D.isoformat()]
        print(f"  [归档] 记录列尾: {rec if rec else '无记录sheet'}"
              f"{'  ←←← 超' + ANCHOR_D.isoformat() + '!' if bad else ''}")
        if f in ANCHORS:
            sn, coord = ANCHORS[f][0]
            wb = load_workbook(p, read_only=True, data_only=True)
            v = wb[sn][coord].value if sn in wb.sheetnames else "缺sheet"
            wb.close()
            ok = getattr(v, "strftime", lambda f=None: None)("%Y-%m-%d") == ANCHOR_D.isoformat() if v is not None else False
            print(f"  [归档] 锚点 {sn}!{coord} = {v}{'  ←←← 应为' + ANCHOR_D.isoformat() + '!' if not ok else ' ✓'}")
    print(f"-- 新版（窗口 ≥{NEW_LO} / 旧月件数应=0影响 / 锚点原样）--")
    for f, key in VERIFY_KEY.items():
        p = os.path.join(WORK, f)
        if not os.path.exists(p):
            continue
        r = _scan_sales_file(p, SALES_SHEET.get(f, "每日销售数据"))
        if r:
            _, dmin, dmax, q9 = r
            note = ""
            if dry_run_hint:
                note = " (dry-run 未删行，残留属预期)"
            elif dmin and dmin >= NEW_LO:
                note = " ✓"
            else:
                note = "  ←←← 窗口外残留!"
            print(f"  [新版] {f}: 范围={dmin}~{dmax} (旧月全月销量现剩{q9}){note}")
        if f in ANCHORS:
            sn, coord = ANCHORS[f][0]
            wb = load_workbook(p, read_only=True, data_only=True)
            v = wb[sn][coord].value if sn in wb.sheetnames else "缺sheet"
            wb.close()
            print(f"  [新版] 锚点 {sn}!{coord} = {v} (应保持原值，由runner日常滚动)")


def backup_mixed():
    os.makedirs(BACKUP_MIXED, exist_ok=True)
    for wb_cfg in C.WORKBOOKS:
        src = os.path.join(WORK, wb_cfg["file"])
        dst = os.path.join(BACKUP_MIXED, wb_cfg["file"])
        if not os.path.exists(dst):
            shutil.copy2(src, dst)
    print(f"混合态备份 → {BACKUP_MIXED}")


def _shift_period(period, n):
    """YYYYMM 前后移动 n 个月。"""
    y, m = int(period[:4]), int(period[4:6])
    idx = y * 12 + (m - 1) + n
    return "%04d%02d" % (idx // 12, idx % 12 + 1)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--apply", action="store_true")
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--phase", default="all",
                    choices=["all", "archive", "october", "daycols", "verify"])
    ap.add_argument("--no-status-fix", action="store_true",
                    help="封存前不做订单状态校正（取消单会留在历史里，日销表偏高约1%%）")
    ap.add_argument("--from", dest="from_period", default="",
                    help="旧月 YYYYMM（默认 202609，即首次换月留档）")
    ap.add_argument("--to", dest="to_period", default="",
                    help="新月 YYYYMM（默认 202610）")
    args = ap.parse_args()

    if args.from_period or args.to_period:
        fp = args.from_period or _shift_period(args.to_period, -1)
        tp = args.to_period or _shift_period(fp, 1)
        _apply_window(fp, tp)
    else:
        print(f"[窗口] 未指定 --from/--to，沿用留档值 {FROM_PERIOD} → {TO_PERIOD}")

    dry = not args.apply
    if args.apply:
        backup_mixed()
        capture_baseline()
    if args.phase in ("all", "archive"):
        build_archive(dry, fix_status=not args.no_status_fix)
    if args.phase in ("all", "october"):
        build_october(dry)
    if args.phase in ("all", "daycols"):
        ensure_day_cols(dry)
    if args.phase in ("all", "verify"):
        verify(dry_run_hint=dry)
    if dry:
        print("\n(dry-run 未写任何文件；确认后加 --apply 实跑)")


if __name__ == "__main__":
    main()
