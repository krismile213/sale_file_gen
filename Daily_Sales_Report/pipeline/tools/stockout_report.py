# -*- coding: utf-8 -*-
"""断货分析 —— 每日把断货 SKU 按地区汇总、查明原因、给采购方案。

用法（在任意目录）：
  python stockout_report.py                                  # 用最新快照
  python stockout_report.py --snap <快照目录>                 # 指定快照
  python stockout_report.py --age <库龄报告.csv>              # 额外指定库龄报告（可多次）
  python stockout_report.py --lead 7 --cover 14              # 后勤周期/目标覆盖天数

输出：
  Daily_Sales_Report/断货分析/亚马逊/断货明细/<YYYYMMDD>/断货明细_<YYYYMMDD>.xlsx
  Daily_Sales_Report/断货分析/亚马逊/断货明细/<YYYYMMDD>/断货报告_<YYYYMMDD>.md

口径：
  可售(亚马逊) = afn-fulfillable-quantity
  在途(亚马逊) = afn-inbound-working + shipped + receiving
  可售(美客多) = Resumen c19 可销售；在途 = c16 待录入 + c17 转账中
  可售(eMAG)  = stock_type=ready_to_sell 的数量之和（无在途概念）
  日均销量     = 近30天有效销量 / 30（有效 = 非 Cancelled / Storno）
  可售天数     = 可售 / 日均

断货分级（2026-10-23 起与「需求B 缺货情况表」统一）：
  严重断货  可售=0 且 在途=0                                  ↔ 需求B 三级
  一般断货  可售=0 且 在途>0                                  ↔ 需求B 二级
  预备断货  可售>0 且 可售天数 + 后勤周期 < 计划备货天数        ↔ 需求B 一级
            （`--no-plan` 可回退旧口径：可售天数 < 后勤周期）
  库龄呆滞  超90天呆滞件>0（独立标记，不覆盖上面的分级）

  计划备货天数读 `断货分析/备货天数.xlsx`（与需求B 同一份）；站点不在表里则回退旧口径。
  补货目标覆盖天数同步取计划备货天数 —— 判定说缺 120 天却只补 14 天会自相矛盾。
"""
import os, sys, csv, math, argparse, re
from collections import defaultdict
from datetime import datetime, timedelta, date

_HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.normpath(os.path.join(_HERE, "..", "core")))
import config as C            # noqa: E402
import runner as R            # noqa: E402
from openpyxl import load_workbook, Workbook   # noqa: E402
from openpyxl.styles import Font, PatternFill, Alignment   # noqa: E402
from openpyxl.utils import get_column_letter   # noqa: E402

# ---------------- 计划备货天数（唯一实现，shortage_report 复用本处） ----------------
PLAN_XLSX = os.path.join(C.DSR, "断货分析", "备货天数.xlsx")
_PLAN_ALIAS = {"沙特阿拉伯": "沙特", "澳大利亚": "澳洲"}


def plan_days_map():
    """读 备货天数.xlsx → {站点: 天数}。文件不在就用内置默认值。"""
    m = {}
    if os.path.exists(PLAN_XLSX):
        try:
            wb = load_workbook(PLAN_XLSX, data_only=True)
            ws = wb[wb.sheetnames[0]]
            for row in ws.iter_rows(min_row=2, values_only=True):
                if row and row[0] and row[1] is not None:
                    try:
                        m[str(row[0]).strip()] = int(float(row[1]))
                    except (TypeError, ValueError):
                        continue
            wb.close()
        except Exception:
            pass
    return m or {"欧洲": 120, "英国": 120, "美国": 95, "菲律宾": 95, "泰国": 95,
                 "加拿大": 75, "阿联酋": 75, "沙特": 75, "印尼": 75, "日本": 60,
                 "澳洲": 60, "罗马尼亚": 60, "墨西哥": 90, "美客多": 60, "印度": 65}


def plan_of(site, acct, m):
    """站点/账号 → 计划备货天数；查不到返回 None（调用方回退旧口径）。"""
    if acct == "美客多":
        return m.get("美客多")
    return m.get(_PLAN_ALIAS.get(site, site))


# ---------------- 站点定义 ----------------
# chans: 该站点对应的 sales-channel；drop: 已知混进来的别家渠道（拒收）
# orders: 订单报告文件名候选（{m}=YYYYMM）
NA = ["亚马逊北美站点{m}订单报告.txt"]
EU = ["亚马逊欧盟站点{m}订单报告.txt", "亚马逊英国站点{m}订单报告.txt"]
EU_CH = {"Amazon.co.uk", "Amazon.de", "Amazon.fr", "Amazon.it", "Amazon.es",
         "Amazon.nl", "Amazon.se", "Amazon.pl", "Amazon.ie", "Amazon.com.be"}

SITES = [
    dict(site="美国", acct="主账号", rep="亚马逊美国管理库存报告{d}.csv", kind="inv",
         chans={"Amazon.com"}, orders=NA),
    dict(site="加拿大", acct="主账号", rep="亚马逊加拿大管理库存报告{d}.csv", kind="inv",
         chans={"Amazon.ca"}, orders=NA),
    dict(site="墨西哥", acct="主账号", rep="亚马逊墨西哥管理库存报告{d}.csv", kind="inv",
         chans={"Amazon.com.mx"}, orders=NA),
    dict(site="英国", acct="主账号", rep="亚马逊英国管理库存报告{d}.csv", kind="inv",
         chans={"Amazon.co.uk"}, orders=EU),
    dict(site="欧洲", acct="主账号", rep="亚马逊欧盟管理库存报告{d}.csv", kind="inv",
         chans=EU_CH - {"Amazon.co.uk"}, orders=EU),
    dict(site="日本", acct="主账号", rep="亚马逊日本管理库存报告{d}.csv", kind="inv",
         colmap="JP24", chans={"Amazon.co.jp"}, orders=["亚马逊日本站点{m}订单报告.txt"]),
    dict(site="沙特阿拉伯", acct="主账号", rep="亚马逊沙特阿拉伯管理库存报告{d}.csv", kind="inv",
         chans={"Amazon.sa"}, orders=["亚马逊沙特阿拉伯站点{m}订单报告.txt"]),
    dict(site="澳大利亚", acct="主账号", rep="亚马逊澳大利亚管理库存报告{d}.csv", kind="inv",
         colmap="JP24", chans={"Amazon.com.au"}, orders=["亚马逊澳大利亚站点{m}订单报告.txt"]),
    dict(site="印度", acct="主账号", rep="亚马逊印度管理库存报告{d}.csv", kind="inv",
         chans={"Amazon.in"}, orders=["亚马逊印度站点{m}订单报告.txt"]),
    dict(site="阿联酋", acct="主账号", rep="亚马逊阿拉伯联合酋长国管理库存报告{d}.csv", kind="inv",
         chans={"Amazon.ae"}, orders=["亚马逊阿拉伯联合酋长国站点{m}订单报告.txt"]),
    # interflow 后台的库存导出**文件名日期不自动更新**（20261110 快照里美国/沙特/英国
    # 三份都还叫 …20261106.csv，内容确实是当天的）。只认 {d} 会判「缺报告」→
    # 该站点库存整片不进断货分析（2026-11-10 又踩一次，与 config.WORKBOOKS 同步加兜底）。
    dict(site="美国", acct="interflow",
         rep=["interflow美国管理库存报告{d}.csv",
              "interflow美国管理库存报告*.csv"], kind="inv",
         chans={"Amazon.com"}, drop={"Amazon.com.mx"}, orders=["interflow北美站点{m}订单报告.txt"]),
    dict(site="加拿大", acct="interflow",
         rep=["interflow加拿大管理库存报告{d}.csv",
              "interflow加拿大管理库存报告*.csv"], kind="inv",
         chans={"Amazon.ca"}, drop={"Amazon.com.mx"}, orders=["interflow北美站点{m}订单报告.txt"]),
    dict(site="澳大利亚", acct="interflow", rep="interflow澳大利亚管理库存报告{d}.csv", kind="inv",
         colmap="JP24", chans={"Amazon.com.au"}, orders=["interflow澳大利亚站点{m}订单报告.txt"]),
    dict(site="阿联酋", acct="interflow", rep="interflow阿拉伯联合酋长国管理库存报告{d}.csv", kind="inv",
         chans={"Amazon.ae"},
         orders=["interflow阿拉伯联合酋长国站点{m}订单报告.txt", "interflow沙特阿拉伯站点{m}订单报告.txt"]),
    # interflow 沙特仓 RUH8（2026-10-24 新增，配合 Interflowd沙特库存销售表）
    # 背景：interflow 账号在沙特有独立仓库，29 SKU / 115 件可售一直没有任何统计出口
    # （阿联酋那条只读阿联酋仓报告）。实测与「主账号-沙特阿拉伯」411 SKU **零重叠**，
    # 是两套独立业务，所以新增这条是纯增量，不影响其它站点的任何数字。
    # 订单：与阿联酋站共用同一份下载（sha256 相同），按 chans={"Amazon.sa"} 切出 5 单。
    dict(site="沙特阿拉伯", acct="interflow",
         # 精确名优先，再回退「任意日期的最新一份」（下载频率低，文件名会滞后）
         rep=["interflow沙特阿拉伯管理库存报告{d}.csv",
              "interflow沙特阿拉伯管理库存报告*.csv"], kind="inv",
         chans={"Amazon.sa"},
         orders=["interflow沙特阿拉伯站点{m}订单报告.txt"]),
]

AGE_TIERS = ["inv-age-0-to-30-days", "inv-age-31-to-60-days", "inv-age-61-to-90-days",
             "inv-age-91-to-180-days", "inv-age-181-to-270-days",
             "inv-age-271-to-365-days", "inv-age-365-plus-days"]
AGE_TIER_LABEL = ["0-30", "31-60", "61-90", "91-180", "181-270", "271-365", "365+"]
# 库龄报告 marketplace 代码 → 站点名（只挂主账号；interflow 账号同名站点无法区分，不自动挂）
AGE_MP = {"US": "美国", "CA": "加拿大", "MX": "墨西哥", "GB": "英国", "JP": "日本",
          "SA": "沙特阿拉伯", "AU": "澳大利亚", "IN": "印度", "AE": "阿联酋",
          "DE": "欧洲", "FR": "欧洲", "IT": "欧洲", "ES": "欧洲", "NL": "欧洲",
          "SE": "欧洲", "PL": "欧洲", "IE": "欧洲", "BE": "欧洲"}

_CANCEL = re.compile(r"cancel|storno|取消", re.I)


def dup_order_reports(snap_dir):
    """同一份订单报告被下载成两个文件名（如 亚马逊欧盟站点 == 亚马逊英国站点）。
    内容完全相同 → 只能取一份，否则销量翻倍。返回 [(保留名, 重复名)]"""
    import hashlib
    seen, dups = {}, []
    if not os.path.isdir(snap_dir):
        return dups
    for fn in sorted(os.listdir(snap_dir)):
        if "订单报告" not in fn:
            continue
        p = os.path.join(snap_dir, fn)
        try:
            h = hashlib.md5(open(p, "rb").read()).hexdigest()
        except OSError:
            continue
        if h in seen:
            dups.append((seen[h], fn))
        else:
            seen[h] = fn
    return dups


# ---------------- 通用读取 ----------------
def _int(v):
    n = R._num(v)
    if isinstance(n, (int, float)):
        return int(n)
    return 0


def _norm_sku(s):
    """`amzn.gr.<SKU>-<hash>[-LN|-VG|-PO]` → `<SKU>`，与 analysis.normalize_sku 同源。

    ⚠️ 库存报告的行也必须归并（2026-11-03 用户实测）：同一 ASIN 在欧盟/英国报告里
    会拆成「裸 SKU + amzn.gr.* 别名行」两条，别名行常常 0 库存，单独成行会被
    判成三级断货、按月度块历史日均算出天价补货单（G520 实测建议 2392 件，
    而同 ASIN 的 G520 行有 677 可售 + 1242 在途 —— 采购看到的就是这种矛盾）。
    月度块/历史销量查询早在用这条规则（2026-10-27 起），库存行不归并就是裂缝。
    analysis 导入失败时原样返回 → 归并自动退化为无操作，行为与旧版一致。
    """
    try:
        import analysis as A
        return A.normalize_sku(s)
    except Exception:
        return str(s or "").strip()


def _read_rows(path, sep):
    for enc in R._enc_order(path):
        try:
            with open(path, encoding=enc, newline="") as f:
                return list(csv.reader(f, delimiter=sep))
        except UnicodeDecodeError:
            continue
    with open(path, encoding="latin-1", newline="") as f:
        return list(csv.reader(f, delimiter=sep))


_OCACHE = {}


def load_orders(path, sep):
    """→ [(date10, sku, qty, channel, cancelled)]"""
    key = (path, sep)
    if key in _OCACHE:
        return _OCACHE[key]
    out = []
    rows = _read_rows(path, sep)
    if rows:
        hdr = [h.strip().lstrip("\ufeff") for h in rows[0]]
        ix = {h: i for i, h in enumerate(hdr)}
        i_sku, i_dt = ix.get("sku"), ix.get("purchase-date")
        i_qty, i_ch, i_st = ix.get("quantity"), ix.get("sales-channel"), ix.get("order-status")
        if i_sku is not None and i_dt is not None:
            for r in rows[1:]:
                if len(r) <= max(i_sku, i_dt):
                    continue
                sku = (r[i_sku] or "").strip()
                if not sku:
                    continue
                out.append((r[i_dt][:10], sku,
                            _int(r[i_qty]) if i_qty is not None and i_qty < len(r) else 0,
                            (r[i_ch] or "").strip() if i_ch is not None and i_ch < len(r) else "",
                            bool(_CANCEL.search(r[i_st] or "")) if i_st is not None and i_st < len(r) else False))
    _OCACHE[key] = out
    return out


def load_emag_orders(path):
    """eMAG 订单 xlsx：r1 表头，产品代码(5)/数量(8)/订单日期(2)/订单状态(13)/产品名称(4)"""
    key = (path, "emag")
    if key in _OCACHE:
        return _OCACHE[key]
    out = []
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or len(row) < 13:
            continue
        code = str(row[4] or "").strip()
        if not code:
            continue
        dt = str(row[1] or "")[:10]
        out.append((dt, code, _int(row[7]), "eMAG",
                    bool(_CANCEL.search(str(row[12] or "")))))
    wb.close()
    _OCACHE[key] = out
    return out


def read_ml_resumen(path):
    """美客多 Resumen：返回 {sku: dict}（列号见文件头注释）"""
    wb = load_workbook(path, read_only=True, data_only=True)
    sn = "Resumen" if "Resumen" in wb.sheetnames else wb.sheetnames[0]
    ws = wb[sn]
    out = {}
    for row in ws.iter_rows(min_row=13, max_col=34, values_only=True):
        sku = row[3]
        if sku is None or not str(sku).strip():
            continue
        sku = str(sku).strip()
        out[sku] = dict(
            name=row[6], status=row[10],
            stale=row[13],           # c14 陈旧库存
            cover_w=row[14],         # c15 销售覆盖周数（'缺货'/'最多5周'/数字）
            pending=row[15],         # c16 待录入
            transfer=row[16],        # c17 转账中
            returned=row[17],        # c18 买家退回
            sellable=row[18],        # c19 可销售
            unsellable=row[19],      # c20 不适合销售
            sold30=row[30],          # c31 最近30天售出件数
            urgency=row[31],         # c32 运输紧急情况
        )
    wb.close()
    return out


def read_emag_stock(path):
    """eMAG 库存：c3 品名 / c7 产品代码 / c10 数量 / c11 stock_type，数据 r6 起。
    ⚠️ 该文件 dimension 失真，不能 read_only。"""
    wb = load_workbook(path, data_only=True)
    sn = "库存" if "库存" in wb.sheetnames else wb.sheetnames[min(2, len(wb.sheetnames) - 1)]
    ws = wb[sn]
    out = {}
    for r in range(6, ws.max_row + 1):
        code = ws.cell(row=r, column=7).value
        if code is None or not str(code).strip():
            continue
        code = str(code).strip()
        qty = _int(ws.cell(row=r, column=10).value)
        st = str(ws.cell(row=r, column=11).value or "").strip()
        d = out.setdefault(code, dict(name=ws.cell(row=r, column=3).value,
                                      sellable=0, other=0))
        if st == "ready_to_sell":
            d["sellable"] += qty
        else:
            d["other"] += qty
    wb.close()
    return out


def read_age(path):
    """库龄报告 → {(marketplace, sku): {tier: qty}}"""
    out = {}
    with open(path, encoding="utf-8-sig", newline="") as f:
        for row in csv.DictReader(f):
            sku = (row.get("sku") or "").strip()
            if not sku:
                continue
            mp = (row.get("marketplace") or "").strip().upper()
            tiers = tuple(_int(row.get(t)) for t in AGE_TIERS)
            out[(mp, sku)] = dict(zip(AGE_TIER_LABEL, tiers))
    return out


# ---- 库龄总表（第二来源，2026-11-03 加）--------------------------------
# trunk/age_report.py 随月销表生成《亚马逊EMAG库龄情况总表*.xlsx》，固定放
# trunk/data/库龄报表/。与 CSV 库龄报告（键=(marketplace代码, sku)，只能挂主账号）
# 互补：总表按（账号, 站点, SKU）分行 —— 亚马逊/interflow/美客多/eMAG 四个账号
# 都有，墨西哥等撞号 SKU 靠账号区分，interflow 与 eMAG 只有从这里才挂得上。
AGE_XLSX_DIR = os.path.join(os.path.dirname(C.SNAP_BASE), "库龄报表")
AGE_XLSX_GLOB = "亚马逊EMAG库龄情况总表*.xlsx"
# ⚠ 总表第 7 档标签是「366+」（= inv-age-366-to-455 + 456-plus 之和；老格式 365-plus），
#   本模块一律按**位置**映射回 AGE_TIER_LABEL（…/365+），否则 aged90 的键会对不上。
_AGE_XLSX_CACHE = {}          # path -> (ages, period)；build 与缺货表共用一次解析


def _find_age_xlsx():
    """库龄报表目录里最新的总表；没有返回 None。忽略「xxx (1).xlsx」重复下载副本。"""
    import glob as _glob
    hits = [h for h in _glob.glob(os.path.join(AGE_XLSX_DIR, AGE_XLSX_GLOB))
            if "(" not in os.path.basename(h)]
    hits = hits or _glob.glob(os.path.join(AGE_XLSX_DIR, AGE_XLSX_GLOB))
    return sorted(hits, key=os.path.getmtime)[-1] if hits else None


def read_age_xlsx(path):
    """库龄总表「本期明细」→ ({(账号, 站点, sku): {档: 件数}}, 期数)。

    行 1 是标题（含期数，如「库龄明细（202610）」），行 2 表头，数据行 3 起；
    列：账号/站点/SKU/品名/七档/合计/180天以上/180+占比 —— 七档按位置取数。
    全 0 行不入 dict（查不到 = 没库龄，与挂不上同一表现）。
    美客多行不进这个 dict：它的呆滞走 Resumen「陈旧库存」（ml_stale），
    长期仓储费口径跟它不是一回事，混进来会把断货明细的美客多数字带歪。
    """
    if path in _AGE_XLSX_CACHE:
        return _AGE_XLSX_CACHE[path]
    wb = load_workbook(path, data_only=True, read_only=True)
    ws = wb["本期明细"]
    out, period = {}, None
    for i, row in enumerate(ws.iter_rows(values_only=True), 1):
        if i == 1:                            # 标题行：库龄明细（202610）
            m = re.search(r"[（(](\d{6})[)）]", str(row[0] or ""))
            period = m.group(1) if m else None
            continue
        if i == 2 or not row or len(row) < 11 or row[0] is None or row[2] is None:
            continue                          # 表头 / 空行 / 尾部填充行
        acct, site, sku = (str(row[0]).strip(), str(row[1]).strip(),
                           str(row[2]).strip())
        tiers = {lab: _int(row[4 + j]) for j, lab in enumerate(AGE_TIER_LABEL)}
        if any(tiers.values()):
            out[(acct, site, sku)] = tiers
    wb.close()
    _AGE_XLSX_CACHE[path] = (out, period)
    return out, period


def age_xlsx_info():
    """→ (总表路径, 期数)。给缺货表写 sheet 说明用（解析走缓存，不重复读）。"""
    p = _find_age_xlsx()
    if not p:
        return None, None
    return p, read_age_xlsx(p)[1]


def _find(snap_dir, patterns, fmt):
    """候选模板 → 第一个命中的路径。支持 `*` / `?` 通配符（2026-10-24 加）。

    为什么需要：interflow 沙特仓的库存报告下载频率低于其它站点，
    快照目录里最新一份可能仍叫 `…20261021.csv`（内容已更新、文件名没变）。
    死认 {d} 会判「缺报告」→ 该站点永远不进断货分析。
    与 runner._resolve_report / precheck._resolve 保持同一套规则：
    glob 取字典序最新，并忽略 `xxx (1).csv` 这类重复下载副本。
    """
    import glob as _glob
    for p in patterns:
        f = os.path.join(snap_dir, p.format(**fmt))
        if "*" in f or "?" in f:
            hits = [h for h in _glob.glob(f) if "(" not in os.path.basename(h)]
            hits = hits or _glob.glob(f)
            if hits:
                return sorted(hits)[-1]
            continue
        if os.path.isfile(f):
            return f
    return None


def _period_of(snap_dir):
    """从快照目录推断数据归属周期：…/数据归属周期/202610/快照日期/20261021 → 202610。"""
    try:
        parts = os.path.normpath(str(snap_dir)).split(os.sep)
        for i, p in enumerate(parts):
            if p == "数据归属周期" and i + 1 < len(parts):
                return parts[i + 1]
    except Exception:
        pass
    return getattr(C, "CURRENT_PERIOD", None)


def _age_dirs(snap_dir):
    """库龄报告的候选目录（按优先级，靠后的同名 sku 覆盖靠前的）：

      ① 数据归属周期/<period>/库龄/ —— 2026-10-22 起：库龄报告统一放这里，可分子目录
      ② 快照目录 —— 老习惯（和库存报告放一起）仍兼容
    """
    dirs = []
    per = _period_of(snap_dir)
    if per:
        d = os.path.join(C.SNAP_BASE, per, "库龄")
        if os.path.isdir(d):
            dirs.append(d)
    dirs.append(snap_dir)
    return dirs


def _autodetect_age(*dirs):
    """目录里任何含 inv-age- 列头的 csv 都当库龄报告（含子目录）。

    同一站点可能堆了多天的报告：按修改时间升序返回，`ages.update()` 逐份覆盖，
    最终生效的是**最新一份**（避免旧库龄盖掉新库龄）。
    """
    hits, seen = [], set()
    for d in dirs:
        if not d or not os.path.isdir(d):
            continue
        for root, _sub, files in os.walk(d):
            for fn in files:
                if not fn.lower().endswith(".csv"):
                    continue
                p = os.path.join(root, fn)
                if p in seen:
                    continue
                try:
                    with open(p, encoding="utf-8-sig", newline="") as f:
                        head = f.readline()
                except Exception:
                    continue
                if "inv-age-" in head:
                    seen.add(p)
                    hits.append(p)
    hits.sort(key=lambda x: os.path.getmtime(x))
    return hits


# ---------------- 主流程 ----------------
def build(snap_dir, snap_day, age_files, lead_days, cover_days, moq=0,
          plan_map=None, use_plan=True):
    """use_plan=True（默认）时，预备断货按「可售天数+后勤周期 < 计划备货天数」判，
    与需求B 一级同口径；补货目标覆盖天数也取计划备货天数。"""
    plan_map = plan_map if plan_map is not None else plan_days_map()
    snap_dt = datetime.strptime(snap_day, "%Y%m%d").date()
    snap_iso = snap_dt.isoformat()
    w7_lo = (snap_dt - timedelta(days=6)).isoformat()
    w30_lo = (snap_dt - timedelta(days=29)).isoformat()
    # 日销表「当月总销量」窗口 = 当月1号 ~ 锚点-1（分母 DAY(锚点-1)，分子必须同宽）。
    # 月初分支（已过天数<8）改用 (已过+7) 天滚动窗，起点 = 锚点-(DAY(锚点)+6)。
    m_lo = snap_dt.replace(day=1).isoformat()
    m_end = (snap_dt - timedelta(days=1)).isoformat()
    roll_lo = (snap_dt - timedelta(days=snap_dt.day + 6)).isoformat()
    fmt = dict(d=snap_day, d4=snap_day[4:], m=C.ORDER_MONTH)

    ages = {}
    for p in age_files:
        ages.update(read_age(p))

    # 库龄总表（第二来源，2026-11-03）：CSV 报告挂不上的 interflow / eMAG 从这里补。
    # 键是（账号, 站点, sku），与 ages 的（marketplace代码, sku）分开，互不覆盖；
    # 优先级：主账号 CSV 报告在 → 用 CSV（同一份原始数据的另一种封装，时点更贴快照），
    # CSV 缺失或该账号本来就不在 CSV 覆盖范围 → 落到总表。
    _age_x = _find_age_xlsx()
    ages2, age_xlsx_per = read_age_xlsx(_age_x) if _age_x else ({}, None)

    rows, missing_rep, sales_only, warn = [], [], [], []

    # ---- 亚马逊系 ----
    for s in SITES:
        _reps = s["rep"] if isinstance(s["rep"], list) else [s["rep"]]
        rep = _find(snap_dir, _reps, fmt)
        if not rep:
            missing_rep.append(f'{s["acct"]}-{s["site"]}: '
                               f'{" / ".join(x.format(**fmt) for x in _reps)}')
            continue
        try:
            data, order = R.read_report(rep, s["kind"], sep=",",
                                        colmap=s.get("colmap"))
        except Exception as e:
            missing_rep.append(f'{s["acct"]}-{s["site"]}: 解析失败 {e}')
            continue
        if not data:
            missing_rep.append(f'{s["acct"]}-{s["site"]}: 报告为空/类型不符')
            continue

        # ---- 归并 amzn.gr.* 别名行（2026-11-03，修「库存 0 却喊补 2392 件」）----
        # 同一 ASIN 拆成「裸 SKU + amzn.gr.<SKU>-<hash> 变体行」时（欧盟 41 组 /
        # 英国 165 组），变体行经常 0 库存。不归并的后果：裸行有货判不出缺货，
        # 别名行 0 库存判成三级、按历史日均算出天价建议量 —— 一件东西两张嘴。
        # 归并 = 数量求和（可售/预留/在途三段）、行名用基础 SKU、成员留档供注记。
        gdata, gorder, gmem = {}, [], {}
        for sku in order:
            b = _norm_sku(sku)
            if b in gdata:
                m = gdata[b]
                for k in ("afn-fulfillable-quantity", "afn-reserved-quantity",
                          "afn-inbound-working-quantity", "afn-inbound-shipped-quantity",
                          "afn-inbound-receiving-quantity"):
                    m[k] = _int(m.get(k)) + _int(data[sku].get(k))
                gmem[b].append(sku)
            else:
                gdata[b] = dict(data[sku])
                gorder.append(b)
                gmem[b] = [sku]
        if len(gorder) < len(order):
            warn.append(f'[归并] {s["acct"]}-{s["site"]}: {len(order)-len(gorder)} 个 '
                        f'amzn.gr.* 别名行并入基础 SKU（库存求和，与月度块同规则）')
        data, order = gdata, gorder

        # 销量：候选报告**只取第一个存在的**。欧盟/英国是同一份报告的两次下载
        # （21862 行 vs 21828 行），若把候选全读一遍，英国与欧洲的销量都会翻倍。
        agg = defaultdict(lambda: [0, 0, 0, 0])       # [近7天, 近30天, 当月累计]
        nosales = False
        op = _find(snap_dir, s["orders"], fmt)
        if op:
            for d, sku, q, ch, cx in load_orders(op, "\t"):
                if ch not in s["chans"] or ch in s.get("drop", ()) or cx:
                    continue
                b = _norm_sku(sku)          # 订单里的 amzn.gr.* 别名也归到基础 SKU
                if w30_lo <= d <= snap_iso:
                    agg[b][1] += q
                if w7_lo <= d <= snap_iso:
                    agg[b][0] += q
                if m_lo <= d <= m_end:
                    agg[b][2] += q
                if roll_lo <= d <= m_end:
                    agg[b][3] += q
            # 匹配率闸：订单报告与库存报告不是同一账号时（interflow 北美实测 0%），
            # 该站点的销量整体不可用 —— 用它算日均会把可售天数算成假数。
            hit = len(set(agg) & set(data))
            rate = hit / len(agg) if agg else 1.0
            if agg and rate < 0.6:
                nosales = True
                warn.append(f'{s["acct"]}-{s["site"]}: 订单报告「{os.path.basename(op)}」'
                            f'与库存 SKU 匹配率仅 {rate:.0%}（{hit}/{len(agg)}），'
                            f'判定为不同账号/站点 → 销量已跳过')
                agg = defaultdict(lambda: [0, 0, 0, 0])
        else:
            nosales = True
            warn.append(f'{s["acct"]}-{s["site"]}: 缺订单报告，销量不可用')

        for sku in order:
            f = data[sku]
            avail = _int(f.get("afn-fulfillable-quantity"))
            inb = (_int(f.get("afn-inbound-working-quantity"))
                   + _int(f.get("afn-inbound-shipped-quantity"))
                   + _int(f.get("afn-inbound-receiving-quantity")))
            q7, q30, q_mtd, q_roll = agg.get(sku, [0, 0, 0, 0])
            # 预留库存：与可售**互斥**（仓 = 可售 + 预留，实测 G415 739=736+3）。
            # 日销表可售天数 =(F+E)/实时速度，所以这里也要带上。
            reserved = _int(f.get("afn-reserved-quantity"))
            # 库龄：按成员 SKU 逐个查、档位求和 —— 裸行有库龄、别名行一般没有，
            # 合并后求和才不丢；两边都没查到才是 None。
            tiers = None
            _acct = "亚马逊" if s["acct"] == "主账号" else s["acct"]
            for msku in gmem[sku]:
                t = (ages.get((_mp_of(s["site"]), msku))
                     if s["acct"] == "主账号" else None)
                if t is None:
                    # 来源②（库龄总表）：按（账号, 站点, sku）精确挂，interflow 也覆盖
                    t = ages2.get((_acct, s["site"], msku))
                if t:
                    if tiers is None:
                        tiers = {lab: 0 for lab in AGE_TIER_LABEL}
                    for lab in AGE_TIER_LABEL:
                        tiers[lab] += _int(t.get(lab))
            mem = gmem[sku]
            rows.append(dict(site=s["site"], acct=s["acct"], sku=sku,
                             name=f.get("product-name"), avail=avail, inbound=inb,
                             q7=q7, q30=q30, q_mtd=q_mtd, q_roll=q_roll,
                             reserved=reserved,
                             tiers=tiers, src="亚马逊",
                             nosales=nosales,
                             # 归并留档：多于 1 个成员时判定依据里要注明（用户可见）
                             aliases=(mem if len(mem) > 1 else None)))
        # 有销量但库存报告里没有
        for sku, (q7, q30, q_mtd, q_roll) in agg.items():
            if sku not in data and q30 > 0:
                sales_only.append(dict(site=s["site"], acct=s["acct"], sku=sku,
                                       name=None, q7=q7, q30=q30,
                                       q_mtd=q_mtd, q_roll=q_roll))

    # ---- 美客多 ----
    p = _find(snap_dir, ["美客多墨西哥Full库存报告{d}.xlsx"], fmt)
    if p:
        ml = read_ml_resumen(p)
        for sku, d in ml.items():
            avail = _int(d["sellable"])
            inb = _int(d["pending"]) + _int(d["transfer"])
            sold30 = _int(d["sold30"])
            rows.append(dict(site="墨西哥", acct="美客多", sku=sku,
                             name=d["name"], avail=avail, inbound=inb,
                             # Resumen 只有 30 天汇总、没有逐日销量 → 当月累计不可得，
                             # 日销表口径（当月累计÷已过天数）对美客多退化为用近30天日均。
                             q7=None, q30=sold30, q_mtd=None, q_roll=None, reserved=0,
                             # 美客多不挂库龄总表：总表的美客多行来自「长期仓储费」报告，
                             # 与 Resumen「陈旧库存」（ml_stale）口径不同，混用会打架
                             tiers=None, src="美客多",
                             ml_stale=_int(d["stale"]), ml_cover=d["cover_w"],
                             ml_urgency=d["urgency"], ml_status=d["status"]))
    else:
        missing_rep.append("美客多: 美客多墨西哥Full库存报告{d}.xlsx".format(**fmt))

    # ---- eMAG ----
    p = _find(snap_dir, ["EMAG罗马尼亚库存报告{d}.xlsx"], fmt)
    if p:
        em = read_emag_stock(p)
        agg = defaultdict(lambda: [0, 0, 0, 0])
        op = _find(snap_dir, ["EMAG全站点{m}订单报告.xlsx"], fmt)
        if op:
            for dt, code, q, ch, cx in load_emag_orders(op):
                if cx:
                    continue
                if w30_lo <= dt <= snap_iso:
                    agg[code][1] += q
                if w7_lo <= dt <= snap_iso:
                    agg[code][0] += q
                if m_lo <= dt <= m_end:
                    agg[code][2] += q
                if roll_lo <= dt <= m_end:
                    agg[code][3] += q
        for sku, d in em.items():
            q7, q30, q_mtd, q_roll = agg.get(sku, [0, 0, 0, 0])
            rows.append(dict(site="罗马尼亚", acct="eMAG", sku=sku, name=d["name"],
                             avail=d["sellable"], inbound=0, q7=q7, q30=q30,
                             q_mtd=q_mtd, q_roll=q_roll, reserved=0,
                             # eMAG 库存报告没有库龄字段，只能从库龄总表挂（2026-11-03）
                             tiers=ages2.get(("eMAG", "罗马尼亚", sku)) or None,
                             src="eMAG"))
    else:
        missing_rep.append("罗马尼亚: EMAG罗马尼亚库存报告{d}.xlsx".format(**fmt))

    # ---- 分级 ----
    for r in rows:
        d30 = r["q30"] / 30.0
        r["daily"] = d30
        if r["avail"] <= 0:
            r["dos"] = 0.0
        elif d30 > 0:
            r["dos"] = r["avail"] / d30
        else:
            r["dos"] = None                      # 不动销 / 销量源缺失
        t = r.get("tiers") or {}
        r["aged90"] = sum(t.get(k, 0) for k in ("91-180", "181-270", "271-365", "365+"))
        r["aged_total"] = sum(t.values()) if t else 0
        r["ml_stale"] = r.get("ml_stale", 0)
        r["stale"] = r["aged90"] or r["ml_stale"]     # 统一呆滞口径
        ns = r.get("nosales", False)

        # 计划备货天数（需求B 同源）—— 站点查不到则 None，回退旧口径
        pdays = plan_of(r["site"], r["acct"], plan_map) if use_plan else None
        r["plan_days"] = pdays

        if r["avail"] <= 0 and r["inbound"] <= 0:
            r["state"] = "严重断货"
        elif r["avail"] <= 0:
            r["state"] = "一般断货"
        elif not ns and r["dos"] is not None and (
                (pdays and r["dos"] + lead_days < pdays)
                or (pdays is None and r["dos"] < lead_days)):
            r["state"] = "预备断货"
        else:
            r["state"] = "正常"

        # 原因
        if ns and r["state"] in ("严重断货", "一般断货"):
            inv_txt = "无在途" if r["inbound"] <= 0 else "在途%d件" % r["inbound"]
            why = "可售为0，%s；销量源缺失，无法判断动销" % inv_txt
        elif ns:
            why = "销量源缺失，无法计算可售天数（需人工核对）"
        elif r["state"] == "严重断货":
            why = "可售为0，无在途" + ("；近30天有销量" if r["q30"] > 0 else "；近30天无销量（疑似已下架/弃售）")
        elif r["state"] == "一般断货":
            cover = (r["inbound"] / r["daily"]) if r["daily"] > 0 else None
            why = f"可售为0，在途{r['inbound']}件"
            why += f"，按日均{r['daily']:.1f}可撑{cover:.0f}天" if cover else "，无动销参照"
        elif r["state"] == "预备断货":
            if pdays:
                why = (f"可售{r['avail']}件，日均{r['daily']:.1f}，仅够{r['dos']:.1f}天；"
                       f"加后勤周期{lead_days}天仍 < 计划备货{pdays}天")
            else:
                why = f"可售{r['avail']}件，日均{r['daily']:.1f}，仅够{r['dos']:.1f}天 < 后勤周期{lead_days}天"
        else:
            why = ""
        if r["stale"] > 0:
            if r["avail"] <= 0:
                why += f"；另有{r['stale']}件超90天但不可售，需核查能否转可售"
            else:
                why += f"；有{r['stale']}件超90天呆滞"
        r["why"] = why

        # 采购方案：目标覆盖天数与判定口径一致（有计划备货天数就按它，否则 --cover）
        tgt_days = pdays if pdays else cover_days
        r["tgt_days"] = tgt_days
        target = math.ceil(r["daily"] * tgt_days) if r["daily"] > 0 else 0
        gap = target - r["avail"] - r["inbound"]
        r["target"] = target
        if r["state"] == "正常":
            r["buy"] = 0
            r["plan"] = "—" if not ns else "销量源缺失，暂不判定"
        elif ns:
            r["buy"] = 0
            r["plan"] = "先补销量数据再定补货量"
        elif r["q30"] == 0 and r["avail"] <= 0:
            r["buy"] = 0
            if r["inbound"] > 0:
                r["plan"] = f"不追加补货：近30天无动销，在途{r['inbound']}件到货后观察"
            else:
                r["plan"] = "不补货：无动销，先确认是否下架/弃售"
        elif r["avail"] <= 0 and r["stale"] > 0:
            # 断货但库里有呆滞件 —— 货卡在不可售状态，先核查再谈补货
            r["buy"] = max(0, gap)
            r["plan"] = (f"先核查{r['stale']}件超90天不可售库存能否转可售；"
                         + (f"仍需补货 {r['buy']} 件" if r["buy"] else "可覆盖则无需补货"))
        elif r["stale"] > 0 and r["daily"] < 0.5:
            r["buy"] = 0
            r["plan"] = f"暂停补货：{r['stale']}件超90天呆滞，先清库存"
        else:
            r["buy"] = max(0, gap)
            if r["buy"] == 0:
                r["plan"] = "在途已覆盖，等到货"
            elif r["state"] == "严重断货":
                r["plan"] = f"立即补货 {r['buy']} 件"
            elif r["state"] == "一般断货":
                r["plan"] = f"补货 {r['buy']} 件（在途不足）"
            elif pdays:
                r["plan"] = f"排入下周三补货 {r['buy']} 件（按计划备货{pdays}天）"
            else:
                r["plan"] = f"排入下周三补货 {r['buy']} 件"
            if r["stale"] > 0:
                r["plan"] += f"；同步清{r['stale']}件呆滞"
        # 优先级：P0=严重断货且近30天真在卖（正在丢单）；P1=断货且有动销；P2=预备断货；P3=无动销断货
        if r["state"] == "正常":
            r["prio"] = "—"
        elif r["state"] == "严重断货" and r["q30"] >= 5:
            r["prio"] = "P0"
        elif r["state"] in ("严重断货", "一般断货") and r["q30"] >= 1:
            r["prio"] = "P1"
        elif r["state"] in ("严重断货", "一般断货"):
            r["prio"] = "P3"
        else:
            r["prio"] = "P2"

        if moq > 1 and r["buy"] > 0:
            r["buy"] = math.ceil(r["buy"] / moq) * moq
            r["plan"] = r["plan"].split("；")[0] + f"（按 {moq} 件/箱取整）"

    rows.sort(key=lambda r: ({"严重断货": 0, "一般断货": 1, "预备断货": 2, "正常": 3}[r["state"]],
                             -r["q30"], r["site"], r["acct"], r["sku"]))
    sales_only.sort(key=lambda r: -r["q30"])
    return rows, sales_only, missing_rep, warn, snap_dt, (w7_lo, w30_lo)


def _mp_of(site):
    for mp, s in AGE_MP.items():
        if s == site:
            return mp
    return ""


# ---------------- 输出 ----------------
HDR_FILL = PatternFill("solid", fgColor="1F4E79")
HDR_FONT = Font(color="FFFFFF", bold=True, size=10)
FILL = {"严重断货": PatternFill("solid", fgColor="FFC7CE"),
        "一般断货": PatternFill("solid", fgColor="FFEB9C"),
        "预备断货": PatternFill("solid", fgColor="DDEBF7")}

COLS = ["地区", "账号", "SKU", "品名", "可售", "在途", "近7天", "近30天", "日均",
        "可售天数"] + AGE_TIER_LABEL + ["呆滞>90天", "优先级", "状态", "原因", "采购方案"]
KEYS = ["site", "acct", "sku", "name", "avail", "inbound", "q7", "q30", "daily",
        "dos"] + [f"tier_{i}" for i in range(7)] + ["aged90", "prio", "state", "why", "plan"]
WIDTH = [9, 10, 10, 42, 7, 7, 7, 8, 7, 9] + [7] * 7 + [9, 7, 10, 58, 34]


def _val(r, k):
    if k.startswith("tier_"):
        i = int(k[5:])
        t = r.get("tiers") or {}
        if not t:
            return None
        return t.get(AGE_TIER_LABEL[i], 0)
    v = r.get(k)
    if k == "daily":
        return round(v, 2) if v else 0
    if k == "dos":
        return None if v is None else round(v, 1)
    return v


def write_sheet(wb, title, rows, note=None):
    ws = wb.create_sheet(title)
    r0 = 1
    if note:
        ws.cell(row=1, column=1).value = note
        ws.cell(row=1, column=1).font = Font(italic=True, size=9, color="666666")
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(COLS))
        r0 = 2
    for j, (c, w) in enumerate(zip(COLS, WIDTH), 1):
        cell = ws.cell(row=r0, column=j, value=c)
        cell.fill, cell.font = HDR_FILL, HDR_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.column_dimensions[get_column_letter(j)].width = w
    for i, r in enumerate(rows, r0 + 1):
        for j, k in enumerate(KEYS, 1):
            v = _val(r, k)
            cell = ws.cell(row=i, column=j, value=v)
            cell.font = Font(size=9)
            if k in ("avail", "inbound", "q7", "q30", "aged90") or k.startswith("tier_"):
                cell.alignment = Alignment(horizontal="center")
            if k == "name":
                cell.alignment = Alignment(wrap_text=False)
            if k in ("why", "plan"):
                cell.alignment = Alignment(wrap_text=True, vertical="top")
        f = FILL.get(r["state"])
        if f:
            for j in (1, 2, 3, len(COLS) - 1):
                ws.cell(row=i, column=j).fill = f
    ws.freeze_panes = ws.cell(row=r0 + 1, column=4)
    ws.auto_filter.ref = f"A{r0}:{get_column_letter(len(COLS))}{r0 + len(rows)}"
    return ws


def write_xlsx(path, rows, sales_only, summary, meta):
    wb = Workbook()
    wb.remove(wb.active)

    # 汇总
    ws = wb.create_sheet("汇总")
    ws["A1"] = f"断货分析汇总 · 快照 {meta['snap']}"
    ws["A1"].font = Font(bold=True, size=13)
    ws["A2"] = (f"后勤周期 {meta['lead']} 天（每周三进货）｜目标覆盖 {meta['cover']} 天｜"
                f"窗口 近7天={meta['w7']}起，近30天={meta['w30']}起")
    ws["A2"].font = Font(size=9, color="666666")
    hdr = ["地区", "账号", "SKU总数", "可售件", "在途件", "P0紧急", "严重断货",
           "一般断货", "预备断货", "呆滞>90天SKU", "待补货件数"]
    for j, h in enumerate(hdr, 1):
        c = ws.cell(row=4, column=j, value=h)
        c.fill, c.font = HDR_FILL, HDR_FONT
        c.alignment = Alignment(horizontal="center")
        ws.column_dimensions[get_column_letter(j)].width = [10, 11, 9, 9, 9, 9, 10, 10, 10, 13, 11][j - 1]
    for i, s in enumerate(summary, 5):
        for j, v in enumerate([s["site"], s["acct"], s["total"], s["avail"], s["inbound"],
                               s["p0"], s["严重断货"], s["一般断货"], s["预备断货"],
                               s["aged"], s["buy"]], 1):
            c = ws.cell(row=i, column=j, value=v)
            c.font = Font(size=9)
            if j >= 3:
                c.alignment = Alignment(horizontal="center")
    tr = 5 + len(summary)
    for j, v in enumerate(["合计", "", sum(s["total"] for s in summary),
                           sum(s["avail"] for s in summary),
                           sum(s["inbound"] for s in summary),
                           sum(s["p0"] for s in summary),
                           sum(s["严重断货"] for s in summary),
                           sum(s["一般断货"] for s in summary),
                           sum(s["预备断货"] for s in summary),
                           sum(s["aged"] for s in summary),
                           sum(s["buy"] for s in summary)], 1):
        c = ws.cell(row=tr, column=j, value=v)
        c.font = Font(size=9, bold=True)
        if j >= 3:
            c.alignment = Alignment(horizontal="center")

    write_sheet(wb, "严重断货", [r for r in rows if r["state"] == "严重断货"],
                "可售=0 且 在途=0 —— 前台已无货可卖，需立即处理")
    write_sheet(wb, "一般断货", [r for r in rows if r["state"] == "一般断货"],
                "可售=0 但 在途>0 —— 已断货，货在路上")
    write_sheet(wb, "预备断货", [r for r in rows if r["state"] == "预备断货"],
                f"可售>0 但 可售天数 < 后勤周期 {meta['lead']} 天 —— 下周三前会断")
    write_sheet(wb, "库龄呆滞", [r for r in rows if r["stale"] > 0],
                "超90天呆滞件>0（亚马逊取库龄报告分档；美客多取 Resumen「陈旧库存」）")
    buy_rows = sorted([r for r in rows if r["buy"] > 0],
                      key=lambda r: (r["prio"], -r["buy"], -r["q30"]))
    # ⚠️ 目标覆盖天数不是固定的 --cover：build() 里是 `pdays if pdays else cover_days`，
    # pdays = 该站点在 备货天数.xlsx 里的计划备货天数（美客多 60、欧洲 120…）。
    # 2026-11-03 修：原文案写死「ceil(日均×14)」，与实际算出的件数对不上（如美客多 156=2.6×60）。
    write_sheet(wb, "采购建议", buy_rows,
                f"建议补货量 = ceil(日均 × 目标覆盖天数) − 可售 − 在途；"
                f"目标覆盖天数取各站点「计划备货天数」"
                f"（--no-plan 或该站点查不到时用 {meta['cover']} 天）；呆滞品优先清仓")

    if sales_only:
        ws = wb.create_sheet("无库存记录")
        ws["A1"] = "近30天有销量、但库存报告里查无此 SKU —— 疑似已下架/换 SKU/账号不符"
        ws["A1"].font = Font(italic=True, size=9, color="666666")
        for j, h in enumerate(["地区", "账号", "SKU", "近7天", "近30天"], 1):
            c = ws.cell(row=2, column=j, value=h)
            c.fill, c.font = HDR_FILL, HDR_FONT
        for i, r in enumerate(sales_only, 3):
            for j, v in enumerate([r["site"], r["acct"], r["sku"], r["q7"], r["q30"]], 1):
                ws.cell(row=i, column=j, value=v).font = Font(size=9)
        for j, w in enumerate([10, 11, 11, 8, 8], 1):
            ws.column_dimensions[get_column_letter(j)].width = w

    wb.save(path)


def write_md(path, rows, sales_only, summary, meta):
    L = []
    A = L.append
    A(f"# 断货分析 · {meta['snap']}\n")
    A(f"> 后勤周期 **{meta['lead']} 天**（每周三进货）｜目标覆盖 **{meta['cover']} 天**｜"
      f"近7天窗口 {meta['w7']} 起、近30天窗口 {meta['w30']} 起\n")
    cnt = {k: sum(1 for r in rows if r["state"] == k) for k in ("严重断货", "一般断货", "预备断货")}
    A(f"## 一、总览\n")
    A(f"- 在册 SKU **{len(rows)}** 个：严重断货 **{cnt['严重断货']}**、"
      f"一般断货 **{cnt['一般断货']}**、预备断货 **{cnt['预备断货']}**")
    A(f"- 待补货件数合计 **{sum(r['buy'] for r in rows)}** 件")
    A(f"- 超 90 天呆滞 SKU **{sum(1 for r in rows if r['stale'] > 0)}** 个\n")

    A("### 按地区\n")
    A("| 地区 | 账号 | SKU | 可售件 | 在途件 | P0 | 严重 | 一般 | 预备 | 呆滞 | 待补货 |")
    A("|---|---|---:|---:|---:|---:|---:|---:|---:|---:|---:|")
    for s in summary:
        A(f"| {s['site']} | {s['acct']} | {s['total']} | {s['avail']} | {s['inbound']} | "
          f"{s['p0']} | {s['严重断货']} | {s['一般断货']} | {s['预备断货']} | "
          f"{s['aged']} | {s['buy']} |")
    A("")

    def block(title, sub, rs, n=40):
        A(f"## {title}\n")
        if sub:
            A(f"> {sub}\n")
        if not rs:
            A("（无）\n")
            return
        A("| 优先级 | 地区 | 账号 | SKU | 可售 | 在途 | 近30天 | 日均 | 可售天数 | 呆滞>90天 | 原因 | 采购方案 |")
        A("|---|---|---|---|---:|---:|---:|---:|---:|---:|---|---|")
        for r in rs[:n]:
            dos = "—" if r["dos"] is None else f"{r['dos']:.1f}"
            A(f"| {r['prio']} | {r['site']} | {r['acct']} | {r['sku']} | {r['avail']} | {r['inbound']} | "
              f"{r['q30']} | {r['daily']:.1f} | {dos} | {r['stale']} | "
              f"{r['why']} | {r['plan']} |")
        if len(rs) > n:
            A(f"\n…另有 {len(rs) - n} 行，见 xlsx\n")
        A("")

    block("二、严重断货（可售 0 / 在途 0）", "前台已无货可卖，需立即处理", 
          [r for r in rows if r["state"] == "严重断货"])
    block("三、一般断货（可售 0 / 在途 >0）", "已断货，货在路上",
          [r for r in rows if r["state"] == "一般断货"])
    block("四、预备断货（可售天数 < 后勤周期）", "下周三前会断，排入补货计划",
          [r for r in rows if r["state"] == "预备断货"])

    A("## 五、库龄呆滞\n")
    aged = sorted([r for r in rows if r["stale"] > 0],
                  key=lambda r: -r["stale"])
    if not aged:
        A("（无）\n")
    else:
        A("| 地区 | 账号 | SKU | 呆滞件 | 365+ | 可售 | 近30天 | 建议 |")
        A("|---|---|---|---:|---:|---:|---:|---|")
        for r in aged[:30]:
            t = r.get("tiers") or {}
            A(f"| {r['site']} | {r['acct']} | {r['sku']} | {r['stale']} | "
              f"{t.get('365+', '—')} | {r['avail']} | {r['q30']} | "
              f"{'清仓，暂停补货' if r['daily'] < 0.5 else '补货同时清呆滞'} |")
        A("")

    if sales_only:
        A("## 六、异常：近30天有销量但库存报告无此 SKU\n")
        A("| 地区 | 账号 | SKU | 近7天 | 近30天 |")
        A("|---|---|---|---:|---:|")
        for r in sales_only[:20]:
            A(f"| {r['site']} | {r['acct']} | {r['sku']} | {r['q7']} | {r['q30']} |")
        A("")

    A("## 七、口径说明\n")
    A("- 可售(亚马逊)=`afn-fulfillable-quantity`；在途=`inbound-working+shipped+receiving`")
    A("- 可售(美客多)=Resumen `可销售`；在途=`待录入+转账中`；日均取 `最近30天售出件数/30`")
    A("- 可售(eMAG)=`stock_type=ready_to_sell` 数量之和；无在途数据")
    A("- 有效销量=订单状态非 Cancelled/Storno；日期按 `purchase-date` 前 10 位（当地时间）")
    if meta["missing"]:
        A("")
        A("### 缺报告\n")
        for m in meta["missing"]:
            A(f"- {m}")
    if meta.get("warn"):
        _merge = [m for m in meta["warn"] if m.startswith("[归并]")]
        _anom = [m for m in meta["warn"] if not m.startswith("[归并]")]
        if _anom:
            A("")
            A("### 数据异常（销量已跳过）\n")
            for m in _anom:
                A(f"- ⚠️ {m}")
        if _merge:
            A("")
            A("### 别名归并（amzn.gr.* → 基础 SKU，正常动作）\n")
            for m in _merge:
                A(f"- {m}")
    if meta.get("dups"):
        A("")
        A("### 重复下载（同内容不同文件名，只取一份）\n")
        for keep, dup in meta["dups"]:
            A(f"- `{dup}` 与 `{keep}` 内容完全相同")
    if meta["age_files"]:
        A("")
        A(f"### 库龄数据源\n")
        for f in meta["age_files"]:
            A(f"- `{os.path.basename(f)}`")
    else:
        A("")
        A("> ⚠️ 本次无库龄报告，库龄列全空。把 `*库龄*.csv` 放进快照目录即可自动识别。")

    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(L))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--snap", default=None)
    ap.add_argument("--age", action="append", default=[])
    ap.add_argument("--lead", type=int, default=7)
    ap.add_argument("--cover", type=int, default=14)
    ap.add_argument("--moq", type=int, default=0, help="最小补货批量（件/箱），0=不取整")
    ap.add_argument("--no-plan", action="store_true",
                    help="回退旧口径：预备断货按「可售天数 < 后勤周期」判，补货目标用 --cover")
    ap.add_argument("--out", default=None)
    a = ap.parse_args()

    if a.snap:
        snap_dir = os.path.abspath(a.snap)
        snap_day = os.path.basename(snap_dir.rstrip("\\/"))
    else:
        snap_dir, snap_day = C.latest_snapshot()
    snap_day = re.sub(r"\D", "", snap_day)[:8]
    if not snap_day:
        sys.exit("无法确定快照日")

    age_files = [os.path.abspath(x) for x in a.age if os.path.isfile(x)]
    age_files += [p for p in _autodetect_age(*_age_dirs(snap_dir)) if p not in age_files]

    print(f"快照: {snap_dir}")
    print(f"库龄报告: {len(age_files)} 份" + ("".join(f"\n  - {os.path.basename(x)}" for x in age_files)))
    _ax = _find_age_xlsx()
    if _ax:
        _, _ax_per = read_age_xlsx(_ax)
        print(f"库龄总表: {os.path.basename(_ax)}"
              + (f"（期数 {_ax_per}，CSV 缺失时用它挂库龄，interflow/eMAG 只此一源）" if _ax_per else ""))
    else:
        print("库龄总表: 未找到（trunk/data/库龄报表/ 下没有《亚马逊EMAG库龄情况总表*.xlsx》）")
    dups = dup_order_reports(snap_dir)
    rows, sales_only, missing, warn, snap_dt, (w7, w30) = build(
        snap_dir, snap_day, age_files, a.lead, a.cover, a.moq,
        use_plan=not a.no_plan)
    if not a.no_plan:
        print(f"计划备货天数: 读 {PLAN_XLSX}"
              f"{'' if os.path.exists(PLAN_XLSX) else '（文件不存在，用内置默认值）'}")

    # 按地区汇总
    summary, seen = [], {}
    for r in rows:
        k = (r["site"], r["acct"])
        if k not in seen:
            seen[k] = dict(site=r["site"], acct=r["acct"], total=0, 严重断货=0,
                           一般断货=0, 预备断货=0, aged=0, buy=0, p0=0,
                           avail=0, inbound=0)
            summary.append(seen[k])
        s = seen[k]
        s["total"] += 1
        if r["state"] in ("严重断货", "一般断货", "预备断货"):
            s[r["state"]] += 1
        if r["stale"] > 0:
            s["aged"] += 1
        if r.get("prio") == "P0":
            s["p0"] += 1
        s["avail"] += r["avail"]
        s["inbound"] += r["inbound"]
        s["buy"] += r["buy"]

    # 断货相关产出统一收在 断货分析/ 下，一级按业务块分（亚马逊 / 东南亚）
    out = a.out or os.path.join(C.DSR, "断货分析", "亚马逊", "断货明细", snap_day)
    os.makedirs(out, exist_ok=True)
    meta = dict(snap=snap_day, lead=a.lead, cover=a.cover, w7=w7, w30=w30,
                missing=missing, age_files=age_files, warn=warn, dups=dups)
    xp = os.path.join(out, f"断货明细_{snap_day}.xlsx")
    mp = os.path.join(out, f"断货报告_{snap_day}.md")
    write_xlsx(xp, rows, sales_only, summary, meta)
    write_md(mp, rows, sales_only, summary, meta)

    print(f"\n在册 SKU {len(rows)}｜严重 {sum(1 for r in rows if r['state']=='严重断货')}"
          f"｜一般 {sum(1 for r in rows if r['state']=='一般断货')}"
          f"｜预备 {sum(1 for r in rows if r['state']=='预备断货')}"
          f"｜呆滞 {sum(1 for r in rows if r['stale']>0)}")
    print(f"待补货 {sum(r['buy'] for r in rows)} 件")
    if missing:
        print("\n[缺报告]")
        for m in missing:
            print("  -", m)
    if warn:
        _merge = [m for m in warn if m.startswith("[归并]")]
        _anom = [m for m in warn if not m.startswith("[归并]")]
        if _anom:
            print("\n[数据异常 · 销量已跳过]")
            for m in _anom:
                print("  -", m)
        if _merge:
            # 归并是正常动作，不是异常，单独归类免得误导看日志的人
            print("\n[别名归并]")
            for m in _merge:
                print("  -", m)
    if dups:
        print("\n[重复下载]")
        for keep, dup in dups:
            print(f"  - {dup}  ==  {keep}")
    if sales_only:
        print(f"\n[异常] 近30天有销量但无库存记录: {len(sales_only)} 个 SKU")
    print(f"\n已输出:\n  {xp}\n  {mp}")


if __name__ == "__main__":
    main()
