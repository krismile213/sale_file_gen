# -*- coding: utf-8 -*-
"""亚马逊 + EMAG 库龄情况表（2026-10-22 需求，11-10 前交付）。

数据源（用户给的导出路径）：
  亚马逊  后台 → 供应链 → 供应链报告 → 库存 → 显示更多 → 亚马逊物流库存
          ⚠ 真实报告是 **6 个基础档 + 一组细分档**，不是简单的「七档」：
            基础 inv-age-0-to-90 / 91-to-180 / 181-to-270 / 271-to-365 /
                 **366-to-455** / **456-plus**
            细分 inv-age-0-to-30 / 31-to-60 / 61-to-90（0-90 拆三份）
                 inv-age-181-to-330 / 331-to-365（181-365 拆两份）
            最后一档两种写法并存（2026-10 实测）：新格式 `inv-age-366-to-455-days`
            + `inv-age-456-plus-days`（美国/英国/欧盟），老格式 `inv-age-365-plus-days`
            （印度/加拿大/沙特/墨西哥/阿联酋）。原代码只写了 365-plus → 新格式站点
            第 7 档恒为 0，「180 天以上」漏掉全部 366 天以上库存且不报错（2026-10-29 修正，
            详见 TIER_SRC / TIER_ALT 注释）
          ※ 日本/澳大利亚导出的是中文版「库龄」报告（表头 `库龄 0-90 天`…`库龄 365 天以上`
            + 同结构细分档），也支持；站点从文件名取（该版没有 marketplace 列）
  EMAG    Fulfilment by eMag → FBE Inventory（状态=全部、时效原因=全部）→ Export

账号维度：亚马逊 / interflow 是两套卖家账号，SKU 会撞号（如沙特 G596 两边都有），
输出按（账号, 站点）分组，不混行。账号从文件名前缀识别。

输出：`亚马逊EMAG库龄情况总表.xlsx`（**总表，逐期累积**，不是每期一个新文件）
  sheet「本期明细」  站点 / SKU / 品名 / 七档件数 / 合计 / 180天以上 / 180+占比
  sheet「站点汇总」  每站点一行：SKU 数 / 总件数 / 各档 / 180+ 占比 / 呆滞等级
  sheet「历史汇总」  每期每站点一行，追加写入（跨期看库龄是否在恶化）

用法：
    python trunk/age_report.py                     # 自动找「最新有库龄报告」的周期
    python trunk/age_report.py --age a.csv --age b.csv
    python trunk/age_report.py --out <目录>
    python trunk/age_report.py --period 202610
"""
import os
import re
import sys
import csv
import argparse
from collections import defaultdict, Counter
from datetime import datetime

_HERE = os.path.dirname(os.path.abspath(__file__))
_PROJ = os.path.dirname(_HERE)
_CORE = os.path.join(_PROJ, "Daily_Sales_Report", "pipeline", "core")
for _p in (_CORE, _HERE):
    if _p not in sys.path:
        sys.path.insert(0, _p)

from openpyxl import Workbook, load_workbook              # noqa: E402
from openpyxl.styles import Font, PatternFill, Alignment  # noqa: E402
from openpyxl.utils import get_column_letter              # noqa: E402

SNAP_BASE = os.path.join(_PROJ, "trunk", "data", "数据归属周期")
OUT_NAME = "亚马逊EMAG库龄情况总表.xlsx"

# 输出固定 7 档（对外口径，「历史汇总」按位置跨期累积，档数不能动）。
# 每档 = 报告里 1..N 个源列之和。
#
# ⚠ 2026-10-29 修（原代码静默算错，必须记住）：
#   原 TIERS 写的是 inv-age-0-to-30 / 31-to-60 / 61-to-90 / 91-to-180 /
#   181-to-270 / 271-to-365 / **inv-age-365-plus-days**。
#   但新版「亚马逊物流库存」报告里**没有 365-plus 这一列**，超 365 天被拆成
#   `inv-age-366-to-455-days` + `inv-age-456-plus-days` 两列 →
#   第 7 档恒为 0，而「180 天以上」= 第 5..7 档之和，等于**整整漏掉 366 天以上的库存**，
#   且 row.get() 取不到列只会返回 None → _int(None) = 0，**一句报错都没有**。
#   实测 202610：美国站 366 天以上有货，旧口径下 180+ 占比被低估。
#   现在改成「一档对多列求和」，并加 _missing_cols() 在缺列时打印告警，
#   以后亚马逊再改列名不会再静默归零。
TIER_LABEL = ["0-30", "31-60", "61-90", "91-180", "181-270", "271-365", "366+"]
TIER_SRC = [["inv-age-0-to-30-days"],
            ["inv-age-31-to-60-days"],
            ["inv-age-61-to-90-days"],
            ["inv-age-91-to-180-days"],
            ["inv-age-181-to-270-days"],
            ["inv-age-271-to-365-days"],
            ["inv-age-366-to-455-days", "inv-age-456-plus-days"]]
# 第 7 档的**老格式备选**：202610 实测同一个后台导出的 16 份报告里两种格式并存 ——
#   美国/英国/欧盟/interflow美国：366-to-455 + 456-plus（新）
#   印度/加拿大/沙特/墨西哥/阿联酋：365-plus（老）
# 所以按「主列有就用主列，主列全缺才用备选」取，两者绝不会同时命中，不会重复计数。
TIER_ALT = {6: ["inv-age-365-plus-days"]}
NTIER = len(TIER_LABEL)
AGED_FROM = 4          # 181 天以上算呆滞（索引 4 起）

# marketplace 代码 → 站点名（与 stockout_report.AGE_MP 一致；UK 是英国报告的写法）
AGE_MP = {"US": "美国", "CA": "加拿大", "MX": "墨西哥", "GB": "英国", "UK": "英国",
          "JP": "日本",
          "SA": "沙特阿拉伯", "AU": "澳大利亚", "IN": "印度", "AE": "阿联酋",
          "DE": "欧洲", "FR": "欧洲", "IT": "欧洲", "ES": "欧洲", "NL": "欧洲",
          "SE": "欧洲", "PL": "欧洲", "IE": "欧洲", "BE": "欧洲", "TR": "欧洲"}

# 文件名里的站点写法 → 标准站点名（中文版报告没有 marketplace 列，只能靠文件名）
_SITE_ALIAS = {"阿拉伯联合酋长国": "阿联酋", "欧盟": "欧洲", "沙特": "沙特阿拉伯"}

# 中文版「库龄」报告的 7 档（与英文 TIER_SRC 一一对应）。
# 注意中文版也有「库龄 0-90 天」这个合并档，但既然 0-30/31-60/61-90 三列都在，就用拆分的；
# 万一某次导出只剩合并档，read_cn_age 会退化到把 0-90 全塞进第 1 档并告警。
CN_TIER_SRC = [["库龄 0-30 天"], ["库龄 31-60 天"], ["库龄 61-90 天"],
               ["库龄 91-180 天"], ["库龄 181-270 天"], ["库龄 271-365 天"],
               ["库龄 365 天以上"]]
CN_TIER_ALT = {6: ["库龄 366 天以上"]}

_HDR_FILL = PatternFill("solid", fgColor="1F4E79")
_HDR_FONT = Font(color="FFFFFF", bold=True, size=10)


def _int(v):
    try:
        return int(float(str(v).replace(",", "").strip()))
    except (TypeError, ValueError):
        return 0


def _sum_spec(row, spec):
    """按档位定义求和：一档可能对应多列（如 366+ = 366-455 + 456-plus）。"""
    return sum(_int(row.get(c)) for c in spec)


def _missing_cols(have, spec_list):
    """档位定义里**报告里不存在**的列名（去重、保序）。空列表 = 全对得上。"""
    miss, seen = [], set()
    for spec in spec_list:
        for c in spec:
            if c not in have and c not in seen:
                seen.add(c)
                miss.append(c)
    return miss


def _warn_missing(fn, miss, hint=""):
    if miss:
        print(f"⚠️ {fn}：报告里没有这些列 → {miss}，对应档位按 0 计。{hint}")


def _resolve_spec(have, spec_list, alt_map=None):
    """按实际表头定档位源列：主列有就用主列；主列全缺且有备选列就用备选列。"""
    spec = [list(s) for s in spec_list]
    for i, alt in (alt_map or {}).items():
        if not any(c in have for c in spec[i]) and any(c in have for c in alt):
            spec[i] = list(alt)
    return spec


def _tiers_of(row, spec_list, merged_col=None, split_n=0):
    """按 spec_list 算各档件数。

    merged_col/split_n 是**老格式兜底**：如果报告只剩「0-90 天」合并档、没有 0-30/31-60/61-90
    三个拆分列（split_n=3），就把合并档整块放进第 1 档，避免总量凭空少掉一大块。
    """
    tiers = [_sum_spec(row, spec) for spec in spec_list]
    if merged_col and split_n:
        if not any(row.get(c) not in (None, "") for spec in spec_list[:split_n] for c in spec):
            tiers[0] = _int(row.get(merged_col))
    return tiers


def _period_of(snap_dir):
    parts = os.path.normpath(str(snap_dir)).split(os.sep)
    for i, p in enumerate(parts):
        if p == "数据归属周期" and i + 1 < len(parts):
            return parts[i + 1]
    return None


AGE_EXTS = (".csv", ".xlsx", ".xlsm")


def _is_age_file(p):
    """按**内容**判断是不是库龄报告，不认文件名。
    csv：看首行表头；xlsx：看有没有 FBE 表或「库龄天数」表头（2026-10-27 新增）。"""
    ext = os.path.splitext(p)[1].lower()
    if ext == ".csv":
        try:
            with open(p, encoding="utf-8-sig", errors="ignore", newline="") as f:
                head = f.readline()
        except Exception:
            return False
        return ("inv-age-" in head or "库龄 0-30 天" in head
                or ("快照日期" in head and "库龄" in head))
    if ext in (".xlsx", ".xlsm"):
        return _xlsx_kind(p) is not None
    return False


def latest_age_period():
    """默认周期 = 「数据归属周期」下**有非空 库龄 目录**的最新一期。

    2026-10-29 修：原来 main() 里写死 `a.period or "202610"`。换到 11 月后
    不带 --period 直接跑，会去读 202610/库龄 的旧文件、把结果按 202610 写回 ——
    新导的库龄报告一份都不会被读到，而「本期明细」被旧数据覆盖后看起来还像跑成功了。
    改成自动探测后，放好文件直接跑即可；显式传 --period 仍然优先。
    """
    if not os.path.isdir(SNAP_BASE):
        return None
    for per in sorted(os.listdir(SNAP_BASE), reverse=True):
        d = os.path.join(SNAP_BASE, per, "库龄")
        if not os.path.isdir(d):
            continue
        for _root, _dirs, files in os.walk(d):
            if any(f.lower().endswith(AGE_EXTS) for f in files):
                return per
    return None


def find_age_files(period=None):
    """库龄报告候选：数据归属周期/<period>/库龄/ 下所有库龄报告（csv + xlsx，含子目录）。

    2026-10-27 扩：EMAG「FBE Inventory」和美客多「长期仓储费」都是 xlsx，
    原实现只收 .csv，这两份永远进不来。改为 csv/xlsx 一起收、按内容识别。

    2026-10-29 去掉「本期没文件就去别的期找」的兜底：原来指定期的目录为空时会扫所有周期，
    于是 `--period 202611` 读到的是 202610 的文件，还被按 202611 写进「历史汇总」——
    期号标错、而且一句报错都没有。现在只认指定期，没有就返回空（调用方去提示）。
    """
    hits = []
    if not period:
        return hits
    d = os.path.join(SNAP_BASE, period, "库龄")
    if os.path.isdir(d):
        for root, _s, files in os.walk(d):
            for fn in files:
                if fn.lower().endswith(AGE_EXTS):
                    hits.append(os.path.join(root, fn))
    ok = [p for p in hits if _is_age_file(p)]
    ok.sort(key=lambda x: os.path.getmtime(x))
    return ok


def _account_of(path):
    """文件名 → 账号。亚马逊/interflow 是两套卖家账号，SKU 会撞号，必须分开统计。"""
    fn = os.path.basename(str(path))
    if fn.startswith("interflow"):
        return "interflow"
    if "EMAG" in fn.upper():
        return "EMAG"
    return "亚马逊"


def _site_from_filename(path):
    """`亚马逊沙特阿拉伯库龄报告20261023.csv` → 沙特阿拉伯（中文版报告没有 marketplace 列）"""
    fn = os.path.basename(str(path))
    fn = re.sub(r"\.(csv|xlsx|xlsm)$", "", fn, flags=re.I)
    for pre in ("亚马逊", "interflow", "Interflow", "EMAG"):
        if fn.startswith(pre):
            fn = fn[len(pre):]
            break
    fn = re.sub(r"(库龄报告|库龄|age report|age_report|报告)[\s_\-]*\d*$", "", fn, flags=re.I)
    fn = fn.strip(" _-")
    return _SITE_ALIAS.get(fn, fn) or "未知"


def read_amazon_age(path):
    """亚马逊英文版「亚马逊物流库存」csv → [(账号, site, sku, name, snapshot_date, [七档件数])]"""
    out = []
    acc = _account_of(path)
    with open(path, encoding="utf-8-sig", newline="") as f:
        rd = csv.DictReader(f)
        have = set(rd.fieldnames or [])
        spec = _resolve_spec(have, TIER_SRC, TIER_ALT)
        miss = _missing_cols(have, spec)
        _warn_missing(os.path.basename(path), miss,
                      "可能是亚马逊又改了列名，请把新表头发我；"
                      "缺的档位会算成 0，180+ 占比会偏低。")
        for row in rd:
            sku = (row.get("sku") or "").strip()
            if not sku:
                continue
            mp = (row.get("marketplace") or "").strip().upper()
            site = AGE_MP.get(mp, mp or _site_from_filename(path))
            tiers = _tiers_of(row, spec,
                              merged_col="inv-age-0-to-90-days", split_n=3)
            out.append((acc, site, sku, (row.get("product-name") or "").strip(),
                        (row.get("snapshot-date") or "").strip()[:10], tiers))
    return out


def read_cn_age(path):
    """中文版「库龄」报告 → 同结构。表头是 `库龄 0-30 天` … `库龄 365 天以上`，
    没有 marketplace 列（一个站点一份），站点从文件名取。"""
    out = []
    acc = _account_of(path)
    site = _site_from_filename(path)
    with open(path, encoding="utf-8-sig", newline="") as f:
        rd = csv.DictReader(f)
        have = set(rd.fieldnames or [])
        spec = _resolve_spec(have, CN_TIER_SRC, CN_TIER_ALT)
        miss = _missing_cols(have, spec)
        _warn_missing(os.path.basename(path), miss,
                      "中文版也可能只剩「库龄 0-90 天」合并档；"
                      "缺的档位会算成 0，180+ 占比会偏低。")
        for row in rd:
            sku = (row.get("sku") or "").strip()
            if not sku:
                continue
            tiers = _tiers_of(row, spec, merged_col="库龄 0-90 天", split_n=3)
            out.append((acc, site, sku, (row.get("商品名称") or "").strip(),
                        (row.get("快照日期") or "").strip()[:10], tiers))
    return out


def _xlsx_rows(path, sheet_hint=None):
    """读 xlsx 全表（含表头行）。

    ⚠ 必须用**非 read_only**：EMAG「FBE Inventory」导出里 `<dimension ref="A1">` 是错的
    （实际 A1:AC323），openpyxl read_only 会尊重 dimension → 只拿到 1 行，静默返回 0 条。
    """
    wb = load_workbook(path, data_only=True)
    try:
        sn = None
        if sheet_hint:
            for s in wb.sheetnames:
                if sheet_hint.lower() in s.lower():
                    sn = s
                    break
        ws = wb[sn or wb.sheetnames[0]]
        return [[c for c in r] for r in ws.iter_rows(values_only=True)]
    finally:
        wb.close()


def _xlsx_kind(path):
    """xlsx 库龄报告类型：'emag' | 'ml' | None。

    EMAG  FBE Inventory：sheet 名含 FBE + Inventory（另有 Data definition / _errors 干扰页）
    美客多 长期仓储费：sheet 名含 Long term storage，或表头含「库龄天数」
    """
    try:
        wb = load_workbook(path, read_only=True, data_only=True)
        try:
            names = list(wb.sheetnames)
            for s in names:
                low = s.lower()
                if "fbe" in low and "inventor" in low and "error" not in low:
                    return "emag"
            for s in names:
                if "long term storage" in s.lower():
                    return "ml"
            # sheet 名不典型 → 扫前 3 行表头
            for s in names[:2]:
                ws = wb[s]
                for row in ws.iter_rows(min_row=1, max_row=3, values_only=True):
                    cells = [str(c or "") for c in row]
                    if "库龄天数" in cells and "SKU" in cells:
                        return "ml"
                    if "part_number" in cells:
                        return "emag"
            return None
        finally:
            wb.close()
    except Exception:
        return None


# EMAG FBE Inventory 的列名（英文，来自「Data definition EN」页，不是猜的）
EMAG_TIER_FIELDS = ["interval_1", "interval_2", "interval_3", "interval_4",
                    "interval_5", "interval_6", "interval_7"]


def _snap_from_name(path):
    """从文件名抓日期：`fbe-inventory-2026-10-27-05-38-59-xxx.xlsx` / `26-10-2026_Long_term...`"""
    fn = os.path.basename(str(path))
    m = re.search(r"(20\d{2})[-_](\d{2})[-_](\d{2})", fn)
    if m:
        return "%s-%s-%s" % m.groups()
    m = re.search(r"(\d{2})[-_](\d{2})[-_](20\d{2})", fn)   # 26-10-2026
    if m:
        return "%s-%s-%s" % (m.group(3), m.group(2), m.group(1))
    return datetime.now().date().isoformat()


def read_emag_age(path):
    """EMAG FBE Inventory → [(账号, 站点, sku, name, snapshot_date, [七档件数])]。

    **2026-10-27 拿到真实样例后改成精确映射**（原来那版是按列名宽松猜的，且读错了 sheet）：
      表结构：r1 分组表头 / r2 空 / r3 英文字段名 / r4 人读表头 / r5 起数据
      定位：找含 `part_number` 的那一行当字段名行，**再跳过 1 行人读表头**才是数据首行
      取值：SKU=part_number（G列，如 G243）、品名=name、数量=stock_available（M列）
            七档=interval_1..interval_7（T..Z列），与亚马逊 inv-age 七档一一对应
      ⚠ 不能用 read_only（dimension 写错），见 _xlsx_rows 注释
    兜底：若找不到 interval_* 列，退回旧的「在库天数/入库日期」猜列逻辑。
    """
    rows = _xlsx_rows(path, sheet_hint="FBE - Inventory")
    if not rows:
        return []

    hdr_i = None
    for i, r in enumerate(rows[:12]):
        if any(str(c or "").strip() == "part_number" for c in r):
            hdr_i = i
            break
    if hdr_i is None:
        return _read_emag_age_legacy(rows)

    hdr = [str(c or "").strip() for c in rows[hdr_i]]
    pos = {k: j for j, k in enumerate(hdr) if k}

    def col(name):
        return pos.get(name)

    c_sku, c_name = col("part_number"), col("name")
    c_qty = col("stock_available")
    c_tiers = [col(t) for t in EMAG_TIER_FIELDS]
    if c_sku is None:
        return []

    has_tiers = all(c is not None for c in c_tiers)
    snap = _snap_from_name(path)
    out = []
    # hdr_i 是字段名行，hdr_i+1 是人读表头行，数据从 hdr_i+2 开始
    for r in rows[hdr_i + 2:]:
        if c_sku >= len(r):
            continue
        sku = str(r[c_sku] or "").strip()
        if not sku or sku in ("Product code", "part_number"):
            continue
        if has_tiers:
            tiers = [_int(r[c]) if c < len(r) else 0 for c in c_tiers]
            if sum(tiers) <= 0:
                continue
        else:
            qty = _int(r[c_qty]) if c_qty is not None and c_qty < len(r) else 0
            if qty <= 0:
                continue
            tiers = [0] * NTIER
            tiers[0] = qty           # 没有分档列时只能全塞第一档（并在日志里提示）
        name = str(r[c_name] or "").strip() if c_name is not None and c_name < len(r) else ""
        out.append(("eMAG", "罗马尼亚", sku, name, snap, tiers))
    return out


def _read_emag_age_legacy(rows):
    """旧兜底：没有 interval_* 列时，按「在库天数/入库日期」归并。"""
    if not rows:
        return []
    hdr = [str(c or "").strip() for c in rows[0]]
    dicts = [dict(zip(hdr, r)) for r in rows[1:]]
    keys = [k for k in hdr if k]

    def pick(*cands):
        for c in cands:
            for k in keys:
                if c in str(k).lower():
                    return k
        return None

    k_sku = pick("sku", "product code", "part number", "cod produs")
    k_name = pick("product name", "name", "denumire")
    k_days = pick("days", "age", "zile")
    k_date = pick("intake", "entry", "inbound date", "data")
    k_qty = pick("quantity", "qty", "stock", "cantitate", "available")
    if not k_sku:
        return []

    today = datetime.now().date()
    agg = defaultdict(lambda: [0] * NTIER)
    names = {}
    for r in dicts:
        sku = str(r.get(k_sku) or "").strip()
        if not sku:
            continue
        qty = _int(r.get(k_qty)) if k_qty else 1
        if qty <= 0:
            continue
        days = None
        if k_days:
            try:
                days = int(float(str(r.get(k_days)).strip()))
            except (TypeError, ValueError):
                days = None
        if days is None and k_date:
            try:
                d = datetime.strptime(str(r.get(k_date)).strip()[:10], "%Y-%m-%d").date()
                days = (today - d).days
            except ValueError:
                days = None
        if days is None:
            continue
        idx = (0 if days <= 30 else 1 if days <= 60 else 2 if days <= 90 else
               3 if days <= 180 else 4 if days <= 270 else 5 if days <= 365 else 6)
        agg[sku][idx] += qty
        if k_name:
            names[sku] = str(r.get(k_name) or "").strip()
    return [("eMAG", "罗马尼亚", sku, names.get(sku, ""), today.isoformat(), t)
            for sku, t in agg.items()]


def read_ml_age(path):
    """美客多「长期仓储费」报告 → [(账号, 站点, sku, name, snapshot_date, [七档件数])]。

    表结构（2026-10-27 真实样例）：r1 标题「您可能需要支付$…」/ r2 中文表头 / r3 起数据
    表头：ML编码 通用编码 SKU 商品编号 产品 产品分类 尺寸 库龄 库龄天数 可售商品
          转运中可售商品 不可售商品 库存数量 产生费用商品 单位费用 长期仓储总费用

    ⚠ 这份报告**只含产生长期仓储费的商品**（202610 实测 30 个 SKU / 70 行），
    不是全量库龄 → 站点汇总里美客多的「总件数」是**费用口径**，不是库存口径，
    输出时会在「取数清单」里标红说明。要全量得让美客多导出完整库存库龄明细。
    """
    rows = _xlsx_rows(path, sheet_hint="Long term storage")
    if not rows:
        return []
    hdr_i = None
    for i, r in enumerate(rows[:8]):
        cells = [str(c or "").strip() for c in r]
        if "SKU" in cells and "库龄天数" in cells:
            hdr_i = i
            break
    if hdr_i is None:
        return []
    hdr = [str(c or "").strip() for c in rows[hdr_i]]
    pos = {k: j for j, k in enumerate(hdr) if k}
    c_sku, c_name = pos.get("SKU"), pos.get("产品")
    c_days = pos.get("库龄天数")
    c_qty = pos.get("库存数量") if "库存数量" in pos else pos.get("可售商品")
    if c_sku is None or c_days is None:
        return []

    snap = _snap_from_name(path)
    agg = defaultdict(lambda: [0] * NTIER)
    names = {}
    for r in rows[hdr_i + 1:]:
        if c_sku >= len(r):
            continue
        sku = str(r[c_sku] or "").strip()
        if not sku or sku == "SKU":
            continue
        try:
            days = int(float(str(r[c_days]).strip()))
        except (TypeError, ValueError, IndexError):
            continue
        qty = _int(r[c_qty]) if c_qty is not None and c_qty < len(r) else 0
        if qty <= 0:
            continue
        idx = (0 if days <= 30 else 1 if days <= 60 else 2 if days <= 90 else
               3 if days <= 180 else 4 if days <= 270 else 5 if days <= 365 else 6)
        agg[sku][idx] += qty
        if c_name is not None and c_name < len(r):
            names[sku] = str(r[c_name] or "").strip()
    return [("美客多", "墨西哥", sku, names.get(sku, ""), snap, t)
            for sku, t in agg.items()]


# ---------------- 汇总 ----------------
def latest_snapshot_only(items):
    """同一（账号, 站点）混进多个快照日时，只保留最新那一天 → (保留的 items, 说明列表)。

    2026-10-29 加。背景：`month_sale_file.py` 里的 belonging_date 是**手改的硬编码**，
    11 月初跑 10 月的月销表时它仍是 202610，任务⑪ 于是去读 `202610/库龄/`。
    这时如果新导的 11 月库龄报告也扔进 `202610/库龄/`，`find_age_files` 会把新旧两份
    一起收进来、`summarize` 直接**求和** —— 件数翻倍，且一句报错都没有。
    实测 202610 的 14 份 csv 单文件内 SKU 全唯一，所以「同站点同 SKU 只在一个快照里」成立，
    按（账号,站点）取最新快照日是安全的。
    """
    dates = defaultdict(set)
    for acc, site, _sku, _n, snap, _t in items:
        if snap:
            dates[(acc, site)].add(snap)
    keep = {k: max(v) for k, v in dates.items()}
    notes = []
    for (acc, site), ds in sorted(dates.items()):
        if len(ds) > 1:
            older = sorted(ds - {keep[(acc, site)]})
            notes.append(f"※ {acc} {site}：目录里有多个快照日 {sorted(ds)}，"
                         f"只取最新 {keep[(acc, site)]}，已忽略 {older}。"
                         f"旧快照请移出该期目录，别和新快照混放（混放会翻倍计数）。")
    kept = [it for it in items
            if not it[4]
            or len(dates.get((it[0], it[1]), ())) <= 1
            or it[4] == keep[(it[0], it[1])]]
    return kept, notes


def summarize(items):
    """→ (detail_rows, site_rows)。按（账号, 站点）分组——亚马逊和 interflow 是两套账号，
    SKU 会撞号（如沙特 G596 两边都有），混在一行会串数据。"""
    detail = []
    for acc, site, sku, name, snap, tiers in items:
        tot = sum(tiers)
        if tot <= 0:
            continue
        aged = sum(tiers[AGED_FROM:])
        detail.append(dict(acc=acc, site=site, sku=sku, name=name, snap=snap,
                           tiers=tiers, total=tot, aged=aged,
                           aged_pct=(aged / tot if tot else 0)))
    by_site = defaultdict(lambda: dict(n=0, total=0, tiers=[0] * NTIER, aged=0))
    for d in detail:
        a = by_site[(d["acc"], d["site"])]
        a["n"] += 1
        a["total"] += d["total"]
        a["aged"] += d["aged"]
        for i, v in enumerate(d["tiers"]):
            a["tiers"][i] += v
    site_rows = []
    for (acc, site), a in by_site.items():
        pct = a["aged"] / a["total"] if a["total"] else 0
        lv = "严重呆滞" if pct >= 0.4 else "偏高" if pct >= 0.2 else "正常"
        site_rows.append(dict(acc=acc, site=site, n=a["n"], total=a["total"],
                              tiers=a["tiers"], aged=a["aged"], aged_pct=pct, level=lv))
    site_rows.sort(key=lambda x: (-x["aged_pct"], x["acc"], x["site"]))
    detail.sort(key=lambda x: (-x["aged_pct"], x["acc"], x["site"], x["sku"]))
    return detail, site_rows


# ---------------- 输出 ----------------
_DET_COLS = ([("acc", "账号"), ("site", "站点"), ("sku", "SKU"), ("name", "品名")]
             + [(f"t{i}", TIER_LABEL[i]) for i in range(NTIER)]
             + [("total", "合计"), ("aged", "180天以上"), ("aged_pct", "180+占比")])
_SITE_COLS = ([("acc", "账号"), ("site", "站点"), ("n", "SKU数")]
              + [(f"t{i}", TIER_LABEL[i]) for i in range(NTIER)]
              + [("total", "总件数"), ("aged", "180天以上"), ("aged_pct", "180+占比"),
                 ("level", "呆滞等级")])


def _put(ws, title, cols, rows, row_mapper):
    ws.cell(row=1, column=1, value=title).font = Font(bold=True, size=12)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(cols))
    for j, (_k, label) in enumerate(cols, start=1):
        c = ws.cell(row=2, column=j, value=label)
        c.fill, c.font = _HDR_FILL, _HDR_FONT
        c.alignment = Alignment(horizontal="center", vertical="center")
    for i, r in enumerate(rows, start=3):
        vals = row_mapper(r)
        for j, (_k, _l) in enumerate(cols, start=1):
            ws.cell(row=i, column=j, value=vals.get(_k))
    ws.freeze_panes = "A3"
    ws.auto_filter.ref = f"A2:{get_column_letter(len(cols))}{max(len(rows) + 2, 3)}"
    for j, (k, label) in enumerate(cols, start=1):
        w = {"name": 40, "sku": 13, "site": 12, "level": 10}.get(k, 9)
        ws.column_dimensions[get_column_letter(j)].width = w


def write(items, out_dir, period, sources=None, extra_notes=None):
    os.makedirs(out_dir, exist_ok=True)
    out = os.path.join(out_dir, OUT_NAME)
    detail, site_rows = summarize(items)

    wb = Workbook()
    ws = wb.active
    ws.title = "本期明细"
    _put(ws, f"库龄明细（{period}）", _DET_COLS, detail,
         lambda r: dict(acc=r["acc"], site=r["site"], sku=r["sku"], name=r["name"],
                        total=r["total"], aged=r["aged"],
                        aged_pct=round(r["aged_pct"], 4),
                        **{f"t{i}": r["tiers"][i] for i in range(NTIER)}))

    ws2 = wb.create_sheet("站点汇总")
    _put(ws2, f"站点库龄汇总（{period}）", _SITE_COLS, site_rows,
         lambda r: dict(acc=r["acc"], site=r["site"], n=r["n"], total=r["total"],
                        aged=r["aged"],
                        aged_pct=round(r["aged_pct"], 4), level=r["level"],
                        **{f"t{i}": r["tiers"][i] for i in range(NTIER)}))

    # 历史汇总：若已有总表则追加（同 period 先删旧行，保证幂等）
    hist = []
    if os.path.exists(out):
        try:
            old = load_workbook(out, data_only=True)
            if "历史汇总" in old.sheetnames:
                ows = old["历史汇总"]
                for row in ows.iter_rows(min_row=3, values_only=True):
                    if row and row[0]:
                        hist.append(list(row))
            old.close()
        except Exception:
            pass
    hist = [h for h in hist if str(h[0]) != str(period)]
    for r in site_rows:
        hist.append([period, r["acc"], r["site"], r["n"], r["total"], r["aged"],
                     round(r["aged_pct"], 4), r["level"]]
                    + r["tiers"])
    hist.sort(key=lambda x: (str(x[0]), str(x[1]), str(x[2])))
    ws3 = wb.create_sheet("历史汇总")
    cols = [("p", "期数"), ("acc", "账号"), ("site", "站点"), ("n", "SKU数"),
            ("total", "总件数"),
            ("aged", "180天以上"), ("aged_pct", "180+占比"), ("level", "呆滞等级")] \
        + [(f"t{i}", TIER_LABEL[i]) for i in range(NTIER)]
    _put(ws3, "跨期库龄趋势（每期每站点一行，看是否在恶化）", cols, hist,
         lambda r: dict(zip([c[0] for c in cols], r)))

    # 取数清单：这一期到底吃了哪些文件、各解析多少行、缺什么
    ws4 = wb.create_sheet("取数清单")
    src_cols = [("fn", "文件"), ("fmt", "格式"), ("rows", "解析行数")]
    _put(ws4, f"取数清单（{period}）—— 缺的文件不会出现在上面两张表里", src_cols,
         sources or [], lambda r: r)
    # 取数清单底部说明：按本期实际吃到的格式动态生成（2026-10-27 改）
    fmts = {str(s.get("fmt", "")) for s in (sources or [])}
    notes = list(extra_notes or [])
    if not any("EMAG" in f for f in fmts):
        notes.append("※ EMAG「FBE Inventory」本期未提供 → 表内无 EMAG 行；"
                     "导出后放到 库龄/<日期>/ 再跑一次即可（按表头识别，不认文件名）。")
    if any("美客多" in f for f in fmts):
        notes.append("※ 美客多数据源是「长期仓储费」报告，只含**产生长期仓储费**的商品，"
                     "件数是费用口径而非全量库存；要全量库龄需美客多导出完整库存库龄明细。")
    if not any("美客多" in f for f in fmts):
        notes.append("※ 美客多本期未提供库龄报告 → 表内无美客多行。")
    note_row = max(len(sources or []) + 4, 5)
    for i, txt in enumerate(notes):
        ws4.cell(row=note_row + i, column=1, value=txt).font = Font(color="C00000", size=10)

    wb.save(out)
    return out, detail, site_rows


def main(argv=None):
    ap = argparse.ArgumentParser()
    ap.add_argument("--age", action="append", default=[], help="库龄报告路径（可多次）")
    ap.add_argument("--period", default=None, help="数据归属周期 YYYYMM")
    ap.add_argument("--out", default=None)
    a = ap.parse_args(argv)

    period = a.period or latest_age_period() or "202610"
    files = a.age or find_age_files(period)
    if not files:
        print(f"没有找到库龄报告。请把亚马逊「亚马逊物流库存」导出（csv）放到：\n"
              f"  {os.path.join(SNAP_BASE, period, '库龄')}")
        return 1

    items = []
    sources = []
    file_items = []          # 2026-11-02 新增：逐文件留存解析结果，供「取数清单」按实际入表行数重算
    for p in files:
        fn = os.path.basename(p)
        ext = os.path.splitext(p)[1].lower()
        if ext in (".xlsx", ".xlsm"):
            kind = _xlsx_kind(p)
            if kind == "emag":
                got, fmt = read_emag_age(p), "EMAG·FBE Inventory"
            elif kind == "ml":
                got, fmt = read_ml_age(p), "美客多·长期仓储费"
            else:
                print(f"[跳过] {fn} —— xlsx 但不是库龄报告")
                continue
        else:
            with open(p, encoding="utf-8-sig", errors="ignore") as f:
                head = f.readline()
            if "inv-age-" in head:
                got, fmt = read_amazon_age(p), "亚马逊·英文版"
            elif "库龄 0-30 天" in head or ("快照日期" in head and "库龄" in head):
                got, fmt = read_cn_age(p), "亚马逊·中文版"
            else:
                print(f"[跳过] {fn} —— csv 表头不认识")
                continue
        print(f"[{fmt}] {fn} → {len(got)} 行")
        items += got
        file_items.append((fn, fmt, got))
    if not items:
        print("库龄报告解析出 0 行，请检查表头。")
        return 1

    _all_files = list(file_items)
    items, snap_notes = latest_snapshot_only(items)

    # ⚠️ 2026-11-02 修：以前 sources 直接把**所有**解析到的文件写进「取数清单」，
    #    而 items 已经被 latest_snapshot_only 按最新快照过滤过 —— 于是清单里会列出
    #    「解析了但一件都没进表」的旧快照文件（实测 202610：清单 50 个文件 9,975 行，
    #    实际入表 17 个文件）。清单是给人核对口径用的，必须与实际入表一致。
    _kept_ids = {id(it) for it in items}
    sources = []
    for fn, fmt, got in _all_files:
        _kept = [g for g in got if id(g) in _kept_ids]
        if got and not _kept:
            print(f"[忽略] {fn} —— 同站点存在更新的快照，本文件未入表")
            continue
        sources.append(dict(fn=fn, fmt=fmt, rows=len(_kept)))
    for line in snap_notes:
        print(f"⚠️ {line}")

    out_dir = a.out or os.path.join(_PROJ, "trunk", "data", "库龄报表")
    out, detail, site_rows = write(items, out_dir, period, sources, snap_notes)
    aged = sum(r["aged"] for r in site_rows)
    total = sum(r["total"] for r in site_rows)
    print(f"\n{period}：{len(detail)} 个 SKU，库存 {total} 件，其中 180 天以上 {aged} 件"
          f"（{aged / total:.1%}）")
    for r in site_rows:
        print(f"  {r['acc']:9s} {r['site']:6s} SKU {r['n']:4d}  件 {r['total']:6d}  "
              f"180+ {r['aged']:5d} ({r['aged_pct']:.1%})  {r['level']}")
    print(f"已输出: {out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
