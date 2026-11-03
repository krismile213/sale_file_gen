# -*- coding: utf-8 -*-
"""日销表程序化更新 - 配置

统一时间戳原则：所有库存报告无论实际导出日期，第一律按 SNAPSHOT_DATE 处理
（用户后续把报告时间戳统一后重跑即可）。

报告格式 kind：
  inv   = 管理库存报告（26 列 afn-* 标准格式，tab 分隔 txt 用 sep='\t'）
  age   = 库龄库存报告（90 列，available/Total Reserved Quantity/inbound-* 字段名）
"""
import os
import re
import calendar
from datetime import date, timedelta

# ---------- 路径 ----------
# （2026-10-12 自会话区回传）原先写死 C:\Users\dell\Desktop\work-main\... 只在装机时有效，
# 换机器后 DSR/WORK_DIR/SNAP_BASE 三个目录全判 False、脚本一步跑不动
# → 改为按 config.py 自身位置推导，跟着工作区走、不再依赖盘符与用户名。
_HERE = os.path.dirname(os.path.abspath(__file__))            # …\Daily_Sales_Report\pipeline\core


def _find_dsr(start):
    """向上找「含 1.每日销售数据表 的目录」作为 DSR 根。

    不再依赖本文件处在第几层（2026-10-14 目录收敛到 pipeline/core/ 后，
    旧的「上溯两级」写法会少一级）。找不到就退回旧布局假设。
    """
    p = start
    for _ in range(8):
        if os.path.isdir(os.path.join(p, "1.每日销售数据表")):
            return p
        parent = os.path.dirname(p)
        if parent == p:
            break
        p = parent
    return os.path.dirname(os.path.dirname(start))


DSR = _find_dsr(_HERE)                                        # …\Daily_Sales_Report
PROJECT = os.path.dirname(DSR)                                # …\sale_file_gen
WORK_DIR = os.path.join(DSR, "1.每日销售数据表")            # 12 份工作副本

# ---------- 数据归属周期 / 快照：全自动取最新 ----------
# 每天导出的报告放 trunk\data\数据归属周期\{YYYYMM}\快照日期\{YYYYMMDD}\
# 2026-10-14 起 CURRENT_PERIOD / ORDER_MONTH / WINDOW_LO / SNAP 全部自动推导，
# 换月、换快照日都不必再手改 config；--period / --snap 仍可手工指定。
SNAP_BASE = os.path.join(PROJECT, "trunk", "data", "数据归属周期")


def _report_count(d):
    """目录里的报告文件数（.csv/.txt/.xlsx，只看本层、不进子目录）。

    用途：同一天在两个周期目录下都有快照时，挑「真放了数据的那一个」。
    """
    n = 0
    try:
        for f in os.listdir(d):
            if f.startswith("~$"):
                continue
            if os.path.isfile(os.path.join(d, f)) and \
                    f.lower().endswith((".csv", ".txt", ".xlsx", ".xls")):
                n += 1
    except OSError:
        pass
    return n


def _discover_latest():
    """扫 SNAP_BASE/{YYYYMM}/快照日期/{YYYYMMDD}，取**日期最大**的 (period, day)。

    按 YYYYMMDD 全局比较（不按目录名），跨月切换时 11-01 的快照会正确
    取代 10-31（即使两个周期目录同时存在）。

    日期相同时（跨月那几天两个周期目录并存）：先比**报告文件数**，多者胜；
    再相同则周期名更大（新月份）者胜。
    ⚠️ 2026-11-02 踩过：`202610/快照日期/20261102` 被手工建了个空壳（只有月销售
       总表 + 几个空目录），而真数据在 `202611/快照日期/20261102`。旧逻辑只比
       日期、相等不替换，于是选中先扫到的空壳 → 38 份报告全部「缺失」。
    """
    best = None
    best_key = None
    if os.path.isdir(SNAP_BASE):
        for per in os.listdir(SNAP_BASE):
            if not re.match(r"^\d{6}$", per):
                continue
            sd = os.path.join(SNAP_BASE, per, "快照日期")
            if not os.path.isdir(sd):
                continue
            for d in os.listdir(sd):
                if not re.match(r"^\d{8}$", d):
                    continue
                key = (d, _report_count(os.path.join(sd, d)), per)
                if best_key is None or key > best_key:
                    best_key = key
                    best = (per, d)
    return best


_LATEST = _discover_latest()
CURRENT_PERIOD = _LATEST[0] if _LATEST else "202610"   # 当前数据归属周期（月）目录名
ORDER_MONTH = CURRENT_PERIOD                           # 订单报告月份（{m} → …站点202610订单报告.txt）
_LATEST_DAY = _LATEST[1] if _LATEST else "20261006"

# SNAP 现指「最新快照目录」：既是 latest_snapshot() 的兜底，也是 sep_compare 等
# 历史工具的默认源。（旧版曾写死 202609/20261006 基线，2026-10-14 起自动。）
SNAP = os.path.join(SNAP_BASE, CURRENT_PERIOD, "快照日期", _LATEST_DAY)
BACKUP_DIR = os.path.join(DSR, ".workbuddy", "backup_20261008")
# 货件报告目录（亚马逊「货件详情报告」TSV，用户陆续上传）：{货件编号}.tsv
SHIP_DIR = os.path.join(SNAP_BASE, CURRENT_PERIOD, "货件")


def latest_snapshot(period=None):
    """返回 (最新快照目录, 8位日期字符串)；指定周期无快照目录则退回全局最新。"""
    base = os.path.join(SNAP_BASE, period or CURRENT_PERIOD, "快照日期")
    if os.path.isdir(base):
        ds = sorted(x for x in os.listdir(base) if re.match(r"^\d{8}$", x))
        if ds:
            return os.path.join(base, ds[-1]), ds[-1]
    return SNAP, _LATEST_DAY


def prev_snapshot(day, period=None):
    """返回比 day 更早的最近一个快照日（同周期内）；没有则返回 None。"""
    base = os.path.join(SNAP_BASE, period or CURRENT_PERIOD, "快照日期")
    if not os.path.isdir(base):
        return None
    ds = sorted(x for x in os.listdir(base) if re.match(r"^\d{8}$", x) and x < day)
    return ds[-1] if ds else None


# ---------- 参数 ----------
SNAPSHOT_DATE = "2026-10-06"     # 统一时间戳：库存记录补列补到这一天（9月期；日常跑时由快照目录名动态取）
EXT_DAYS = True                  # 中间日（含周末）全部占位列+公式（与手工一致）

# ---------- 上架三态（2026-10-23 用户拍板）----------
# 原话：「listing和库存的判定可以用来额外辅助，一旦有库存那天就可以当作是这个SKU的起售日期」
# True  → 新口径：库存说话。可售>0 记「在售」，可售=0 且在途>0 记「在途」，
#         两者都 0 且 listing 不通记「未上架」；**补行不再要求 listing 活着**，
#         也不再要求「首次出现 >= 当月 1 号」（唯一例外见 TRISTATE_KEEP_MONTH_GATE）。
# False → 完整回退旧行为（listing 闸 + 首现月份闸，状态只有在售/未上架），
#         代码里所有三态分支都走 else，不影响任何既有逻辑。
TRISTATE_ENABLED = True

# 三态下是否仍然只补「本月新出现」的 SKU（True = 保留月份闸，False = 只要在表外就补）。
# ⚠️ 2026-10-24 用户拍板改为 False，改用 TRISTATE_REQUIRE_EVIDENCE 的「有货/有销量」闸。
# 原因：月份闸按「首现日期」判新旧，会把「9 月首现、本月刚到货」的 SKU 挡在表外，
# 与「日销表要及时更新新品」直接冲突。保留这个开关只为出问题时能一键回退。
TRISTATE_KEEP_MONTH_GATE = False

# 补行证据闸（2026-10-24 用户原话：「正则可以打开，前提是销量对得上」）。
# True → 补行必须有硬证据，满足任一条即可：
#        ① 本站可售 > 0    ② 本站有在途 > 0    ③ 本站本月有销量 > 0
#        没库存、没销量的 SKU（哪怕 listing 通、哪怕首现是本月）一律不补。
#        这样放宽正则、废弃月份闸都不会引入垃圾行 —— 证据闸兜底。
# False → 回到旧的「listing 活着 + 首现 >= 当月 1 号」判据。
TRISTATE_REQUIRE_EVIDENCE = True

# 「零证据也补行」的 SKU 白名单（2026-10-26 用户拍板）。
# 用户原话：「只要识别到 G599 或者 G602 的数据的时候就可以列进来，也方便后面货件列的比对」
# 背景：G599 在库存报告里 afn-listing-exists=Yes，但 可售/在途/本月销量 全 0，
#       被上面的证据闸挡住；用户希望这类「listing 已建、货还没到」的真 SKU 先进表，
#       这样后面货件列才能对得上。
# 判据：SKU 必须命中**真编号**正则（G599 这种），且 listing 不是明确 False。
# ⚠️ 正则闸不能省：欧盟报告里「listing 通但不在表里」的 42 个 SKU 中，41 个是
#    amzn.gr.G117-UOFBnm4I1… 这类亚马逊系统临时编号，只有 G599 是真编号。
#    不加闸会一次塞进 41 个垃圾行。
# 置空字符串 "" 即可完全关闭（回到「必须 有货/有在途/有销量」的老闸）。
ZERO_EVIDENCE_SKU_RE = r"^[A-Z]{1,3}\d{3,4}[A-Z]{0,3}$"

# 三态下，哪些状态算「有货可卖」——用来写「起售日期」。
# 用户口径：有库存那天 = 起售日期，所以在途不算。改成 ("在售","在途") 可把发货日也算起售。
TRISTATE_START_STATES = ("在售",)
# 当月表窗口下界 = 上月7天起点（月度订单报告按"最后更新日期"导出，会混入更老的
# 订单行——2024/2025 退款散单、上月早期的散单；append 前按 purchase-date 过滤掉，
# 否则会破坏"上月7天+当月"窗口口径）。2026-10-14 起由 CURRENT_PERIOD 自动推导。
def _window_lo(period):
    """上月最后 7 天起点 = 上月最后一天 − 6 天（202610 → 2026-09-24）。"""
    y, m = int(period[:4]), int(period[4:6])
    py, pm = (y - 1, 12) if m == 1 else (y, m - 1)
    last = date(py, pm, calendar.monthrange(py, pm)[1])
    return (last - timedelta(days=6)).isoformat()


WINDOW_LO = _window_lo(CURRENT_PERIOD)

# ---------- 站点表月度封存（2026-10-27 用户拍板）----------
# 用户原话：「日销表的数据很重要，而且以后跑，这几个国家表也是要更新的」
# 背景：欧盟英国簿里除了「英国/欧洲-每日库存分析表」，还有 10 张站点表
#       （意大利/法国/西班牙/德国/土耳其/荷兰/瑞典/波兰/比利时/爱尔兰）。
#       apply_month_summary 原来只遍历名字含「分析」的 sheet，站点表从来没被月度封存
#       覆盖过 → 8/9/10 月整列是 0（2026-10-27 手工补过一次，现在改成每期自动）。
# True  → 除分析表外，任何「scan_analysis 能解出 SUMIFS 口径 + 12 个月标签列」的
#         sheet 也参与月度封存（站点表天然满足；不满足的自动跳过，无需逐簿配置）。
# False → 完整回退旧行为：只封存分析表。
SITE_MONTHLY_FREEZE = True

# 站点表「上月日均」列是否同步改写为活公式（=IF(封存列="","",ROUND(封存列/天数,2))）。
# 站点表 G 列就是「上月日均」，且是染色规则的基准（如意大利
# `AND(P7>$G7*2.5,$G7>1.5)`），静态值换月后不会更新 → 默认一并刷新。
# 置 False 只封存月度列、不动日均列（出问题时的降级开关）。
SITE_MONTHLY_FILL_AVG = True

# 站点表补行（2026-10-27 用户拍板：「销量是最核心的东西，要保证准确」）。
# 背景：G598（Amazon.it 10月 1 件）和 K001（Amazon.se 10月 1 件）在 10 张站点表里
#       一行都没有，站点表合计比分析表少 2 件 —— 补行闸原本只作用于分析表。
# True  → 站点表按「本站本月销量 > 0」补行（证据闸，先于一切；不看库存/首现日期），
#         复用 append_sku_rows_auto（克隆列最完整行、公式平移、CF 范围跟随），
#         在 apply_month_summary 之前执行，保证新行当期就拿到月度值。
# False → 完整回退：站点表永不自动补行（2026-10-27 之前的行为）。
SITE_APPEND_ENABLED = True

# 站点表名 → 该站订单渠道（month_sales 的 channel 参数）。只处理簿内实际存在的表，
# 其它簿没有这些 sheet 名，天然跳过。amzn.gr.* 别名由 normalize_sku 归并，
# G261IT 这类带站点后缀的行名是订单原始 SKU 本身，无需特殊处理。
SITE_SHEET_CHANNELS = {
    "意大利": "Amazon.it",
    "法国": "Amazon.fr",
    "西班牙": "Amazon.es",
    "德国": "Amazon.de",
    "土耳其": "Amazon.com.tr",
    "荷兰": "Amazon.nl",
    "瑞典": "Amazon.se",
    "波兰": "Amazon.pl",
    "比利时": "Amazon.be",
    "爱尔兰": "Amazon.ie",
}

# 编码尝试顺序（亚马逊 csv 有 utf-8 / gb18030 / 日文 shift_jis）
ENC_CANDIDATES = ["utf-8-sig", "gb18030", "shift_jis", "latin-1"]

# ⚠️ gb18030 几乎不会解码失败，若排在 shift_jis 之前，日文报告会被"成功"读成乱码
# （如 新品→怴昳、対応→懳墳），且数值列正常、只有文本列坏 → 静默出错。
# 故按报告文件名给"首选编码"，命中则提到最前（其它候选仍保留兜底）。
REPORT_ENC = {
    "日本": "shift_jis",        # 亚马逊日本站点订单/管理库存报告
    "澳大利亚": "shift_jis",    # 亚马逊澳大利亚 + interflow 澳大利亚（均为日文界面导出）
}

# 日文界面导出的管理库存报告（24 列，列头日文乱码但数据正常）：按位置映射标准字段
JP24 = [
    "sku", "fnsku", "asin", "product-name", "condition", "your-price",
    "mfn-listing-exists", "mfn-fulfillable-quantity", "afn-listing-exists",
    "afn-warehouse-quantity", "afn-fulfillable-quantity", "afn-unsellable-quantity",
    "afn-reserved-quantity", "afn-total-quantity", "per-unit-volume",
    "afn-inbound-working-quantity", "afn-inbound-shipped-quantity",
    "afn-inbound-receiving-quantity", "afn-researching-quantity",
    "afn-reserved-future-supply", "afn-future-supply-buyable",
    "afn-fc-transfer-quantity", "afn-onhand-buyable-quantity", "store",
]
# ⚠️ 位置16(shipped)/17(receiving) 的日文语义存疑，两者之和不受影响（在途=Y=P+Q+R），单列值待人工核对

# 每个工作簿的"下载日期"锚点（滚动到快照日）。⚠️ 必须列全**所有**带日期的库存 sheet：
# 2026-10-12 发现 ANCHORS 每簿只登记一个（第一个库存 sheet）时，「在途库存!C1」系列
# 从不滚动，4 个簿停在 2026-09-18；Int澳则登记了在途、漏了可售（B1 停在 09-14）。
# 现按"该 sheet 数据真实入库日"逐格列全；美客多 库存明细!A1 在 update_ml_stock_detail 内自行滚动（2937处引用）。
ANCHORS = {
    "日本库存销售表.xlsx":                 [("可售+在途库存", "C1")],
    "沙特库存销售表.xlsx":                 [("沙特阿拉伯可售库存", "B1"), ("沙特阿拉伯在途库存", "C1")],
    "澳大利亚库存销售表.xlsx":              [("可售库存", "B1"), ("在途库存", "C1")],
    "印度库存销售表.xlsx":                 [("可售库存", "B1"), ("在途库存", "C1")],
    "阿联酋库存销售表.xlsx":                [("阿联酋可售库存", "B1"), ("阿联酋在途库存", "C1")],
    "Interflowd澳大利亚库存销售表.xlsx":    [("在途库存", "C1"), ("可售库存", "B1")],
    "Interflowd阿联酋库存销售表.xlsx":      [("阿联酋可售库存", "B1"), ("阿联酋在途库存", "C1")],
    "Interflowd沙特库存销售表.xlsx":        [("沙特可售库存", "B1"), ("沙特在途库存", "C1")],
    # ↓ 北美系/欧盟系的各库存 sheet 不各自存日期，C1 是公式指向同一个手工静态日期格
    # （北美: 数据透视图!H5；欧盟: 数据透视表!B11）。因为是**共享格**，runner 只在
    # 该簿全部库存报告都入库时才滚动（见 runner 的 stock_skipped 判定），缺报告就不动。
    "北美加拿大库存销售表.xlsx":             [("数据透视图", "H5")],
    "Interflowd北美加拿大库存销售表.xlsx":   [("数据透视图", "H5")],
    "欧盟英国库存销售表.xlsx":               [("数据透视表", "B11")],
}

# ---------- 每份工作簿的数据源映射 ----------
# stocks : 库存 sheet ← 报告
# records: 库存记录 sheet（右侧加列）；没有此配置则跳过
# sales  : 每日销售数据 sheet ← 订单报告
# 报告文件名模板：{d}=快照日YYYYMMDD {d4}=MMDD {m}=订单报告月YYYYMM

# 库存记录「在售块」取值字段覆盖（默认 afn-onhand-buyable-quantity）。
# 墨西哥记录表在售取 afn-fulfillable-quantity（结余式「墨西哥可售库存」口径，
# 与 2026-10-13 在售块改造 verify_full.py 实测一致）；在途块所有表统一
# = inbound-working + inbound-shipped + inbound-receiving。
RECORD_INSALE_FIELD = {
    "墨西哥库存记录": "afn-fulfillable-quantity",
}
# ---- sales-channel 白名单（2026-10-16 新增）------------------------------------
# `runner.append_sales` 据此拒收「渠道属于**其它簿**」的行：源报告错位（如把沙特数据
# 下载成阿联酋报告名）时不再把脏行灌进来。判定只在「已知属于别的簿」时触发 ——
# 完全没见过的渠道（如新开的欧盟站点）**保留并告警**，避免误杀合法行。
# `Non-Amazon*` 一律放行（分析表有效性公式已排除它们）。
EU_CHANNELS = ["Amazon.co.uk", "Amazon.de", "Amazon.fr", "Amazon.it", "Amazon.es",
               "Amazon.nl", "Amazon.se", "Amazon.pl", "Amazon.ie", "Amazon.com.be"]
NA_CHANNELS = ["Amazon.com", "Amazon.ca", "Amazon.com.mx"]

WORKBOOKS = [
    dict(
        file="欧盟英国库存销售表.xlsx",
        stocks=[
            dict(sheet="英国可售+在途库存", report="亚马逊英国管理库存报告{d}.csv", kind="inv"),
            # 「欧洲可售+在途库存」2026-10-13 起改用快照内的「亚马逊欧盟管理库存报告」
            # （26 列 afn-* inv 式），替代原先要手工下载改名放到 DSR 根目录的
            # 「欧洲库存{d4}.csv」（93 列库龄 age 式）。两者口径等价：
            #   inv 原生 afn-onhand-buyable-quantity == fulfillable + fc-transfer（实测 495/495 行一致）
            # 且 SKU 覆盖更全（EU 表 458 个 SKU：age 源命中 434、24 个每天被清空；
            # 新报告命中 457）。副作用：新报告另含 38 个英国 SKU，会按 append 逻辑
            # 追加成 EU 表新行（无害，欧洲记录表按 SKU SUMIF 不会串）。
            dict(sheet="欧洲可售+在途库存", report="亚马逊欧盟管理库存报告{d}.csv", kind="inv"),
            # 「多国库存」= 亚马逊「多国库存报告」原样粘贴（Pan-EU 多国本地履约量）。
            # 报告下载名不带日期（欧洲多国库存.csv / 多国库存报告.csv——20261013 快照里
            # 用户放的是后者）；带日期的候选留给改名后的情形。按顺序取第一个存在的。
            dict(sheet="多国库存",
                 report=["欧洲多国库存{d}.csv", "欧洲多国库存.csv",
                         "多国库存报告{d}.csv", "多国库存报告.csv"],
                 kind="multi"),
        ],
        records=["英国库存记录 ", "欧洲库存记录"],
        # ⚠️ 亚马逊后台该报告的下改名会漂移：10-13 快照是「亚马逊**欧洲**站点」，
        # 10-14 快照变成「亚马逊**欧盟**站点」（内容同一份泛欧站点订单报告）。
        # report 支持候选列表（runner 按顺序取第一个存在的），两种命名都兜住。
        sales=dict(sheet="每日销售数据",
                   report=["亚马逊欧洲站点{m}订单报告.txt", "亚马逊欧盟站点{m}订单报告.txt"],
                   sep="\t", expect_channel=EU_CHANNELS),
    ),
    dict(
        file="北美加拿大库存销售表.xlsx",
        stocks=[
            dict(sheet="美国可售+在途库存", report="亚马逊美国管理库存报告{d}.csv", kind="inv"),
            dict(sheet="加拿大可售+在途库存", report="亚马逊加拿大管理库存报告{d}.csv", kind="inv"),
            dict(sheet="墨西哥可售+在途库存", report="亚马逊墨西哥管理库存报告{d}.csv", kind="inv"),
            # 结余式可售库存 sheet：墨西哥真数据在「墨西哥可售库存」（366 行）；
            # 美国/加拿大这两张是空壳模板（8月版起就没有），保留映射以便将来填数
            dict(sheet="美国", report="亚马逊美国管理库存报告{d}.csv", kind="balance"),
            dict(sheet="加拿大", report="亚马逊加拿大管理库存报告{d}.csv", kind="balance"),
            dict(sheet="墨西哥可售库存", report="亚马逊墨西哥管理库存报告{d}.csv", kind="balance"),
        ],
        records=["美国库存记录", "加拿大库存记录", "墨西哥库存记录"],
        sales=dict(sheet="每日销售数据", report="亚马逊北美站点{m}订单报告.txt", sep="\t",
                   expect_channel=NA_CHANNELS),
    ),
    dict(
        file="日本库存销售表.xlsx",
        stocks=[dict(sheet="可售+在途库存", report="亚马逊日本管理库存报告{d}.csv", kind="inv",
                     colmap="JP24")],
        records=[],   # 日本无库存记录 sheet
        sales=dict(sheet="每日销售数据", report="亚马逊日本站点{m}订单报告.txt", sep="\t",
                   expect_channel=["Amazon.co.jp"]),
    ),
    dict(
        file="沙特库存销售表.xlsx",
        stocks=[
            dict(sheet="沙特阿拉伯可售库存", report="亚马逊沙特阿拉伯管理库存报告{d}.csv", kind="balance"),
            dict(sheet="沙特阿拉伯在途库存", report="亚马逊沙特阿拉伯管理库存报告{d}.csv", kind="inv"),
        ],
        records=[],
        sales=dict(sheet="每日销售数据沙特", report="亚马逊沙特阿拉伯站点{m}订单报告.txt", sep="\t",
                   expect_channel=["Amazon.sa"]),
    ),
    dict(
        file="澳大利亚库存销售表.xlsx",
        stocks=[
            dict(sheet="可售库存", report="亚马逊澳大利亚管理库存报告{d}.csv", kind="balance",
                 colmap="JP24"),
            dict(sheet="在途库存", report="亚马逊澳大利亚管理库存报告{d}.csv", kind="inv",
                 colmap="JP24"),   # 澳洲报告也是日文界面导出（24列，shift_jis）
        ],
        records=[],
        sales=dict(sheet="每日销售数据", report="亚马逊澳大利亚站点{m}订单报告.txt", sep="\t",
                   expect_channel=["Amazon.com.au"]),
    ),
    dict(
        file="印度库存销售表.xlsx",
        stocks=[
            dict(sheet="可售库存", report="亚马逊印度管理库存报告{d}.csv", kind="balance"),
            dict(sheet="在途库存", report="亚马逊印度管理库存报告{d}.csv", kind="inv"),
        ],
        records=[],
        sales=dict(sheet="每日销售数据", report="亚马逊印度站点{m}订单报告.txt", sep="\t",
                   expect_channel=["Amazon.in"]),
    ),
    dict(
        file="阿联酋库存销售表.xlsx",
        stocks=[
            dict(sheet="阿联酋可售库存", report="亚马逊阿拉伯联合酋长国管理库存报告{d}.csv", kind="balance"),
            dict(sheet="阿联酋在途库存", report="亚马逊阿拉伯联合酋长国管理库存报告{d}.csv", kind="inv"),
        ],
        records=[],
        sales=dict(sheet="每日销售数据阿联酋", report="亚马逊阿拉伯联合酋长国站点{m}订单报告.txt", sep="\t",
                   expect_channel=["Amazon.ae"]),
    ),
    dict(
        file="Interflowd北美加拿大库存销售表.xlsx",
        stocks=[
            dict(sheet="美国可售+在途库存", report="interflow美国管理库存报告{d}.csv", kind="inv"),
            dict(sheet="加拿大可售+在途库存", report="interflow加拿大管理库存报告{d}.csv", kind="inv"),
        ],
        records=["美国库存记录", "加拿大库存记录", "美国预留库存记录", "加拿大预留库存记录"],
        sales=dict(sheet="每日销售数据", report="interflow北美站点{m}订单报告.txt", sep="\t",
                   expect_channel=["Amazon.com", "Amazon.ca"]),
    ),
    dict(
        file="Interflowd澳大利亚库存销售表.xlsx",
        stocks=[
            dict(sheet="可售库存", report="interflow澳大利亚管理库存报告{d}.csv", kind="balance",
                 colmap="JP24"),
            dict(sheet="在途库存", report="interflow澳大利亚管理库存报告{d}.csv", kind="inv",
                 colmap="JP24"),   # 同为日文界面导出
        ],
        records=[],
        sales=dict(sheet="每日销售数据", report="interflow澳大利亚站点{m}订单报告.txt", sep="\t",
                   expect_channel=["Amazon.com.au"]),
    ),
    dict(
        file="Interflowd阿联酋库存销售表.xlsx",
        stocks=[
            dict(sheet="阿联酋可售库存", report="interflow阿拉伯联合酋长国管理库存报告{d}.csv", kind="balance"),
            dict(sheet="阿联酋在途库存", report="interflow阿拉伯联合酋长国管理库存报告{d}.csv", kind="inv"),
        ],
        records=[],
        sales=dict(sheet="每日销售数据阿联酋", report="interflow阿拉伯联合酋长国站点{m}订单报告.txt", sep="\t",
                   expect_channel=["Amazon.ae"]),
    ),
    # Interflowd沙特（2026-10-24 新建，方案 A）
    # 起因：interflow 账号有阿联酋、沙特两个物理仓，沙特仓 RUH8 的 116 件可售一直没有任何簿接收
    # （29 SKU，含 G596=44 / G594=8 / G592=7）。订单报告两份哈希相同，含 Amazon.ae 58 件
    # + Amazon.sa 5 件；本簿按 expect_channel=["Amazon.sa"] 只收 sa，与阿联酋簿互补不重复。
    dict(
        file="Interflowd沙特库存销售表.xlsx",
        # 库存报告：精确名优先，再回退到「任意日期的最新一份」。
        # interflow 沙特仓的下载频率低于其它站点，20261023 快照里最新一份仍叫 …20261021.csv
        # （内容已更新、文件名没变），死认 {d} 会天天判缺报告 → 库存永不更新。
        stocks=[
            dict(sheet="沙特可售库存",
                 report=["interflow沙特阿拉伯管理库存报告{d}.csv",
                         "interflow沙特阿拉伯管理库存报告*.csv"], kind="balance"),
            dict(sheet="沙特在途库存",
                 report=["interflow沙特阿拉伯管理库存报告{d}.csv",
                         "interflow沙特阿拉伯管理库存报告*.csv"], kind="inv"),
        ],
        records=[],
        # allow_mixed_channel：这份订单报告与阿联酋站是**同一份下载**（sha256 相同），
        # 天然混着 Amazon.ae 58 件 + Amazon.sa 5 件。本簿只收 sa（逐行过滤），
        # 不升级成"整份错位拒收"，否则 5 件沙特订单永远进不来。
        sales=dict(sheet="每日销售数据沙特",
                   report=["interflow沙特阿拉伯站点{m}订单报告.txt",
                           "interflow沙特阿拉伯站点*订单报告.txt"], sep="\t",
                   expect_channel=["Amazon.sa"], allow_mixed_channel=True,
                   # sku_gate_min：体检第 ⑩ 关的 SKU 归属阈值（默认 0.6）。
                   # 本簿库存只覆盖沙特仓 29 个 SKU，而 Amazon.sa 渠道有 19 个 SKU
                   # —— 多数从阿联酋仓发货，天然只有 32% 重合，不是导错账号
                   # （错账号实测 10~15%）。下调到 0.2 保留抓错能力。
                   sku_gate_min=0.2),
    ),
    # 美客多：库存明细 ← ML后台"库存总览(Resumen)"报告（Full仓状态，5个sheet取Resumen页），
    # 整块替换+锚点A1滚动，已接入；订单数据 ← ML后台月度订单报告（xlsx）。
    # 订单报告按月下载，当月报告每天重复导入（# de venta 去重，幂等）；上月终版报告绝对路径常驻。
    # 2026-10-12 接入：销售日期/状态/描述按西语正典规范化（ML后台语言漂移，202608/202610 中文、202609 西语），
    # 否则 c64 purchase-date 公式（查透视图O1:P12西语月名）和 c65 有效性公式（查透视图K2:L38西语关键词）全 #N/A，
    # 分析表 SUMIFS(BL,BM) 会整列漏计。有效性语义已用 backup_daily_20261010 缓存值 1183 行复刻验证 0 差异。
    dict(
        file="美客多库存销售表.xlsx",
        stocks=[dict(sheet="库存明细", report="美客多墨西哥Full库存报告{d}.xlsx", kind="ml_full")],
        records=[],
        sales=[
            # 上月终版：9月报告（西语，1375行，工作簿订单止于9-1，9-2~9-30 缺口在此补齐）
            dict(sheet="订单数据",
                 report=os.path.join(PROJECT, "trunk", "data", "数据归属周期",
                                     "202609", "快照日期", "20261006",
                                     "美客多墨西哥站点202609订单报告.xlsx"),
                 kind="ml_sales", window_lo="2026-09-01"),
            # 当月：{m}=ORDER_MONTH，每天导出的月度累计报告
            dict(sheet="订单数据", report="美客多墨西哥站点{m}订单报告.xlsx",
                 kind="ml_sales", window_lo="2026-09-01"),
        ],
    ),
    dict(
        file="罗马尼亚库存销售表.xlsx",
        # 库存明细 ← eMAG后台预填的优惠更新文件（sheet3「库存」，2026-10-13 接入）：
        # c6=PNK c7=产品代码 c10=数量 c11=stock_type → 工作簿 c4/c5/c27/c28，整块重贴+A1锚点滚动。
        # ⚠️ 文件 dimension 记录失真，read_only 会漏读（runner 已按普通模式处理）。
        # 之前快照里那份"空库存sheet"是没带数据的空白模板，要导"带库存数字"的那份。
        stocks=[dict(sheet="库存明细", report="EMAG罗马尼亚库存报告{d}.xlsx", kind="emag_stock")],
        records=[],
        sales=dict(sheet="订单数据", report="EMAG全站点{m}订单报告.xlsx",
                   sep=None, key="订单编号"),   # xlsx 报告；订单编号唯一可做去重键
    ),
]

# 库龄报告(age)字段 → 管理库存(inv)字段 的映射（写值时统一成 inv 名）
AGE_FIELD_MAP = {
    "available": "afn-fulfillable-quantity",
    "Total Reserved Quantity": "afn-reserved-quantity",
    "unfulfillable-quantity": "afn-unsellable-quantity",
    "inbound-working": "afn-inbound-working-quantity",
    "inbound-shipped": "afn-inbound-shipped-quantity",
    "inbound-received": "afn-inbound-receiving-quantity",
    "fc-transfer": "afn-fc-transfer-quantity",
    # 以下 inv 字段在 age 报告无来源，置 0
    "afn-researching-quantity": 0,
    "afn-reserved-future-supply": 0,
    "afn-future-supply-buyable": 0,
}

# ---------- 美客多订单报告 中文→西语正典 规范化 ----------
# ML 后台语言漂移：同一份月度订单报告，202608/202610 导出为中文界面、202609 为西语。
# 工作簿「订单数据」正典 = 西语（手工粘贴期的导出语言），且 c64 purchase-date /
# c65 有效性两个数组公式直接在文本上查「墨西哥透视图」的西语月名表(O1:P12)和西语
# 关键词表(K2:L38)——中文文本进来全 #N/A，分析表按 BL/BM 的 SUMIFS 整列漏计。
# 因此 append 前把语言相关列规范化回西语正典（字符串取自工作簿既有 1183 行 / 202609 报告）。
ML_MONTHS = {
    "enero": "enero", "febrero": "febrero", "marzo": "marzo", "abril": "abril",
    "mayo": "mayo", "junio": "junio", "julio": "julio", "agosto": "agosto",
    "septiembre": "septiembre", "octubre": "octubre", "noviembre": "noviembre",
    "diciembre": "diciembre",
    "一月": "enero", "二月": "febrero", "三月": "marzo", "四月": "abril",
    "五月": "mayo", "六月": "junio", "七月": "julio", "八月": "agosto",
    "九月": "septiembre", "十月": "octubre", "十一月": "noviembre", "十二月": "diciembre",
}
# 状态(c3) 中文 → 西语正典（静态映射）。目标有效性=工作簿缓存实况：
# Entregado/En camino/Procesando/En punto de retiro→1（靠D2命中Llegó/Llega/retirar el），
# Cancelada/Venta cancelada/Devolución*/Paquete cancelado→0（C2直接命中K20/K22/K21/K26）
ML_STATUS_MAP = {
    "已送达": "Entregado",
    "在路上": "En camino",
    "在仓库处理中": "Procesando en la bodega",
    "买家取消": "Cancelada por el comprador",
    "在取件点": "En punto de retiro",
    "销售已取消": "Venta cancelada",
    "退货已审核。商品已重新上架。": "Devolución revisada. Pusimos el producto de nuevo a la venta",  # 0 (K21)
    "退货准备中": "Devolución en preparación",      # 0 (K21)
    "退货运输中": "Devolución en camino. Revisaremos el producto",              # 0 (K21)
    "退货运输中。我们将检查商品。": "Devolución en camino. Revisaremos el producto",  # 0 (K21)
    "配送已重新安排": "Envío reprogramado",          # 1 (K25)
    "Mercado Libre取消的包裹": "Paquete cancelado por Mercado Libre",            # 0 (K26 cancelado)
    "包含 2 件商品的包裹": "Paquete de 2 productos",   # 正典实况（14 行）: 非取消态，SKU=' ' 由公式守卫判 0
}
# 状态(c3) 动态模板：{d}=文本中的数字, {snap}=快照日西语日期('10 de octubre')
# 有效性与正典对照：'Enviaremos el paquete el 21 de septiembre'缓存=1（靠D2 'Para organizar los
# envíos…'命中K4 envío）；'Para enviar el 24 de septiembre'缓存=1（同D2族）
ML_STATUS_PATTERNS = [
    (r"^我们将在(?P<d>\d{1,2}) 十月发出包裹$", "Enviaremos el paquete el {d} de octubre"),
    (r"^标签已准备好，可打印$", "Para enviar el {snap}"),
    (r"^已延迟 (?P<d>\d{1,2}) 天$", "Está demorado {d} día{s}"),   # 缓存实况=0（D2命中K36）
]
# 状态描述(c4) 正则 → 西语模板（命名组 {d}{d2}{wd}{zhm}，{m}=行内中文/西语月名→西语，{snap}=快照日）。
# 未命中的 c4 保留中文（c65 会回退查 c3）；已列入的规则保证有效性=该状态的缓存语义。
_ZHM = r"(?P<zhm>十月|十一月|十二月|一月|二月|三月|四月|五月|六月|七月|八月|九月)"
ML_DESC_PATTERNS = [
    # 已送达: '该包裹于5月十月日送达。' → 'Llegó el 5 de octubre' → K3 'Llegó' → 1
    (r"^该包裹于(?P<d>\d{1,2})月" + _ZHM + r"日送达。?$", "Llegó el {d} de {m}"),
    # 在途送达时间 → K5 'Llega' → 1（正典 'Llega mañana'/'Llega el lunes 21 de septiembre'）
    (r"^明日送达。?$", "Llega mañana"),
    (r"^今日送达。?$", "Llega hoy"),
    (r"^今明两天内送达。?$", "Llega entre hoy y mañana"),
    # '十月10日至11日期间送达。' → 'Llega entre el 10 y el 11 de octubre'（正典 'Llega entre el 21 y 22 de septiembre'）
    (r"^" + _ZHM + r"(?P<d>\d{1,2})日至(?P<d2>\d{1,2})日期间送达。?$", "Llega entre el {d} y el {d2} de {m}"),
    # '星期一月12日，十月送达。'/'星期日月11日送达。' → 'Llega el lunes 12 de octubre'
    (r"^星期(?P<wd>.)月(?P<d>\d{1,2})日，?" + _ZHM + r"?送达。?$", "Llega el {wd} {d} de {m}"),
    # '将于星期一月12日抵达发货中心。' → 'Llegará…' 前缀命中 K5 → 1
    (r"^将于星期(?P<wd>.)月(?P<d>\d{1,2})日抵达发货中心。?$", "Llegará al centro de despacho el {wd} {d} de {m}"),
    # 在途-抵达发货中心家族（正典 r1786/1801 'Llegará al centro de despacho el …' → K5 'Llega' → 1）
    (r"^今日抵达发货中心。?$", "Llega hoy al centro de despacho"),
    (r"^明日抵达发货中心。?$", "Llega mañana al centro de despacho"),
    # '预计十月13日至14日抵达发货中心。' → 'Llegará entre el 13 y el 14 de octubre al centro de despacho'
    (r"^预计" + _ZHM + r"(?P<d>\d{1,2})日至(?P<d2>\d{1,2})日抵达发货中心。?$",
     "Llegará entre el {d} y el {d2} de {m} al centro de despacho"),
    # '预计今日 16 点至 21 点送达。' → 'Llega hoy entre las 16 y las 21'（K5 'Llega' → 1）
    (r"^预计今日 (?P<h1>\d{1,2}) 点至 (?P<h2>\d{1,2}) 点送达。?$",
     "Llega hoy entre las {h1} y las {h2}"),
    # '买家选择了标准运输，标签将…'（截断）→ 正典 'Para organizar los envíos…'（K4 'envío' → 1）
    (r"^买家选择了标准运输，标签将.*$",
     "Para organizar los envíos, el comprador eligió el envío estándar."),
    # '最晚 15 de 九月 送达。' → 'Llega como máximo el 15 de septiembre.'（正典 r591）
    (r"^最晚 (?P<d>\d{1,2}) de " + _ZHM + r" 送达。?$", "Llega como máximo el {d} de {m}."),
    # 在取件点: '您的买家需在 Agencia Mercado Libre - X 取件。' → K12 'retirar el' → 1（正典 202609）
    (r"^您的买家需在 Agencia Mercado Libre - (?P<d>.+?) 取件。?$",
     "Tu comprador debe retirar el paquete en Agencia Mercado Libre - {d}."),
    # 我们将在N日发出包裹: '为协调发货，我们将于12月的十月日寄出。' → K4 'envío' → 1（正典 r423）
    (r"^为协调发货，我们将于(?P<d>\d{1,2})月的" + _ZHM + r"日寄出。?$",
     "Para organizar los envíos, enviaremos el {d} de {m}."),
    # 标签已准备好: '您需要在今天 15 点之前将包裹送至 Mercado Libre 服务点，以免延误。'
    #   → 正典 r671 'Para organizar los envíos, podrás imprimir la etiqueta el …'(K4 → 1)
    (r"^您需要在今天 \d{1,2} 点之前将包裹送至 Mercado Libre 服务点，以免延误。?$",
     "Para organizar los envíos, podrás imprimir la etiqueta el {snap}."),
    # 已延迟(卖家未发货, 缓存=0): 两种中文变体 → K36 'Despacha el paquete cuanto antes en una agencia de' → 0
    (r"^请尽快(将包裹送至|在\s?Mercado ?Libre物流点发出包裹).{0,30}(影响您的信誉|信誉已不再受影响)。?$",
     "Despacha el paquete cuanto antes en una agencia de Mercado Libre. La demora está afectando tu reputación."),
    # 买家取消描述 → K29 'La canceló' → 0（C2 已兜底，翻译只为展示一致）
    (r"^(买家)?已取消因买家取消购买$", "La canceló porque se arrepintió de comprarlo."),
    (r"^(买家)?已取消并指定了其他问题。?$", "La canceló y especificó otro problema."),
    (r"^(买家)?已取消因无法等待商品$", "La canceló porque no podía esperar al producto."),
    # ML取消包裹描述（C2 'Paquete cancelado por Mercado Libre' 兜底 K26 → 0）
    (r"^我们已取消此交易。如有疑问请查看帮助$",
     "Tuvimos que cancelar esta venta. Si tienes dudas, revisa la ayuda."),
    # 退货审核描述 → K24 'reembolsamos' → 0（C2 K21 兜底）
    (r"^我们已验证商品正确且状况良好。已向买家退款。$",
     "Validamos que es el producto correcto y está en buenas condiciones. Le reembolsamos el dinero al comprador."),
]
ML_WEEKDAYS = {"一": "lunes", "二": "martes", "三": "miércoles", "四": "jueves",
               "五": "viernes", "六": "sábado", "日": "domingo"}
# 简单值列映射（c5 多产品组合装 / c6 套装 / c18 广告 / c25 发票 / c41 运输方式）
ML_VALUE_MAP = {
    "是": "Sí", "内": "No",       # ML中文导出把 'No' 机翻成 '内'
    "未附发票": "Factura no adjunta",
    "Mercado Envios Full全渠道物流": "Mercado Envíos Full",
    "邮局和派送点": "Correo y puntos de despacho",
}
