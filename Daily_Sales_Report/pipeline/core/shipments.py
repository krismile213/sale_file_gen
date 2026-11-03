# -*- coding: utf-8 -*-
"""货件报告（TSV）→ 各簿每日库存分析表「货件列」匹配与填写

背景与规则（用户 2026-10-13/14 口述）：
  · 货件报告放 trunk\\data\\数据归属周期\\{YYYYMM}\\货件\\ 下，可按地区分子文件夹
    （2026-10-14 起用户实际结构：欧盟+英国 / 北美 / interflow / 印度 / 日本 / 澳大利亚），
    递归扫描所有 *.tsv（亚马逊「货件详情报告」，tab 分隔：头部 key-value 区含
     「货件编号 / 名称(含发货日期与仓库代码) / 配送地址(仓库代码) / SKU 总数 / 商品总数」，
    空行后明细区表头为「卖家 SKU|商品名称|ASIN|FNSKU|外部编号|状况|备货方|预处理类型|贴标方|已发货」，
    数量取第 10 列「已发货」）。
  · 子文件夹（或仓库代码）→ 目标簿+分析表：
      欧盟+英国                → 欧盟英国簿，仓库分 英国（BHX4/LBA4/...）/ 欧洲（HAJ1/DTM2/...）
      北美                     → 北美加拿大簿，仓库分 美国/加拿大（YYZ*/YYC*/YOW*/YHM*）/墨西哥（MEX*）
      interflow                → Interflowd 三簿，仓库分 美（默认）/加/澳（XSY*/BWU*）/阿联酋（DXB* 等）
      日本/印度/澳大利亚        → 对应单站点簿
      根目录/未知文件夹         → 按仓库回落欧盟英国簿（旧约定）
  · 发货日期格式随来源：欧/印/澳=DD/MM/YYYY，北美/interflow=MM/DD/YYYY，日本=YYYY/MM/DD；
    中东按仓库分：沙特 RUH/JED/DMM=MM/DD/YYYY，阿联酋 AUH/DXB 等=DD/MM/YYYY；
    两种两位段都 ≤12 的歧义日期按上述约定解析。
  · 运输方式（按 书+分析表 定，不再按仓库猜）：
      欧盟英国簿：点名清单=空运，其余英国/德国=卡航
      北美簿 美/加=海运、墨=空运；Interflowd 美=空运、加=海运；其余站点=空运
  · 表结构：行1=货件号、行2=发货日期(公式解析行5)、行3=预计时效(公式)、
    行4=预计到货(公式)、行5=标签「M/D运输(+M/D到货)」、
    数据行(行7+)按 A 列 SKU 填该货件的数量；行6=SUBTOTAL 合计。
    空列位（行1 无货件号 且 行5 空或占位公式 =TEXT(...)）只需填行1/行5/数据行，
    行2/3/4 公式已预置会自动从行5 标签解析。
  · 对应月份只放对应月份的货件：按发货日期先后落到货件列（新列接在已有货件列之后）。
  · AFN 结论（2026-10-14 实证）：库存报告 afn-inbound-working/shipped/receiving
    只覆盖「亚马逊已登记入网」的货件（海运未入网段缺失、到 FC 未收完部分多计），
    不能替代手工货件列；G(在途=AFN) 与 H(调拨在途合计=手工未到货) 双轨保留，
    在途差额 AC=G-H 即两者对账。
"""
import os
import re
import csv
import datetime


# ---------------- 常量规则 ----------------

# 用户点名的空运货件号（其余按规则推断）
AIR_SHIPMENTS = {
    "FBA15MGF23RQ", "FBA15MHKDZC2", "FBA15MKTKXXD",          # 德国空运
    "FBA15M9V5DK8", "FBA15MGF5JS5", "FBA15MHJMWC9", "FBA15MKT97C8",  # 英国空运
}

# 英国仓库代码（其余欧洲仓归「欧洲-每日库存分析表」）
UK_WAREHOUSES = {
    "BHX1", "BHX2", "BHX3", "BHX4", "BHX7", "BHX8",          # 伯明翰
    "LBA4", "LBA7",                                          # 利兹
    "MAN1", "MAN2", "MAN3", "MAN4",                          # 曼彻斯特
    "LTN1", "LTN4", "LTN7",                                  # 卢顿
    "EMA1", "EMA2",                                          # 东米德兰兹
    "CWL1", "CWL2", "NCL1", "EUK5", "XUKA",
}

# 表名（欧盟英国簿）
UK_SHEET = "英国-每日库存分析表"
EU_SHEET = "欧洲-每日库存分析表"

# 目标簿文件名
BOOK_EU = "欧盟英国库存销售表.xlsx"
BOOK_NA = "北美加拿大库存销售表.xlsx"
BOOK_INT_NA = "Interflowd北美加拿大库存销售表.xlsx"
BOOK_INT_AE = "Interflowd阿联酋库存销售表.xlsx"
BOOK_INT_AU = "Interflowd澳大利亚库存销售表.xlsx"
BOOK_INT_SA = "Interflowd沙特库存销售表.xlsx"   # 2026-10-24：沙特仓 RUH8 独立成簿
BOOK_JP = "日本库存销售表.xlsx"
BOOK_IN = "印度库存销售表.xlsx"
BOOK_AU = "澳大利亚库存销售表.xlsx"
BOOK_SA = "沙特库存销售表.xlsx"
BOOK_AE = "阿联酋库存销售表.xlsx"
BOOK_ML = "美客多库存销售表.xlsx"        # 2026-10-14：美客多（MercadoLibre）PDF 入库准备单
BOOK_RO = "罗马尼亚库存销售表.xlsx"      # 2026-10-14：eMAG 罗马尼亚 PDF 送货单

# 「无对应簿」哨兵：该票货件不该落任何簿（2026-10-20）。
# 起因：interflow 账户发往沙特 RUH8 的 FBA15MHHH99G（77 件）被 int_target_of 兜底成"美国"，
# 写进 Interflowd北美簿的「美国-每日库存分析表」L 列 —— 美国库存报告里这些 SKU 的在途为 0，
# 这 77 件永远不会"到货"，在途永久挂账。
# 2026-10-20：沙特站业务没有统计出口 ⇒ 明确跳过并告警，不再静默算到别的簿。
# 2026-10-24：已建 Interflowd沙特库存销售表.xlsx（沙特仓 RUH8 有 116 件可售一直看不见），
#             int_target_of 的沙特分支改为返回 BOOK_INT_SA，SKIP 仅留给其它未建簿站点。
SKIP = "__NO_BOOK__"

ML_SHEET = "墨西哥-每日库存分析表"
RO_SHEET = "罗马尼亚-每日库存分析表"

# 仓库代码前缀 → 站点（北美/interflow 用）
# 2026-10-30：FBA19VFTFF83（配送地址 YXU1，加拿大伦敦仓）因缺 YXU 前缀被兜底成"美国"，
#             545 件写进美国表 I 列、加拿大的在途合计漏这 545 件。补齐其余加拿大机场码。
_CA_PREFIX = ("YYZ", "YYC", "YOW", "YHM", "YVR", "YEG", "YOO", "YGK", "YUL", "YXX",
              "YQR", "YWG", "YXU", "YKF", "YXE", "YYJ", "YQT", "YXS", "YXY")
_MX_PREFIX = ("MEX",)
_AU_PREFIX = ("XSY", "BWU", "MEL", "PER", "ADL", "OAK")
_AE_PREFIX = ("DXB", "DXF", "AUH", "XJB", "FJR")
_SA_PREFIX = ("RUH", "JED", "DMM")

# 货件列区间（行1=货件号、行5=标签、行6=SUBTOTAL、行7+=数量）——兜底默认，运行时按「在途差额」列实测
UK_FIRST_COL, UK_LAST_COL = 9, 28          # I .. AB
EU_FIRST_COL, EU_LAST_COL = 9, 38          # I .. AL
DATA_FIRST_ROW = 7

# 子文件夹 → (书文件名, 分析表名|None=按仓库解析, 发货日期格式)
FOLDER_RULES = {
    "北美": (BOOK_NA, None, "MDY"),
    "印度": (BOOK_IN, "印度-每日库存分析表", "DMY"),
    "日本": (BOOK_JP, "日本-每日库存分析表", "YMD"),
    "澳大利亚": (BOOK_AU, "澳大利亚-每日库存分析表", "DMY"),
    "欧盟+英国": (BOOK_EU, None, "DMY"),
    "欧盟英国": (BOOK_EU, None, "DMY"),
    "interflow": (None, None, "MDY"),
    # 10-14 新增：沙特/阿联酋独立簿；「中东」文件夹在 resolve_target 按仓库分流（RUH/JED/DMM=沙特，DXB*=阿联酋）
    "沙特": (BOOK_SA, "沙特阿拉伯-每日库存分析表", "MDY"),
    "阿联酋": (BOOK_AE, "阿联酋-每日库存分析表", "MDY"),
    # 10-14 新增：两类 PDF 货件（非亚马逊，无仓库代码，按文件夹直接落簿）
    "美客多": (BOOK_ML, ML_SHEET, "MDY"),      # 美客多墨西哥入库准备单（PDF，无发货日期）
    "emag": (BOOK_RO, RO_SHEET, "DMY"),       # eMAG 罗马尼亚送货单（PDF，日期在单据内）
}

# 海运的 (书, 分析表)；其余（除欧盟簿点名空运外）= 空运
SEA_SHEETS = {
    (BOOK_NA, "美国-每日库存分析表"),
    (BOOK_NA, "加拿大-每日库存分析表"),
    (BOOK_INT_NA, "加拿大-每日库存分析表"),
}

# 明细区数量所在列（0-based，"已发货"）
QTY_IDX = 9
SKU_RE = re.compile(r"^[A-Z]+\d+[A-Z]*(-\d+)?$")   # G415 / K013-016 / G261IT


def col_letter(n):
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def parse_ship_date(name, fmt="DMY"):
    """从「名称」字段提取发货日期 → (月, 日)。
    支持三种格式：YYYY/MM/DD（日本）、DD/MM/YYYY（欧/印/澳）、MM/DD/YYYY（北美/interflow）。
    两位段均 ≤12 的歧义日期按 fmt 解析（默认 DMY）；含 >12 的段可自证。
    """
    if not name:
        return None
    m = re.search(r"\((\d{4})/(\d{1,2})/(\d{1,2})", name)
    if m:
        return int(m.group(2)), int(m.group(3))
    m = re.search(r"\((\d{1,2})/(\d{1,2})/(\d{4})", name)
    if not m:
        return None
    a, b = int(m.group(1)), int(m.group(2))
    if a > 12:
        return b, a          # DD/MM 明确
    if b > 12:
        return a, b          # MM/DD 明确
    if fmt == "MDY":
        return a, b
    return b, a              # 默认 DMY（欧洲/印度/澳/日本以外的兜底）


def read_shipment(path, fmt="DMY"):
    """解析单个货件 TSV → dict(id, wh, name, ship_md, nsku, nqty, items{sku:qty})。"""
    rows = []
    for enc in ("utf-8-sig", "gb18030", "latin-1"):
        try:
            with open(path, encoding=enc, newline="") as fh:
                rows = list(csv.reader(fh, delimiter="\t"))
            break
        except Exception:
            continue
    d = dict(id=os.path.splitext(os.path.basename(path))[0], wh="", name="",
             ship_md=None, nsku=None, nqty=None, items={}, names={})
    in_detail = False
    for r in rows:
        if not r:
            continue
        k = r[0]
        if k == "货件编号":
            d["id"] = (r[1] if len(r) > 1 else "").strip() or d["id"]
        elif k == "名称":
            d["name"] = r[1] if len(r) > 1 else ""
            d["ship_md"] = parse_ship_date(d["name"], fmt)
        elif k == "配送地址":
            d["wh"] = (r[1] if len(r) > 1 else "").strip()
        elif k == "SKU 总数":
            d["nsku"] = r[1] if len(r) > 1 else None
        elif k == "商品总数":
            d["nqty"] = r[1] if len(r) > 1 else None
        elif k == "卖家 SKU":
            in_detail = True
            continue
        if in_detail and SKU_RE.match(k.strip()):
            try:
                q = float(r[QTY_IDX])
            except (IndexError, ValueError):
                q = 0
            d["items"][k.strip()] = d["items"].get(k.strip(), 0) + q
            if len(r) > 1 and r[1].strip():
                d["names"].setdefault(k.strip(), r[1].strip())
    return d


def target_scope(wh):
    """仓库代码 → ('英国'|'欧洲', 分析表名)——欧盟英国簿专用。"""
    if wh.upper() in UK_WAREHOUSES:
        return "英国", UK_SHEET
    return "欧洲", EU_SHEET


def _wh_of(wh):
    return (wh or "").upper().strip()


# ---------------- 文件夹 × 仓库 交叉校验（2026-10-20） ----------------
# 目的：抓「TSV 放错文件夹」—— 这类错误比"仓库代码不认识"更常见，而且以前会静默落错簿
# （如沙特 RUH8 放进 interflow → 兜底成美国）。只做「仓库明确属于别国」的反向校验，
# 美国/欧洲仓代码无法穷举（未知前缀一律不判），因此不会误报。
_REGION_PREFIX = [
    ("CA", _CA_PREFIX),
    ("MX", _MX_PREFIX),
    ("AU", _AU_PREFIX),
    ("AE", _AE_PREFIX),
    ("SA", _SA_PREFIX),
    ("JP", ("XJE", "NRT", "KIX", "FSZ", "OKJ")),
    ("IN", ("BOM", "ISK", "DEL", "BLR", "HYD", "MAA", "PNQ")),
]

# 文件夹 → 允许的地区集合（未登记的文件夹不校验）
_FOLDER_REGIONS = {
    "北美": {"US", "CA", "MX"},
    "interflow": {"US", "CA", "AU", "AE", "SA"},
    "中东": {"AE", "SA"},
    "欧盟+英国": {"UK", "EU"},
    "欧盟英国": {"UK", "EU"},
    "日本": {"JP"},
    "印度": {"IN"},
    "澳大利亚": {"AU"},
    "沙特": {"SA"},
    "阿联酋": {"AE"},
    "美客多": {"MX"},
    "emag": {"RO"},
}


def region_of(wh):
    """仓库代码 → 地区标签（CA/MX/AU/AE/SA/JP/IN/UK）；无法判定返回 None。"""
    w = _wh_of(wh)
    if w in UK_WAREHOUSES:
        return "UK"
    for tag, prefixes in _REGION_PREFIX:
        if w.startswith(prefixes):
            return tag
    return None


_REGION_CN = {"US": "美国", "CA": "加拿大", "MX": "墨西哥", "AU": "澳大利亚",
              "AE": "阿联酋", "SA": "沙特", "JP": "日本", "IN": "印度",
              "UK": "英国", "EU": "欧洲", "RO": "罗马尼亚"}


def sid_label(s):
    """货件号 + 地区说明（**仅特殊情况用**，正常票仍只写货件号）：
    `FBA15MHHH99G（RUH8沙特）`。仓库无法判国时只带仓库代码，如 `FBA19QQ8H4L2（HEA2）`。"""
    wh = _wh_of(s.get("wh"))
    if not wh:
        return str(s.get("id"))
    cn = _REGION_CN.get(region_of(wh) or "", "")
    return f"{s.get('id')}（{wh}{cn}）"


def cross_region_warn(wh, folder):
    """文件夹与仓库地区不符 → 警告文本；正常返回 None。只警告，不改归属。
    文本刻意不重复仓库/地区（调用方用 sid_label 已经带上了）。"""
    w = _wh_of(wh)
    # OAK 歧义：_AU_PREFIX 里含 OAK，但美国奥克兰仓也是 OAK3/OAK4。
    # 实测 202608–202610 全部货件无 OAK*，暂不影响；出现时提醒人工确认。
    if w.startswith("OAK"):
        return f"仓库 {wh} 前缀 OAK 有歧义（澳洲 OAK* vs 美国奥克兰 OAK3/4），当前按澳洲处理，请确认"
    fd = (folder or "").strip().lower()
    allowed = _FOLDER_REGIONS.get(fd)
    if not allowed:
        return None
    reg = region_of(wh)
    if reg and reg not in allowed:
        return f"放在「{folder or '根目录'}」文件夹，但仓库属于{_REGION_CN.get(reg, reg)} → 可能放错文件夹"
    return None


def na_sheet_of(wh):
    """北美簿仓库 → 分析表名。"""
    w = _wh_of(wh)
    if w.startswith(_MX_PREFIX):
        return "墨西哥-每日库存分析表"
    if w.startswith(_CA_PREFIX):
        return "加拿大-每日库存分析表"
    return "美国-每日库存分析表"


def int_target_of(wh):
    """interflow 仓库 → (书文件名, 分析表名)。无对应簿时返回 (SKIP, None)。

    ⚠️ 沙特前缀（RUH/JED/DMM）必须显式拦下：Interflowd 没有沙特簿，若落到兜底分支
    会被算成"美国"（2026-10-20 FBA15MHHH99G 事故）。返回 SKIP 由 load_shipments 跳过。
    """
    w = _wh_of(wh)
    if w.startswith(_CA_PREFIX):
        return BOOK_INT_NA, "加拿大-每日库存分析表"
    if w.startswith(_AU_PREFIX):
        return BOOK_INT_AU, "澳大利亚-每日库存分析表"
    if w.startswith(_AE_PREFIX):
        return BOOK_INT_AE, "阿联酋-每日库存分析表"
    if w.startswith(_SA_PREFIX):
        return BOOK_INT_SA, "沙特-每日库存分析表"
    return BOOK_INT_NA, "美国-每日库存分析表"


def resolve_target(folder, wh):
    """子文件夹 + 仓库代码 → (书文件名|None, 分析表名|None, 发货日期格式)。
    未知文件夹/根目录 → (None, None, "DMY")，由调用方回落欧盟英国簿旧规则。
    """
    fd = (folder or "").strip().lower()
    if fd == "中东":
        w = _wh_of(wh)
        if w.startswith(_AE_PREFIX):
            # 阿联酋（AUH/DXB/DXF/XJB/FJR）源报告用 DD/MM/YYYY（实测 29/09、06/10、13/10）
            return BOOK_AE, "阿联酋-每日库存分析表", "DMY"
        # 沙特（RUH/JED/DMM）源报告用 MM/DD/YYYY（实测 09/22、09/29、10/06、10/13）
        return BOOK_SA, "沙特阿拉伯-每日库存分析表", "MDY"
    rule = FOLDER_RULES.get(fd)
    if rule is None:
        return None, None, "DMY"
    book, sheet, fmt = rule
    if book is not None and sheet is not None:
        return book, sheet, fmt
    if book == BOOK_NA:
        return BOOK_NA, na_sheet_of(wh), fmt
    if book == BOOK_EU:
        return BOOK_EU, target_scope(wh)[1], fmt
    # book 为 None：interflow，按仓库落到三本 Int 簿
    b2, s2 = int_target_of(wh)
    return b2, s2, fmt


def transport_for(book, sheet, sid):
    """(书, 分析表, 货件号) → (运输方式, 备注)。"""
    if book == BOOK_EU:
        return guess_transport(sid, None, None)
    if (book, (sheet or "").strip()) in SEA_SHEETS:
        return "海运", ""
    return "空运", ""


def guess_transport(sid, scope, total_qty):
    """运输方式：点名清单=空运；英国/欧洲（德）其余=卡航；数量启发式给出提示。"""
    if sid in AIR_SHIPMENTS:
        return "空运", "点名"
    note = ""
    if total_qty and total_qty < 500:
        note = "件数少(疑似空运?)"
    return "卡航", note


def label_of(ship, transport, arrive_md=None, arrived=False):
    """行5 标签：'M/D运输'（到货后补 '到货'）。
    · arrive_md=(月,日) → 'M/D运输M/D到货'
    · arrived=True（PDF 货件用「到货.txt」标记，无确切到货日）→ 'M/D运输到货'
    """
    if not ship["ship_md"]:
        base = transport
    else:
        base = f"{ship['ship_md'][0]}/{ship['ship_md'][1]}{transport}"
    if arrive_md:
        return base + f"{arrive_md[0]}/{arrive_md[1]}到货"
    if arrived:
        return base + "到货"
    return base


def read_arrived(ship_dir):
    """读取「到货.txt」标记：货件目录（含子目录）下任一 `到货.txt`，每行一个条目，
    可为 文件名 / 文件名去扩展名 / 货件号，`#` 开头与空行忽略。返回小写条目集合。
    用途：PDF 货件（美客多/eMAG）文字层无「到货」字段，由该标记驱动行5 补「到货」。
    """
    keys = set()
    if not os.path.isdir(ship_dir):
        return keys
    for root, _dirs, files in os.walk(ship_dir):
        for fn in files:
            if fn.strip().lower() != "到货.txt":
                continue
            try:
                with open(os.path.join(root, fn), encoding="utf-8-sig") as fh:
                    for line in fh:
                        s = line.strip()
                        if not s or s.startswith("#"):
                            continue
                        keys.add(s.lower())
                        keys.add(os.path.splitext(s)[0].lower())
            except OSError:
                continue
    return keys


def load_shipments(ship_dir):
    """递归读取货件目录（含地区子文件夹）所有 TSV **与 PDF**，
    解析并映射到 书+分析表+运输方式。每个 dict 附 folder/book/sheet/total/transport/src/arrived。

    ⚠️ 同一票货件重复下载会得到「FBAxxxx.tsv」+「FBAxxxx (1).tsv」两个文件，id（取 TSV 里的
    「货件编号」字段）相同 → 若不去重，plan_columns 会把它当两票各占一个空列位，数量被复制
    到两列、合计翻倍（2026-10-14 实测 interflow 下确有 2 对）。按 id 去重，优先保留不带
    「 (1)」后缀的原文件名。
    """
    out = []
    if not os.path.isdir(ship_dir):
        return out
    arrived_keys = read_arrived(ship_dir)
    for root, _dirs, files in os.walk(ship_dir):
        folder = os.path.relpath(root, ship_dir)
        if folder == ".":
            folder = ""
        for fn in sorted(files):
            low = fn.lower()
            is_tsv = low.endswith(".tsv")
            is_pdf = low.endswith(".pdf")
            if not (is_tsv or is_pdf):
                continue
            p = os.path.join(root, fn)
            fmt0 = FOLDER_RULES.get(folder.strip().lower(), (None, None, "DMY"))[2]
            if is_tsv:
                s = read_shipment(p, fmt=fmt0)
            else:
                import ship_pdf
                s = ship_pdf.read_pdf_shipment(p)
                if s is None:
                    continue
            if not (s["id"] and s["items"]):
                continue
            s["src"] = fn
            s["folder"] = folder
            book, sheet, fmt = resolve_target(folder, s["wh"])
            if book == SKIP:
                # 无对应簿（interflow 账户发往沙特 RUH/JED/DMM）：不落任何簿。
                # 仍保留在返回列表里（book=SKIP），供 runner 告警与 precheck 统计。
                s.update(book=SKIP, sheet=None, total=sum(s["items"].values()),
                         transport="", note="无对应簿（沙特站）", arrived=False,
                         warn=cross_region_warn(s["wh"], folder))
                out.append(s)
                continue
            if book is None:
                # 根目录/未知文件夹：回落欧盟英国簿旧规则（按仓库分英/欧）
                book, sheet = BOOK_EU, target_scope(s["wh"])[1]
            s["book"], s["sheet"] = book, sheet
            s["warn"] = cross_region_warn(s["wh"], folder)
            # TSV：用「子文件夹 + 仓库」解析出的格式重算发货日期。
            # ⚠️ 不能只在 ship_md 为 None 时补算：`中东` 文件夹按仓库分格式（沙特 MM/DD、
            # 阿联酋 DD/MM），而上面的 fmt0 取的是文件夹级默认（中东未登记 → DMY），
            # 两段均 ≤12 的歧义日期会按错格式定格（如 RUH8 的 10/06 被读成 6/10）。
            # PDF 的日期来自单据本身，仅在缺时兜底，不重算。
            if is_tsv:
                s["ship_md"] = parse_ship_date(s["name"], fmt)
            elif s["ship_md"] is None:
                s["ship_md"] = parse_ship_date(s["name"], fmt)
            s["total"] = sum(s["items"].values())
            s["transport"], s["note"] = transport_for(book, sheet, s["id"])
            s["arrived"] = (fn.lower() in arrived_keys
                            or os.path.splitext(fn)[0].lower() in arrived_keys
                            or str(s["id"]).strip().lower() in arrived_keys)
            out.append(s)
    # 同一 id 去重：优先「无 (1) 后缀」的原文件名（空格+括号是系统重名下载的产物）
    uniq = {}
    for s in out:
        k = str(s["id"]).strip()
        prev = uniq.get(k)
        if prev is None or ("(" in prev["src"] and "(" not in s["src"]):
            uniq[k] = s
    return list(uniq.values())


def sanitize_shipment_cols(wb):
    """货件列数据区（行 7+）只应放数量；任何残留文本都会把 H=调拨在途合计 打成 #VALUE!
    （H 公式对每个货件列做 `IF(列5含"到货",0,IF(格="",0,格))` 累加，文本 + 数字 = #VALUE!，
    且 H6 的 SUBTOTAL 会跟着整列变 #VALUE!）。

    实例：2026-10-14 Interflowd北美加拿大簿的 美国/加拿大 分析表 V77:V80 混入文本「钢化膜」
    （该列本是空货件列），导致 H77:H80 与 H6/AC6 全 #VALUE!。全 12 本实测仅此 7 格，
    货件列数据区不存在合法文本 ⇒ 一律清空是安全的。
    列范围按行 5「在途差额」实测（与 plan_columns 同源）。返回清理格数。
    """
    n = 0
    for sn in list(wb.sheetnames):
        if "分析表" not in sn:
            continue
        ws = wb[sn]
        fc, lc = detect_sheet_cols(ws)
        for r in range(DATA_FIRST_ROW, ws.max_row + 1):
            for c in range(fc, lc + 1):
                v = ws.cell(row=r, column=c).value
                if isinstance(v, str) and v.strip() != "":
                    ws.cell(row=r, column=c).value = None
                    n += 1
    return n


def detect_sheet_cols(ws, default=(9, 28)):
    """按行5「在途差额」标签定位货件列范围 → (first, last=差额列-1)。找不到用 default。"""
    for c in range(9, 60):
        v = ws.cell(5, c).value
        if isinstance(v, str) and "在途差额" in v:
            return 9, c - 1
    return default


def plan_columns(ws, ships, first_col, last_col):
    """为每个货件确定目标列：行1 已有货件号 → 复用；否则顺序占用空列位。

    空列位判据（2026-10-14 用户拍板方案① 起）：
      ① 行1 无货件号 且（行5 无标签 或 行5 是占位公式）—— 真空白列位；
      ② **行5 含「到货」的已完成货件列** —— 该列早已从 G/H 剔除（H 公式对含「到货」
         的列计 0），使命完成 → 回收复用（覆盖行1/行5/数据）。
    ②的排除项：若该列旧货件号**仍是本批要落的票**（用户重放了同一票 TSV），则该列按
    「按 id 复用」处理，不计入空位，避免同一列被两张票抢用。

    返回 (assign{货件号:列号}, existing{货件号:列号}, new{货件号:列号}, overflow[..],
         recycle{列号})。recycle = 被回收的「已到货」列，apply_sheet 需先清其数据区
    （旧票的 SKU 行与新票不同，不清会把旧数量残留在新票没覆盖到的行里）。

    分配顺序（2026-10-14）：**先占真空白列位，再回收「已到货」列**——空位够时不动历史
    货件列（保留台账），只有空位用尽才覆盖最早的已完成列。
    """
    row1 = {}                       # 列号 -> 货件号（统一成 str：货件号在行1 可能是 int）
    row5 = {}
    for c in range(first_col, last_col + 1):
        v = ws.cell(1, c).value
        if v is not None and str(v).strip():
            row1[c] = str(v).strip()
        row5[c] = ws.cell(5, c).value
    id_to_col = {v: c for c, v in row1.items()}
    cur_ids = {str(s["id"]).strip() for s in ships}
    existing, free_empty, free_recycle, arrived_cols = {}, [], [], set()
    for c in range(first_col, last_col + 1):
        v5 = row5.get(c)
        # 占位公式列（如 =TEXT(数据透视图!$G$1,"M/D海运")）视同空列位
        placeholder = isinstance(v5, str) and v5.startswith("=")
        arrived = isinstance(v5, str) and "到货" in v5
        old_id = row1.get(c)
        if old_id is not None and old_id in cur_ids:
            continue                # 本批要复用的在途票 → 已用，不算空位
        if arrived:
            free_recycle.append(c)  # 已到货旧列 → 回收复用（备用池）
            arrived_cols.add(c)
        elif old_id is None and (not v5 or placeholder):
            free_empty.append(c)    # 真空白列位（优先用）
    assign, new, overflow = {}, {}, []
    free = list(free_empty) + list(free_recycle)
    free_iter = iter(free)
    # 按发货日期排序分配（对应月份先后落列）
    for s in sorted(ships, key=lambda x: (x["ship_md"] or (99, 99), x["id"])):
        if s["id"] in id_to_col:
            assign[s["id"]] = id_to_col[s["id"]]
            existing[s["id"]] = id_to_col[s["id"]]
        else:
            try:
                c = next(free_iter)
            except StopIteration:
                overflow.append(s["id"])
                continue
            assign[s["id"]] = c
            new[s["id"]] = c
    # recycle 只报「真正被本批新票占用的已到货列」（未用到的历史列不动、不清）
    recycle = {c for c in new.values() if c in arrived_cols}
    return assign, existing, new, overflow, recycle


def _sku_rows(ws, max_row=None):
    """A 列 SKU → 行号。"""
    m = {}
    last = max_row or ws.max_row
    for r in range(DATA_FIRST_ROW, last + 1):
        v = ws.cell(r, 1).value
        if v is not None and str(v).strip():
            m[str(v).strip()] = r
    return m


def compare_existing(ws, ship, col):
    """比对「表内人工值(col 列)」vs「TSV 解析值」。返回 (n_both, diffs, only_sheet, only_tsv)。"""
    sku_rows = _sku_rows(ws)
    sheet_vals = {}
    for sku, r in sku_rows.items():
        v = ws.cell(r, col).value
        if isinstance(v, (int, float)) and v:
            sheet_vals[sku] = float(v)
    mine = ship["items"]
    both = set(mine) & set(sheet_vals)
    diffs = [(k, mine[k], sheet_vals[k]) for k in both if abs(mine[k] - sheet_vals[k]) > 1e-6]
    return len(both), diffs, sorted(set(sheet_vals) - set(mine)), sorted(set(mine) - set(sheet_vals))


def _subtotal_end_row(ws):
    """从已有的行6 SUBTOTAL 公式里取覆盖末行（O6 `=SUBTOTAL(9,O7:O2838)` → 2838）。"""
    for c in range(1, min(ws.max_column, 60) + 1):
        v = ws.cell(6, c).value
        if isinstance(v, str):
            m = re.search(r"SUBTOTAL\(\s*9\s*,\s*[A-Z]{1,2}\d+\s*:\s*[A-Z]{1,2}(\d+)\s*\)", v)
            if m:
                return int(m.group(1))
    return 3000


def _ensure_subtotal(ws, c, last_row):
    """保证货件列行6 有合计公式。

    openpyxl 不支持插列平移公式，历史扩容是手工插列 → 新列没有 SUBTOTAL，
    合计行会漏计这批货件（2026-10-30 实测：美国簿 I/J 两列只有数据没有合计）。
    """
    if ws.cell(6, c).value not in (None, ""):
        return False
    from openpyxl.utils import get_column_letter
    end = max(_subtotal_end_row(ws), last_row + 500)
    col = get_column_letter(c)
    ws.cell(6, c).value = f"=SUBTOTAL(9,{col}{DATA_FIRST_ROW}:{col}{end})"
    return True


def apply_sheet(ws, ships, assign, new, recycle=()):
    """填写货件列：行1=货件号、行5=标签、数据行=数量。
    新列：行1/行5/数据行全写；已存在列：**只补空格**（不覆盖任何已有值，保护人工修正）。

    ⚠️ 「只补空格」是必需的：SKU 行常在货件列写好之后才补上（非欧盟簿不自动追行），
    若复用列一律跳过，后补行的数量就永远填不进来（2026-10-14 实测：补完 G594-G598 行后
    重跑，票数量仍空）。空格补写不覆盖任何已有值，故幂等且安全。

    ⚠️ recycle（复用「已到货」列，方案①）例外：这些列要**先清数据区再写**。旧票的
    SKU 行与新票不同，不清会把旧数量残留在新票没覆盖到的行里（脏值）。
    返回 [(货件号, 列, 标签, 写入格数, 未命中SKU数, 回填格数)]。
    """
    recycle = set(recycle or ())
    sku_rows = _sku_rows(ws)
    last_row = max(sku_rows.values()) if sku_rows else DATA_FIRST_ROW
    out = []
    for s in ships:
        c = assign.get(s["id"])
        if not c:
            continue
        is_new = s["id"] in new
        if is_new:
            if c in recycle:
                for r in range(DATA_FIRST_ROW, last_row + 1):
                    if ws.cell(r, c).value is not None:
                        ws.cell(r, c).value = None
            ws.cell(1, c).value = s["id"]
            ws.cell(5, c).value = s["label"]
            _ensure_subtotal(ws, c, last_row)
        filled = miss = refill = 0
        for sku, q in s["items"].items():
            r = sku_rows.get(sku)
            if r is None:
                miss += 1
                continue
            if is_new:
                ws.cell(r, c).value = q
                filled += 1
            else:
                cur = ws.cell(r, c).value
                if cur is None or (isinstance(cur, str) and cur.strip() == ""):
                    ws.cell(r, c).value = q
                    refill += 1
        out.append((s["id"], c, s["label"], filled, miss, refill))
    return out

