# -*- coding: utf-8 -*-
"""每日库存分析表维护：缺失 SKU 追加行 / 月度销量封存 / 「当月总销量」表头

规则（2026-10-13 用户确认）：
  · 分析表（英国-每日库存分析表 / 欧洲-每日库存分析表）结构：
      行1=货件号，行2=发货日期(公式)，行3=预计时效(公式)，行4=预计到货(公式)，
      行5=标签/列头，行6=合计，行7+=数据（A=SKU）
      E=预留库存 F=可售库存 G=在途 H=调拨在途合计
      I..=货件列（英国到 AB / 欧洲到 AL）  AC/AM=在途差额  AD/AN=实时速度
      AE/AO=上月日均  AF/AP=在售可售天数  AG/AQ=在售在途可售天数
      AH/AR=「当月总销量」(公式随锚点滚)  AI../AS..=每日销量
      右侧月度汇总区：英国 BV(种类)+BW..CH(12月..1月)；欧洲 CF+CG..CR
  · 9 月已结束 → 把 9 月销量**封存**进月度区的「9月」列（英国 BZ / 欧洲 CJ）。
    口径与日销量一致：有效性(AQ=1) 且 英国按站点 G="Amazon.co.uk"、欧洲按地区 AR="欧洲"。
  · AH/AR 表头原写死「9月总销量」但公式实际随 数据透视表!B11 滚动（现算 10 月）→ 改为
    `=MONTH(数据透视表!B11-1)&"月总销量"` 自动跟随。
"""
import datetime
import re as _re2
from openpyxl.utils import get_column_letter as _gl

UK_SHEET = "英国-每日库存分析表"
EU_SHEET = "欧洲-每日库存分析表"

# 库存源 sheet 里这些列名视为「listing 是否活着」
# 2026-10-23 补中文列名：亚马逊日本/澳大利亚、interflow澳大利亚 的管理库存报告
# 表头是中文，但**这两份报告本身没有 listing 列**（是 afn-onhand-buyable 式），
# 所以实际仍走「无列 → 出现即视为在售」。这里补候选是为将来中文版标准报告留位，
# 避免哪天真出现「上架状态」列却被当成无列、把未上架的 SKU 也补进来。
_LISTING_COLS = ("afn-listing-exists", "afn_listing_exists",
                 "上架状态", "是否上架", "listing状态", "listing 状态")
# SKU 形态：G598 / K013-018 / SKU139 这类，用来过滤「Info」「Mandatory」说明行
# SKU 格式：1~3 个大写字母 + 3~4 位数字 + **可选的单个字母后缀** + 可选后缀段。
# 2026-10-24 放宽：原正则不认 G136R / G136B / G590J 这类「字母后缀」，
# 导致库存源里出现这种 SKU 时永远补不进表（用户拍板「正则可以打开」）。
# 安全性由补行闸兜底：没库存、没销量的 SKU 一律不补，放宽正则不会引入垃圾行。
_SKU_RE = _re2.compile(r"^[A-Z]{1,3}\d{3,4}[A-Z]{0,3}([-_ ].*)?$")

DATA_FIRST_ROW = 7

# ---------- amzn.gr 别名 SKU 归并（2026-10-27）----------
# 亚马逊「全球开店」跨站 listing 会在订单/库存报告里冒出
#   amzn.gr.G192-0Zq_pBgolYmPYQlNGUC8BwDm-LN
# 这类**系统临时编号**：实物就是主 SKU G192 卖的。不归并的话：
#   · 月度封存把它当成「表里没有的 SKU」直接丢掉（实测 2026-09 英国少 8 件）；
#   · 补行证据闸也看不到这部分销量。
# 归并规则：去掉 `amzn.gr.` 前缀，再用项目自己的 SKU 形态取头部
#   （`[-_ ]\d+` 段要求后面是分隔符或结尾，避免把 `G192-0Zq…` 的 `-0` 当成 K013-016 式后缀）。
# 置 False 完全回退旧行为（别名当独立 SKU 处理）。
AMZN_GR_MERGE = True
_AMZN_GR_RE = _re2.compile(r"^amzn\.gr\.", _re2.I)
_SKU_HEAD_RE = _re2.compile(r"^([A-Z]{1,3}\d{3,4}[A-Z]{0,3}(?:[-_ ]\d+(?=[-_ ]|$))?)")


def normalize_sku(s):
    """`amzn.gr.<SKU>-<hash>[-LN|-VG]` → `<SKU>`；其余原样（去空白）。"""
    t = str(s or "").strip()
    if not AMZN_GR_MERGE or not _AMZN_GR_RE.match(t):
        return t
    m = _SKU_HEAD_RE.match(t[8:])
    return m.group(1) if m else t


# 月度汇总区的「9月」列
UK_MONTH9_COL = 78      # BZ
EU_MONTH9_COL = 88      # CJ
# 「当月总销量」列
UK_CUR_COL = 34         # AH
EU_CUR_COL = 44         # AR


def month_col(month, uk=True):
    """月度汇总区「N月」列号：英国 12月=BW(75)..1月=CH(86)；欧洲 12月=CG(85)..1月=CR(96)。"""
    return (75 if uk else 85) + (12 - int(month))

# 追加新行时要复制的公式列（相对引用，直接同列复制即可）
FORMULA_COLS = ["E", "F", "G", "H", "AC", "AD", "AF", "AG", "AH", "AI"]
# E/F/G/H/AC/AD/AF/AG/AH/AI.. 直到日列末（AI..BU / AS..CE）
UK_FORMULA_TAIL = (35, 73)     # AI .. BU
EU_FORMULA_TAIL = (45, 83)     # AS .. CE
UK_EXTRA_SINGLE = ["BV"]       # 种类
EU_EXTRA_SINGLE = ["CF"]


# 数据区「连续多少个空行」才算结束。
# 原来是「遇空即停」，2026-10-24 实测踩坑：用户手工清空中间某一行 A 列后，
# sku_set 只读到那一行就停 → 后面 375 个 SKU 全被当成「表里没有」→
# 下次跑批重复补 375 行。改成允许空洞，连续 30 行空才收尾。
_BLANK_STOP = 30


def _last_data_row(ws, max_scan=3000):
    """最后一个有 SKU 的数据行（允许中间有空洞）。"""
    last = DATA_FIRST_ROW - 1
    blank = 0
    end = min(ws.max_row or 0, DATA_FIRST_ROW + max_scan)
    for r in range(DATA_FIRST_ROW, end + 1):
        if ws.cell(r, 1).value is None:
            blank += 1
            if blank >= _BLANK_STOP:
                break
            continue
        blank = 0
        last = r
    return last


def sku_set(ws):
    """A 列已有 SKU 集合（允许中间有空洞，连续 30 行空才收尾）。"""
    out = set()
    blank = 0
    end = min(ws.max_row or 0, DATA_FIRST_ROW + 3000)
    for r in range(DATA_FIRST_ROW, end + 1):
        v = ws.cell(r, 1).value
        if v is None:
            blank += 1
            if blank >= _BLANK_STOP:
                break
            continue
        blank = 0
        sv = str(v).strip()
        if sv:
            out.add(sv)
    return out


def _plain_formula(v):
    """复制公式到新行时去掉「数组公式」属性。

    openpyxl 的 ArrayFormula 带着 ref（指向原格），整对象复制到新行会让 ref 失效
    （2026-10-16：阿联酋簿 F/G 被存成数组公式 → VLOOKUP(A:A,…) 退化成用第 1 行去查 → 全 0）。
    """
    try:
        from openpyxl.worksheet.formula import ArrayFormula
    except Exception:
        return v
    if isinstance(v, ArrayFormula):
        return "=" + str(v.text or "").lstrip("=")
    return v


# ---------------------------------------------------------------- 通用补行（全 12 簿）
# 2026-10-16 用户拍板：货件里出现、分析表 A 列没有的 SKU **一律自动补行**，
# 补在数据区末尾、按票内出现顺序（与既有约定一致：新 SKU 批量落表尾，如 G594..G598）。
# 旧 append_sku_rows 直接照抄公式字符串，行号不平移 → 仅欧盟簿在用，已弃用。
BRANDS = ("Sony", "Poco", "Redmi", "Xiaomi", "Motorola", "Samsung", "Apple",
          "iPhone", "Honor", "Oppo", "Vivo", "Realme", "Google", "OnePlus", "Nokia")


def _brand_col(ws, max_scan=140):
    """品牌列 = 行5 标签为「品牌」的那一列（BZ/其它），找不到返回 None。"""
    for c in range(1, min(ws.max_column, max_scan) + 1):
        v = ws.cell(5, c).value
        if isinstance(v, str) and v.strip() == "品牌":
            return c
    return None


def brand_of(name):
    """从品名猜品牌（英文优先，其次中文）。猜不到返回 ''。"""
    n = name or ""
    for b in ("Sony", "Poco", "Motorola", "Samsung", "Apple", "Honor", "Oppo",
              "Vivo", "Realme", "Google", "OnePlus", "Nokia", "iPhone"):
        if b.lower() in n.lower():
            return b
    if "Xiaomi" in n or "Redmi" in n or "红米" in n:
        return "Redmi"
    if "小米" in n:
        return "小米"
    if "苹果" in n:
        return "苹果"
    return ""


def _row_width(ws, r, maxc):
    """某行实际有多少列「有内容或样式」。"""
    n = 0
    for c in range(1, maxc + 1):
        cell = ws.cell(r, c)
        if cell.value is not None or cell.has_style:
            n = c
    return n


def pick_template_row(ws, last, maxc=None):
    """挑一个**列最完整**的行当克隆模板，而不是死认末数据行。

    2026-10-23 踩坑：补行照抄末行，而末行本身列残缺（英国表 G212/K001/K004/K005
    那批老行只有 55 格，只到 BV；欧洲表 K001 那群只到 CF/CT）→ 新补的 G598 也就只有
    50~54 格，缺了 AD/AE（在途差额/实时速度）、AF/AG（可售天数）、AS..CE（日销列）、
    CF（种类）以及 BW..DF / CG..DP 整片月度汇总区。表面看行在表里，实际上算不出数、
    条件格式也不覆盖。
    做法：从末行往上扫最多 400 行，取「有 SKU 且列数最多 / 基本打平」的第一行。
    """
    maxc = maxc or min(ws.max_column, 300)
    best, best_n = last, -1
    floor = DATA_FIRST_ROW
    for r in range(last, max(floor, last - 400), -1):
        if ws.cell(r, 1).value is None:
            continue
        n = _row_width(ws, r, maxc)
        if n > best_n:
            best, best_n = r, n
    return best


def clone_row(ws, src, dst, clear_range=None, max_col=None):
    """整行克隆 src→dst：
      · 公式 → Translator 平移行号（数组公式保留数组形态，ref 改指向新格）
      · **文本型**静态值 → 照抄（属性列：分类 C/状态 D/膜类型/品牌/判/机会产品）
      · **数值型**静态值 → 留空（月度销量 12月..1月、月初基线：新 SKU 无历史，不能继承）
      · clear_range=(fc,lc) → 该区间（货件列）留空，等 apply_sheet 回填
    样式一并照抄。返回 (公式格数, 文本格数)。
    """
    from copy import copy as _copy
    from openpyxl.formula.translate import Translator
    from openpyxl.worksheet.formula import ArrayFormula

    fc, lc = (clear_range or (None, None))
    maxc = max_col or min(ws.max_column, 300)
    nf = ns = 0
    for c in range(1, maxc + 1):
        tc, nc = ws.cell(src, c), ws.cell(dst, c)
        v = tc.value
        if fc is not None and fc <= c <= lc:
            nc.value = None
        elif isinstance(v, ArrayFormula):
            txt = Translator(v.text, origin=tc.coordinate).translate_formula(nc.coordinate)
            nc.value = ArrayFormula(ref=nc.coordinate, text=txt)
            nf += 1
        elif isinstance(v, str) and v.startswith("="):
            nc.value = Translator(v, origin=tc.coordinate).translate_formula(nc.coordinate)
            nf += 1
        elif isinstance(v, str):
            nc.value = v
            ns += 1
        else:
            nc.value = None
        try:
            nc._style = _copy(tc._style)
        except Exception:
            pass
    return nf, ns


def grow_cf_last_row(ws, old_end, new_end):
    """补行后把条件格式规则的**范围末行** old_end → new_end。

    openpyxl 对已存在的规则只能改 formula/priority，**改不动 sqref**（2026-10-16 踩坑：
    补 G598 行后粉色 CF 的 `A7:BK80` 没跟到 81，新行落在规则外）→ 必须重建容器。
    只动「末行正好等于 old_end」的区间，其它规则原样保留。返回改动数。
    """
    if new_end <= old_end:
        return 0
    from openpyxl.formatting.formatting import ConditionalFormattingList
    from openpyxl.worksheet.cell_range import CellRange

    changed = 0
    new_cfl = ConditionalFormattingList()
    for rng in list(ws.conditional_formatting):
        parts = []
        for r in str(rng.sqref).split():
            try:
                cr = CellRange(r)
            except Exception:
                parts.append(r)
                continue
            if cr.max_row == old_end:
                cr.max_row = new_end
                parts.append(str(cr))
                changed += 1
            else:
                parts.append(r)
        for rule in rng.rules:
            new_cfl.add(" ".join(parts), rule)
    if changed:
        ws.conditional_formatting = new_cfl
    return changed


def append_sku_rows_auto(ws, new_skus, name_map=None, clear_range=None, state_map=None):
    """全 12 簿通用补行：数据区末尾按给定顺序追加缺失 SKU 行。

    · 克隆**列最完整的那一行**（pick_template_row，不是死认末行）→ 保证新行带齐
      AD/AE/AF/AG、日销列、月度汇总区；公式平移 / 文本属性照抄 / 货件列清空
    · A=SKU，B=品名（name_map），品牌列=brand_of(品名)（可判时覆盖克隆来的旧品牌）
    · state_map={sku: "在售"/"在途"/"未上架"} → 覆盖 D 列状态
      （clone_row 会照抄模板行，必须显式改）
    · 条件格式范围末行自动跟随（grow_cf_last_row）
    · 幂等：A 列已存在该 SKU → 跳过
    返回 [(行号, sku), ...]
    """
    last = _last_data_row(ws)
    before = last
    bcol = _brand_col(ws)
    have = sku_set(ws)
    tpl = pick_template_row(ws, last)
    out = []
    for sku in new_skus:
        if sku in have:
            continue
        src = tpl
        last += 1
        clone_row(ws, src, last, clear_range=clear_range)
        ws.cell(last, 1).value = sku
        nm = (name_map or {}).get(sku)
        ws.cell(last, 2).value = nm if nm else None
        if state_map and sku in state_map:
            # D=状态。clone_row 已把末行状态抄过来，这里按 listing 实况覆盖
            ws.cell(last, 4).value = state_map[sku]
        if bcol:
            b = brand_of(nm)
            if b:
                ws.cell(last, bcol).value = b
        have.add(sku)
        out.append((last, sku))
    if out:
        grow_cf_last_row(ws, before, last)
    return out


def repair_row_width(ws, clear_range=None):
    """把「列残缺」的已有数据行补齐。

    判定：某数据行有 SKU，但列数 < 完整行（模板行）列数的 75% → 视为残缺。
    处理：用 clone_row 克隆模板行到该行（公式平移 / 文本属性照抄 / 数值留空 /
          货件列清空），再把 A=原SKU、B=原品名、C/D=原值 写回 —— **不删行、不改行序**，
          所以外部引用（库存记录、数据透视表、断货分析）不受影响。
    用途：修历史遗留（2026-10-23 前补行照抄末行留下的 G598/K001/G600/G601 等）。
    返回 [(行号, sku, 原列数, 新列数), ...]
    """
    maxc = min(ws.max_column, 300)
    last = _last_data_row(ws)
    tpl = pick_template_row(ws, last, maxc)
    full = _row_width(ws, tpl, maxc)
    if not full:
        return []
    out = []
    for r in range(DATA_FIRST_ROW, last + 1):
        sku = ws.cell(r, 1).value
        if sku is None:
            continue
        n = _row_width(ws, r, maxc)
        if n >= full * 0.75:
            continue
        old = (str(sku), ws.cell(r, 2).value, ws.cell(r, 3).value, ws.cell(r, 4).value)
        clone_row(ws, tpl, r, clear_range=clear_range, max_col=maxc)
        ws.cell(r, 1).value = old[0]
        ws.cell(r, 2).value = old[1]
        ws.cell(r, 3).value = old[2]
        ws.cell(r, 4).value = old[3]
        out.append((r, old[0], n, _row_width(ws, r, maxc)))
    return out


def source_sheets_of(ws, wb):
    """分析表对应的库存源 sheet 列表（可能多个）。

    按命名规律找，不用解析公式 —— 公式有 `'xx'!` 带引号和不带引号两种写法，
    且 data_only 加载时读不到公式，不可靠。规律：
        <站点>-每日库存分析表  →  <站点>可售+在途库存 / <站点>可售库存 / <站点>在途库存
                                 → 可售+在途库存 / 可售库存 / 在途库存 / 库存明细
    多源是常态（如 Interflowd澳 = 可售库存 + 在途库存，要合并）。
    """
    name = str(ws.title or "").strip()
    site = name.split("-")[0].strip() if "-" in name else ""
    cands = []
    if site:
        cands += [f"{site}可售+在途库存", f"{site}可售库存", f"{site}在途库存"]
    cands += ["可售+在途库存", "可售库存", "在途库存", "库存明细"]
    out = []
    for c in cands:
        if c in wb.sheetnames and c not in out:
            out.append(c)
    return out


def _find_sku_col(hdr):
    """表头里找卖家 SKU 列。避开 FNSKU。"""
    # 精确优先
    for k in ("sku", "msku", "seller-sku", "vendor_ext_id"):
        if k in hdr:
            return hdr[k]
    for k, c in hdr.items():
        if "sku" in k and "fnsku" not in k and "fn sku" not in k:
            return c
    return None


def listing_skus(ws, start_row=2):
    """从库存源 sheet 取 [(sku, 品名, listing是否活着)]。

    两种源结构都要认：
    · 亚马逊标准报告（行2 表头，有 afn-listing-exists）→ 按该列判在售
    · 用户自建「可售库存 / 在途库存」（行2 表头，无 listing 列）→ 出现即视为在售
    返回的 SKU 会过格式校验，跳过 'Info' / 'Mandatory' 这类说明行。
    """
    hdr = {}
    for c in range(1, min(ws.max_column, 60) + 1):
        v = ws.cell(2, c).value
        if v:
            hdr[str(v).strip().lower()] = c
    sku_c = _find_sku_col(hdr)
    if not sku_c:
        return []
    name_c = hdr.get("product-name") or hdr.get("product_name") or hdr.get("producto") or 4
    lis_c = next((hdr[k] for k in _LISTING_COLS if k in hdr), None)

    out = []
    for r in range(start_row, min(ws.max_row, 5000) + 1):
        sku = ws.cell(r, sku_c).value
        if sku is None:
            continue
        sku = str(sku).strip()
        if not _SKU_RE.match(sku):
            continue
        if lis_c is not None:
            # 只有明确的「是/Yes/True/1」才算在售；No / 空 / Unknown 一律不算。
            # 实测 Amazon 报告只有 Yes/No 两种值，No 表示该账号在该站点没上架这个 SKU。
            lis = str(ws.cell(r, lis_c).value or "").strip().lower()
            # 「在售」是本项目 D 列自己的写法（与「未上架」成对），也要认
            alive = lis in ("yes", "是", "true", "1", "在售")
        else:
            # 源里没有 listing 列 → 出现即视为在售。
            # 适用：用户自建的「可售库存/在途库存/多国库存」，以及 afn-onhand-buyable 式的
            # 管理库存报告（日本/澳大利亚/interflow澳大利亚）。这类报告只列「已入仓」的 SKU，
            # 能出现在里面本身就说明 listing 是通的。
            alive = True
        nm = ws.cell(r, name_c).value
        out.append((sku, str(nm).strip() if nm else "", alive))
    return out


def listing_skus_merged(wb, sheet_names):
    """合并多个源 sheet 的 listing_skus，同一 SKU 取第一次出现（品名/在售状态取 or）。"""
    seen = {}
    for sn in sheet_names:
        if sn not in wb.sheetnames:
            continue
        for sku, nm, alive in listing_skus(wb[sn]):
            if sku not in seen:
                seen[sku] = [nm, alive]
            else:
                if not seen[sku][0] and nm:
                    seen[sku][0] = nm
                seen[sku][1] = seen[sku][1] or alive
    return [(k, v[0], v[1]) for k, v in seen.items()]


def append_sku_rows(ws, new_skus, tail_range, extra_single=(), name_map=None):
    """【已弃用，改用 append_sku_rows_auto】在数据区末尾追加新 SKU 行：
    复制末行公式（**不平移行号**，仅欧盟簿历史兼容），A=SKU，B=品名，其余静态值留空。
    返回 [新增行号...]。
    """
    last = _last_data_row(ws)
    src = last
    rows = []
    for sku in new_skus:
        last += 1
        # A 列 SKU
        ws.cell(last, 1).value = sku
        # B 列品名（可选）
        nm = (name_map or {}).get(sku)
        ws.cell(last, 2).value = nm if nm else None
        # C/D 列（品类/状态）沿用末行，便于筛选保持一致
        for c in (3, 4):
            ws.cell(last, c).value = ws.cell(src, c).value
        # 公式列（E/F/G/H/AC/AD/AF/AG/AH）
        for col in FORMULA_COLS:
            cn = _colnum(col)
            v = ws.cell(src, cn).value
            ws.cell(last, cn).value = _plain_formula(v)
        # 日销量列 AI..BU / AS..CE（公式）
        for cn in range(tail_range[0], tail_range[1] + 1):
            ws.cell(last, cn).value = _plain_formula(ws.cell(src, cn).value)
        # 种类列
        for col in extra_single:
            cn = _colnum(col)
            ws.cell(last, cn).value = _plain_formula(ws.cell(src, cn).value)
        rows.append(last)
    return rows


def _colnum(s):
    n = 0
    for ch in s:
        n = n * 26 + (ord(ch) - 64)
    return n


def month_sales(daily_ws, ym, region=None, channel=None, max_row=None):
    """从「每日销售数据」复现日销量口径，汇总指定月份（yyy-mm）各 SKU 销量。
    条件：有效性(状态<>Cancelled 且 MID(sales-channel,1,7)=='Amazon.') 且
          英国用 channel='Amazon.co.uk' / 欧洲用 region='欧洲'。
    列：C=purchase-date(3) E=order-status(5) G=sales-channel(7) L=sku(12) O=quantity(15)
    兼容普通与 read_only 工作表（iter_rows 顺序读；read_only 版供归档簿封存用）。
    """
    acc = {}
    for row in daily_ws.iter_rows(min_row=3, max_row=max_row, min_col=1, max_col=15, values_only=True):
        sku = row[11]
        if not sku:
            continue
        purch = row[2]
        if not purch or str(purch)[:7] != ym:
            continue
        status = row[4]
        chan = row[6] or ""
        if str(status).strip() == "Cancelled":
            continue
        if str(chan)[:7] != "Amazon.":
            continue
        if channel is not None and str(chan) != channel:
            continue
        if region is not None:
            reg = "英国" if str(chan) == "Amazon.co.uk" else "欧洲"
            if reg != region:
                continue
        q = row[14] or 0
        try:
            q = float(q)
        except (TypeError, ValueError):
            q = 0
        k = normalize_sku(sku)
        acc[k] = acc.get(k, 0) + q
    return acc


def freeze_month(ws, col, values):
    """把 {sku: qty} 封存到月度区的指定列（只写有数据的 SKU 行）。返回写入格数。
    先清该列数据区残留再写：当月列每日滚动时，"上月有余值本月归零"的 SKU 不能留旧值。"""
    n = 0
    for r in range(DATA_FIRST_ROW, DATA_FIRST_ROW + 3000):
        sku = ws.cell(r, 1).value
        if sku is None:
            break
        k = str(sku).strip()
        v = values.get(k)
        cell = ws.cell(r, col)
        if v:
            cell.value = v
            n += 1
        elif cell.value is not None:
            cell.value = None
    return n


def rename_current_header(ws, col):
    """「当月总销量」表头改为动态：=MONTH(数据透视表!B11-1)&"月总销量"（欧盟簿专用）。"""
    ws.cell(5, col).value = '=MONTH(数据透视表!B11-1)&"月总销量"'


_CUR_SUM_RE = _re2.compile(r'^=SUMPRODUCT\(')
# 锚点 = r7 里带 -1 的那个 MONTH(...) 实参。用 [^()] 防止跨括号误吞
# （`MONTH($Y$5:$BK$5)=MONTH(库存明细!$A$1-1)` 若用 .+? 会吞成 "$Y$5:$BK$5)=MONTH(库存明细!$A$1"）
_ANCHOR_IN_SUM = _re2.compile(r'MONTH\(\s*([^()]+?)\s*-\s*1\s*\)')


def dynamic_current_headers(ws):
    """把本表所有「当月总销量」列的表头改成**动态跟随锚点**（全 12 簿通用）。

    识别：数据行 r7 是 `=SUMPRODUCT((MONTH($Y$5:$BK$5)=MONTH(<锚点>-1))*Y7:BK7)`
    形态的列即为「当月总销量」列。锚点直接从该公式里抽出（比硬编码每簿锚点更自洽：
    锚点一旦调整，表头自动跟随）。表头写 `=MONTH(<锚点>-1)&"月总销量"`。
    返回 [(列号, 旧表头, 新表头)]（仅返回真正发生变化的列）。
    """
    changed = []
    for c in range(1, ws.max_column + 1):
        f7 = ws.cell(7, c).value
        if not (isinstance(f7, str) and _CUR_SUM_RE.match(f7.strip())):
            continue
        m = _ANCHOR_IN_SUM.search(f7)
        if not m:
            continue
        anchor = m.group(1)
        new = f'=MONTH({anchor}-1)&"月总销量"'
        cur = ws.cell(5, c).value
        if cur != new:
            ws.cell(5, c).value = new
            changed.append((c, cur, new))
    return changed


# ======================================================================
# 日列「组合」（列 outline 折叠）按快照日同步 —— 2026-10-14 用户口述
#   「1014 就要只用从 14 号往后到月底组合了，1013 就取消」
#   规则：快照日 .. 当月月末 的日列整段折叠（outlineLevel=1, hidden=1），
#   折叠按钮落在该段后一列（collapsed=1）；其余日列一律展开。
#   只动日列（行5 为日期），左侧固定列上的历史分组（E/H/I/L 等）不碰。
# ======================================================================
PREV_AVG_LABEL = "上月日均"


def day_columns(ws, row=5, max_scan=200):
    """行5 为日期（datetime）的列 → [(列号, date)]，按列号升序。"""
    import datetime as _dt
    out = []
    for c in range(1, min(ws.max_column, max_scan) + 1):
        v = ws.cell(row, c).value
        if isinstance(v, _dt.datetime):
            out.append((c, v.date()))
    return out


def _set_col(ws, c, outline=None, hidden=None, collapsed=None):
    """按需设置某列的 outline 属性（None=不动）。

    ⚠️ 调用方须先经 `_materialize_col_ranges` 把该列所在的**区间**列定义拆成单列，
    否则这里 `ws.column_dimensions[letter]` 会新建一个默认 width=13 的 ColumnDimension，
    与残留的区间定义重叠 → Excel 取单列定义 → 列宽从原值突变成 13（表格宽度不一致）。
    """
    from openpyxl.utils import get_column_letter
    cd = ws.column_dimensions[get_column_letter(c)]
    if outline is not None:
        cd.outlineLevel = outline
    if hidden is not None:
        cd.hidden = hidden
    if collapsed is not None:
        cd.collapsed = collapsed


def _materialize_col_ranges(ws, cols):
    """把覆盖 cols 中任一列的**多列区间**列定义拆成单列定义（逐列继承原宽度/样式/分组）。

    为什么需要（2026-10-14 修「表格宽度不一样」）：原表把日列写成区间
    （如 `<col min=35 max=42 width=4.75 style=6/>`）。按列设 outline 时
    `ws.column_dimensions[letter]` 会为单列新建 ColumnDimension（openpyxl 默认 width=13），
    与区间定义重叠；Excel 对重叠取更具体的单列 → 这些列宽由 4.75 变 13。
    先把区间拆成单列（宽度/样式照抄），再设 outline，即不再产生重叠。
    返回拆分的区间个数。
    """
    from openpyxl.utils import get_column_letter
    from copy import copy as _copy
    holder = ws.column_dimensions
    want = set(cols)
    splits = []
    for key, cd in list(holder.items()):
        mn, mx = cd.min or 0, cd.max or 0
        if mx > mn and any(mn <= c <= mx for c in want):
            splits.append((key, mn, mx, cd))
    for key, mn, mx, cd in splits:
        try:
            del holder[key]
        except KeyError:
            pass
        for c in range(mn, mx + 1):
            nd = holder[get_column_letter(c)]
            if nd.min is None:
                nd.min = c
            if nd.max is None:
                nd.max = c
            for k in ("width", "bestFit", "hidden", "outlineLevel", "collapsed"):
                v = getattr(cd, k, None)
                if v is not None:
                    setattr(nd, k, v)
            # 列默认样式（<col style="n"/>）经 _style 复制：
            # ColumnDimension.style 是只读 Alias('style_id')，不能直接 setattr。
            try:
                if getattr(cd, "_style", None) is not None:
                    nd._style = _copy(cd._style)
            except Exception:
                pass
    return len(splits)


def sync_day_outline(ws, snap_date):
    """按快照日同步日列「组合」。幂等。返回 (hidden_cols[list], button_col|None)。"""
    import datetime as _dt, calendar
    d0 = _dt.date.fromisoformat(snap_date) if isinstance(snap_date, str) else snap_date
    month_end = _dt.date(d0.year, d0.month, calendar.monthrange(d0.year, d0.month)[1])
    days = day_columns(ws)
    if not days:
        return [], None
    hidden = [c for c, d in days if d0 <= d <= month_end]
    button = max(hidden) + 1 if hidden else None
    # 先把日列（含折叠按钮列）所在的区间列定义拆成单列 —— 否则列宽会被重置（见 _materialize_col_ranges）
    _materialize_col_ranges(ws, [c for c, _ in days] + ([button] if button else []))
    for c, d in days:
        if d0 <= d <= month_end:
            _set_col(ws, c, outline=1, hidden=True, collapsed=False)
        else:
            _set_col(ws, c, outline=0, hidden=False, collapsed=False)
    if button:
        _set_col(ws, button, collapsed=True)
    return hidden, button


def fill_prev_avg(ws, values, col=None, month_col=None, days=None, decimals=2):
    """把「上月日均」列改写为**活公式**：=IF(封存列="","",ROUND(封存列/天数,decimals))。
    列缺省按行5 标签「上月日均」定位；month_col=「N月」封存列列号（由 caller 传入），
    days=上月日历天数。month_col/days 缺失时退回写静态值（兼容旧调用）。
    封存列为空的行 → 清空（与日销口径同步，避免上月有余值本月归零的 SKU 留旧值）。
    2026-10-15 起由静态值改为公式：换月后程序只刷封存列，日均自动联动；
    round 2 位保住小销量 SKU（round 1 会把月销 1 件的 0.03 抹成 0.0，
    蓝色规则 AI<AE/2.5 且 AI>0 对 AE=0 永假）。"""
    if col is None:
        col = next((c for c in range(1, 80)
                    if str(ws.cell(5, c).value).strip() == PREV_AVG_LABEL), None)
    if col is None:
        return None, 0
    n = 0
    for r in range(DATA_FIRST_ROW, DATA_FIRST_ROW + 3000):
        sku = ws.cell(r, 1).value
        if sku is None:
            break
        k = str(sku).strip()
        v = values.get(k)
        cell = ws.cell(r, col)
        if v is None:
            if cell.value is not None:
                cell.value = None
            continue
        if month_col and days:
            m = ws.cell(r, month_col)
            if m.value in (None, "", 0):
                # 封存列为空：不应有日均（防御 freeze_month 与此处口径漂移）
                if cell.value is not None:
                    cell.value = None
                continue
            ref = m.coordinate
            cell.value = f'=IF({ref}="","",ROUND({ref}/{int(days)},{decimals}))'
        else:
            cell.value = round(float(v), decimals)
        n += 1
    return col, n


# ======================================================================
# 月汇总封存（全簿通用，2026-10-14）：口径自动提取自各分析表日销列 SUMIFS
# 公式，与表内口径永远一致。封存一律用源表「原始列」重算（C/B 日期、E/M 状态、
# G 渠道），不读公式辅助列——openpyxl 保存会清公式缓存，辅助列缓存不可靠。
# ======================================================================
_RANGE_RE2 = _re2.compile(r"^(?:'([^']+)'|([^'!]+))!\$([A-Z]+):\$([A-Z]+)$")
_ES_MONTHS = {'enero': '01', 'febrero': '02', 'marzo': '03', 'abril': '04',
              'mayo': '05', 'junio': '06', 'julio': '07', 'agosto': '08',
              'septiembre': '09', 'octubre': '10', 'noviembre': '11',
              'diciembre': '12'}


def parse_daily_sumifs(formula):
    """解析日销 SUMIFS 公式 → 封存口径 spec。
    形态：=SUMIFS(源!$O:$O, 源!$L:$L,$A7, 源!$AP:$AP,BB$5, 源!$G:$G,$D$5, 源!$AQ:$AQ,$D$4)
    返回 {'src','qty','sku','date','valid','extra'(可选)}（列字母），解析失败返回 None。"""
    body = str(formula).strip()
    if not body.upper().startswith('=SUMIFS('):
        return None
    parts = body[len('=SUMIFS('):].rstrip(')').split(',')
    if len(parts) % 2 != 1:
        return None

    def rng(p):
        m = _RANGE_RE2.match(p.strip())
        return (m.group(1) or m.group(2), m.group(3)) if m else None

    q = rng(parts[0])
    if not q:
        return None
    spec = {'src': q[0], 'qty': q[1]}
    for i in range(1, len(parts), 2):
        kr = rng(parts[i])
        v = parts[i + 1].strip()
        if not kr or kr[0] != spec['src']:
            return None
        k = kr[1]
        if v == '$A7':
            spec['sku'] = k
        elif _re2.match(r'^[A-Z]{1,3}\$5$', v):
            spec['date'] = k
        elif v == '$D$4':
            spec['valid'] = k
        elif v == '$D$5':
            spec['extra'] = k
    if 'sku' not in spec or 'date' not in spec or 'valid' not in spec:
        return None
    return spec


def _col_from_ref(ref):
    m = _re2.match(r'\$?([A-Z]+)\$?\d+', str(ref or ''))
    return m.group(1) if m else None


def _last_sumifs_col(ws, max_scan=140):
    """行7 里最右侧的 `=SUMIFS(` 单元格列号；没有返回 None。

    站点表（意大利/法国/…）行5 的日列头是**公式链**（`=AB5+1`），openpyxl 非
    data_only 模式读到的是公式字符串、不是 datetime，`scan_analysis` 的 datetime
    扫描会全落空。这些表的口径与主分析表完全一致（同一个 SUMIFS 模板），
    所以退回「行7 最右的 SUMIFS」即可拿到同样的 spec。
    """
    top = min(ws.max_column or 0, max_scan)
    for ci in range(top, 0, -1):
        f = ws.cell(7, ci).value
        if isinstance(f, str) and f.strip().upper().startswith('=SUMIFS('):
            return ci
    return None


def scan_analysis(ws):
    """扫描分析表（read_only 可用）：最新日销列 SUMIFS 口径 + D4/D5 条件值 + 月标签列。
    月列定位：行5 标签 == "N月" 的最小列号（第一块=封存块，第二块列号更大）。
    返回 dict(spec, d4, d5, month_cols={m: col})，无法解析返回 None。"""
    import datetime as _dt
    dt_cols, lab = set(), {}
    for row in ws.iter_rows(min_row=5, max_row=5, max_col=ws.max_column):
        for c in row:
            v = c.value
            if isinstance(v, _dt.datetime):
                dt_cols.add(c.column)
            elif isinstance(v, str):
                s = v.strip()
                if s.endswith('月') and s[:-1].isdigit():
                    lab.setdefault(int(s[:-1]), []).append(c.column)
    dcol = None
    for ci in sorted(dt_cols, reverse=True):
        f7 = ws.cell(7, ci).value
        if isinstance(f7, str) and f7.strip().upper().startswith('=SUMIFS('):
            dcol = ci
            break
    if dcol is None:
        # 站点表兜底（2026-10-27）：行5 日列是公式 → 没有 datetime
        dcol = _last_sumifs_col(ws)
    if dcol is None:
        return None
    spec = parse_daily_sumifs(ws.cell(7, dcol).value)
    if not spec:
        return None
    d4, d5 = ws['D4'].value, ws['D5'].value
    if d5 is not None and isinstance(d5, str) and d5.startswith('='):
        return None
    month_cols = {m: min(cols) for m, cols in lab.items() if cols}
    return {'spec': spec, 'd4': d4, 'd5': d5, 'month_cols': month_cols}


def _es_ym(s):
    """美客多西语日期 '31 de agosto de 2026 23:18 hs.' → '2026-08'。"""
    m = _re2.search(r'(\d{1,2}) de ([A-Za-záéíóúÁÉÍÓÚ]+) de (\d{4})', str(s or ''))
    if not m:
        return None
    mm = _ES_MONTHS.get(m.group(2).lower())
    return f'{m.group(3)}-{mm}' if mm else None


def _first_formula(wsrc, col, drow, n=5):
    """数据区前 n 行里该列第一个公式文本（str 或 ArrayFormula.text）。"""
    from openpyxl.utils import column_index_from_string
    ci = column_index_from_string(col)
    for row in wsrc.iter_rows(min_row=drow, max_row=drow + n, min_col=ci, max_col=ci, values_only=True):
        v = row[0]
        if hasattr(v, 'text'):
            return v.text
        if isinstance(v, str) and v.startswith('='):
            return v
    return None


def _region_pairs(eform):
    """地区派生列公式 =IF(G3="Amazon.ca","加拿大",IF(...,"墨西哥","美国"))
    → (原始列, [(渠道, 地区)...], 兜底地区)。非派生列返回 (col, None, None)。"""
    col = _col_from_ref(_re2.search(r'(?:IF|LEFT|MID)\(\$?([A-Z]+)\$?\d', eform or ''))
    if not eform or 'IF(' not in eform:
        return None, None, None
    g = _re2.match(r'=IF\(\$?([A-Z]+)\$?\d', eform)
    gc = g.group(1) if g else 'G'
    quotes = _re2.findall(r'"([^"]*)"', eform)
    if len(quotes) >= 2 and len(quotes) % 2 == 1:
        pairs = list(zip(quotes[0:-1:2], quotes[1:-1:2]))
        return gc, pairs, quotes[-1]
    return gc, None, None


def resolve_month_sales(wsrc, ym, spec, d4, d5=None):
    """从源表（read_only，公式句柄）按「日销 SUMIFS 等价」口径重算 ym 各 SKU 销量。
    自动识别三类簿：
      amazon  有效性=E<>Cancelled 且 MID(G,1,7)="Amazon."（北美系额外放行 B 列含 TEMU）
      emag    有效性=M 列不含 Canceled/Storno 子串（罗马尼亚）
      ml      有效性=透视图关键词映射 SWITCH/XLOOKUP/LOOKUP（美客多，从透视图 K/L 静态值复刻）
    日期：LEFT(C,10)/MID(B,..) → 原始列前 7 位；美客多数组 → 西语文本解析。
    extra：原始列直接比 d5；派生地区列按公式映射后比 d5。"""
    from openpyxl.utils import column_index_from_string

    # 表头行：行1 有值=1（罗马尼亚/美客多），全空=2（亚马逊系）
    r1 = next(wsrc.iter_rows(min_row=1, max_row=1, max_col=5, values_only=True))
    hrow = 1 if any(v is not None for v in r1) else 2
    drow = hrow + 1

    # ---- 日期 ----
    dform = _first_formula(wsrc, spec['date'], drow) or ''
    m = _re2.search(r'(?:LEFT|MID)\(\$?([A-Z]+)\$?\d', dform)
    if m:
        dcol_raw, es_mode = m.group(1), False
    elif 'TEXTBEFORE' in dform:          # 美客多：NUMBERVALUE(年&"/"&月&"/"&TEXTBEFORE(B," de"))
        dcol_raw, es_mode = _col_from_ref(_re2.search(r'\(([A-Z]+)\d', dform)) or 'B', True
    else:
        dcol_raw, es_mode = spec['date'], False

    # ---- 有效性 ----
    vform = _first_formula(wsrc, spec['valid'], drow) or ''
    if 'XLOOKUP' in vform or ('LOOKUP(' in vform and '透视图' in vform):
        mode = 'ml'
    elif 'Canceled' in vform:
        mode = 'emag'
    else:
        mode = 'amazon'
    temu = 'TEMU' in vform
    stat_col = _col_from_ref(_re2.search(r'(?:E|F)\d+<>"Cancelled"', vform)) or 'E'
    if mode == 'emag':
        stat_col = _col_from_ref(_re2.search(r'FIND\("Canceled",\$?([A-Z]+)\$?\d', vform)) or 'M'
    ml, ml_cols = [], ('D', 'S', 'C')
    if mode == 'ml':
        try:
            desc_col = _col_from_ref(_re2.search(r'XLOOKUP\(\$?([A-Z]+)\$?\d', vform)) or 'D'
            s_col = _col_from_ref(_re2.search(r'\$?([A-Z]+)\$?\d+=" "', vform)) or 'S'
            pv = _re2.search(r"([A-Za-z0-9一-龥]+透视图)", vform)
            psheet = pv.group(1) if pv else None
            # 透视图列：K:K/L:L/$K$2:$K$38 混合形态 → 展平去重保序，前两个=关键词列/值列
            seq = []
            for m2 in _re2.finditer(r'透视图!\$?([A-Z]+)\$?\d*:\$?([A-Z]+)\$?\d*', vform):
                for g in m2.groups():
                    if g not in seq:
                        seq.append(g)
            kc = seq[0] if seq else 'K'
            vc = seq[1] if len(seq) > 1 else 'L'
            if psheet and wsrc.parent and psheet in wsrc.parent.sheetnames:
                pws = wsrc.parent[psheet]
                kci, vci = column_index_from_string(kc), column_index_from_string(vc)
                for row in pws.iter_rows(min_row=1, max_row=60, min_col=min(kci, vci),
                                         max_col=max(kci, vci), values_only=True):
                    kv = row[0] if kc <= vc else row[1]
                    vv = row[1] if kc <= vc else row[0]
                    if kv is not None and str(kv).strip() != '' and str(kv).strip().lower() != '关键词':
                        ml.append((str(kv), vv))
            ml_cols = (desc_col, s_col, 'C')
        except Exception:
            ml = []

    # ---- extra（站点/地区）----
    ecol = eorig = spec.get('extra')
    epairs = efb = None
    if ecol:
        eform = _first_formula(wsrc, ecol, drow) or ''
        if 'IF(' in eform:
            eorig, epairs, efb = _region_pairs(eform)
            if not eorig:
                eorig = ecol

    # ---- 汇总 ----
    need = {spec['qty'], spec['sku'], dcol_raw, stat_col}
    if mode == 'amazon':
        need.add('G')
    if temu:
        need.add('B')
    if mode == 'ml':
        need.update(ml_cols)
    if eorig:
        need.add(eorig)
    idx = {L: column_index_from_string(L) for L in need}
    maxc = max(idx.values())
    acc = {}

    def gv(row, L):
        i = idx[L]
        return row[i - 1] if i <= len(row) else None

    for row in wsrc.iter_rows(min_row=drow, max_col=maxc, values_only=True):
        # 月份
        if es_mode:
            if _es_ym(gv(row, dcol_raw)) != ym:
                continue
        else:
            if str(gv(row, dcol_raw) or '')[:7] != ym:
                continue
        # 有效性 == d4
        if mode == 'amazon':
            st = str(gv(row, stat_col) or '').strip()
            ch = str(gv(row, 'G') or '')
            ok = st != 'Cancelled' and ch[:7] == 'Amazon.'
            if temu and not ok:
                ok = st != 'Cancelled' and 'TEMU' in str(gv(row, 'B') or '')
            val = 1 if ok else 0
        elif mode == 'emag':
            st = str(gv(row, stat_col) or '')
            val = 0 if ('Canceled' in st or 'Storno' in st) else 1
        else:  # ml
            D, S, C = (gv(row, ml_cols[0]), gv(row, ml_cols[1]), gv(row, ml_cols[2]))
            if str(S or '') == ' ':
                val = 0
            else:
                val = None
                Ds, Cs = str(D or ''), str(C or '')
                for k, v in ml:                      # XLOOKUP 精确
                    if Ds.strip() == str(k).strip():
                        val = v
                        break
                if val is None:
                    for k, v in reversed(ml):        # LOOKUP(FIND) 最后一个匹配
                        if str(k) in Ds:
                            val = v
                            break
                if val is None:
                    for k, v in reversed(ml):
                        if str(k) in Cs:
                            val = v
                            break
        if val != d4:
            continue
        # extra（站点/地区）
        if eorig:
            ev = gv(row, eorig)
            if epairs is not None:
                ev_s = str(ev or '')
                got = efb
                for ch2, rg in epairs:
                    if ev_s == ch2:
                        got = rg
                        break
                ev = got
            if d5 is not None and str(ev or '') != str(d5):
                continue
        # 数量累加
        sku = gv(row, spec['sku'])
        if not sku:
            continue
        q = gv(row, spec['qty']) or 0
        try:
            q = float(q)
        except (TypeError, ValueError):
            q = 0
        k = normalize_sku(sku)
        acc[k] = acc.get(k, 0) + q
    return acc


# ============================================================================
# 上架三态（2026-10-23 用户拍板）
# ----------------------------------------------------------------------------
# 用户原话：
#   「listing和库存的判定可以用来额外辅助，一旦有库存那天就可以当作是这个SKU的起售日期」
# 判定顺序（库存优先，listing 只做「未上架」的辅助判据）：
#   可售 > 0                        → 在售
#   可售 = 0 且 在途 > 0             → 在途
#   两者都 0 且 listing 活着         → 在售（listing 在但库存为 0，仍算在售）
#   两者都 0 且 listing 不通/未知    → 未上架
# 「未知」指源里压根没有 listing 列（自建库存 sheet / 部分管理库存报告），
# 此时不能因为拿不到信息就判未上架 —— 交给库存判定，库存也 0 才落到未上架。
# ============================================================================

# 可售列（亚马逊标准报告 / interflow 管理库存报告 / 用户自建 sheet 的读法）
# EMAG 库存明细（罗马尼亚簿）：current_stock=当前库存，stock=Stock
# 美客多「库存明细」是**三行合并表头**：行1 分组名（可售库存/在途库存/预留库存），
# 行2 西语字段名，行3 子字段名 → 这里写西语/分组名，_pick_header_row 会沿列向下找子字段。
_AVAIL_COLS = (
    "afn-fulfillable-quantity", "afn_fulfillable_quantity",
    "afn-onhand-buyable-quantity", "afn_onhand_buyable_quantity",
    "可售数量", "可售库存", "可售", "proposed_quantity",
    "在库商品可售数量", "期末仓库余额",
    "current_stock", "current stock", "当前库存", "stock", "库存",
    "unidades en full", "aptas para vender",
)
# SKU 列候选（EMAG 用 part_number=Product code；亚马逊用 seller-sku/msku）
# ⚠️ vendor_ext_id 在 EMAG 里是 **Offer ID**（不是 SKU），必须排在 part_number 之后。
_SKU_COLS = ("sku", "msku", "seller-sku", "product code", "part_number",
             "商品编号", "vendor_ext_id")
# 在途列（亚马逊三段在途 / 自建 sheet 的「在途库存」）
_INB_COLS = (
    "afn-inbound-working-quantity", "afn-inbound-shipped-quantity",
    "afn-inbound-receiving-quantity",
    "afn_inbound_working_quantity", "afn_inbound_shipped_quantity",
    "afn_inbound_receiving_quantity",
    "在途库存", "在途数量", "在途",
    "unidades en camino a full", "pendientes de ingreso",
)
_NAME_COLS = ("product-name", "product_name", "商品名称", "产品", "producto")

STATE_ON_SALE = "在售"
STATE_IN_TRANSIT = "在途"
STATE_NOT_LISTED = "未上架"


def _num0(v):
    """宽松取数：'1,234' / '12.0' / None / '' → float。取不到返回 0.0。"""
    if v is None:
        return 0.0
    s = str(v).strip().replace(",", "").replace("，", "")
    if not s:
        return 0.0
    try:
        return float(s)
    except ValueError:
        return 0.0


def _pick_header_row(ws, probe_rows=(2, 1, 3, 4, 5)):
    """找表头 → (hdr_dict, header_row, sku_col, sub_rows)。

    · 标准亚马逊报告 / 自建 sheet / EMAG「库存明细」：表头在第 2 行
    · 美客多「库存明细」：**三行合并表头** —— 行1 分组（可售库存/在途库存）、
      行2 西语字段、行3 子字段。行2 有 SKU 列但没有任何数量列，
      数量在行3 的子字段里（Aptas para vender / Pendientes de ingreso…）。
      做法：选中表头行后，**沿每列往下再找 3 行**，把子字段名也并进该列的别名表。

    hdr 键 → 列号；同名列出现两次只保留第一次。
    sub_rows 给调用方判断「数据从第几行开始」（合并表头要跳过子字段行）。
    """
    for hr in probe_rows:
        if hr > (ws.max_row or 1):
            continue
        rows = {}
        for rr in range(hr, min(hr + 3, ws.max_row or hr) + 1):
            d = {}
            for c in range(1, min(ws.max_column, 80) + 1):
                v = ws.cell(rr, c).value
                if v is None:
                    continue
                k = str(v).strip().lower()
                if not k:
                    continue
                d.setdefault(k, []).append(c)
            rows[rr] = d
        main = rows.get(hr) or {}
        if not main:
            continue
        # SKU 列：先按显式候选名精确查（EMAG 的 part_number 必须压过 vendor_ext_id=Offer ID）
        sku_c = None
        for k in _SKU_COLS:
            if k in main:
                sku_c = main[k][0]
                break
        if not sku_c:
            sku_c = _find_sku_col({k: v[0] for k, v in main.items()})
        if not sku_c:
            continue

        # 合并向下最多 3 行的子字段（只用于数量列识别，SKU 列仍取主表头行）
        # ⚠️ 只有当该行**在 SKU 列上是空**时才算子表头行。日本那种标准报告
        #    第 3 行就是数据（A3=C001），若照收会把 start_row 推到第 6 行，
        #    前 3 个 SKU 直接读不到（2026-10-23 踩坑：G597 因此判成「在售」）。
        hdr = {k: v[0] for k, v in main.items()}
        last_sub = hr
        for rr in sorted(rows):
            if rr <= hr:
                continue
            cell_sku = ws.cell(rr, sku_c).value
            if cell_sku is not None and str(cell_sku).strip():
                break                      # 已经是数据行了，停
            last_sub = rr
            for k, cols in rows[rr].items():
                if k not in hdr:
                    hdr[k] = cols[0]
        if not any(k in hdr for k in _AVAIL_COLS + _INB_COLS):
            continue
        return hdr, hr, sku_c, (hr, last_sub)
    return None, None, None, None


def stock_snapshot(ws, start_row=None, max_row=5000):
    """从**一个**库存源 sheet 取 {sku: {name, avail, inb, listing}}。

    · avail = 可售列（多候选取第一个命中的列）
    · inb   = 三段在途之和（不足三段就有什么算什么）
    · listing = True / False / None（None = 源里没有 listing 列，信息未知）
    自建 sheet（可售库存 / 在途库存 / 可售+在途库存）常常没有 listing 列，
    也可能只有「在途库存」这种纯在途 sheet —— avail 列找不到就只算 inb。
    表头行自动探测（2026-10-23）：EMAG「库存明细」SKU 列叫 part_number、
    数量列叫 current_stock，与亚马逊报告完全不同，不能死认第 2 行。
    """
    hdr, hr, sku_c, sub = _pick_header_row(ws)
    if hdr is None:
        return {}
    if start_row is None:
        start_row = (sub[1] if sub else hr) + 1

    def _first(cols):
        for k in cols:
            if k in hdr:
                return hdr[k]
        return None

    av_c = _first(_AVAIL_COLS)
    # 在途列：注意 "stock" 已在 _AVAIL_COLS 里，别把可售列又当在途加一遍
    inb_cs = []
    for k in _INB_COLS:
        if k in hdr and hdr[k] != av_c:
            inb_cs.append(hdr[k])
    lis_c = next((hdr[k] for k in _LISTING_COLS if k in hdr), None)
    name_c = _first(_NAME_COLS)

    out = {}
    for r in range(start_row, min(ws.max_row, max_row) + 1):
        sku = ws.cell(r, sku_c).value
        if sku is None:
            continue
        sku = str(sku).strip()
        if not _SKU_RE.match(sku):
            continue
        avail = _num0(ws.cell(r, av_c).value) if av_c else 0.0
        inb = sum(_num0(ws.cell(r, c).value) for c in inb_cs) if inb_cs else 0.0
        if lis_c is not None:
            v = str(ws.cell(r, lis_c).value or "").strip().lower()
            listing = v in ("yes", "是", "true", "1", STATE_ON_SALE)
        else:
            listing = None
        nm = ws.cell(r, name_c).value if name_c else None
        cur = out.get(sku)
        if cur is None:
            out[sku] = {"name": str(nm).strip() if nm else "",
                        "avail": avail, "inb": inb, "listing": listing}
        else:
            # 同一 SKU 在一个源 sheet 里可能多行（多 condition / 多 FC）→ 累加，listing 取 or
            cur["avail"] += avail
            cur["inb"] += inb
            if listing is not None:
                cur["listing"] = bool(cur["listing"]) or listing
            if not cur["name"] and nm:
                cur["name"] = str(nm).strip()
    return out


def stock_snapshot_merged(wb, sheet_names):
    """合并多个库存源 sheet → {sku: {name, avail, inb, listing}}。

    跨 sheet 累加 avail/inb（Interflowd澳 = 可售库存 + 在途库存 这类要合并），
    listing 三值合并：任一 True → True；全 None → None；其余 False。
    """
    out = {}
    for sn in (sheet_names or []):
        if sn not in wb.sheetnames:
            continue
        for sku, rec in stock_snapshot(wb[sn]).items():
            cur = out.get(sku)
            if cur is None:
                out[sku] = dict(rec)
                continue
            cur["avail"] += rec["avail"]
            cur["inb"] += rec["inb"]
            a, b = cur["listing"], rec["listing"]
            if a is True or b is True:
                cur["listing"] = True
            elif a is None and b is None:
                cur["listing"] = None
            else:
                cur["listing"] = False
            if not cur["name"] and rec["name"]:
                cur["name"] = rec["name"]
    return out


def tri_state(rec):
    """库存记录 → 三态之一。rec 来自 stock_snapshot(_merged)。"""
    if not rec:
        return STATE_NOT_LISTED
    if (rec.get("avail") or 0) > 0:
        return STATE_ON_SALE
    if (rec.get("inb") or 0) > 0:
        return STATE_IN_TRANSIT
    if rec.get("listing") is False:
        return STATE_NOT_LISTED
    # listing 为 True（有 listing 无库存）或 None（未知）→ 按在售处理，
    # 与旧逻辑「无 listing 列 = 出现即在售」保持一致，不制造新的未上架。
    return STATE_ON_SALE


def tri_state_of_sku(stock, sku):
    """便捷：直接给 SKU 要状态。"""
    return tri_state((stock or {}).get(sku))


def start_sale_date(site_rows):
    """从「型号首库存登记.xlsx → 站点明细」的某 SKU 行集里算**起售日期**。

    site_rows = [(站点, 首可售日, 首在途日), ...]
    口径（用户拍板）：**有库存那天 = 起售日期**。这里的「有库存」取最早的可售日；
    只有在途、从没可售 → 返回 ("", 沿用首在途日)，由调用方决定要不要退而用发货日。
    返回 (start_date, basis)，basis ∈ {"可售", "在途", ""}。
    """
    av = [d for _, d, _ in site_rows if d]
    if av:
        return min(av), "可售"
    inb = [d for _, _, d in site_rows if d]
    if inb:
        return min(inb), "在途"
    return "", ""


def sync_state_cf(ws):
    """让「断货粉/浅粉」条件格式排除 **在途**（在途 SKU 不是断货）。

    背景：2026-10-23 起 D 列是三态。原来的粉色断货规则只看 F(可售)=0，
    在途 SKU（F=0、G=350）会被标成断货 —— 但它货在路上了，不是断货。

    ⚠️ 只认**断货粉规则**：区间落在 F:H（预留/可售/在途三列）且公式里带
    `N($F{n})<=0`（可售<=0）。别的 AND 规则一律不碰（2026-10-23 踩坑：
    第一版按「AND( 开头」全匹配，把竞品 T7:T1175、日销 Y7:BK1297、
    机会产品 AV7:AV14743 等无关规则也加了 `<>"在途"`）。
    幂等：已带 `<>"在途"` 的跳过。返回改动规则数。
    """
    touched = 0
    for rng in ws.conditional_formatting:
        sq = str(rng.sqref)
        # 只认 F:H 那三列（断货粉 / 浅粉）：区间必须完整落在 F..H 内
        ok = False
        for part in sq.split():
            m = _re2.match(r"^\$?([A-Z]{1,3})\$?(\d+)(?::\$?([A-Z]{1,3})\$?(\d+))?$", part)
            if not m:
                continue
            c1 = m.group(1)
            c2 = m.group(3) or m.group(1)
            if c1 == "F" and c2 in ("F", "G", "H"):
                ok = True
                break
        if not ok:
            continue
        for rule in rng.rules:
            f = getattr(rule, "formula", None)
            if not f:
                continue
            new = []
            changed = False
            for expr in f:
                t = str(expr or "")
                # 已经追加过在途排除（公式末尾是 ,$D7<>"在途") → 跳过
                if 'N($F' not in t or _re2.search(r'\$D\d+<>"在途"', t):
                    new.append(t)
                    continue
                # 浅粉那条（OR(N($G)>0,N($H)>0) = 有在途/有调拨）本来就是在途专属，
                # 加了 <>在途 会永不成立 → 必须跳过，否则浅粉标记全丢。
                if _re2.search(r"OR\(\s*N\(\$G", t):
                    new.append(t)
                    continue
                m = _re2.search(r"\$?F\$?(\d+)", t)
                if not m:
                    new.append(t)
                    continue
                row = m.group(1)
                inner = t[4:-1] if t.startswith("AND(") and t.endswith(")") else t
                new.append(f'AND({inner},$D{row}<>"{STATE_IN_TRANSIT}")')
                changed = True
            if changed:
                rule.formula = new
                touched += 1
    return touched
