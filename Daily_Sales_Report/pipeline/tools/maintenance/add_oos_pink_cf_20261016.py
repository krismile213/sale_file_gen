# -*- coding: utf-8 -*-
"""新增/更新「缺货·要补货」两档粉色条件格式（全量：自动发现所有 *每日库存分析表）。

判定（行级）：
  必须同时满足
    ① A 有 SKU
    ② 可售库存 F <= 0            —— 关键触发（在售为 0）
    ③ SUM(日销列) > 0            —— 窗口内卖过，排除「从没卖过的死 SKU」
                                    （需求方原话：代表要补货，**而不是没有销量**）
    ④ 连续无销量天数 >= 3
       连续无销量 = (锚点(快照日) − 1) − 该 SKU 最后一个「日销>0」的日列日期
       （**快照日当天不计入**：快照日是早上导数，当天数据未满）
  分级（在途/调拨不要求为 0）
    深粉 断货·无补给 ： 在途 G = 0 且 调拨 H = 0   —— 最紧急
    浅粉 断货·补给在途： 在途 G > 0 或 调拨 H > 0  —— 别重复下单，盯到货时效

着色范围（2026-10-16 晚 用户要求，**不再整行**）：
  · 库存三列 **F:H**（可售 / 在途 / 调拨）
  · 日列区 `<日销首列>7:<日销末列><末数据行>`，但**只标「最后销售之后 ~ 快照日前一天」**
    那段无销量的格（如 G444 最后销售 10/8 → 标 10/9..10/15）
  ⇒ 两档 × 两范围 = **4 条规则**（同日销列上原有的 红/蓝 染色带仍会盖住「当天有销量」格）。

自动发现：
  · 分析表 = sheet 名以「每日库存分析表」结尾
  · 日销列 = 行 5 里真正的日期格（isinstance datetime）
  · 锚点   = 日销首列左邻列(L1-1) 的行 5 表头 `=MONTH(<锚点>-1)&"月总销量"` 里正则取
             （与 pipeline/analysis.py 的 _ANCHOR_IN_SUM 同口径）

同时写入 `fullCalcOnLoad=1` + `calcMode=auto`：
  日销列是 SUMIFS 公式，openpyxl 存盘会丢公式缓存值 ⇒ 打开时必须全量重算，
  否则某些阅读器/关闭了自动重算的 Excel/WPS 里粉标不显示。

优先级压到最低：日销列上原有的 红/蓝 染色带仍会盖住「当天有销量」的格。
重复运行会**就地更新**同色规则的公式（按 dxf 填充色识别），不会重复叠加。

管线：strip_pivot_records → openpyxl 编辑 → restore_parts → testzip → 备份 → 落盘。
用法：python add_oos_pink_cf_20261016.py           # 干跑
     python add_oos_pink_cf_20261016.py --apply
"""
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

import os, re, sys, shutil, zipfile, traceback
from datetime import datetime, date
from openpyxl.utils import get_column_letter as gl
from openpyxl.formatting.rule import FormulaRule
from openpyxl.formatting.formatting import ConditionalFormattingList
from openpyxl.styles import PatternFill

BOOKS_DIR = os.path.join(DSR, "1.每日销售数据表")
BK = os.path.join(DSR, ".workbuddy", "backup_oos_pink_20261016")
TMP = os.path.join(DSR, ".workbuddy", "_tmp_oos_pink")
OUT = os.path.join(DSR, ".workbuddy", "add_oos_pink_out.txt")
import runner as R

APPLY = "--apply" in sys.argv
TIER1 = "FFFF8FA8"          # 深粉 断货·无补给
TIER2 = "FFFFDDE7"          # 浅粉 断货·补给在途
THRESH = 3
SHEET_SUFFIX = "每日库存分析表"
RE_ANCHOR = re.compile(r'MONTH\(\s*([^()]+?)\s*-\s*1\s*\)')


def _base(anchor, L1, L2):
    """行级基础条件（①②③④）：有 SKU、可售≤0、窗口内卖过、连续无销量≥THRESH 天。

    ⚠️ ④ 从 **锚点格 − 1 天**（= 最后一个有数据的日子，如 10/15）起算 —— 用户 2026-10-16 拍板：
    「比如 1016 导的，要 1013-1015 三天没有才行；1016 早上导数，就不算了」。
    即 `(锚点−1) − 最后销售日 >= THRESH` ⇒ 最后销售日 ≤ 10/12；**10/16 当天不计入**。
    这样「无销量天数」与日列实际染色段长度**严格相等**（10/13..10/15 = 3 天）。
    """
    return (f'AND($A7<>"",N($F7)<=0,SUM(${L1}7:${L2}7)>0,'
            f'{anchor}-1-LOOKUP(2,1/(${L1}7:${L2}7>0),${L1}$5:${L2}$5)>={THRESH}')


def _tier_tail(tier):
    if tier == 1:
        return ',N($G7)=0,N($H7)=0)'
    return ',OR(N($G7)>0,N($H7)>0))'


def rule_formula_stock(anchor, L1, L2, tier):
    """库存三列 F:H —— 只判行级条件（可售/在途/调拨）。"""
    return _base(anchor, L1, L2) + _tier_tail(tier)


def rule_formula_date(anchor, L1, L2, tier):
    """日列区 —— 只标「最后销售之后、快照日前一天(含)」那段无销量的格
    （如 G444 最后销售 10/8 → 标 10/9..10/15）。{L1}$5 为该列行5 的日期（列相对、行绝对）。

    2026-10-23 加上界保险：`<{锚点}` → `<MIN({锚点},TODAY()+1)`
      · 锚点正常（= 快照日）时 MIN 取锚点 → **与原公式完全等价**，不改任何染色结果；
      · 锚点若被误设成未来日期 → MIN 封顶到今天，**不会把还没到的日期标红**；
      · 打开历史归档表时 TODAY() 变大 → MIN 仍取锚点 → 历史染色不受影响。
    """
    last = f'LOOKUP(2,1/(${L1}7:${L2}7>0),${L1}$5:${L2}$5)'
    ceil = f'MIN({anchor},TODAY()+1)'
    return (_base(anchor, L1, L2)
            + f',{L1}$5>{last},{L1}$5<{ceil}'
            + _tier_tail(tier))


def color_of(rule):
    d = rule.dxf
    if d is None or d.fill is None:
        return None
    return (getattr(d.fill.bgColor, "rgb", None)
            or getattr(d.fill.fgColor, "rgb", None))


def txt(v):
    return str(getattr(v, "text", v) or "")


def enable_full_recalc(wb):
    c = wb.calculation
    if c is None:
        from openpyxl.workbook.properties import CalcProperties
        c = CalcProperties()
        wb.calculation = c
    c.fullCalcOnLoad = True
    c.calcMode = "auto"
    try:
        c.forceFullCalc = True
    except Exception:
        pass


os.makedirs(BK, exist_ok=True)
os.makedirs(TMP, exist_ok=True)
_lines = []


def emit(s=''):
    _lines.append(str(s))
    print(s)


books = sorted(f for f in os.listdir(BOOKS_DIR) if f.lower().endswith(".xlsx"))
emit(f"扫描 {len(books)} 本簿 …\n")
stats = {"sheets": 0, "new": 0, "updated": 0, "skip": 0, "fail": 0}

for fn in books:
    cur = os.path.join(BOOKS_DIR, fn)
    emit("=" * 78)
    emit(f"# {fn}")
    try:
        n = R.strip_pivot_records(cur, os.path.join(TMP, "stripped.xlsx"))
        src = os.path.join(TMP, "stripped.xlsx") if n else cur
        wb = R.load_workbook(src)
        targets = [s for s in wb.sheetnames if s.strip().endswith(SHEET_SUFFIX)]
        if not targets:
            emit("  （无 *每日库存分析表，跳过）")
            wb.close()
            continue

        touched = False
        for sheet in targets:
            ws = wb[sheet]
            days = [c.column for c in ws[5] if isinstance(c.value, (datetime, date))]
            if not days:
                emit(f"  [{sheet}] 行5 无日期列，跳过")
                stats["skip"] += 1
                continue
            L1, L2 = gl(days[0]), gl(days[-1])
            cont = days == list(range(days[0], days[-1] + 1))
            anchor = None
            for cand in (ws.cell(5, days[0] - 1).value, ws.cell(5, days[0]).value):
                m = RE_ANCHOR.search(txt(cand))
                if m:
                    anchor = m.group(1)
                    break
            if not anchor:
                emit(f"  [{sheet}] ⚠️ 取不到锚点（表头={txt(ws.cell(5, days[0] - 1).value)[:50]!r}），跳过")
                stats["skip"] += 1
                continue
            end = max((r for r in range(7, ws.max_row + 1)
                       if ws.cell(r, 1).value not in (None, "")), default=ws.max_row)
            stock_rng = f"F7:H{end}"          # 库存三列：可售 F / 在途 G / 调拨 H
            date_rng = f"{L1}7:{L2}{end}"     # 日列区：只标「最后销售之后 ~ 快照日前一天」
            before = sum(len(r.rules) for r in ws.conditional_formatting)
            old_pink = sum(1 for r in ws.conditional_formatting for rule in r.rules
                           if color_of(rule) in (TIER1, TIER2))

            mx = 0                            # 只统计「非粉色」规则 → 粉色始终优先级最低
            for r in ws.conditional_formatting:
                for rule in r.rules:
                    if (color_of(rule) not in (TIER1, TIER2)
                            and rule.priority and rule.priority > mx):
                        mx = rule.priority

            specs = [                         # (色, 公式, 优先级, 范围, 说明)
                (TIER1, rule_formula_stock(anchor, L1, L2, 1), mx + 1, stock_rng, "库存F:H·深粉"),
                (TIER2, rule_formula_stock(anchor, L1, L2, 2), mx + 2, stock_rng, "库存F:H·浅粉"),
                (TIER1, rule_formula_date(anchor, L1, L2, 1), mx + 3, date_rng, "日列无销量段·深粉"),
                (TIER2, rule_formula_date(anchor, L1, L2, 2), mx + 4, date_rng, "日列无销量段·浅粉"),
            ]
            # 重建 CF 容器：先丢掉旧粉色规则（含旧版「整行 A7:xx」范围），再按新范围/公式写入；
            # 其它规则原样保留。（openpyxl 对已存在规则只改 formula/priority、不动 sqref，
            # 改范围必须重建容器；补行后末行增长也靠这里自动跟上。）
            new_cfl = ConditionalFormattingList()
            for r in list(ws.conditional_formatting):
                for rule in r.rules:
                    if color_of(rule) not in (TIER1, TIER2):
                        new_cfl.add(str(r.sqref), rule)
            acts = []
            for color, fml, pr, rng, tag in specs:
                nr = FormulaRule(formula=[fml], stopIfTrue=False,
                                 fill=PatternFill(start_color=color, end_color=color,
                                                  fill_type="solid"))
                nr.priority = pr
                new_cfl.add(rng, nr)
                acts.append(f"{tag} pri={pr} @{rng}")
            ws.conditional_formatting = new_cfl
            touched = True
            stats["updated"] += 1
            emit(f"  [{sheet}] 日销列 {L1}..{L2}({len(days)}列,连续={cont}) 锚点={anchor} "
                 f"范围={stock_rng} + {date_rng} 规则 {before}→{before - old_pink + 4}")
            emit(f"       {' / '.join(acts)}")

        if not touched:
            emit("  （无改动）")
            wb.close()
            continue

        enable_full_recalc(wb)
        if not APPLY:
            emit("  [dry-run] 未落盘")
            wb.close()
            continue

        saved = os.path.join(TMP, "saved_%s" % fn)
        wb.save(saved)
        wb.close()
        R.restore_parts(cur, saved)
        with zipfile.ZipFile(saved) as z:
            bad = z.testzip()
        if bad:
            raise RuntimeError(f"zip 损坏: {bad}")
        dst = os.path.join(BK, fn)
        if not os.path.exists(dst):
            shutil.copy2(cur, dst)
        try:
            os.replace(saved, cur)
        except PermissionError:
            shutil.copyfile(saved, cur)
        stats["sheets"] += len(targets)
        emit(f"  [OK] 已落盘（含 fullCalcOnLoad=1）备份→{BK}")
    except Exception as e:
        stats["fail"] += 1
        emit(f"  [FAIL] {e!r}")
        emit(traceback.format_exc())

emit("")
emit(f"汇总：处理分析表 {stats['sheets']} 张；新增规则 {stats['new']} 条、"
     f"更新 {stats['updated']} 条、跳过 {stats['skip']} 张、失败 {stats['fail']} 本")

with open(OUT, "w", encoding="utf-8") as f:
    f.write("\n".join(_lines) + "\n")
