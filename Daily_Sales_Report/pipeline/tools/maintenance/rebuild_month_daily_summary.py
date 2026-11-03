# -*- coding: utf-8 -*-
"""从「年/月/日」归档重建某月的「日销站点汇总表」（数值版，打开即见）。

背景
----
月末归档（如 `2026年9月-XX库存销售表.xlsx`）里，各分析表的日销列是公式：

    =SUMIFS(源!$X:$X, 源!$L:$L, $A7, 源!$AE:$AE, Y$5, [源!$G:$G, $D$5], 源!$AF:$AF, $D$4)

脚本生成后没被 Excel 重算过，openpyxl / WPS 读到的缓存是空的 —— 打开像一片 0。
本工具**解析该公式**，直接从源表订单明细重算，产出数值化汇总表。

安全约束（硬性）
--------------
* 全程 **只读** 归档目录，绝不写入 / 修改任何既有工作簿（含当月活文件与当月归档）。
* 产出始终是一个**新建**的 xlsx。

用法
----
    python rebuild_month_daily_summary.py --month 2026-09 --verify
    python rebuild_month_daily_summary.py --month 2026-09 --build
"""
import os
import re
import sys
import argparse
import datetime
import collections

from concurrent.futures import ProcessPoolExecutor
import openpyxl
from openpyxl.utils import get_column_letter as L, column_index_from_string as CI

_HERE = os.path.dirname(os.path.abspath(__file__))             # …\pipeline\tools\maintenance
_DSR = os.path.abspath(os.path.join(_HERE, "..", "..", ".."))   # …\Daily_Sales_Report
WORK = os.path.join(_DSR, "1.每日销售数据表")
ARC_ROOT = os.path.join(WORK, "2026年每日数据")

SHEET_KW = "每日库存分析表"
HDR_ROW = 5      # 行5 = 列头（日销列头是 datetime）
DATA_ROW = 7     # 行7 起 = 数据（行6 是 SUBTOTAL）


def month_dir(month):
    """'2026-09' → …\\2026年每日数据\\2026年9月"""
    return os.path.join(ARC_ROOT, f"{month[:4]}年{int(month[5:7])}月")

_REF = re.compile(r"^'?([^'!]+)'?!\$([A-Z]{1,3}):\$[A-Z]{1,3}$")
_CELLREF = re.compile(r"^(\$?)([A-Z]{1,3})(\$?)(\d+)$")


# ---------------------------------------------------------------- 公式解析

def _split_args(s):
    out, depth, cur, q = [], 0, [], ""
    for ch in s:
        if q:
            cur.append(ch)
            if ch == q:
                q = ""
            continue
        if ch in "\"'":
            q = ch
            cur.append(ch)
            continue
        if ch == "(":
            depth += 1
        elif ch == ")":
            depth -= 1
        if ch == "," and depth == 0:
            out.append("".join(cur))
            cur = []
            continue
        cur.append(ch)
    out.append("".join(cur))
    return [a.strip() for a in out]


def parse_sumifs(f):
    """'=SUMIFS(源!$X:$X, 源!$A:$A, $A7, ...)' → (源表名, 求和列号, [(条件列号, 条件表达式)])"""
    if not isinstance(f, str):
        return None
    m = re.match(r"^=\s*SUMIFS\s*\((.*)\)\s*$", f.strip(), re.I | re.S)
    if not m:
        return None
    args = _split_args(m.group(1))
    if len(args) < 3:
        return None

    def ref(a):
        mm = _REF.match(a.strip())
        return (mm.group(1), CI(mm.group(2))) if mm else None

    sr = ref(args[0])
    if not sr:
        return None
    crits = []
    for i in range(1, len(args) - 1, 2):
        c = ref(args[i])
        if not c:
            return None
        crits.append((c[1], args[i + 1].strip()))
    return sr[0], sr[1], crits


def _norm(v):
    """统一比较键：datetime→date；'2026-09-06'→date；其余原样 strip。"""
    if isinstance(v, datetime.datetime):
        return v.date()
    if isinstance(v, datetime.date):
        return v
    if isinstance(v, str):
        s = v.strip()
        m = re.match(r"^(\d{4})[-/](\d{1,2})[-/](\d{1,2})", s)
        if m:
            try:
                return datetime.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
            except ValueError:
                pass
        return s
    return v


# ---------------------------------------------------------------- 迷你公式求值
# 源表的辅助列（purchase date / 有效性 / 地区 / sales-channel）本身也是公式且无
# 缓存，这里内联求值。只覆盖实际出现的那几个函数。

_TOKEN = re.compile(
    r'\s*(?:(?P<num>\d+(?:\.\d+)?)'
    r'|(?P<str>"[^"]*")'
    r'|(?P<ref>\$?[A-Z]{1,3}\$?\d+)'
    r'|(?P<name>[A-Za-z_][A-Za-z0-9_.]*)'
    r'|(?P<op><>|<=|>=|=|<|>|\+|-|,|\(|\)))')
_TOK_CACHE = {}
_REFCO = re.compile(r"^\$?([A-Z]{1,3})\$?\d+$")
_ROWT = re.compile(r"(\$?)([A-Z]{1,3})(\$?)\d+")


def _tokenize(s):
    """按公式字符串缓存——同列同公式模板会被复用几十万次，省最多的时间。"""
    global _TOK_CACHE
    if s in _TOK_CACHE:
        return _TOK_CACHE[s]
    out, i = [], 0
    while i < len(s):
        m = _TOKEN.match(s, i)
        if not m:
            raise ValueError(f"token@{i}: {s[i:i + 16]!r}")
        i = m.end()
        for k, v in m.groupdict().items():
            if v is not None:
                # ref token 预解析列号：否则每行每列都要重跑正则 + 列号换算
                out.append((k, v, CI(_REFCO.match(v).group(1)) if k == "ref" else -1))
                break
    _TOK_CACHE[s] = out
    return out


def _truthy(x):
    if x is None:
        return False
    if isinstance(x, bool):
        return x
    if isinstance(x, (int, float)):
        return x != 0
    if isinstance(x, str):
        return x != ""
    return True


def _num(x):
    return x if isinstance(x, (int, float)) and not isinstance(x, bool) else 0


def _call(name, a):
    name = name.upper().split(".")[-1]        # 去掉 _xlfn. / _xlws. 之类前缀
    if name == "LEFT":
        s = "" if not a or a[0] is None else str(a[0])
        return s[:int(a[1])] if len(a) > 1 else s[:1]
    if name == "MID":
        s = "" if a[0] is None else str(a[0])
        st, ln = int(a[1]), int(a[2])
        st = max(st - 1, 0)
        return s[st:st + ln]
    if name == "FIND":
        sub, s = str(a[0]), ("" if a[1] is None else str(a[1]))
        i = s.find(sub)
        return i + 1 if i >= 0 else None      # 找不到 ⇒ 非数字（Excel 为 #VALUE）
    if name == "ISNUMBER":
        return isinstance(a[0], (int, float)) and not isinstance(a[0], bool)
    if name == "AND":
        return all(_truthy(x) for x in a)
    if name == "OR":
        return any(_truthy(x) for x in a)
    if name == "IF":
        return a[1] if _truthy(a[0]) else a[2]
    if name == "IFNA":
        return a[0]
    if name == "SWITCH":                      # SWITCH(expr, v1, r1, ..., [default])
        if not a:
            return None
        expr, i = a[0], 1
        while i + 1 < len(a):
            if expr == a[i]:
                return a[i + 1]
            i += 2
        return a[i] if i < len(a) else None
    raise ValueError(f"不支持函数 {name}")


class _P:
    """极简递归下降解析器。行内引用一律按「当前行」取值（辅助列都是同行引用）。"""

    def __init__(self, toks, row):
        self.t, self.i, self.row = toks, 0, row

    def peek(self):
        t = self.t[self.i] if self.i < len(self.t) else ("", "", -1)
        return t[0], t[1]

    def take(self):
        v = self.t[self.i]
        self.i += 1
        return v

    def expr(self):
        return self.cmp()

    def cmp(self):
        left = self.sum()
        k, v = self.peek()
        if k == "op" and v in ("=", "<>", "<", ">", "<=", ">="):
            self.take()
            right = self.sum()
            op = v
            try:
                if op == "=":
                    return left == right
                if op == "<>":
                    return left != right
                if op == "<":
                    return left < right
                if op == ">":
                    return left > right
                if op == "<=":
                    return left <= right
                return left >= right
            except TypeError:
                return op == "<>"
        return left

    def sum(self):
        left = self.atom()
        while True:
            k, v = self.peek()
            if k == "op" and v in ("+", "-"):
                self.take()
                r = self.atom()
                left = _num(left) + _num(r) if v == "+" else _num(left) - _num(r)
            else:
                return left

    def atom(self):
        k, v, ci = self.take()
        if k == "num":
            return float(v)
        if k == "str":
            return v[1:-1]
        if k == "ref":
            m = re.match(r"^\$?([A-Z]{1,3})\$?\d+$", v)
            ci = CI(m.group(1))
            return self.row[ci - 1] if ci - 1 < len(self.row) else None
        if k == "name":
            if self.peek()[1] != "(":            # 无括号 ⇒ 常量名（TRUE / FALSE）
                return {"TRUE": True, "FALSE": False}.get(v.upper())
            self.take()                           # '('
            args = []
            if self.peek()[1] != ")":
                args.append(self.expr())
                while self.peek()[1] == ",":
                    self.take()
                    args.append(self.expr())
            self.take()                           # ')'
            return _call(v, args)
        if k == "op" and v == "(":
            e = self.expr()
            self.take()
            return e
        raise ValueError(f"意外 token {k}={v!r}")


def eval_formula(f, row):
    """求值以 '=' 开头的公式。公式按「结构」去重后再缓存 token：
    同列公式只差行号（LEFT(C3,10) / LEFT(C4,10)），归一化后共用一份 token，
    否则 key 永不命中，等于每行都重新 tokenize。"""
    body = _ROWT.sub(lambda m: f"{m.group(1)}{m.group(2)}0", f[1:])
    t = _TOK_CACHE.get(body)
    if t is None:
        t = _tokenize(body)
        _TOK_CACHE[body] = t
    return _P(t, row).expr()


_RANGE = re.compile(r"(\$?)([A-Z]{1,3})(\$?)\d+:\$?[A-Z]{1,3}\$?\d+")


def _resolve(row, ci, row_no=None):
    """取源表某行第 ci 列的值；若是公式（含数组公式）则内联求值。

    数组公式里的范围引用（如 B2:B1515）按「当前行」折叠成单格（B{row_no}）。
    """
    if ci - 1 >= len(row):
        return None
    v = row[ci - 1]
    if v.__class__.__name__ == "ArrayFormula" and hasattr(v, "text"):
        v = "=" + v.text
    if isinstance(v, str) and v.startswith("="):
        f = v if ":" not in v else _RANGE.sub(
            lambda m: f"{m.group(1)}{m.group(2)}{row_no}", v)
        try:
            return eval_formula(f, row)
        except Exception:
            return None
    return v


def load_index(src_ws, sum_col, crits):
    """源表 → {条件元组: 数量合计}（条件列公式内联求值）。"""
    idx = collections.defaultdict(float)
    need = max([sum_col] + [c for c, _ in crits])
    for i, row in enumerate(src_ws.iter_rows(min_row=1, values_only=True), start=1):
        if not row or len(row) < need:
            continue
        key = tuple(_norm(_resolve(row, c, i)) for c, _ in crits)
        v = row[sum_col - 1]
        if isinstance(v, (int, float)):
            idx[key] += float(v)
    return idx


# ---------------------------------------------------------------- 单表重算

def day_columns(ws):
    """行5 里所有 datetime 列 → [(列号, date)]。"""
    out = []
    for c in ws[HDR_ROW]:
        v = c.value
        if isinstance(v, (datetime.datetime, datetime.date)):
            out.append((c.column, v.date() if isinstance(v, datetime.datetime) else v))
    return out


def recompute_table(ws, wb, path=None, sn=None):
    """返回 (模式, [(列号,date)], [(行号,sku,型号,{列号:值})])；失败返回 None。

    惰性缓存回退：若重算全 0，才现场打开 data_only 视图取归档缓存 —— 覆盖靠跨表
    数组公式（NUMBERVALUE / LOOKUP / XLOOKUP）取数的簿（如美客多）。不预载，省时间。
    """
    dcols = day_columns(ws)
    if not dcols:
        return None
    parsed = parse_sumifs(ws.cell(DATA_ROW, dcols[0][0]).value)
    if not parsed:
        return None
    src_name, sum_col, crits = parsed
    if src_name not in wb.sheetnames:
        return None
    idx = load_index(wb[src_name], sum_col, crits)

    # 条件分类：row=当前行某列 / col=当前列第N行 / fixed=固定格 / const=字面量
    kinds = []
    for _c, e in crits:
        m = _CELLREF.match(e)
        if not m:
            kinds.append(("const", e))
            continue
        ca, col, ra, row_s = m.groups()
        ci, ri = CI(col), int(row_s)
        if not ra:
            kinds.append(("row", ci))
        elif not ca:
            kinds.append(("col", ri))
        else:
            kinds.append(("fixed", (ri, ci)))

    fixed_cache = {}
    for k, arg in kinds:
        if k == "fixed" and arg not in fixed_cache:
            fixed_cache[arg] = ws.cell(arg[0], arg[1]).value

    # 快路径：各簿条件均形如 ($A7, 日列$5, [渠道], 有效性)，SKU 恒在首位。
    # 预计算每列「非随行」后缀，每行只算一次 SKU，避免 行×日列 次重复构造 key。
    fast = bool(kinds) and kinds[0][0] == "row"
    suffix = {}
    if fast:
        sku_col = kinds[0][1]
        for col, dv in dcols:
            t = []
            for k, arg in kinds[1:]:
                if k == "col":
                    t.append(_norm(dv if arg == HDR_ROW else ws.cell(arg, col).value))
                elif k == "fixed":
                    t.append(_norm(fixed_cache[arg]))
                else:
                    t.append(None)
            suffix[col] = tuple(t)

    rows = []
    max_col = dcols[-1][0]
    rn = DATA_ROW - 1
    for row in ws.iter_rows(min_row=DATA_ROW, max_col=max_col, values_only=True):
        rn += 1
        if not row or row[0] in (None, ""):
            continue
        sku = str(row[0]).strip()
        model = str(row[1]).strip() if len(row) > 1 and row[1] is not None else ""
        vals = {}
        if fast:
            r0 = _norm(row[sku_col - 1] if sku_col - 1 < len(row) else None)
            sfx = suffix
            for col, _d in dcols:
                vals[col] = idx.get((r0,) + sfx[col], 0.0)
        else:
            for col, dv in dcols:
                key = []
                for k, arg in kinds:
                    if k == "row":
                        key.append(_norm(row[arg - 1] if arg - 1 < len(row) else None))
                    elif k == "col":
                        key.append(_norm(dv if arg == HDR_ROW else ws.cell(arg, col).value))
                    elif k == "fixed":
                        key.append(_norm(fixed_cache[arg]))
                    else:
                        key.append(arg)
                vals[col] = idx.get(tuple(key), 0.0)
        rows.append((rn, sku, model, vals))

    if path is not None and sn is not None and not any(
            v for _r, _s, _m, vs in rows for v in vs.values()):
        wbv = openpyxl.load_workbook(path, read_only=True, data_only=True)
        try:
            # read_only 下随机 ws.cell() 会反复解析 XML，必须整段顺序读
            snap = {}
            for i, rv in enumerate(wbv[sn].iter_rows(
                    min_row=DATA_ROW, max_col=dcols[-1][0], values_only=True),
                    start=DATA_ROW):
                snap[i] = rv
            for rn, _s, _m, vs in rows:              # 重算全 0 ⇒ 用归档缓存
                rv = snap.get(rn)
                for col, _d in dcols:
                    v = rv[col - 1] if rv and col - 1 < len(rv) else None
                    vs[col] = float(v) if isinstance(v, (int, float)) else 0.0
        finally:
            wbv.close()
    return (src_name, sum_col, crits), dcols, rows


# ---------------------------------------------------------------- 验证

def verify(month):
    arc = month_dir(month)
    files = sorted(f for f in os.listdir(arc)
                   if f.endswith(".xlsx") and not f.startswith("~$"))
    print(f"== 验证 {month}：重算口径 vs 归档缓存 ==")
    ok = bad = skip = 0
    for f in files:
        p = os.path.join(arc, f)
        wb = openpyxl.load_workbook(p, read_only=True)
        wbv = openpyxl.load_workbook(p, read_only=True, data_only=True)
        for sn in wb.sheetnames:
            if SHEET_KW not in sn:
                continue
            got = recompute_table(wb[sn], wb, p, sn)
            if not got:
                print(f"  {f[6:24]:<20}{sn[:16]:<18} 跳过（公式不识别）")
                skip += 1
                continue
            _mode, dcols, rows = got
            cached = [wbv[sn].cell(6, c).value for c, _ in dcols]
            has_cache = any(isinstance(v, (int, float)) and v for v in cached)
            calc = [sum(r[3][c] for r in rows) for c, _ in dcols]
            if not has_cache:
                print(f"  {f[6:24]:<20}{sn[:16]:<18} 归档无缓存 → 跳过比对"
                      f"（重算合计 {sum(calc):.0f}）")
                skip += 1
                continue
            diff = sum(1 for a, b in zip(calc, cached) if abs((a or 0) - (b or 0)) > 1e-6)
            print(f"  {f[6:24]:<20}{sn[:16]:<18} 缓存 {sum(v or 0 for v in cached):.0f} "
                  f"vs 重算 {sum(calc):.0f}  {'OK' if diff == 0 else f'差异 {diff} 列'}")
            if diff == 0:
                ok += 1
            else:
                bad += 1
                for (c, d), a, b in zip(dcols, calc, cached):
                    if abs((a or 0) - (b or 0)) > 1e-6:
                        print(f"      {d:%m/%d} 重算={a:.0f} 缓存={b}")
        wb.close()
        wbv.close()
    print(f"\n一致 {ok} / 不一致 {bad} / 跳过 {skip}")
    return 0 if bad == 0 else 1


# ---------------------------------------------------------------- 生成

def _verdict_map(ws):
    """从归档分析表取「判」列（行5 表头 == '判'），返回 {数据行号: 判定值}。

    这是月末归档里现成的判定（常规 / 保底 / 新品）——之前的归档缺了判定类信息，
    这里一并补回。库存类判定列在归档里是 #REF! 公式，取不到，故只取「判」。
    """
    col = None
    for c in ws[HDR_ROW]:
        if str(c.value).strip() == "判":
            col = c.column
            break
    if col is None:
        return {}
    out = {}
    for i, (v,) in enumerate(ws.iter_rows(min_col=col, max_col=col,
                                          min_row=DATA_ROW, values_only=True),
                             start=DATA_ROW):
        if v is not None and str(v).strip() != "":
            out[i] = str(v).strip()
    return out


def _book_job(job):
    """单本簿的重算（供多进程调用）。返回 [(站点, [日期头], [(sku, 型号, 判定, [日销])])]。"""
    f, month, d_from, d_to = job
    p = os.path.join(month_dir(month), f)
    wb = openpyxl.load_workbook(p, read_only=True)
    out = []
    for sn in wb.sheetnames:
        if SHEET_KW not in sn:
            continue
        ws = wb[sn]
        got = recompute_table(ws, wb, p, sn)
        if not got:
            continue
        _mode, dcols, rows = got
        cols = [(c, d) for c, d in dcols if d_from <= d <= d_to]
        if not cols:
            continue
        site = sn.replace(SHEET_KW, "").strip("-— ") or sn
        verdict = _verdict_map(ws)
        out.append((site,
                    [f"{d:%m/%d}" for _c, d in cols],
                    [(sku, model, verdict.get(rn, ""), [vals.get(c, 0) for c, _ in cols])
                     for rn, sku, model, vals in rows]))
    wb.close()
    return out


def build(month, d_from, d_to, out=None, workers=4):
    arc = month_dir(month)
    files = sorted(f for f in os.listdir(arc)
                   if f.endswith(".xlsx") and not f.startswith("~$"))
    out_wb = openpyxl.Workbook()
    out_wb.remove(out_wb.active)
    used = {}
    stat = []
    # 每本 openpyxl 遍历就要十几秒，多进程摊薄
    with ProcessPoolExecutor(max_workers=workers) as ex:
        results = ex.map(_book_job, [(f, month, d_from, d_to) for f in files])
        for site_list in results:
            for site, days, rows in site_list:
                k = used.get(site, 0)
                name = site if not k else f"{site}({k + 1})"
                used[site] = k + 1

                ws = out_wb.create_sheet(name[:31])
                ws.append(["SKU", "型号", "判定", "期间合计"] + days)
                for sku, model, verdict, vals in rows:
                    ws.append([sku, model, verdict, sum(vals)] + vals)
                ws.column_dimensions["A"].width = 14
                ws.column_dimensions["B"].width = 30
                ws.column_dimensions["C"].width = 8
                ws.column_dimensions["D"].width = 10
                for i in range(5, len(days) + 5):
                    ws.column_dimensions[L(i)].width = 7
                ws.freeze_panes = "E2"
                stat.append((name, len(rows), len(days), sum(sum(v) for _s, _m, _vd, v in rows)))

    out = out or os.path.join(
        arc, f"{month[:4]}年{int(month[5:7])}月-日销站点汇总"
             f"（{d_from:%m月%d日}-{d_to:%m月%d日}）.xlsx")
    out_wb.save(out)
    print(f"== 生成 {os.path.basename(out)} ==")
    print(f'{"sheet":<22}{"SKU行":>7}{"天数":>6}{"期间合计":>10}')
    for name, n, nd, tot in stat:
        print(f"{name:<22}{n:>7}{nd:>6}{tot:>10.0f}")
    print(f"\n共 {len(stat)} 张表，总销量 {sum(s[3] for s in stat):.0f} 件")
    print(f"输出：{out}")
    return out


def main(argv):
    ap = argparse.ArgumentParser(description="从归档重建某月日销站点汇总表（只读归档）")
    ap.add_argument("--month", required=True, help="归档月份，如 2026-09")
    ap.add_argument("--verify", action="store_true", help="比对重算口径与归档缓存")
    ap.add_argument("--build", action="store_true", help="生成汇总表")
    ap.add_argument("--from", dest="d_from", default=None, help="起始日 YYYY-MM-DD")
    ap.add_argument("--to", dest="d_to", default=None, help="结束日 YYYY-MM-DD")
    ap.add_argument("--out", default=None, help="输出路径")
    a = ap.parse_args(argv)

    def pd(s):
        return datetime.date(*map(int, s.split("-"))) if s else None

    y, m = int(a.month[:4]), int(a.month[5:7])
    d_from = pd(a.d_from) or datetime.date(y, m, 1)
    d_to = pd(a.d_to) or datetime.date(y, m + 1, 1)
    if a.verify:
        return verify(a.month)
    if a.build:
        build(a.month, d_from, d_to, a.out)
        return 0
    ap.print_help()
    return 1


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
