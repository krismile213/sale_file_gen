# -*- coding: utf-8 -*-
"""
mst_read.py —— 月销总表只读工具（XML 直读 + 表头自动定位分组）

为什么单独一个模块
    ① 月销总表【{年}年】sheet 是 297 列宽表，`pd.read_excel` 会系统性错列（印尼读成 3792、
       合计读成 11228），必须 XML 直读；
    ② 「分组组首列」写死在好几处（东南亚脚本 [5]、mst_xml.py、verify 脚本）—— 模板列结构一变
       就集体静默读错列。这里改成**从行 1 表头自动识别**，并用「亚马逊 == 十国之和」等恒等式
       自检；识别不出来就抛异常，不静默。

口径
    · 第 1 行 = 分组名（欧洲 / 英国 / 美国 … 总计），第 2 行 = 1月..12月
    · 第 3 行 = 公式合计（**读不得**），故列求和从第 4 行起
    · 某组第 M 月列 = 组首列 + (M − 1)

对外
    group_bases(path)                  -> {"欧洲": 3, ...}           组首列（1-based）
    group_month_totals(path, month)    -> {"欧洲": 24543, ...}       该组第 month 月合计
    group_map_warnings(totals)         -> ["..."]                    恒等式自检告警
"""
import os
import re
import zipfile

from openpyxl.utils import get_column_letter

# 月销总表里真实存在的分组名（顺序即表内顺序）
KNOWN = ["欧洲", "英国", "美国", "加拿大", "日本", "澳洲", "阿联酋", "沙特", "墨西哥", "印度",
         "亚马逊", "菲律宾", "泰国", "印尼", "新加坡", "马来", "越南", "虾皮", "东南亚",
         "京东", "eMAG", "TEMU", "美客多", "总计"]

# 别名：下游脚本/文档里常见的另一种叫法
ALIAS = {
    "虾皮境内": "虾皮", "虾皮跨境": "虾皮", "Shopee": "虾皮",
    "东南亚合计": "东南亚", "SEA": "东南亚",
    "eMag": "eMAG", "Emag": "eMAG", "TEMU平台": "TEMU", "Temu": "TEMU",
    "美客多墨西哥仓": "美客多", "Mercado": "美客多",
}

AMAZON_10 = ["欧洲", "英国", "美国", "加拿大", "日本", "澳洲", "阿联酋", "沙特", "墨西哥", "印度"]
SEA_8 = ["菲律宾", "泰国", "印尼", "新加坡", "马来", "越南", "虾皮", "东南亚"]
OTHER_4 = ["京东", "eMAG", "TEMU", "美客多"]


# ---------------------------------------------------------------- 基础
def sheet_xml(path, sheet_name):
    """返回某个 sheet 的 XML 文本（不经过 openpyxl，保留全部缓存值）"""
    with zipfile.ZipFile(path) as z:
        wbx = z.read("xl/workbook.xml").decode("utf-8")
        rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
        rm = {}
        for m in re.finditer(r"<Relationship\b[^>]*>", rels):
            t = m.group(0)
            i = re.search(r'Id="([^"]+)"', t)
            g = re.search(r'Target="([^"]+)"', t)
            if i and g:
                rm[i.group(1)] = g.group(1).lstrip("/")
        target = None
        for m in re.finditer(r"<sheet\b[^>]*>", wbx):
            t = m.group(0)
            n = re.search(r'name="([^"]+)"', t)
            r = re.search(r'r:id="([^"]+)"', t)
            if n and r and n.group(1) == sheet_name:
                target = rm[r.group(1)]
                break
        if target is None:
            raise KeyError(f"{os.path.basename(path)} 里找不到 sheet「{sheet_name}」")
        if not target.startswith("xl/"):
            target = "xl/" + target
        return z.read(target).decode("utf-8")


def col_sum(xml, col, min_row=4):
    """某一列从 min_row 起所有数值格之和（跳过公式合计行 / 空壳格）"""
    tot = 0.0
    for m in re.finditer(r'<c r="%s(\d+)"[^>]*?(?:/>|><v>([^<]*)</v></c>)' % col, xml):
        if int(m.group(1)) < min_row or m.group(2) in (None, ""):
            continue
        try:
            tot += float(m.group(2))
        except ValueError:
            continue
    return tot


# ---------------------------------------------------------------- 表头定位
_YEAR_SHEET_CACHE = {}


def year_sheet(path, year=None):
    """月销总表里该年份的 sheet（形如「2026年」）。

    ⚠️ 这张簿里同时存在「2025年」和「2026年」（还有两张 (备) 备份表），
       **按名字取第一个会取到 2025 年**（2026-11-10 实测踩到：总计读成 70,301）。
       所以：给了年份就用该年份；没给就取**最大年份**（= 最新那张），绝不取第一个。
    """
    key = (os.path.abspath(path), year)
    if key in _YEAR_SHEET_CACHE:
        return _YEAR_SHEET_CACHE[key]
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True)
    try:
        names = list(wb.sheetnames)
    finally:
        wb.close()
    years = [int(n[:4]) for n in names if re.fullmatch(r"\d{4}年", n)]
    if not years:
        raise KeyError(f"{os.path.basename(path)} 里没有「YYYY年」这样的 sheet（现有：{names}）")
    want = int(year) if year else max(years)
    hit = f"{want}年"
    if hit not in names:
        raise KeyError(f"{os.path.basename(path)} 里没有 sheet「{hit}」（现有：{names}）")
    _YEAR_SHEET_CACHE[key] = hit
    return hit


def group_bases(path, sheet_name=None, year=None):
    """从行 1 表头识别各分组组首列（1-based）。识别不全或缺关键组 → 抛异常。"""
    import openpyxl

    if sheet_name is None:
        sheet_name = year_sheet(path, year)
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    try:
        if sheet_name not in wb.sheetnames:
            raise KeyError(f"{os.path.basename(path)} 里找不到 sheet「{sheet_name}」")
        ws = wb[sheet_name]
        row1 = next(ws.iter_rows(min_row=1, max_row=1, max_col=400, values_only=True))
    finally:
        wb.close()

    bases = {}
    for i, v in enumerate(row1):
        if v is None:
            continue
        name = str(v).strip().replace("\n", "")
        if not name:
            continue
        name = ALIAS.get(name, name)
        if name in KNOWN and name not in bases:
            bases[name] = i + 1

    missing = [g for g in KNOWN if g not in bases]
    if missing:
        raise ValueError(
            f"月销总表行 1 表头里识别不到这些分组：{missing}（已识别 {list(bases)}）"
            f"—— 模板列结构可能变了，请核对并更新 mst_read.KNOWN")
    # 组间距必须 ≥12 列（一组一年 12 个月），否则说明识别错位
    seq = sorted(bases.items(), key=lambda kv: kv[1])
    for (g1, c1), (g2, c2) in zip(seq, seq[1:]):
        if c2 - c1 < 12:
            raise ValueError(f"分组列位可疑：{g1}@{c1} 与 {g2}@{c2} 间距 <12 列")
    return bases


def group_month_totals(path, month, sheet_name=None, year=None):
    """各分组第 month 月的合计：{组名: 数值}"""
    if not (1 <= int(month) <= 12):
        raise ValueError(f"月份必须是 1~12，收到 {month}")
    bases = group_bases(path, sheet_name, year)
    x = sheet_xml(path, sheet_name or year_sheet(path, year))
    out = {}
    for g, base in bases.items():
        out[g] = col_sum(x, get_column_letter(base + int(month) - 1))
    return out


# ---------------------------------------------------------------- 恒等式自检
def group_map_warnings(t, tol=0.5):
    """用恒等式检查「读出来的分组列位」是否自洽；返回告警列表（空 = 自洽）"""
    w = []

    def chk(label, a, b):
        if abs(a - b) > tol:
            w.append(f"{label}：{a:.0f} ≠ {b:.0f}（差 {a - b:+.0f}）")

    chk("亚马逊 == 十国之和", t.get("亚马逊", 0), sum(t.get(g, 0) for g in AMAZON_10))
    chk("东南亚 == 八国之和", t.get("东南亚", 0), sum(t.get(g, 0) for g in SEA_8[:7]))
    chk("总计 == 亚马逊+东南亚+其他四渠道", t.get("总计", 0),
        t.get("亚马逊", 0) + t.get("东南亚", 0) + sum(t.get(g, 0) for g in OTHER_4))
    return w


def totals_or_none(path, month, sheet_name=None, year=None):
    """容错版：文件不存在/结构异常时返回 (None, 原因)，供下游脚本降级使用"""
    if not path or not os.path.isfile(path):
        return None, f"月销总表不存在：{path}"
    try:
        return group_month_totals(path, month, sheet_name, year), None
    except Exception as e:                                    # noqa: BLE001
        return None, f"月销总表读取失败：{e}"


if __name__ == "__main__":
    import sys
    sys.stdout.reconfigure(encoding="utf-8")
    p = sys.argv[1] if len(sys.argv) > 1 else \
        os.path.join(os.path.dirname(os.path.abspath(__file__)),
                     "result", "数据归属周期", "202610", "月销售总表.xlsx")
    mth = int(sys.argv[2]) if len(sys.argv) > 2 else 10
    bs = group_bases(p)
    tt = group_month_totals(p, mth)
    print(f"[{os.path.basename(p)}] {mth} 月")
    for g in KNOWN:
        print(f"  {g:<8} 组首 {get_column_letter(bs[g]):<4} {tt[g]:>8.0f}")
    ws = group_map_warnings(tt)
    print("-" * 46)
    print("恒等式自检：" + ("全部通过 ✅" if not ws else "❌ " + "；".join(ws)))
