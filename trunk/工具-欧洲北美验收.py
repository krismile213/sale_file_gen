# -*- coding: utf-8 -*-
"""
欧洲版 / 北美版销量表 · 完整体检（只读，不改任何文件）

用法
----
    PYTHONPATH="C:\\Users\\dell\\AppData\\Roaming\\Python\\Python314\\site-packages"
        "C:\\Python314\\python.exe" 工具-欧洲北美验收.py [归属周期] [快照日期]

检查 6 组（口径与东南亚版对齐）：

  A. 包完整性   zip 可解 / 无缩略图残留 / WPS 部件在 / 零悬空关系
  B. 观感       每 sheet 冻结窗格、无 topLeftCell 冲突、迷你图、条件格式、**只展开当月**的折叠
  C. 结构       行1 有列头 / A2=='SKU'（列头+合计混合行）/ 行3 起数据 / A 列 SKU 与 B 列型号一一对应
  D. 数值       月合计=各国之和；欧洲 9 国和=主表「欧洲」块、英国=「英国」块；
                北美逐国=主表对应分组块（美国含 TEMU、墨西哥含美客多）
  E. 历史表     按行1 表头定位当月列，行2(total) = 各国行之和 = Total 表当月合计
  F. 交付件     份数 / 表头 / 行数

⚠️ 三条容易写错的判据（本工具第一版就栽在这里，别再改回去）：
  ① **各 sheet 的基础列数不同**：国家 sheet = 2（A=SKU、B=型号，C 列起即月块），Total 表 = 6（A~F）。
     判断"行2 是不是合计行"不能只看 A2 —— 它是**列头+合计的混合行**，母本 SEA Total 也是 A2='SKU'。
  ② **月合计列永远可见**，"只展开当月"针对的是**渠道列**；拿整块去判折叠会全部误报。
  ③ **历史表是"行=国家、列=月"**，当月列要按行1 表头找，不能按 month_offset 直接算列号。

月份与块位一律用 `get_column_letter` 现算，不手推。
"""
import io
import os
import re
import posixpath
import sys
import zipfile

import openpyxl
from openpyxl.utils import get_column_letter as gl

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)

# ── 月销总表【当年】的分组块（顺序即列序，勿动）
MAIN_GROUPS = ['欧洲', '英国', '美国', '加拿大', '日本', '澳洲', '阿联酋', '沙特', '墨西哥',
               '印度', '亚马逊', '菲律宾', '泰国', '印尼', '新加坡', '马来', '越南', '虾皮境内',
               '东南亚', '京东', 'eMAG', 'TEMU平台', '美客多', '总计']
GI = {g: i for i, g in enumerate(MAIN_GROUPS)}
MAIN_BASE, MAIN_BW = 3, 12            # 【当年】base=3 列、月块 12 列

# ── 主表「欧洲各国」sheet：月块 11 列（10 国 + 合计），国家偏移
EU_SHEET = "2026年-欧洲各国每月销售情况"
EU_OFF = {'比利时': 0, '德国': 2, '西班牙': 3, '法国': 4, '意大利': 5,
          '荷兰': 6, '波兰': 7, '瑞典': 8, '爱尔兰': 9}     # 1=土耳其（本表无数据）
EU_SHEET_BASE, EU_SHEET_BW = 3, 11

def _expect_total(script_name):
    """从生成脚本里读 `EXPECT_TOTAL = <n>`（2026-11-02 改：原来本工具把合计写死成
    31837 / 9443，只对 202609 成立，换月后即使成品完全正确也会报假红 —— 已踩过一次）。
    脚本里 0 表示「不硬断言」，此时退回 0 由调用处跳过合计校验。"""
    try:
        src = io.open(os.path.join(HERE, script_name), encoding="utf-8").read()
        m = re.search(r"^EXPECT_TOTAL\s*=\s*(\d+)", src, re.M)
        return int(m.group(1)) if m else 0
    except Exception:
        return 0


EU = dict(
    tag="欧洲", sub="欧洲销量表", en_suffix="Europe",
    tsheet="EU Total 欧洲总销量表", n=10, base=6, hist_sheet=None,
    expect=_expect_total("europe_sale_file.py"),
    expect_src="生成脚本 EXPECT_TOTAL（202610 起动态读取，勿再写死）",
    # 欧洲主表只有「欧洲」一个大块（9 国合并），逐国无法单独对撞 → 走"9 国和 vs 欧洲块"
    pairs=[("前 9 国之和", ["欧洲"], 9), ("英国", ["英国"], 1)],
)
NA = dict(
    tag="北美", sub="北美销量表", en_suffix="North America",
    tsheet="NA Total 北美总销量表", n=3, base=6, hist_sheet=None,
    expect=_expect_total("north_america_sale_file.py"),
    expect_src="生成脚本 EXPECT_TOTAL（202610 起动态读取，勿再写死）",
    pairs=[("美国", ["美国", "TEMU平台"], 1), ("加拿大", ["加拿大"], 1), ("墨西哥", ["墨西哥", "美客多"], 1)],
)

PASS, FAIL = [], []


def ck(ok, label, detail=""):
    (PASS if ok else FAIL).append(label)
    print(f"    {'✅' if ok else '❌'} {label}" + (f"   {detail}" if detail else ""))


def norm_target(relspath, t):
    """rels 里的 Target → 包内路径（基准 = 剥掉 _rels/ 那一层：a/b/_rels/c.rels → a/b/）"""
    if t.startswith('/'):
        return t.lstrip('/')
    return posixpath.normpath(posixpath.join(posixpath.dirname(posixpath.dirname(relspath)), t))


def sheet_files(z):
    """按 workbook.xml 声明顺序返回 [(sheet 名, 包内路径)]"""
    names = re.findall(r'<sheet name="([^"]+)"', z.read('xl/workbook.xml').decode('utf-8'))
    rels = re.findall(r'Id="[^"]+"[^>]*Target="([^"]+)"',
                      z.read('xl/_rels/workbook.xml.rels').decode('utf-8'))
    tgt = {r.split('/')[-1]: r for r in rels}
    return [(nm, 'xl/' + tgt.get(f'sheet{i}.xml', f'worksheets/sheet{i}.xml'))
            for i, nm in enumerate(names, 1)]


def hidden_cols(xml):
    m = re.search(r'<cols>.*?</cols>', xml, re.S)
    if not m:
        return set()
    out = set()
    for a in re.findall(r'<col ([^/>]*)/?>', m.group(0)):
        if 'hidden="1"' not in a and 'hidden="true"' not in a:
            continue
        lo = re.search(r'min="(\d+)"', a)
        hi = re.search(r'max="(\d+)"', a)
        if lo:
            out.update(range(int(lo.group(1)), (int(hi.group(1)) if hi else int(lo.group(1))) + 1))
    return out


def check_month_blocks(z, key, base, nch, bw, month_offset, label):
    """只展开当月：当月渠道列可见；其余 11 个月渠道列必须全隐藏；月合计列恒可见"""
    hid = hidden_cols(z.read(key).decode('utf-8'))
    bad = []
    for m in range(1, 13):
        c0 = base + (m - 1) * bw
        chans = list(range(c0, c0 + nch))
        totalc = c0 + nch
        if m == month_offset:
            if any(c in hid for c in chans):
                bad.append(f"{m}月渠道列被折叠")
        else:
            shown = [gl(c) for c in chans if c not in hid]
            if shown:
                bad.append(f"{m}月渠道列未折叠{shown[:3]}")
        if totalc in hid:
            bad.append(f"{m}月合计列被折叠")
    ck(not bad, f"{label} 折叠正确（仅展开 {month_offset} 月）", "; ".join(bad[:4]))


def check_group(tag, path, cfg, month_offset, src_main):
    print("\n" + "=" * 78)
    print(f"### {cfg['tag']}版  {os.path.basename(path)}")
    if not os.path.exists(path):
        ck(False, f"文件存在：{path}")
        return
    z = zipfile.ZipFile(path)

    # ── A 包完整性
    print("  ── A. 包完整性")
    ck(z.testzip() is None, "zip 可完整解压")
    names = set(z.namelist())
    ck(not [n for n in names if n.lower().startswith('docprops/thumbnail')], "无缩略图残留")
    for p in ['customXml/item1.xml', 'customXml/item2.xml', 'docProps/custom.xml']:
        ck(p in names, f"WPS 部件 {p}")
    ck('KSOReadingLayout' in z.read('docProps/custom.xml').decode('utf-8'), "阅读模式 KSOReadingLayout=true")
    dangling = [(n, t) for n in sorted(x for x in names if x.endswith('.rels'))
                for t in re.findall(r'Target="([^"]+)"', z.read(n).decode('utf-8'))
                if norm_target(n, t) not in names]
    ck(not dangling, "零悬空关系", str(dangling[:2]))

    # ── B 观感
    print("  ── B. 观感（冻结 / 折叠 / 迷你图 / 条件格式）")
    sf = sheet_files(z)
    n_pane = n_bad = n_spark = n_cf = 0
    for _nm, key in sf:
        x = z.read(key).decode('utf-8')
        n_pane += ('<pane ' in x and 'state="frozen"' in x)
        n_bad += bool(re.search(r'<sheetView\b[^>]*\btopLeftCell=', x))
        n_spark += ('x14:sparklineGroups' in x)
        n_cf += x.count('<conditionalFormatting')
    ck(n_pane == len(sf), f"全部 {len(sf)} 张 sheet 都有冻结窗格", f"实际 {n_pane}")
    ck(n_bad == 0, "无 sheetView/@topLeftCell 冲突残留", f"残留 {n_bad} 处")
    ck(n_spark >= 1, f"迷你图 {n_spark} 张表")
    ck(n_cf > 0, f"条件格式 {n_cf} 条")
    n, base, bw = cfg['n'], cfg['base'], cfg['n'] + 1
    check_month_blocks(z, dict(sf)[cfg['tsheet']], base + n + 1, n, bw, month_offset,
                       f"{cfg['tsheet']}")
    # 国家 sheet（基础列 2、块 = 各渠道 + 合计列）
    # ⚠️ sheet 顺序固定为 [历史表, Total, 国家…] → 国家表从第 3 个起；写成 2+j 会把 Total 当国家表查
    for j in range(n):
        nm = wb_sheet_by_idx(sf, 3 + j)
        nch = cfg['nch_per'][j]
        check_month_blocks(z, dict(sf)[nm], 3, nch, nch + 1, month_offset, f"{nm}")

    # ── C 结构
    print("  ── C. 结构与基础列")
    wb = openpyxl.load_workbook(path)
    nohead = [nm for nm in wb.sheetnames
              if not (str(wb[nm].cell(1, 1).value or '').strip() or str(wb[nm].cell(1, 2).value or '').strip())]
    ck(not nohead, "每张 sheet 行1 有列头", str(nohead))
    ws = wb[cfg['tsheet']]
    ck(str(ws.cell(2, 1).value or '').strip() == 'SKU',
       "Total 行2 = 列头+合计混合行（A2='SKU'，与母本 SEA Total 一致）", f"A2={ws.cell(2,1).value!r}")
    a3 = str(ws.cell(3, 1).value or '').strip()
    ck(bool(a3) and a3 != 'SKU', "行3 起为 SKU 数据行", f"A3={a3!r}")
    miss_t = [ws.cell(r, 1).value for r in range(3, ws.max_row + 1)
              if ws.cell(r, 1).value and not ws.cell(r, 2).value]
    ck(not miss_t, f"{cfg['tsheet']} A 列 SKU 与 B 列型号一一对应",
       f"{ws.max_row - 2 - len(miss_t)}/{ws.max_row - 2} 有型号" + (f"，缺 {miss_t[:6]}" if miss_t else ""))
    miss_c = []
    for nm in wb.sheetnames[2:]:
        w = wb[nm]
        miss_c += [w.cell(r, 1).value for r in range(3, w.max_row + 1)
                   if w.cell(r, 1).value and not w.cell(r, 2).value]
    ck(not miss_c, "各国 sheet 缺型号 = 0", str(miss_c[:6]))

    # ── D 数值
    print("  ── D. 数值 vs 月销总表")
    src = src_main
    ws_main = src['2026年']

    # 主表行3 是 =SUM() 公式；脚本产物无缓存 → 直读为 0。改为对数据行(4~main_last)累加。
    main_last = max((r for r in range(4, ws_main.max_row + 1)
                     if ws_main.cell(r, 1).value not in (None, '')), default=4)

    def col_sum(sheet, col, last):
        return sum(int(sheet.cell(r, col).value) for r in range(4, last + 1)
                   if isinstance(sheet.cell(r, col).value, (int, float)))

    def blk(g, m=month_offset):
        return col_sum(ws_main, MAIN_BASE + GI[g] * MAIN_BW + m - 1, main_last)

    wv = openpyxl.load_workbook(path, data_only=True)[cfg['tsheet']]
    c0 = base + n + 1 + (month_offset - 1) * bw
    vals = [int(wv.cell(2, c0 + j).value or 0) for j in range(n)]
    pos = 0
    for label, groups, take in cfg['pairs']:      # pairs 按列序依次取，pos 顺推
        got = sum(vals[pos:pos + take])
        pos += take
        exp = sum(blk(g) for g in groups)
        ck(got == exp, f"{label} {got:>6}", "= " + "+".join(f"{g}{blk(g)}" for g in groups))
    tot = sum(vals)
    ck(int(wv.cell(2, c0 + n).value or 0) == tot, f"月合计列 = 各国之和 {tot}")
    if cfg['expect']:
        ck(tot == cfg['expect'], f"{cfg['tag']}合计 = {cfg['expect']}", f"实际 {tot}（{cfg['expect_src']}）")
    else:
        print(f"    ℹ️ {cfg['tag']}合计 {tot}（生成脚本未硬断言 EXPECT_TOTAL=0，跳过合计校验）")

    if cfg['tag'] == '欧洲':                     # 参照：主表「欧洲各国」sheet 逐国值
        es = src[EU_SHEET]
        c0e = EU_SHEET_BASE + (month_offset - 1) * EU_SHEET_BW
        eu_last = max((r for r in range(4, es.max_row + 1)
                       if es.cell(r, 1).value not in (None, '')), default=4)
        ref = {cn: col_sum(es, c0e + off, eu_last) for cn, off in EU_OFF.items()}
        mine = {cn: vals[i] for i, cn in enumerate(list(EU_OFF))}
        gap = sum(mine.values()) - sum(ref.values())
        print(f"    ℹ️ 主表「欧洲各国」sheet {month_offset} 月逐国：{ref}")
        print(f"       本表 9 国逐国：{mine}")
        print(f"       本表比主表多 {gap} 件 —— 主表该 sheet 只写固定行数（历史天花板，行4~{eu_last}），"
              f"新增 SKU 不落入；本表已按长表补齐")

    # ── E 历史表
    print("  ── E. 历史表")
    wh = wb[wb.sheetnames[0]]
    want = f"2026年{month_offset}月"
    col = None
    for c in range(2, wh.max_column + 1):
        if str(wh.cell(1, c).value or '').strip() == want:
            col = c
            break
    if col is None:
        ck(False, f"历史表找到 {want} 列")
    else:
        s = sum(wh.cell(r, col).value or 0 for r in range(3, wh.max_row + 1))
        ck(int(s) == tot, f"历史表「{want}」列（行3~{wh.max_row}）各国之和 = {tot}", f"{gl(col)} 列，实际 {int(s)}")
        ck(int(wh.cell(2, col).value or 0) == tot, "历史表 total 行 = 各国之和", f"实际 {wh.cell(2, col).value}")

    # ── F 交付件
    print("  ── F. 交付件")
    d = os.path.dirname(path)
    fs = sorted(f for f in os.listdir(d) if f.endswith('_sales.xlsx'))
    ck(len(fs) == n, f"份数 {len(fs)}（应 {n}）")
    if fs:
        w2 = openpyxl.load_workbook(os.path.join(d, fs[0]), data_only=True)
        hd = [w2.active.cell(1, c).value for c in range(1, w2.active.max_column + 1)]
        ck(all(k in hd for k in ('SKU', 'Model', 'Type', 'Brand')), "表头含 SKU/Model/Type/Brand", str(hd))


def wb_sheet_by_idx(sf, idx):
    return sf[idx - 1][0]


def main(year_period='202609', snap='20261006'):
    main_xlsx = os.path.join(ROOT, 'trunk', 'result', '数据归属周期', year_period, '月销售总表.xlsx')
    if not os.path.exists(main_xlsx):
        print(f"⚠️ 找不到月销总表：{main_xlsx}")
        return 1
    src = openpyxl.load_workbook(main_xlsx, data_only=True)
    month_offset = int(year_period[-2:]) % 100

    EU['nch_per'] = [1] * 10
    NA['nch_per'] = [2, 1, 2]                   # 美国=亚马逊+TEMU；加拿大=亚马逊；墨西哥=亚马逊+美客多
    for cfg in (EU, NA):
        cfg['path'] = os.path.join(ROOT, 'trunk', 'result', cfg['sub'], year_period,
                                   f"2026年{cfg['tag']}销量表{cfg['en_suffix']} Sales table.xlsx")
        check_group(cfg['tag'], cfg['path'], cfg, month_offset, src)

    print("\n" + "=" * 78)
    print(f"通过 {len(PASS)} 项 / 失败 {len(FAIL)} 项")
    if FAIL:
        print("❌ 未通过：")
        for f in FAIL:
            print("   -", f)
    else:
        print("✅ 全部通过")
    return 1 if FAIL else 0


if __name__ == '__main__':
    sys.exit(main(*sys.argv[1:3]))
