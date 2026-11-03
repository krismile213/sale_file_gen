# -*- coding: utf-8 -*-
"""向各簿分析表注入「红色小基数档」条件格式（2026-10-15 启用）。

规则形态逐表从**现有红色规则**解析得到（不写死列号/倍数/门槛）：
    现有（大基数档，一字不改）: AND(日销>上周均*倍数, 上周均>门槛)
    新增（小基数档）        : AND(上周均>0, 上周均<=门槛, 日销>=2, 日销>上周均*4)

两档按「上周均<=门槛 / >门槛」互斥；与蓝色规则无交集（蓝色要求上周均>2.5），
因此不抢色、不改现有任何一格的颜色。复用现有红色 dxf（不新增颜色），
priority 取本表现有最大值+1（最低优先级，先匹配大基数档）。

用法：
    python pipeline/tools/inject_small_base_cf.py              # 注入 12 簿全部分析表
    python pipeline/tools/inject_small_base_cf.py --dry-run    # 只报告不写文件
    python pipeline/tools/inject_small_base_cf.py --only 日本   # 只跑文件名含关键字的簿
    python pipeline/tools/inject_small_base_cf.py --mul 4 --min-qty 2   # 调灵敏度
"""
import os
import re
import sys
import zipfile

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'core'))
import config as C  # noqa: E402

RED_RGB = 'FF0000'   # dxf 红色特征：bgColor rgb="FFFF0000" 含 FF0000（粉 FFFBB2B2/蓝 FF51D9FF 不含）

_CF_RE = re.compile(r'<conditionalFormatting sqref="([^"]+)">(.*?)</conditionalFormatting>', re.S)
_RULE_RE = re.compile(r'<cfRule\b([^>]*?)(?:/>|>(.*?)</cfRule>)', re.S)
_FORMULA_RE = re.compile(r'<formula>(.*?)</formula>', re.S)
_B1_RE = re.compile(r'(<c r="B1"[^>]*t="inlineStr"><is><t>)(.*?)(</t></is></c>)', re.S)
_RED_CLAUSE = re.compile(r'(?<!粉)红色：[^；]*(?:；|$)')
_ADDITION = re.compile(r'或上月日均≤[\d.]+且日销≥\d+件、超过上月日均[\d.]+倍；')
_MUL_TERM = re.compile(r'(\$?[A-Z]{1,3}\$?\d+)\s*>\s*(\$?[A-Z]{1,3}\$?\d+)\s*\*\s*([\d.]+)')
_GATE_TERM = re.compile(r'(?<![A-Z0-9$])(\$?[A-Z]{1,3}\$?\d+)\s*>\s*([\d.]+)(?![0-9.])')


def esc(s):
    """公式进 XML：只转 < 与 &（> 按本仓表现有风格保持原样）。"""
    return s.replace('&', '&amp;').replace('<', '&lt;')


def sheet_map(parts):
    """{sheet名: 'xl/worksheets/sheetN.xml'}（兼容 Target 带 / 前缀或相对路径）"""
    wbxml = parts['xl/workbook.xml'].decode('utf-8', 'ignore')
    rels = {}
    for tag in re.findall(r'<Relationship\b[^>]*?/?>',
                          parts['xl/_rels/workbook.xml.rels'].decode('utf-8', 'ignore')):
        rid = re.search(r'\bId="(rId\d+)"', tag)
        tgt = re.search(r'\bTarget="([^"]+)"', tag)
        if rid and tgt:
            rels[rid.group(1)] = tgt.group(1)
    out = {}
    for tag in re.findall(r'<sheet\b[^>]*?/?>', wbxml):
        nm = re.search(r'\bname="([^"]+)"', tag)
        rid = re.search(r'\br:id="(rId\d+)"', tag)
        if not nm or not rid:
            continue
        tgt = rels.get(rid.group(1), '')
        if tgt.startswith('/'):
            tgt = tgt.lstrip('/')
        elif not tgt.startswith('xl/'):
            tgt = 'xl/' + tgt
        if tgt in parts:
            out[nm.group(1)] = tgt
    return out


def unesc(s):
    """XML 实体还原（公式在 sheet XML 里是 &gt;/&lt; 转义态）。"""
    return (s.replace('&lt;', '<').replace('&gt;', '>')
             .replace('&quot;', '"').replace('&apos;', "'").replace('&amp;', '&'))


def find_red_rule(xml, dxfs):
    """找「日销 vs 上周均」红色规则：dxf 含 FF0000，且公式同时含
    *项（X>Y*倍数）与门槛项（Y>门槛，无数值运算）。取 priority 最小者。
    返回 dict(sqref, dxf_id, daily_ref, avg_ref, mul, gate, priority)；找不到返回 None。"""
    best = None
    for m in _CF_RE.finditer(xml):
        sq, body = m.group(1), m.group(2)
        for r in _RULE_RE.finditer(body):
            attrs, inner = r.group(1), (r.group(2) or '')
            dxf_m = re.search(r'dxfId="(\d+)"', attrs)
            fm = _FORMULA_RE.search(inner)
            if not dxf_m or not fm:
                continue
            dxf_id = int(dxf_m.group(1))
            if dxf_id >= len(dxfs) or RED_RGB not in dxfs[dxf_id]:
                continue
            formula = unesc(fm.group(1))
            body = formula.strip()
            if body.upper().startswith('AND(') and body.endswith(')'):
                body = body[4:-1]
            terms = [t.strip() for t in body.split(',')]
            mul_m = None
            for t in terms:
                mm = _MUL_TERM.search(t)
                if mm:
                    mul_m = mm
                    break
            if not mul_m:
                continue
            avg_ref = mul_m.group(2)
            gate = None
            for t in terms:
                gm = _GATE_TERM.search(t)
                if gm and gm.group(1) == avg_ref:
                    gate = gm.group(2)
                    break
            if gate is None:
                continue
            pm = re.search(r'priority="(\d+)"', attrs)
            cand = dict(sqref=sq, dxf_id=dxf_id, daily_ref=mul_m.group(1),
                        avg_ref=avg_ref, mul=mul_m.group(3), gate=gate,
                        priority=int(pm.group(1)) if pm else 0)
            if best is None or cand['priority'] < best['priority']:
                best = cand
    return best


def update_legend(xml, mul, min_qty, gate):
    """B1 图例（inlineStr）在「红色」分句后追加小基数档说明：只追加、不改原句。
    先清除既有追加句（含历史上错位插入的）再重插 → 自愈且幂等。
    返回 (新xml, 是否改动)。"""
    m = _B1_RE.search(xml)
    if not m:
        return xml, False
    orig = unesc(m.group(2))
    txt = _ADDITION.sub('', orig)
    clause = _RED_CLAUSE.search(txt)
    if not clause:
        return xml, False
    addition = f'或上月日均≤{gate}且日销≥{min_qty}件、超过上月日均{mul:g}倍；'
    txt2 = txt[:clause.end()] + addition + txt[clause.end():]
    if txt2 == orig:
        return xml, False
    return xml[:m.start(2)] + esc(txt2) + xml[m.end(2):], True


def inject_book(path, mul=4, min_qty=2, dry=False):
    """注入一本簿。返回 [(sheet, sqref, dxf_id, daily_ref, avg_ref, mul_old, gate, formula)]"""
    with zipfile.ZipFile(path) as z:
        order = z.namelist()
        parts = {n: z.read(n) for n in order}
    smap = sheet_map(parts)
    styles = parts.get('xl/styles.xml', b'').decode('utf-8', 'ignore')
    dxfs = re.findall(r'<dxf>.*?</dxf>', styles, re.S)
    out = []
    for sn, tgt in sorted(smap.items()):
        if '分析' not in sn:
            continue
        xml = parts[tgt].decode('utf-8')
        red = find_red_rule(xml, dxfs)
        if not red:
            print(f'    [{sn}] 未找到「日销>上周均*倍数 且 上周均>门槛」红色规则，跳过（需人工核对）')
            continue
        formula = (f'AND({red["avg_ref"]}>0,{red["avg_ref"]}<={red["gate"]},'
                   f'{red["daily_ref"]}>={min_qty},{red["daily_ref"]}>{red["avg_ref"]}*{mul:g})')
        if esc(formula) in xml:
            print(f'    [{sn}] 小基数档已存在（{formula}），跳过')
            out.append((sn, red['sqref'], red['dxf_id'], red['daily_ref'],
                        red['avg_ref'], red['mul'], red['gate'], formula))
            xml, leg = update_legend(xml, mul, min_qty, red['gate'])
            if leg:
                parts[tgt] = xml.encode('utf-8')
                print(f'    [{sn}] B1 图例已追加小基数档说明')
            continue
        new_prio = max([int(p) for p in re.findall(r'priority="(\d+)"', xml)] or [0]) + 1
        block = (f'<conditionalFormatting sqref="{red["sqref"]}">'
                 f'<cfRule type="expression" dxfId="{red["dxf_id"]}" priority="{new_prio}">'
                 f'<formula>{esc(formula)}</formula></cfRule></conditionalFormatting>')
        idx = xml.find('<conditionalFormatting')
        if idx < 0:
            print(f'    [{sn}] 表内无 conditionalFormatting 块，跳过')
            continue
        xml = xml[:idx] + block + xml[idx:]
        xml, leg = update_legend(xml, mul, min_qty, red['gate'])
        parts[tgt] = xml.encode('utf-8')
        out.append((sn, red['sqref'], red['dxf_id'], red['daily_ref'],
                    red['avg_ref'], red['mul'], red['gate'], formula))
        print(f'    [{sn}] 红色小基数档: {formula}')
        print(f'        区域 {red["sqref"]} | 复用 dxf {red["dxf_id"]} | '
              f'priority {new_prio} | 原大基数档 日销{red["daily_ref"]} vs 上周均{red["avg_ref"]}'
              f' ×{red["mul"]} 门槛{red["gate"]}'
              + (' | B1 图例已追加' if leg else ''))
    if not out or dry:
        return out
    tmp = path + '.tmp'
    with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zo:
        for n in order:
            zo.writestr(n, parts[n])
    try:
        os.replace(tmp, path)
    except OSError:
        import shutil
        shutil.copyfile(tmp, path)
        os.remove(tmp)
    return out


def main(argv):
    mul, min_qty, dry, only = 4, 2, False, None
    i = 0
    while i < len(argv):
        a = argv[i]
        if a == '--dry-run':
            dry = True
        elif a == '--mul':
            i += 1
            mul = float(argv[i])
        elif a == '--min-qty':
            i += 1
            min_qty = int(argv[i])
        elif a == '--only':
            i += 1
            only = argv[i]
        else:
            raise SystemExit(f'未知参数 {a}')
        i += 1
    books = [w['file'] for w in C.WORKBOOKS if not only or only in w['file']]
    total = 0
    for f in books:
        path = os.path.join(C.WORK_DIR, f)
        if not os.path.exists(path):
            print(f'[跳过] {f} 不存在')
            continue
        print(f'[注入] {f}' + ('（dry-run）' if dry else ''))
        got = inject_book(path, mul=mul, min_qty=min_qty, dry=dry)
        total += len(got)
    print(f'\n合计 {total} 张分析表{"已具备" if dry else "已注入"}小基数档'
          f'（×{mul:g}、≥{min_qty} 件）' + ('；dry-run 未写文件' if dry else ''))
    return 0


if __name__ == '__main__':
    sys.exit(main(sys.argv[1:]))
