# -*- coding: utf-8 -*-
"""
trunk 成品对齐 3,485 口径 —— 直接复制 + 零损失补回
策略：5 个副本文件整表复制到 trunk/result，然后把"复制会丢的 4 处"从 trunk 备份补回：
  1. 三个次表的 Jan-Sep 2026 Cumulative 累计列（副本缺末列）→ 按 SKU 差值传播补回
  2. 东南亚主表年度 sheet AM~AT（1~8 月历史销量，副本为空）→ XML 手术搬回，不动 AU 缓存
  3. 三个次表多带的中文 sheet（冗余，与 Sheet1 逐 SKU 一致）→ 删除，保持 trunk 单 Sheet1 结构
幂等可重跑；每步改前 assert 防呆。
"""
import os, re, shutil, zipfile
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

REPO = r'D:\Desktop\work-main\sale_file_gen'
TR_SEA = os.path.join(REPO, 'trunk', 'result', '东南亚销量表', '202609')
TR_MON = os.path.join(REPO, 'trunk', 'result', '数据归属周期', '202609')
CP = os.path.join(REPO, '菲律宾重算_202609')
BAK = r'D:\dev-cache\sale_file_gen\backup\trunk对齐3,485前_20261016'

def n0(v):
    if v is None: return 0
    if isinstance(v, str):
        s = v.strip()
        if s in ('00', ''): return 0
        try: return float(s)
        except ValueError: return 0
    return v if isinstance(v, (int, float)) else 0

FILES = [
    (os.path.join(CP, '月销售总表.xlsx'), os.path.join(TR_MON, '月销售总表.xlsx')),
    (os.path.join(CP, '2026年东南亚销量表Southeast Asia Sales table.xlsx'), os.path.join(TR_SEA, '2026年东南亚销量表Southeast Asia Sales table.xlsx')),
    (os.path.join(CP, 'Philippines_202609_sales.xlsx'), os.path.join(TR_SEA, 'Philippines_202609_sales.xlsx')),
    (os.path.join(CP, 'Indonesia_202609_sales.xlsx'), os.path.join(TR_SEA, 'Indonesia_202609_sales.xlsx')),
    (os.path.join(CP, 'Thailand_202609_sales.xlsx'), os.path.join(TR_SEA, 'Thailand_202609_sales.xlsx')),
]

# ============ Step 0 · 复制前 sanity：副本关键值必须已是 3,485 口径 ============
print('========== Step 0 · 副本 sanity ==========')
# 月销总表：数据行求和（行3 是 SUM 公式无缓存）
wm = load_workbook(FILES[0][0], data_only=True)
ws = wm['2026年']
em = sum(n0(ws.cell(r, 143).value) for r in range(4, ws.max_row + 1))  # EM 菲9月
hs = sum(n0(ws.cell(r, 227).value) for r in range(4, ws.max_row + 1))  # HS 东南亚9月
ka = sum(n0(ws.cell(r, 287).value) for r in range(4, ws.max_row + 1))  # KA 总计9月
assert (em, hs, ka) == (3485, 11310, 63199), f'副本月销总表异常: {em}/{hs}/{ka}'
print(f'  ✓ 副本月销总表 数据行求和 菲={em:g} 东南亚={hs:g} 总计={ka:g}')
wm.close()
# 东南亚主表
wsea = load_workbook(FILES[1][0], data_only=True)
ph = wsea['Philippines 菲律宾销量表']
assert n0(ph['BD2'].value) == 3485, f"副本菲律宾 BD2={ph['BD2'].value}"
sea_t = wsea['SEA Total 东南亚总销量表']
assert n0(sea_t['CI2'].value) == 11310, f"副本 SEA Total CI2={sea_t['CI2'].value}"
yr = None
for sn in wsea.sheetnames:
    if '销量情况' in sn: yr = wsea[sn]; break
assert n0(yr['AU2'].value) == 11310 and n0(yr['AU3'].value) == 3485, f"副本年度 AU2/AU3={yr['AU2'].value}/{yr['AU3'].value}"
# AM~AT 必须为空（否则说明副本已有历史值，无需补回）
hist_nonzero = sum(1 for r in range(2, 16) for c in range(39, 47) if n0(yr.cell(r, c).value))
assert hist_nonzero == 0, f'副本 AM~AT 已有 {hist_nonzero} 个非零值，需重新评估'
print(f'  ✓ 副本东南亚主表 菲律宾 BD2=3485 SEA Total CI2=11310 年度 AU2=11310/AU3=3485，AM~AT 空（待补回）')
wsea.close()
# 次表 Sep Total 总和
for f, exp in [('Philippines_202609_sales.xlsx', 3485), ('Indonesia_202609_sales.xlsx', 4026), ('Thailand_202609_sales.xlsx', 3238)]:
    w = load_workbook(os.path.join(CP, f), data_only=True)
    s = w['Sheet1']
    tot = sum(n0(s.cell(r, s.max_column).value) for r in range(2, s.max_row + 1))
    assert tot == exp, f'副本 {f} Sep Total 总和={tot} 应 {exp}'
    assert 'Sheet1' in w.sheetnames and len([x for x in w.sheetnames if '销量表' in x]) == 1
    print(f'  ✓ 副本 {f} Sep Total 总和={tot}')
    w.close()

# ============ Step 1 · 整表复制到 trunk ============
print()
print('========== Step 1 · 整表复制 5 文件到 trunk/result ==========')
for src, dst in FILES:
    shutil.copy2(src, dst)
    assert os.path.getsize(src) == os.path.getsize(dst)
    print(f'  ✓ {os.path.basename(dst)}')

# ============ Step 2 · 次表：删中文 sheet + 补回累计列 ============
print()
print('========== Step 2 · 次表补回累计列（删冗余中文 sheet） ==========')
for f, delta_note in [('Philippines_202609_sales.xlsx', '+362'), ('Indonesia_202609_sales.xlsx', '+0'), ('Thailand_202609_sales.xlsx', '+0')]:
    dst = os.path.join(TR_SEA, f)
    wt = load_workbook(os.path.join(BAK, f), data_only=True)   # trunk 备份（含累计列）
    st = wt['Sheet1']
    # trunk 累计列 = 最后一列；Sep Total = 倒数第二列
    col_cum_t = st.max_column
    col_sep_t = col_cum_t - 1
    assert str(st.cell(1, col_cum_t).value or '').strip() == 'Jan-Sep 2026 Cumulative', f'{f} trunk 末列不是累计列'
    tmap = {}
    for r in range(2, st.max_row + 1):
        sku = st.cell(r, 1).value
        if sku: tmap[str(sku).strip()] = (n0(st.cell(r, col_sep_t).value), n0(st.cell(r, col_cum_t).value))
    wt.close()
    # 打开复制后的副本
    w = load_workbook(dst)
    # 2a. 删中文 sheet（保留 Sheet1）
    for sn in list(w.sheetnames):
        if sn != 'Sheet1':
            w.remove(w[sn])
            print(f'  [{f}] 删除冗余 sheet: {sn!r}')
    s = w['Sheet1']
    assert s.max_column == col_sep_t, f'{f} 副本列数 {s.max_column} != trunk Sep Total 列 {col_sep_t}'
    # 2b. 追加累计列
    cum_col = s.max_column + 1
    hdr_src = load_workbook(os.path.join(BAK, f))['Sheet1'].cell(1, col_cum_t)
    c = s.cell(1, cum_col, 'Jan-Sep 2026 Cumulative')
    c.font = hdr_src.font.copy(); c.fill = hdr_src.fill.copy()
    c.alignment = hdr_src.alignment.copy(); c.border = hdr_src.border.copy()
    n_new = n_miss = 0
    tot_new = tot_old = 0
    for r in range(2, s.max_row + 1):
        sku = s.cell(r, 1).value
        if not sku: continue
        sku = str(sku).strip()
        new_sep = n0(s.cell(r, col_sep_t).value)
        if sku in tmap:
            old_sep, old_cum = tmap[sku]
            new_cum = old_cum + (new_sep - old_sep)   # 差值传播
        else:
            new_cum = new_sep                          # 副本新增 SKU 行（如 G569）：历史为 0
            n_new += 1
            print(f'    [{f}] 新增 SKU 行 {sku}: 累计 = 9月销量 {new_sep:g}')
        cell = s.cell(r, cum_col, new_cum)
        cell.number_format = hdr_src.number_format
        tot_new += new_cum; tot_old += tmap.get(sku, (0, 0))[1]
    w.save(dst); w.close()
    # 2c. assert：列数、表头、总和恒等式
    w2 = load_workbook(dst, data_only=True)
    s2 = w2['Sheet1']
    assert s2.max_column == cum_col and str(s2.cell(1, cum_col).value).strip() == 'Jan-Sep 2026 Cumulative'
    assert w2.sheetnames == ['Sheet1'], f'{f} sheets={w2.sheetnames}'
    tot_re = sum(n0(s2.cell(r, cum_col).value) for r in range(2, s2.max_row + 1))
    assert tot_re == tot_new, f'{f} 累计列写后总和 {tot_re} != {tot_new}'
    print(f'  ✓ {f} 累计列已补回：总和 {tot_old:g} → {tot_new:g}（{delta_note}），sheets={w2.sheetnames}')
    w2.close()

# ============ Step 3 · 东南亚主表：XML 手术搬回 AM2:AT15 历史值（不碰 AU 缓存） ============
print()
print('========== Step 3 · 东南亚主表搬回 1~8 月历史列 AM~AT（XML 手术） ==========')
dst_sea = FILES[1][1]
# 3a. 从 trunk 备份读 AM~AT 静态值
wt = load_workbook(os.path.join(BAK, '2026年东南亚销量表Southeast Asia Sales table.xlsx'), data_only=True)
yt = None
for sn in wt.sheetnames:
    if '销量情况' in sn: yt = wt[sn]; break
hist = {}
for r in range(2, 16):
    for c in range(39, 47):  # AM~AT
        v = yt.cell(r, c).value
        if v is not None:
            hist[(r, c)] = v
wt.close()
print(f'  trunk 备份 AM~AT 非空格: {len(hist)} 个')
# 3a-2. 防呆：trunk 备份 AM~AT 是公式格（XLOOKUP）→ 必须搬公式文本+缓存，不能写死静态值
wf = load_workbook(os.path.join(BAK, '2026年东南亚销量表Southeast Asia Sales table.xlsx'), data_only=False)
yf = None
for sn in wf.sheetnames:
    if '销量情况' in sn: yf = wf[sn]; break
hist_f = {}
n_static = 0
for (r, c) in hist:
    fv = yf.cell(r, c).value
    if isinstance(fv, str) and fv.startswith('='):
        hist_f[(r, c)] = fv[1:]  # XML 里公式不带前导 =
    else:
        n_static += 1
        hist_f[(r, c)] = None      # 静态值：走 <v> 分支
print(f'  ✓ AM~AT: 公式格 {len(hist_f)-n_static} 个（搬公式+缓存）、静态格 {n_static} 个（搬值）')
wf.close()
# 3b. 定位年度 sheet 的 sheetN.xml
with zipfile.ZipFile(dst_sea) as z:
    wbxml = z.read('xl/workbook.xml').decode('utf-8')
    rels = z.read('xl/_rels/workbook.xml.rels').decode('utf-8')
    sheet_xml = z.read('xl/worksheets/sheet1.xml').decode('utf-8')  # 先占位，下面按名字找
m = re.search(r'<sheet[^>]*name="([^"]*销量情况[^"]*)"[^>]*r:id="(rId\d+)"', wbxml)
if not m:
    m = re.search(r'<sheet[^>]*r:id="(rId\d+)"[^>]*name="([^"]*销量情况[^"]*)"', wbxml)
    sheet_name, rid = m.group(2), m.group(1)
else:
    sheet_name, rid = m.group(1), m.group(2)
m2 = re.search(r'<Relationship[^>]*Id="%s"[^>]*Target="([^"]+)"' % rid, rels)
if not m2:
    m2 = re.search(r'<Relationship[^>]*Target="([^"]+)"[^>]*Id="%s"' % rid, rels)
target = m2.group(1).lstrip('/')          # 可能是 '/xl/worksheets/sheetN.xml' 或 'worksheets/sheetN.xml'
if not target.startswith('xl/'):
    target = 'xl/' + target
print(f'  年度 sheet: {sheet_name!r} → {target}')
with zipfile.ZipFile(dst_sea) as z:
    xml = z.read(target).decode('utf-8')

# 3c. 逐格写入（公式格写 <f>+缓存<v>，静态格写 <v>；保留原 s 样式属性）
from xml.sax.saxutils import escape as xesc

def set_cell(xml, ref, formula, cache):
    # 构造新单元格内容
    if formula is not None:
        inner = f'<f>{xesc(formula)}</f>' + (f'<v>{cache}</v>' if cache is not None else '')
    else:
        inner = f'<v>{cache}</v>' if cache is not None else ''
    # 先查成对格 <c r="REF" attrs>...</c>（属性组排除 / 与 >，防自闭合误配）
    m = re.search(r'<c r="%s"([^>/]*)>' % ref, xml)
    if m:
        start = m.end()
        end = xml.find('</c>', start)
        assert end != -1, f'{ref} 找不到闭合 </c>'
        seg = xml[start:end]
        assert '<c ' not in seg, f'{ref} 段跨格: {seg[:80]}'
        return xml[:start] + inner + xml[end:]
    # 自闭合 <c r="REF" attrs/>
    m2 = re.search(r'<c r="%s"([^>]*?)/>' % ref, xml)
    if m2:
        return xml[:m2.start()] + f'<c r="{ref}"{m2.group(1)}>{inner}</c>' + xml[m2.end():]
    # 完全没有该 <c> 元素 → 在 <row r="N"> 行首插入
    rownum = re.match(r'[A-Z]+(\d+)', ref).group(1)
    pat_row = re.compile(r'(<row r="%s"[^>]*>)' % rownum)
    mm = pat_row.search(xml)
    if not mm: raise AssertionError(f'{ref} 所在行 {rownum} 不存在')
    return pat_row.sub(lambda x: x.group(1) + f'<c r="{ref}">{inner}</c>', xml, count=1)

n_done = 0
for (r, c) in sorted(hist):
    ref = f'{get_column_letter(c)}{r}'
    xml = set_cell(xml, ref, hist_f.get((r, c)), hist[(r, c)])
    n_done += 1
# 3d. AU 列公式缓存必须仍在（防止误伤）
au_cache_ok = len(re.findall(r'<c r="AU\d+"[^>]*>(?:(?!</c>).)*?<v>', xml, re.S))
assert au_cache_ok >= 5, f'AU 缓存疑似丢失（只剩 {au_cache_ok} 个）'
print(f'  ✓ 已搬回 {n_done} 个历史值；AU 列公式缓存仍在（{au_cache_ok} 个）')
tmp = dst_sea + '.tmp'
with zipfile.ZipFile(dst_sea) as zin, zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
    for item in zin.infolist():
        data = zin.read(item.filename)
        if item.filename == target:
            data = xml.encode('utf-8')
        zout.writestr(item, data)
os.replace(tmp, dst_sea)
print('  已写回', dst_sea)

# ============ Step 4 · 终验自证（复读关键值） ============
print()
print('========== Step 4 · 终验自证 ==========')
# 东南亚主表
w = load_workbook(dst_sea, data_only=True)
ph = w['Philippines 菲律宾销量表']; stl = w['SEA Total 东南亚总销量表']; yr = None
for sn in w.sheetnames:
    if '销量情况' in sn: yr = w[sn]; break
assert n0(ph['BD2'].value) == 3485, f"菲律宾 BD2={ph['BD2'].value}"
assert n0(stl['CI2'].value) == 11310, f"SEA Total CI2={stl['CI2'].value}"
assert n0(yr['AU2'].value) == 11310 and n0(yr['AU3'].value) == 3485, f"AU2/AU3={yr['AU2'].value}/{yr['AU3'].value}"
hist_ok = all(n0(yr.cell(r, c).value) == n0(v) for (r, c), v in hist.items())
assert hist_ok, 'AM~AT 搬回后复读不一致'
print(f'  ✓ 东南亚主表: BD2=3485｜CI2=11310｜AU2=11310/AU3=3485｜AM~AT {len(hist)} 格历史值复读一致')
w.close()
# 次表
for f, exp_sep, exp_cum in [('Philippines_202609_sales.xlsx', 3485, 40442 + 362),
                            ('Indonesia_202609_sales.xlsx', 4026, None),
                            ('Thailand_202609_sales.xlsx', 3238, None)]:
    w = load_workbook(os.path.join(TR_SEA, f), data_only=True)
    s = w['Sheet1']
    tot = sum(n0(s.cell(r, s.max_column - 1).value) for r in range(2, s.max_row + 1))
    cum = sum(n0(s.cell(r, s.max_column).value) for r in range(2, s.max_row + 1))
    assert tot == exp_sep, f'{f} Sep Total={tot}'
    line = f'  ✓ {f}: Sep Total={tot}｜累计列总和={cum:g}｜sheets={w.sheetnames}'
    if exp_cum is not None:
        assert cum == exp_cum, f'{f} 累计总和 {cum} != {exp_cum}'
        line += f'（恒等式 40,442+362={exp_cum} ✓）'
    print(line)
    w.close()
# 月销总表
w = load_workbook(FILES[1][1], data_only=True)  # 占位防呆，无实际作用
w.close()
wm = load_workbook(os.path.join(TR_MON, '月销售总表.xlsx'), data_only=True)
ws = wm['2026年']
em = sum(n0(ws.cell(r, 143).value) for r in range(4, ws.max_row + 1))
hs = sum(n0(ws.cell(r, 227).value) for r in range(4, ws.max_row + 1))
ka = sum(n0(ws.cell(r, 287).value) for r in range(4, ws.max_row + 1))
assert (em, hs, ka) == (3485, 11310, 63199), f'月销总表 {em}/{hs}/{ka}'
print(f'  ✓ 月销售总表: 菲={em:g}｜东南亚={hs:g}｜总计={ka:g}')
wm.close()
print()
print('========== 全部完成：trunk 5 表已对齐 3,485 口径，历史列/累计列/单 Sheet1 结构零损失 ==========')
