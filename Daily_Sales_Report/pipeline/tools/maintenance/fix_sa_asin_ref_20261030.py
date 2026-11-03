# -*- coding: utf-8 -*-
"""修复 Interflowd沙特库存销售表.xlsx 分析表 CH 列 ASIN 公式的行号错位。

背景（2026-10-30）：
  该簿由 build_interflow_sa.py 从 Interflowd阿联酋簿派生。派生时「属性列按 SKU 从原行带回」
  把 CH(ASIN) 列的**公式字符串原样抄了过来**，行引用仍是阿联酋簿的原行号（A42/A12/A54…），
  于是：
    · 指向空行 → XLOOKUP 查不到 → 显示 0 或空（用户看到的「ASIN 大量为空」）
    · 指向别人的行 → 显示成**别的 SKU 的 ASIN**（CH8=G583 却缓存了 B0HHMH12QL，那是 G587 的）

修法：XML 手术，只改 sheet6.xml 里 CH7..CH40 的 <f> 行引用 + 纠正旧缓存 <v>，
     不重存整个工作簿（避免丢 richData / vmlDrawing / CF）。
幂等：可重复执行。
"""
import os, re, shutil, zipfile, tempfile

BOOK = r'C:\Users\dell\Desktop\work-main\sale_file_gen\Daily_Sales_Report\1.每日销售数据表'
PATH = os.path.join(BOOK, 'Interflowd沙特库存销售表.xlsx')
SHEET_XML = 'xl/worksheets/sheet6.xml'      # 沙特-每日库存分析表
COL = 'CH'
ROW_LO, ROW_HI = 7, 40                      # 34 个数据行
SRC_SHEET = '沙特在途库存'


def read_zip_map(p):
    with zipfile.ZipFile(p) as z:
        return {n: z.read(n) for n in z.namelist()}


def build_src_map(files):
    """从「沙特在途库存」sheet 读 SKU -> ASIN。"""
    import openpyxl, io
    wb = openpyxl.load_workbook(io.BytesIO(open_bytes(files)), read_only=True, data_only=True)
    ws = wb[SRC_SHEET]
    m = {}
    for r in ws.iter_rows(min_row=3, max_col=3, values_only=True):
        s, a = r[0], r[2]
        if isinstance(s, str) and s.strip() and a:
            m[s.strip()] = str(a).strip()
    wb.close()
    return m


def open_bytes(files):
    import io
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as z:
        for n, b in files.items():
            z.writestr(n, b)
    buf.seek(0)
    return buf.getvalue()


def main():
    files = read_zip_map(PATH)
    src = build_src_map(files)
    print('在途库存 SKU->ASIN:', len(src))

    xml = files[SHEET_XML].decode('utf-8')
    orig = xml

    # 单元格行表：A列 SKU
    sku_of = {}
    for m in re.finditer(r'<c r="A(\d+)"[^>]*t="(?:s|str)"[^>]*>\s*<v>(.*?)</v>', xml, re.S):
        sku_of[int(m.group(1))] = m.group(2)
    # 共享字符串形式需另查 sharedStrings
    ss = []
    if 'xl/sharedStrings.xml' in files:
        sx = files['xl/sharedStrings.xml'].decode('utf-8')
        for m in re.finditer(r'<si>(.*?)</si>', sx, re.S):
            ss.append(''.join(re.findall(r'<t[^>]*>(.*?)</t>', m.group(1), re.S)))
    for m in re.finditer(r'<c r="A(\d+)"[^>]*t="s"[^>]*>\s*<v>(\d+)</v>', xml, re.S):
        i = int(m.group(2))
        if i < len(ss):
            sku_of[int(m.group(1))] = ss[i]

    fixed = filled = cleared = 0
    for r in range(ROW_LO, ROW_HI + 1):
        pat = re.compile(r'(<c r="%s%d"(?P<attr>[^>]*)>)(?P<inner>.*?)(</c>)' % (COL, r), re.S)
        m = pat.search(xml)
        if not m:
            continue
        inner = m.group('inner')
        fm = re.search(r'<f>(.*?)</f>', inner, re.S)
        if not fm:
            continue
        fml = fm.group(1)
        new_fml = re.sub(r'(XLOOKUP\(\s*\$?A)\$?\d+', r'\g<1>%d' % r, fml, count=1)
        if new_fml == fml:
            continue
        inner2 = inner.replace(fml, new_fml)

        sku = sku_of.get(r)
        asin = src.get(sku) if sku else None
        attr = m.group('attr') or ''
        # 纠正缓存值：有正确 ASIN 就写成字符串值，否则删掉旧缓存让 Excel 重算
        inner2 = re.sub(r'<v>.*?</v>', '', inner2, flags=re.S)
        if asin:
            inner2 = '<f>%s</f><v>%s</v>' % (new_fml, asin)
            attr = re.sub(r'\st="[^"]*"', '', attr) + ' t="str"'
            filled += 1
        else:
            inner2 = '<f>%s</f>' % new_fml
            attr = re.sub(r'\st="[^"]*"', '', attr)
            cleared += 1
        xml = xml[:m.start()] + ('<c r="%s%d"%s>' % (COL, r, attr)) + inner2 + '</c>' + xml[m.end():]
        fixed += 1

    print('改写公式 %d 格（补值 %d，留空待算 %d）' % (fixed, filled, cleared))
    if xml == orig:
        print('无改动，退出'); return

    files[SHEET_XML] = xml.encode('utf-8')
    bak = PATH + '.bak_pre_asin_fix'
    if not os.path.exists(bak):
        shutil.copy2(PATH, bak)
        print('备份:', os.path.basename(bak))

    tmp = PATH + '.tmp'
    with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as z:
        for n, b in files.items():
            z.writestr(n, b)
    os.replace(tmp, PATH)
    print('已保存')


if __name__ == '__main__':
    main()
