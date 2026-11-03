# -*- coding: utf-8 -*-
"""修补「列残缺」的分析表数据行 —— XML 手术版（2026-10-23）

## 为什么不用 openpyxl
欧盟簿解压后 137MB（`每日销售数据` 单表 53.7MB / 78 万格），openpyxl 普通模式
`load_workbook` 在这台机器上会直接把进程拖死（SIGTERM / 内存耗尽）。
zip 级复制只要 1.4 秒，分析表 XML 只有 5~6MB —— 所以**只碰这两个 sheet 的 XML**，
其余部件整段照抄，不经过 openpyxl。

## 背景
2026-10-23 前，`append_sku_rows_auto` 克隆的是**末数据行**；而欧盟簿两张表的末行
本身列就残缺（历史遗留），于是新补的行也只长出半截：

| 表 | 残缺行 | 现有列数 | 应有列数 | 缺什么 |
|---|---|---|---|---|
| 英国-每日库存分析表 | r466 G598 | 54 (到 BV) | 110 (到 DF) | 日销列、AD/AE/AF/AG、BW..DF |
| 英国-每日库存分析表 | r467~472 | 55 | 110 | 同上 |
| 欧洲-每日库存分析表 | r465 G598 | 50 (到 CF) | 120 (到 DP) | AS..CE 日销列、AD..AR、CG..DP |
| 欧洲-每日库存分析表 | r466~468 | 51~52 | 120 | 同上 |

后果：这些行没有日销列（算不出当月销量）、没有可售天数、没有月度汇总，
条件格式也覆盖不到 —— 看起来在表里，实际是「半行」。

## 做法
1. 找**列最完整**的数据行当模板（默认 r7）。
2. 对每个残缺行：把模板行的每个 `<c>` 拷过来，`r=` 改成目标行号，
   单元格内的相对行号引用（`I465` → `I<目标行>`）用 `_shift_refs` 平移。
3. 保留原行的 A/B/C/D 值（SKU/型号/分类/状态）。
4. 货件列（I..差额列-1）的值清掉，只留样式 —— 与 openpyxl 版 clone_row 口径一致。
5. **不删行、不改行序** → 库存记录/透视表/断货分析的引用全部不受影响。

## 用法
```bash
python tools/maintenance/repair_row_width_20261023.py            # 干跑，只打印计划
python tools/maintenance/repair_row_width_20261023.py --apply    # 落盘
```
幂等：已是完整列数的行不会再动。
"""
import os as _os, sys as _sys

_here = _os.path.dirname(_os.path.abspath(__file__))
DSR = _here
for _i in range(6):
    if _os.path.isdir(_os.path.join(DSR, "1.每日销售数据表")):
        break
    DSR = _os.path.dirname(DSR)

import os, sys, re, zipfile, shutil, time

# ---- 需要处理的 (簿, 分析表) ----
TARGETS = [
    ("欧盟英国库存销售表.xlsx", "英国-每日库存分析表"),
    ("欧盟英国库存销售表.xlsx", "欧洲-每日库存分析表"),
]
BOOKS_DIR = _os.path.join(DSR, "1.每日销售数据表")
BACKUP = _os.path.join(DSR, ".workbuddy", "backup_rowwidth_20261023")
DATA_FIRST_ROW = 7


# ---------------- zip / sheet 定位 ----------------
def _sheet_parts(path):
    """{sheet name: zip 内路径}"""
    z = zipfile.ZipFile(path)
    wbxml = z.read("xl/workbook.xml").decode("utf-8")
    rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    rmap = {}
    for tag in re.findall(r"<Relationship\b[^>]*/?>", rels):
        rid = re.search(r'\bId="([^"]+)"', tag)
        tgt = re.search(r'\bTarget="([^"]+)"', tag)
        if rid and tgt:
            rmap[rid.group(1)] = tgt.group(1)
    out = {}
    for m in re.finditer(r'<sheet[^>]*name="([^"]+)"[^>]*r:id="([^"]+)"', wbxml):
        t = rmap.get(m.group(2), "").lstrip("/")
        if not t.startswith("xl/"):
            t = "xl/" + t
        out[m.group(1)] = t
    z.close()
    return out


# ---------------- 列号 / 引用平移 ----------------
def ci(col):
    n = 0
    for ch in col:
        n = n * 26 + (ord(ch) - 64)
    return n


def cl(n):
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def shift_refs(text, src_row, dst_row):
    """把公式里第 src_row 行的相对引用平移到 dst_row（保持 $ 语义）。

    只动「行号恰好等于 src_row」的引用；$ 绝对行（$5 表头、$1476 等）不动。
    覆盖：A1 / $A1 / A$1 / $A$1 / 'sheet name'!A1 / 整行 A:A 之外的单列 A1:B2。
    """
    if src_row == dst_row:
        return text
    pat = re.compile(r"(\$?)([A-Z]{1,3})(\$?)(\d+)")

    def rep(m):
        dc, col, dr, row = m.group(1), m.group(2), m.group(3), int(m.group(4))
        if dr == "$":              # 绝对行，不动
            return m.group(0)
        if row != src_row:         # 不是模板行，不动
            return m.group(0)
        return f"{dc}{col}{dr}{dst_row}"

    return pat.sub(rep, text)


# ---------------- 单元格解析 ----------------
CELL_RE = re.compile(r'<c r="([A-Z]+)(\d+)"([^>]*?)(?:/>|>(.*?)</c>)', re.S)


def parse_row(body):
    """返回 [(col_letter, attrs, inner_or_None), ...]，按列序。"""
    out = []
    for m in CELL_RE.finditer(body):
        out.append((m.group(1), m.group(3), m.group(4)))
    return out


def cell_text(attrs, inner):
    """取单元格的静态文本（inlineStr / v）。"""
    if inner is None:
        return ""
    t = re.search(r"<t[^>]*>(.*?)</t>", inner, re.S)
    if t:
        return t.group(1)
    v = re.search(r"<v>(.*?)</v>", inner, re.S)
    if v:
        return v.group(1)
    return ""


def build_cell(col, row, attrs, inner):
    if inner is None:
        return f'<c r="{col}{row}"{attrs}/>'
    return f'<c r="{col}{row}"{attrs}>{inner}</c>'


def template_row(rows, data_first=DATA_FIRST_ROW):
    """列最多的数据行 → (行号, cells)。"""
    best, best_n = None, -1
    for rn, body in rows.items():
        if rn < data_first:
            continue
        cells = parse_row(body)
        if not cells:
            continue
        if ws_has_sku(cells):
            n = len(cells)
            if n > best_n:
                best, best_n = rn, n
    return best, (rows.get(best) if best else None)


def ws_has_sku(cells):
    for col, attrs, inner in cells:
        if col == "A":
            return bool(cell_text(attrs, inner).strip())
    return False


def main():
    apply = "--apply" in sys.argv
    total = 0
    for fn, sn in TARGETS:
        p = os.path.join(BOOKS_DIR, fn)
        print(f"\n===== {fn} | {sn}")
        parts = _sheet_parts(p)
        if sn not in parts:
            print(f"  [!] 找不到 sheet: {sn!r}（现有 {list(parts)[:6]}…）")
            continue
        part = parts[sn]

        z = zipfile.ZipFile(p)
        x = z.read(part).decode("utf-8")
        z.close()

        # 拆行
        rows = {}
        for m in re.finditer(r'<row r="(\d+)"([^>]*)>(.*?)</row>', x, re.S):
            rows[int(m.group(1))] = m.group(3)
        if not rows:
            print("  [!] 没解析到任何 <row>，跳过")
            continue

        tpl_rn, tpl_body = template_row(rows)
        tpl_cells = parse_row(tpl_body)
        full = len(tpl_cells)
        print(f"  模板行 r{tpl_rn}，{full} 列（到 {tpl_cells[-1][0]}）")

        # 找残缺行
        bad = []
        for rn in sorted(rows):
            if rn < DATA_FIRST_ROW:
                continue
            cells = parse_row(rows[rn])
            if not cells or not ws_has_sku(cells):
                continue
            if len(cells) < full * 0.75:
                bad.append((rn, cells))
        if not bad:
            print("  无需修补（所有数据行列完整）")
            continue

        # 货件列范围：第 9 列（I）~ 差额列前 1 列。
        # 差额列 = 模板里「=X7-Y7」这种单行差集公式（英国 AB `=G7-H7` / 欧洲 Z `=G7-H7`），
        # 从左往右第一个命中即为差额列。**不能匹配 `<f>..</f>` 整串**——XML 里还带
        # `<f>` 之外的属性，且 `G7-H7` 前面可能有 `=` 被转义。
        ship_lo = 9
        diff_hi = None
        for col, attrs, inner in tpl_cells:
            if inner and re.search(r"<f>\s*=?\s*\$?[A-Z]+\d+-\$?[A-Z]+\d+\s*</f>", inner):
                diff_hi = ci(col) - 1
                break
        if diff_hi is None:            # 兜底：按英国/欧洲惯例
            diff_hi = 29 if sn.startswith("英国") else 39
        print(f"  货件列清空范围: {cl(ship_lo)}..{cl(diff_hi)}")

        new_rows = {}
        for rn, old_cells in bad:
            old_vals = {}
            for col, attrs, inner in old_cells:
                if col in ("A", "B", "C", "D"):
                    old_vals[col] = (attrs, inner)
            out = []
            for col, attrs, inner in tpl_cells:
                c = ci(col)
                txt = inner
                if txt is not None:
                    txt = shift_refs(txt, tpl_rn, rn)
                if ship_lo <= c <= diff_hi and rn != tpl_rn:
                    # 货件列：只留样式，值清空（与 clone_row(clear_range=) 口径一致）
                    txt = None
                if col in old_vals and col in ("A", "B", "C", "D"):
                    o_attrs, o_inner = old_vals[col]
                    # 只保留文本值，样式跟模板（原行样式往往是残缺的）
                    if o_inner and "<t>" in o_inner:
                        txt = o_inner
                    elif o_inner and "<f>" not in o_inner:
                        txt = o_inner
                    else:
                        txt = None
                out.append(build_cell(col, rn, attrs, txt))
            new_rows[rn] = "".join(out)
            print(f"  r{rn} {cell_text(*[x for x in old_cells if x[0]=='A'][0][1:]) if False else ''}"
                  f"→ {len(out)} 列（原 {len(old_cells)} 列）")

        if not apply:
            total += len(bad)
            continue

        # 回写 XML
        def sub_row(m):
            rn = int(m.group(1))
            if rn in new_rows:
                return '<row r="%d"%s>%s</row>' % (rn, m.group(2), new_rows[rn])
            return m.group(0)

        x2 = re.sub(r'<row r="(\d+)"([^>]*)>(.*?)</row>', sub_row, x, flags=re.S)
        assert x2 != x, "替换未生效"

        # 写回 zip（其余部件整段照抄，1.4 秒）
        os.makedirs(BACKUP, exist_ok=True)
        bak = os.path.join(BACKUP, fn)
        if not os.path.exists(bak):
            shutil.copy2(p, bak)
        tmp = p + ".tmp"
        t0 = time.time()
        zin = zipfile.ZipFile(p)
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zo:
            for i in zin.infolist():
                data = x2.encode("utf-8") if i.filename == part else zin.read(i.filename)
                zo.writestr(i, data)
        zin.close()
        try:
            os.replace(tmp, p)
        except OSError:
            with open(p, "wb") as fw, open(tmp, "rb") as fr:
                fw.write(fr.read())
            os.remove(tmp)
        print(f"  ✓ 已落盘（{time.time()-t0:.1f}s，备份 {os.path.basename(bak)}）")
        total += len(bad)

    print(f"\n合计修补 {total} 行" + ("（已落盘）" if apply else "（干跑，未落盘；加 --apply 生效）"))


if __name__ == "__main__":
    main()
