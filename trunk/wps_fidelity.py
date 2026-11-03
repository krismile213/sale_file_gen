# -*- coding: utf-8 -*-
"""
WPS / Excel 观感保真注入模块（欧洲版、北美版共用）

openpyxl 保存 xlsx 之后调用 apply_fidelity()，把东南亚模板里那些"办公观感"部件补回去：

  1. x14 迷你图         openpyxl 完全不支持 sparkline，读进来就会丢 → 只能在保存后手写进 sheet XML 的 <extLst>
  2. WPS 专有部件        customXml/pixelators + docProps/custom.xml（KSOReadingLayout=true → WPS 阅读模式黄十字）
  3. 文档缩略图          不搬、并清掉残留。东南亚那份缩略图画的是东南亚表，搬进欧洲/北美文件属于"配错图"；
                         清掉之后 WPS 打开时若需要会自己生成（实测：欧洲那份被 WPS 保存后产生了
                         docProps/thumbnail.wmf）。
  4. sheetView 修正      清掉与冻结窗格冲突的 @topLeftCell
                         ← 这是本次"横滚后看不到 SKU/型号"的根因：sheetView/@topLeftCell 与 <pane> 打架，
                           WPS 里冻结窗格直接失效，视图停在 AL 列又没有 A/B 列固定

用法
----
    from wps_fidelity import apply_fidelity
    apply_fidelity(result_path, spark_specs=[
        dict(sheet="2025-2026年销量情况 Sales", col="Z", r1=3, r2=12, c1="B", c2="Y"),
    ])

spark_specs 每项：sheet 名、迷你图所在列、数据行范围、数据列范围（起始/结束列字母）。
"""
import os
import re
import shutil
import zipfile

# 东南亚模板（"母本"）：观感部件的来源
_SE_TEMPLATE = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "data", "数据归属周期", "202609", "快照日期", "20261006",
    "东南亚", "2026年东南亚销量表Southeast Asia Sales table.xlsx",
)

# 需要按字节复制的部件（源路径 → 目标路径）
_BINARY_PARTS = [
    ("customXml/item1.xml", "customXml/item1.xml"),
    ("customXml/item2.xml", "customXml/item2.xml"),
    ("customXml/itemProps1.xml", "customXml/itemProps1.xml"),
    ("customXml/itemProps2.xml", "customXml/itemProps2.xml"),
    ("customXml/_rels/item1.xml.rels", "customXml/_rels/item1.xml.rels"),
    ("customXml/_rels/item2.xml.rels", "customXml/_rels/item2.xml.rels"),
    ("docProps/custom.xml", "docProps/custom.xml"),          # KSOReadingLayout=true → WPS 阅读模式
]

_CT_DEFAULTS = [
    ('Extension="bin"', 'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings"'),
    ('Extension="vml"', 'ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"'),
]
_CT_OVERRIDES = [
    ('/customXml/itemProps1.xml', 'application/vnd.openxmlformats-officedocument.customXmlProperties+xml'),
    ('/customXml/itemProps2.xml', 'application/vnd.openxmlformats-officedocument.customXmlProperties+xml'),
    ('/docProps/custom.xml', 'application/vnd.openxmlformats-officedocument.custom-properties+xml'),
]

_SPARK_URI = "{05C60535-1F16-4fd2-B633-F4F36F0B64E0}"
_X14_NS = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
_XM_NS = "http://schemas.microsoft.com/office/excel/2006/main"

# 迷你图配色（照抄东南亚模板）
_SPARK_COLORS = [
    ("colorSeries", "FF376092"), ("colorNegative", "FFD00000"), ("colorAxis", "FF000000"),
    ("colorMarkers", "FFD00000"), ("colorFirst", "FFD00000"), ("colorLast", "FFD00000"),
    ("colorHigh", "FFD00000"), ("colorLow", "FFD00000"),
]


def _q(sheet_name):
    """工作表名放进公式要加单引号（名里有空格/中文时）"""
    return "'" + sheet_name.replace("'", "''") + "'"


def _spark_ext(spec):
    """生成一个 sheet 的 <extLst>（只含迷你图那一段）"""
    name, col, r1, r2 = spec["sheet"], spec["col"], spec["r1"], spec["r2"]
    c1, c2 = spec["c1"], spec["c2"]
    lines = [
        f'<extLst><ext uri="{_SPARK_URI}" xmlns:x14="{_X14_NS}">',
        f'<x14:sparklineGroups xmlns:xm="{_XM_NS}">',
        '<x14:sparklineGroup displayEmptyCellsAs="gap">',
    ]
    lines += [f'<x14:{tag} rgb="{rgb}"/>' for tag, rgb in _SPARK_COLORS]
    lines.append("<x14:sparklines>")
    for r in range(r1, r2 + 1):
        lines.append(
            f"<x14:sparkline><xm:f>{_q(name)}!${c1}${r}:${c2}${r}</xm:f>"
            f"<xm:sqref>{col}{r}</xm:sqref></x14:sparkline>"
        )
    lines += ["</x14:sparklines>", "</x14:sparklineGroup>", "</x14:sparklineGroups>", "</ext></extLst>"]
    return "".join(lines)


def _fix_sheet_view(xml):
    """修 sheetView：清掉与冻结窗格冲突的 topLeftCell，并把三个 selection 收敛到合法位置。

    1) `<sheetView topLeftCell="AL1">` 与 `<pane state="frozen">` 同时出现时，WPS 会让冻结窗格失效，
       表现就是"横滚到 8 月块后 SKU / 产品型号两列不见了" → 必须删掉。
    2) openpyxl 生成的空视图里三个 selection 都写着 activeCell="A1"，而 A1 属于左上冻结区、
       不属于 topRight / bottomRight 面板 —— WPS 会据此微调视图。按东南亚模板的写法收敛：
         topRight    → 冻结列右侧第一列 + 第 1 行（表头）
         bottomLeft  → A 列 + 第 3 行（第一条数据行）
         bottomRight → 冻结列右侧第一列 + 第 3 行
    """
    xml = re.sub(r'(<sheetView\b[^>]*?)\s+topLeftCell="[^"]*"', r"\1", xml)

    m = re.search(r"<sheetView\b([^>]*)>(.*?)</sheetView>", xml, re.S)
    if not m:
        return xml
    attrs, body = m.group(1), m.group(2)
    pane = re.search(r'<pane\b[^>]*?xSplit="(\d+)"[^>]*?ySplit="(\d+)"', body)
    if not pane:
        return xml
    from openpyxl.utils import get_column_letter as _gl
    rx = int(pane.group(1)) + 1          # 冻结列右侧第一列
    ry = int(pane.group(2)) + 1          # 冻结行下方第一行
    c = _gl(rx)
    fixed = (
        f'<selection pane="topRight" activeCell="{c}1" sqref="{c}1"/>'
        f'<selection pane="bottomLeft" activeCell="A{ry}" sqref="A{ry}"/>'
        f'<selection pane="bottomRight" activeCell="{c}{ry}" sqref="{c}{ry}"/>'
    )
    body = re.sub(r"<selection\b[^>]*/>", "", body)
    return xml[:m.start()] + f"<sheetView{attrs}>" + body + fixed + "</sheetView>" + xml[m.end():]


def _add_sparkline_to_sheet(xml, spec):
    """把迷你图 extLst 追加到 worksheet 末尾（必须在 </worksheet> 之前，且是整个文档的最后一个子元素）"""
    ext = _spark_ext(spec)
    m = re.search(r"<extLst>.*?</extLst>\s*$", xml, re.S)
    if m:                                    # 已有 extLst → 把迷你图并进去（放最后）
        inner = ext[len("<extLst>"):-len("</extLst>")]
        return xml[:m.end() - len("</extLst>")] + inner + "</extLst>"
    return xml.replace("</worksheet>", ext + "</worksheet>")


def _patch_content_types(b):
    s = b.decode("utf-8")
    for ext_key, ctype in _CT_DEFAULTS:
        if ext_key not in s:
            s = s.replace('<Default Extension="rels"',
                          f'<Default {ext_key} {ctype}/><Default Extension="rels"', 1)
    for part, ctype in _CT_OVERRIDES:
        if f'PartName="{part}"' not in s:
            s = s.replace("</Types>", f'<Override PartName="{part}" ContentType="{ctype}"/></Types>')
    return s.encode("utf-8")


def _patch_pkg_rels(b):
    s = b.decode("utf-8")
    if "docProps/custom.xml" not in s:
        s = s.replace("</Relationships>",
                      '<Relationship Id="rIdCustom" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
                      'relationships/custom-properties" Target="docProps/custom.xml"/></Relationships>')
    return s.encode("utf-8")


def _strip_thumbnail(data, order):
    """清掉缩略图部件及其关系/内容类型声明（就地改 data / order）

    为什么清：东南亚模板那份 thumbnail.emf 画的是东南亚表，搬进来是"配错图"；
    WPS 打开保存后也会自己生成缩略图（实测生成 .wmf），所以留着只会误导。
    """
    killed = [n for n in list(data) if n.lower().startswith("docprops/thumbnail")]
    for n in killed:
        data.pop(n, None)
    order[:] = [n for n in order if not n.lower().startswith("docprops/thumbnail")]

    rels = data.get("_rels/.rels")
    if rels:
        s = rels.decode("utf-8")
        s2 = re.sub(r'<Relationship\b[^>]*Type="[^"]*/thumbnail"[^>]*/>', "", s)
        if s2 != s:
            data["_rels/.rels"] = s2.encode("utf-8")

    ct = data.get("[Content_Types].xml")
    if ct:
        s = ct.decode("utf-8")
        s2 = re.sub(r'<Default\s+Extension="(?:emf|wmf)"[^>]*/>', "", s)
        if s2 != s:
            data["[Content_Types].xml"] = s2.encode("utf-8")
    return killed


def apply_fidelity(path, spark_specs=(), src_template=None):
    """给已保存的 xlsx 注入 WPS 观感部件 + 迷你图 + 修 sheetView

    参数
    ----
    path          : 目标 xlsx（就地改写）
    spark_specs   : [dict(sheet=表名, col='Z', r1=3, r2=12, c1='B', c2='Y'), ...]
    src_template  : 观感部件来源（默认东南亚模板）
    """
    src = src_template or _SE_TEMPLATE
    if not os.path.exists(src):
        print(f"⚠️  保真源模板不存在，跳过观感注入：{src}")
        return False

    with zipfile.ZipFile(src) as z:
        parts = {}
        for sp, dp in _BINARY_PARTS:
            try:
                parts[dp] = z.read(sp)
            except KeyError:
                pass

    with zipfile.ZipFile(path) as z:
        infos = z.infolist()
        data = {i.filename: z.read(i.filename) for i in infos}
        order = [i.filename for i in infos]

    # 1) 定位工作表 XML：workbook.xml 里 sheet 的声明顺序 = sheetN.xml
    wb_xml = data["xl/workbook.xml"].decode("utf-8")
    sheet_names = re.findall(r'<sheet name="([^"]+)"', wb_xml)

    # 2) 修 sheetView + 注入迷你图
    specs = {s["sheet"]: s for s in spark_specs}
    for idx, nm in enumerate(sheet_names, 1):
        key = f"xl/worksheets/sheet{idx}.xml"
        if key not in data:
            continue
        xml = data[key].decode("utf-8")
        xml = _fix_sheet_view(xml)
        if nm in specs:
            xml = _add_sparkline_to_sheet(xml, specs[nm])
        data[key] = xml.encode("utf-8")

    # 3) 包级关系 + 内容类型
    data["[Content_Types].xml"] = _patch_content_types(data["[Content_Types].xml"])
    if "_rels/.rels" in data:
        data["_rels/.rels"] = _patch_pkg_rels(data["_rels/.rels"])

    # 3.5) 清掉缩略图残留（东南亚那张缩略图画的是东南亚表，不搬）
    _strip_thumbnail(data, order)

    # 4) 逐字节复制观感部件
    for name, blob in parts.items():
        if name not in data:
            order.append(name)
        data[name] = blob

    # 5) 回写（[Content_Types].xml 放最前，符合惯例）
    order = ["[Content_Types].xml"] + [n for n in order if n != "[Content_Types].xml"]
    tmp = path + ".fi.tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for name in order:
            zout.writestr(name, data[name])
    shutil.move(tmp, path)
    return True


def check_fidelity(path):
    """自检：返回 (结论列表, 是否全通过)"""
    out, ok = [], True
    with zipfile.ZipFile(path) as z:
        names = set(z.namelist())
        wb_xml = z.read("xl/workbook.xml").decode("utf-8")
        sheet_names = re.findall(r'<sheet name="([^"]+)"', wb_xml)

        n_spark = 0
        n_bad_view = 0
        n_pane = 0
        for idx, nm in enumerate(sheet_names, 1):
            key = f"xl/worksheets/sheet{idx}.xml"
            if key not in names:
                continue
            x = z.read(key).decode("utf-8")
            if "x14:sparklineGroups" in x:
                n_spark += 1
            if re.search(r'<sheetView\b[^>]*\btopLeftCell=', x):
                n_bad_view += 1
            if '<pane ' in x and 'state="frozen"' in x:
                n_pane += 1
        test = [
            ("WPS customXml(pixelators)", "customXml/item1.xml" in names and "customXml/item2.xml" in names),
            ("docProps/custom.xml(阅读模式)", "docProps/custom.xml" in names),
            ("无缩略图残留", not any(n.lower().startswith("docprops/thumbnail") for n in names)),
            ("customXml 关系已声明", "docProps/custom.xml" in z.read("_rels/.rels").decode("utf-8")),
            (f"迷你图 {n_spark} 张表", n_spark > 0),
            (f"冻结窗格 {n_pane} 张表", n_pane > 0),
            (f"sheetView 冲突残留 {n_bad_view} 处", n_bad_view == 0),
        ]
        for k, v in test:
            out.append(("✅" if v else "❌") + " " + k)
            ok = ok and bool(v)
    return out, ok


if __name__ == "__main__":
    import sys
    for p in sys.argv[1:]:
        print("=" * 70)
        print(p)
        for line in check_fidelity(p)[0]:
            print("  ", line)
