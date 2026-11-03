# -*- coding: utf-8 -*-
"""PDF 货件报告解析（美客多 Inbound 准备单 / eMAG 罗马尼亚送货单）。

背景（2026-10-14 用户新增）：货件目录下出现了两类 PDF，TSV 解析器读不了：
  · 货件/美客多/Inbound-<运单号>-preparation-instructions.pdf
      美客多「入库准备单」，有中/西两种版式；字段 = ML代码 / 通用代码 / SKU：G###，
      数量在右侧「单位 / UNIDADES」列。文字层顺序固定：先列商品块（含 SKU），
      随后是表头 + 数量列（数量有时与说明文字被合并到同一行，如「5 • Envuélvelos…」）。
      故按**页内顺序**取 SKU 与「行首数字」数量，逐页 zip；以首页 stated 自校验。
  · 货件/Emag/reception_document_*.pdf
      罗马尼亚送货单（Aviz de insotire a marfii），表体每行末尾 `PN:G###`，
      数量在 `buc/pcs <数量> <单价> <金额>`。

返回结构与 shipments.read_shipment 对齐：
  dict(id, wh="", name="", ship_md=(月,日)|None, nsku, nqty, items{sku:qty}, names{}, src)

⚠️ 依赖 pypdf（纯 Python）。主环境 C:\\Python314 若未装，需 `pip install pypdf`
   到该解释器的 site-packages（项目 PYTHONPATH 指向的那个）。
"""
import os
import re

# SKU 值：字母开头，可含数字/连字符；文字层会把「G2 79」拆成带空格，故允许内部空格后统一去空格
_SKU_RE = re.compile(r"SKU[：:]\s*([A-Za-z0-9][A-Za-z0-9\-]*(?:[ \t]+[A-Za-z0-9\-]+)*)")
# 数量列在表头「准备说明 / INSTRUCCIONES DE PREPARACIÓN」之后
_HEAD_RE = re.compile(r"(准备说明|INSTRUCCIONES\s+DE\s+PREPARACI[ÓO]N)")
# 数量行：行首数字，其后是行尾或项目符号「•」（后面可能跟说明文字，被合并到同一行）
_QTY_LINE_RE = re.compile(r"^[ \t]*(\d+)[ \t]*(?:•|$)", re.M)


def _pages_text(path):
    from pypdf import PdfReader
    return [pg.extract_text() or "" for pg in PdfReader(path).pages]


def _norm_sku(s):
    return re.sub(r"\s+", "", str(s or "")).strip()


def parse_ml(path):
    """美客多入库准备单 → dict。id=运单号；无发货日期（PDF 里没有）。

    两种版式（中文 / 西班牙文）文字层一致：先列商品块（含 SKU），随后是数量列。
    数量与 SKU 均按「页内顺序」取，逐页 zip；以首页 stated（商品数 | 总数量）自校验。
    """
    pages = _pages_text(path)
    head = pages[0] if pages else ""
    m = re.search(r"(?:运单号|Env[íi]o)\s*#?\s*(\d+)", head)
    sid = m.group(1) if m else os.path.splitext(os.path.basename(path))[0]
    m = re.search(r"(?:商品|Productos del env[íi]o)\s*[:：]\s*(\d+)\s*\|.*?"
                  r"(?:总数量|Total de unidades)\s*[:：]\s*(\d+)", head)
    stated = (int(m.group(1)), int(m.group(2))) if m else (None, None)

    items, names = {}, {}
    for pg in pages:
        skus = [_norm_sku(s) for s in _SKU_RE.findall(pg)]
        skus = [s for s in skus if s]
        hm = _HEAD_RE.search(pg)
        qtys = [int(x) for x in _QTY_LINE_RE.findall(pg[hm.end():])] if hm else []
        if not skus or len(skus) != len(qtys):
            # 页内数量与 SKU 数不等 → 放弃该页配对（保守，避免错配）
            continue
        for s, q in zip(skus, qtys):
            items[s] = items.get(s, 0) + q
    return dict(id=sid, wh="", name="", ship_md=None,
                nsku=len(items), nqty=sum(items.values()),
                items=items, names=names, src=os.path.basename(path),
                stated=stated)


def parse_emag(path):
    """eMAG 罗马尼亚送货单 → dict。id=单号；ship_md=单据日期(月,日)。"""
    all_txt = "\n".join(_pages_text(path))
    m = re.search(r"Numar/Number:\s*(\S+)", all_txt)
    sid = m.group(1) if m else os.path.splitext(os.path.basename(path))[0]
    m = re.search(r"Date \(D,M,Y\):\s*(\d{1,2})-(\d{1,2})-(\d{4})", all_txt)
    ship_md = (int(m.group(2)), int(m.group(1))) if m else None   # (月, 日)
    items = {}
    for mm in re.finditer(r"PN:\s*([A-Z0-9]+)[\s\S]{0,60}?buc/pcs\s+([\d.,]+)", all_txt):
        sku = mm.group(1).strip()
        qty = round(float(mm.group(2).replace(",", "")))
        items[sku] = items.get(sku, 0) + qty
    return dict(id=sid, wh="", name="", ship_md=ship_md,
                nsku=len(items), nqty=sum(items.values()),
                items=items, names={}, src=os.path.basename(path),
                stated=None)


def read_pdf_shipment(path):
    """按文件特征分派：文件名含 Inbound- → 美客多；含 reception_document → eMAG。"""
    b = os.path.basename(path).lower()
    if b.startswith("inbound-"):
        return parse_ml(path)
    if "reception_document" in b:
        return parse_emag(path)
    return None
