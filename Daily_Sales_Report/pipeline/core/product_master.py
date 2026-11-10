# -*- coding: utf-8 -*-
"""product_master.py —— 商品资料（属性主数据）读取与映射

数据源：`sale_file_gen/商品资料.xlsx` → sheet「商品资料」（r2 表头，r3 起 652 行）

用户口径（2026-11-07，写在资料 r1 里，原文照录）：
    A: 1 个月销量 500 套及以上
    B: 3 个月销量 500 套及以上
    C: 6 个月销量 250 套及以上
    D: 6 个月销量 200-250
    E: 6 个月销量 200 套以下
    X 类为新品：XA 日均≥16.7 / XB 16.7>日均≥5.5 / XC 5.5>日均≥1.38
              XD 1.38>日均≥1.11 / XE 1.11>日均 / X 暂无销售数据
    评级为 E 或 XE → **不要采购**
    停售标准：钢化膜、水凝膜连续 12 个月评级 E → 停售；平板钢化膜一年总销量 < 250 → 停售
    手机壳产品**不参与**以上产品评级规则

对外接口：
    load()                      → {SKU: 属性dict}
    attrs(sku)                  → 属性dict 或 None
    is_stop_sale(sku)           → True/False/None（资料里查不到 → None）
    product_type(sku)           → '钢化膜'/'水凝膜'/'手机壳'/... 或 None
    should_buy(sku)             → False 表示评级 E/XE，建议不采购；None = 无评级
    coverage(skus)              → (命中数, 缺失列表)
"""
import os
import re
from openpyxl import load_workbook

_CACHE = {"mtime": None, "data": None, "path": None}

# 国家/站点后缀：G256IT → G256、G332I → G332、G433J → G433、G535A → G535、G261UK → G261
# 只认 1~2 个大写字母尾巴，且前面必须是数字或字母（不能把 K012 拆坏）。
_SUFFIX_RE = re.compile(r"(?<=[0-9A-Za-z])[A-Z]{1,2}$")


def _candidate_paths():
    """商品资料.xlsx 的可能位置（按优先级）"""
    here = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    # here = ...\Daily_Sales_Report ; 资料在上级 sale_file_gen\
    return [
        os.path.join(os.path.dirname(here), "商品资料.xlsx"),
        os.path.join(here, "商品资料.xlsx"),
        r"C:\Users\dell\Desktop\work-main\sale_file_gen\商品资料.xlsx",
    ]


def find_path():
    for p in _candidate_paths():
        if os.path.isfile(p):
            return p
    return None


def load(force=False):
    """读商品资料 → {SKU: dict}。按 mtime 缓存，改过文件自动重读。"""
    p = find_path()
    if not p:
        return {}
    mt = os.path.getmtime(p)
    if not force and _CACHE["data"] is not None and _CACHE["mtime"] == mt:
        return _CACHE["data"]
    wb = load_workbook(p, read_only=True, data_only=True)
    ws = wb["商品资料"]
    # r2 表头定位
    hdr = [c for c in next(ws.iter_rows(min_row=2, max_row=2, max_col=30, values_only=True))]
    idx = {}
    for j, h in enumerate(hdr):
        if not isinstance(h, str):
            continue
        hs = h.strip()
        if hs == "SKU" and "sku" not in idx:
            idx["sku"] = j
        elif hs == "ASIN":
            idx["asin"] = j
        elif "参考成本价" in hs:
            idx["cost"] = j
        elif hs == "手机价格":
            idx["phone_price"] = j
        elif "销量得分" in hs:
            idx["score"] = j
        elif hs == "类型":
            idx["type"] = j
        elif hs == "中文品名":
            idx["name_cn"] = j
        elif hs == "英文品名":
            idx["name_en"] = j
        elif hs == "品牌":
            idx["brand"] = j
        elif hs == "系列":
            idx["series"] = j
        elif hs == "供应商":
            idx["supplier"] = j
        elif hs == "型号":
            idx["model"] = j
        elif hs == "状态":
            idx["status"] = j
        elif hs == "镜头膜":
            idx["lens"] = j
        elif hs == "EAN码":
            idx["ean"] = j
    data = {}
    for r in ws.iter_rows(min_row=3, max_row=2000, max_col=30, values_only=True):
        s = r[idx["sku"]] if idx.get("sku") is not None and idx["sku"] < len(r) else None
        if not isinstance(s, str) or not s.strip():
            continue
        sku = s.strip()
        rec = {}
        for k, j in idx.items():
            if k == "sku":
                continue
            rec[k] = r[j] if j < len(r) else None
        data[sku] = rec
    wb.close()
    # 国家后缀索引：资料里是裸 SKU（G256），日销表里常带站点后缀（G256IT / G332I / G433J）。
    # 实测 6 个「未收录」里有 5 个属于这种情况，不建索引它们会整批漏掉属性
    # （其中 G256=停售，漏掉就会去补一个已停售的品）。
    _CACHE.update(mtime=mt, data=data, path=p)
    return data


def _bare_index(master=None):
    """裸 SKU → 带后缀的全名列表。资料里若同时有 G261 和 G261IT 不会互串。"""
    m = master if master is not None else load()
    idx = {}
    for s in m:
        head = _SUFFIX_RE.sub("", s)
        if head != s:
            idx.setdefault(head, []).append(s)
    return idx


def _suffix_candidates(s):
    """G256IT / G332I / G433J / G535A / G261IT → [G256, G332, G433, G535, G261]

    规则：去掉末尾 1~2 个大写字母（站点后缀 IT/I/J/A/UK/DE…）。
    ⚠️ 只去掉「纯字母尾巴」，不动 K012-016 这种带数字的后缀。
    """
    out = []
    t = s
    for n in (2, 1):
        if len(t) > n + 2 and t[-n:].isalpha() and t[-n:].isupper():
            out.append(t[:-n])
    return out


def attrs(sku, master=None):
    """取单个 SKU 的属性；查不到返回 None

    匹配顺序：①精确同名 ②去国家后缀后同名（G256IT → G256）。
    第②步是必要的：资料里存裸 SKU，部分日销表用带站点后缀的 SKU。
    若去掉后缀后命中多个（资料里同时有 G261 和 G261IT），保留为精确匹配，不猜。
    """
    if not sku:
        return None
    m = master if master is not None else load()
    s = str(sku).strip()
    if s in m:
        return m[s]
    for cand in _suffix_candidates(s):
        if cand in m:
            hit = dict(m[cand])
            hit["_matched"] = cand        # 记录是靠去后缀命中的，便于报告里标注
            return hit
    return None


def is_stop_sale(sku, master=None):
    """状态是否停售。'停售'/'不销售' → True；'在售' → False；查不到 → None"""
    a = attrs(sku, master)
    if not a:
        return None
    st = (a.get("status") or "").strip()
    if st in ("停售", "不销售"):
        return True
    if st == "在售":
        return False
    return None


def product_type(sku, master=None):
    """类型：钢化膜/水凝膜/手机壳/平板钢化膜/3D曲面/3D热弯/镜头膜"""
    a = attrs(sku, master)
    if not a:
        return None
    return (a.get("type") or "").strip() or None


def should_buy(sku, master=None):
    """评级 E / XE → False（用户口径：不要采购）。无评级 → None

    ⚠️ 手机壳不参与评级规则（用户原文：「手机壳产品不参与以上产品评级规则！」）：
    资料里手机壳那一行即使残留 E/XE 字样，也照样可采购 → 返回 None，不做压制。
    """
    a = attrs(sku, master)
    if not a:
        return None
    if (a.get("type") or "").strip() == "手机壳":
        return None                     # 不参与评级 → 不限制
    sc = (str(a.get("score") or "")).strip().upper()
    if not sc:
        return None
    return sc not in ("E", "XE")


def coverage(skus):
    """→ (命中数, 缺失 SKU 列表)"""
    m = load()
    hit, miss = 0, []
    for s in skus:
        if s in m:
            hit += 1
        else:
            miss.append(s)
    return hit, miss


if __name__ == "__main__":
    m = load()
    print(f"商品资料: {find_path()}")
    print(f"SKU 数: {len(m)}")
    from collections import Counter
    print("类型:", dict(Counter((v.get("type") or "").strip() for v in m.values()).most_common()))
    print("状态:", dict(Counter((v.get("status") or "").strip() for v in m.values()).most_common()))
    print("得分:", dict(Counter((str(v.get("score") or "")).strip() for v in m.values()).most_common()))
    stop = sum(1 for s in m if is_stop_sale(s, m))
    nobuy = sum(1 for s in m if should_buy(s, m) is False)
    print(f"停售/不销售: {stop}   评级 E/XE(不采购): {nobuy}")
    for s in ("G547", "G261IT", "K001", "G535A"):
        a = attrs(s, m)
        print(f"  {s}: type={product_type(s,m)} status={(a or {}).get('status')} "
              f"score={(a or {}).get('score')} should_buy={should_buy(s,m)}")
