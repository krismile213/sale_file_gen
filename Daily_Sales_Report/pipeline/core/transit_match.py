# -*- coding: utf-8 -*-
"""在途对账：用「库存报告」判断货件有没有被亚马逊识别，给出调拨建议。

业务口径（2026-11-06 用户口述 + 当日拍板，本模块唯一依据）：
  · G「在途」= 库存报告 afn-inbound 三段之和 —— **亚马逊口径**，只有「已入网」的货件才有数。
  · H「调拨在途合计」= 手工货件列里没打「到货」的票之和 —— **我方口径**，
    货件一发货（甚至还没被亚马逊识别）就先写进表，好让每日断货测算提前看到补给。
  · 「在途 = 调拨」只在「货件还没被亚马逊识别」时成立。一旦亚马逊识别了，
    那批货进了 G，H 里同票必须打「到货」（或「部分到货」扣数量）退出，否则重复计算。

**2026-11-06 用户拍板的四条口径**（改代码前必读）：
  ① **部分到货**：允许「一票只识别了一部分」。行5 写 `M/D运输部分到货`，该列数量改成
     「票总量 − 已识别件数」。这样 H 里只留尚未识别的部分，不会像整票打到货那样打过头
     （G261 教训：P 票 74 件整票打到货，但亚马逊只认 61 件 → 差额反倒变成 +61）。
  ② **不做「未发」态**：用户明确「未发不用」。未发运的票不特殊标记。
  ③ **日期优先**：判断「该退 H 还是该留」时，**先看发货日期**——
     老票（发得早）优先判定为已到货，新票（最近一两周发的）默认留在 H。
     用户原话：「优先根据日期保证在途的货件正确，未发的那批不强求」。
  ④ **库存 > 在途 = 没更新货件**：若某 SKU 的 G（亚马逊在途）明显大于 H（我方调拨），
     说明这票货已经进亚马逊但货件表没登记 → 这是**漏更新**，要报警提示补货件，
     而不是把差额算成「到货」。

**为什么不逐 SKU 精确匹配**（2026-11-06 全站实测）：
  H 是把各票数量直接相加，没有按 SKU 取小；报告三段在途只覆盖已入网部分，
  两边逐 SKU 根本不可能相等 —— 日本 448 个 SKU 里只有 23 个相等，印度/英国/沙特全是 0。
  逐 SKU 精确匹配在这里是错的方法，不要再用。

判据（从硬到软，日期优先）：
  ① **差额方向**：`G − H` 逐 SKU 累计。
     · G<H = 我方记得多 → 这些票还没被识别，H 该留
     · G>H = 亚马逊有、我方没记 → 这票已经进去了，或货件没登记（见 ④）
  ② **票级差额贡献**：把这票从 H 拿掉，看 Σ|G−H| 是变小还是变大。
     变小 → 这票留着是错的，该退 H。
  ③ **日期闸门**（用户 ③）：只对「发得够早」的票判定退 H。
     最近 NEW_DAYS 天内发的票一律留在 H（哪怕证据指向已到货）——用户每周二/三下新票，
     新票刚写进表、亚马逊还没识别，这时判它「已到货」会把真实的在途补给抹掉。
  ④ 证据不足时报「待人工」，不猜。
"""
import datetime

PARTIAL_MIN = 20       # 差额小于这个数不判，避免把噪声当信号
NEW_DAYS = 10          # 发货日在这之内 = 「新票」，一律留 H（用户口径③）


def _col_sums(ws, sku_rows, col):
    out = {}
    for sku, r in sku_rows.items():
        v = ws.cell(r, col).value
        if isinstance(v, (int, float)) and v:
            out[sku] = int(v)
    return out


def _cols(ws, first_col=None, last_col=None):
    import shipments as Sh
    if first_col is None or last_col is None:
        first_col, last_col = Sh.detect_sheet_cols(ws)
    out = []
    for c in range(first_col, last_col + 1):
        sid = ws.cell(1, c).value
        if sid is None or str(sid).strip() == "":
            continue
        lab = ws.cell(5, c).value
        out.append(dict(col=c, sid=str(sid).strip(), label=lab,
                        arrived=Sh.is_arrived_label(lab),
                        partial=Sh.is_partial_label(lab),
                        unshipped=Sh.is_unshipped_label(lab),
                        md=Sh._label_md(lab)))
    return out


def _age_days(md, ref):
    """行5 的 (月,日) 距参照日多少天。无日期/无法解析返回 None。"""
    if not md or ref is None:
        return None
    for yr in (ref.year, ref.year - 1):
        try:
            d = datetime.date(yr, md[0], md[1])
        except ValueError:
            continue
        if d <= ref:
            return (ref - d).days
    return None


def coverage(ws, inv, sku_rows=None, first_col=None, last_col=None):
    """入网覆盖度 → {col: (has, total)}（只算「还没完全到货」的列）。
    软判据：区分「整票不在报告里」（明显未入网）与「部分在」（在接收中）。"""
    import shipments as Sh
    if sku_rows is None:
        sku_rows = Sh._sku_rows(ws)
    out = {}
    for d in _cols(ws, first_col, last_col):
        if d["arrived"] or d["unshipped"]:
            continue
        skus = list(_col_sums(ws, sku_rows, d["col"]))
        if skus:
            out[d["col"]] = (sum(1 for s in skus if s in inv), len(skus))
    return out


def attribute(ws, inv, sku_rows=None, first_col=None, last_col=None,
              ref_date=None, new_days=NEW_DAYS, min_gap=PARTIAL_MIN):
    """票级差额贡献 + 日期闸门。

    对每票算：Σ|G−H| 在「含这票」与「不含这票」两种情形下的差（delta）。
      · delta < −min_gap → 拿掉后差额变小 → 这票留着是错的 → 候选「退 H」
      · delta > +min_gap → 拿掉后差额变大 → 这票确实还在路上 → 留 H
    再叠加日期闸门：age < new_days 的票强制留 H（reason='new'）。

    返回 dict(per={col: dict(...)}, base=Σ|G−H|, veto=原因|None)
    """
    import shipments as Sh
    if sku_rows is None:
        sku_rows = Sh._sku_rows(ws)
    allc = _cols(ws, first_col, last_col)
    active = [d for d in allc if not d["arrived"] and not d["unshipped"]]
    act_cols = [d["col"] for d in active]
    # 逐票的 SKU 数量表（每列只扫一次）
    mats = {d["col"]: _col_sums(ws, sku_rows, d["col"]) for d in active}

    recs = []
    for sku, ev in inv.items():
        G = int(ev.get("inb") or 0)
        q = {c: mats[c].get(sku, 0) for c in act_cols}
        if G or any(q.values()):
            recs.append((sku, G, q))

    def resid(cols):
        tot = 0
        for _sku, G, q in recs:
            H = sum(q[c] for c in cols)
            tot += abs(G - H)
        return tot

    base = resid(act_cols)
    per = {}
    for d in active:
        c = d["col"]
        qty = sum(q[c] for _s, _g, q in recs)
        delta = resid([x for x in act_cols if x != c]) - base
        age = _age_days(d["md"], ref_date)
        r = dict(col=c, sid=d["sid"], label=d["label"], qty=qty,
                 delta=delta, age=age, md=d["md"])
        if age is not None and age < new_days:
            r["verdict"], r["reason"] = "pending", "new"
        elif delta < -min_gap:
            r["verdict"], r["reason"] = "arrived", "delta"
        elif delta > min_gap:
            r["verdict"], r["reason"] = "pending", "delta"
        else:
            r["verdict"], r["reason"] = "weak", "flat"
        per[c] = r

    # 同批发出的票必须同进同退：同一天里只有部分票判「已到货」→ 整批撤回
    from collections import defaultdict
    byday = defaultdict(list)
    for c, r in per.items():
        if r["md"]:
            byday[r["md"]].append(c)
    bad = []
    for _md, cs in byday.items():
        a = [c for c in cs if per[c]["verdict"] == "arrived"]
        if a and len(a) != len(cs):
            bad.extend(a)
    veto = None
    if bad:
        for c in bad:
            per[c]["verdict"] = "weak"
            per[c]["reason"] = "same_batch"
        veto = ("同一天发出的票判法不一致 → 撤回「已到货」："
                + "、".join(Sh.col_letter(c) for c in sorted(bad)))
    return dict(per=per, base=base, veto=veto)


def judge(cov, attr):
    """综合给每列结论。act ∈ arrived / partial / pending / weak / empty。
    · verdict=arrived 且覆盖率 ≥90% → 整票到货（打「到货」）
    · verdict=arrived 但覆盖率低      → 只退已识别部分（打「部分到货」+ 扣数量）"""
    out = {}
    for c, r in attr["per"].items():
        has, total = cov.get(c, (0, 0))
        share = has / total if total else 0.0
        v = dict(r)
        v.update(has=has, total=total, share=share)
        if total == 0:
            v["act"] = "empty"
        elif r["verdict"] == "arrived":
            v["act"] = "arrived" if share >= 0.9 else "partial"
        elif r["verdict"] == "pending":
            v["act"] = "pending"
        else:
            v["act"] = "weak"
        out[c] = v
    return out


def missing_shipment_alarm(ws, inv, sku_rows=None, first_col=None, last_col=None,
                           min_gap=PARTIAL_MIN):
    """④ 库存 > 在途 → 货件表漏登记告警。

    **只有 `H == 0` 才报「漏更新」**（2026-11-06 收紧）：
      · H == 0 且 G 明显 > 0 → 亚马逊那边有一票在途、我方货件列里**一条记录都没有**
        → 这票的货件报告确实没放进 `货件/` 目录（用户没更新货件）。**这才是真漏更新。**
      · H > 0 但明显小于 G → 货件表里有这票，只是**数量记录偏少**（可能是分批发的、
        或只登记了部分，也可能销量已消耗）。这是「数量口径差」，不是「漏货件」，
        混在一起报会把上面那类真漏更新淹没（实测英国表混报 253 个 SKU、合计 3.6 万件）。

    返回 [(sku, G, H, gap), ...]，按 gap 降序。"""
    import shipments as Sh
    if sku_rows is None:
        sku_rows = Sh._sku_rows(ws)
    allc = _cols(ws, first_col, last_col)
    active = [d["col"] for d in allc if not d["arrived"] and not d["unshipped"]]
    mats = {c: _col_sums(ws, sku_rows, c) for c in active}
    out = []
    for sku, ev in inv.items():
        G = int(ev.get("inb") or 0)
        if G <= min_gap:
            continue
        H = sum(mats[c].get(sku, 0) for c in active)
        if H != 0:
            continue                        # 有货件记录 → 不是漏货件（见 docstring）
        out.append((sku, G, H, G))
    out.sort(key=lambda x: -x[3])
    return out


def report_skus_present(ws, inv, col, sku_rows=None):
    """某列有量 SKU 中「报告里有记录」的列表（给人工核对用的明细）。"""
    import shipments as Sh
    if sku_rows is None:
        sku_rows = Sh._sku_rows(ws)
    return sorted(s for s in _col_sums(ws, sku_rows, col) if s in inv)
