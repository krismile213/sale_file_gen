# -*- coding: utf-8 -*-
"""在途对账 + 调拨建议（只读报告，默认不落盘）。

用法：
  python pipeline/tools/transit_audit.py                       # 全簿预演（打印建议）
  python pipeline/tools/transit_audit.py --book 日本            # 只看一本（模糊匹配）
  python pipeline/tools/transit_audit.py --detail               # 附带证据明细
  python pipeline/tools/transit_audit.py --alarm                # 只看「库存>在途=货件没更新」
  python pipeline/tools/transit_audit.py --recon                # 逐站 G（亚马逊）vs H（我方）对账表
  python pipeline/tools/transit_audit.py --apply --backup       # 落盘
  python pipeline/tools/transit_audit.py --apply --prune        # 落盘 + 清掉已到货列
  环境变量 TRANSIT_SNAP / TRANSIT_PERIOD 可指定快照日与周期。

**按 sheet 收窄（2026-11-06）**：`inv_for()` 原先合并整簿所有库存报告，多站点簿会串味
（欧盟英国簿把 40895 件套到英国表上 → 英国差异率虚报 +79.6%，真实 52.9%）。
现在 `inv_map_for_book()` 让每个分析表只吃自己站点的 kind="inv" 报告，
「分析表名 → 库存 sheet 名」靠 `config.stock_sheet_for()`。

落盘做什么（2026-11-06 用户拍板口径）：
  · act=arrived → 行5 追加「到货+打标日」（如 `9/22海运到货11/6`），整列从 H 剔除；
    **保留 5 天才清**（KEEP_ARRIVED_DAYS，用户「到货的保留5天吧」）
  · act=partial → 行5 追加「部分到货」，该列数量扣成「总量 × (1 − 覆盖率)」
  · **--prune**：把「已到货且已过保留期」的列货件号/标签/数量一起清空，只留未到货的票
    （用户要求「文件太多了，只保留没到货的」）。默认不开启，需显式加 --prune。
    保留期内的到货列会打印 `[留] … 还剩 N 天`。
撤销：把行5 里的「到货/部分到货」删掉，再跑一次日销批（货件数量会自动回填）。
"""
import os
import sys
import glob
import shutil
import datetime
import csv
import io

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))       # Daily_Sales_Report
sys.path.insert(0, os.path.join(ROOT, "pipeline", "core"))

import shipments as Sh          # noqa: E402
import transit_match as TM      # noqa: E402
import config as C              # noqa: E402

from openpyxl import load_workbook   # noqa: E402


def _dec(p):
    raw = open(p, "rb").read()
    for enc in ("utf-8-sig", "utf-8", "gb18030", "shift_jis"):
        try:
            return raw.decode(enc)
        except Exception:
            continue
    return raw.decode("utf-8", "replace")


def read_inv_report(path, jp=False):
    """库存报告 → {sku: {avail, inb}}。日文 24 列按 config.JP24 的位置映射。"""
    txt = _dec(path).replace("\r\n", "\n").replace("\r", "\n")
    rows = list(csv.reader(io.StringIO(txt)))
    if not rows:
        return {}
    if jp:
        h = {k: i for i, k in enumerate(C.JP24)}
    else:
        hdr = [x.strip().lower() for x in rows[0]]
        h = {k: i for i, k in enumerate(hdr)}
        if "sku" not in h:
            for cand in ("seller-sku", "msku"):
                if cand in h:
                    h["sku"] = h[cand]
                    break
    cs, ca = h.get("sku"), h.get("afn-fulfillable-quantity")
    ci = [h.get(k) for k in ("afn-inbound-working-quantity",
                             "afn-inbound-shipped-quantity",
                             "afn-inbound-receiving-quantity")]
    if cs is None or ca is None:
        return {}
    out = {}
    for r in rows[1:]:
        need = [cs, ca] + [x for x in ci if x is not None]
        if len(r) <= max(need):
            continue
        sku = r[cs].strip()
        if not Sh.SKU_RE.match(sku):
            continue
        try:
            out[sku] = dict(avail=int(float(r[ca] or 0)),
                            inb=sum(int(float(r[x] or 0)) for x in ci if x is not None))
        except ValueError:
            continue
    return out


def _resolve(snap_dir, cands, fmt):
    for t in cands:
        p = os.path.join(snap_dir, t.format(**fmt))
        if os.path.exists(p):
            return p
    return None


def inv_for(wb_cfg, snap_dir, fmt, sheet=None):
    """本簿库存报告 → {sku: {avail, inb}}，同一份报告只算一次。

    **按 sheet 收窄**（2026-11-06）：传了 `sheet`（库存 sheet 名，来自
    `config.stock_sheet_for(分析表名)`）就**只吃这一条** kind="inv" 的 stocks。
    不传则退回旧的整簿合并口径（会给多站点簿串味，仅为兼容旧调用保留）。

    为什么必须收窄：欧盟英国簿把欧盟 40895 件套到英国表头上 → 英国差异率
    从真实 52.9% 虚报成 +79.6%；北美簿墨西哥从 99.9% 虚报成 +89.6%。
    · kind="balance" 是同一份报告的「可售库存」副本 → 永远排除（否则在途数两遍）
    · kind="multi"（多国库存）不是 inv 报告 → 永远排除
    """
    cands = []
    for s in wb_cfg.get("stocks", []):
        if s.get("kind") != "inv":
            continue
        if sheet is not None and str(s.get("sheet", "")).strip() != str(sheet).strip():
            continue
        cands.append(s)
    merged, seen = {}, set()
    for s in cands:
        tmpls = s["report"] if isinstance(s["report"], list) else [s["report"]]
        p = _resolve(snap_dir, tmpls, fmt)
        if not p or p in seen:
            continue
        seen.add(p)
        jp = s.get("colmap") == "JP24" or "日本" in os.path.basename(p)
        for sku, ev in read_inv_report(p, jp).items():
            cur = merged.setdefault(sku, dict(avail=0, inb=0))
            cur["avail"] += ev["avail"]
            cur["inb"] += ev["inb"]
    return merged


def inv_map_for_book(wb_cfg, snap_dir, fmt):
    """整簿 → {库存 sheet 名: inv 字典}，只收 kind="inv"（逐 sheet 各配自己的报告）。

    没有对应库存报告的（美客多 / eMAG / 罗马尼亚）返回空 dict。
    """
    out = {}
    for s in wb_cfg.get("stocks", []):
        if s.get("kind") != "inv":
            continue
        sh = str(s.get("sheet", "")).strip()
        if not sh or sh in out:
            continue
        d = inv_for(wb_cfg, snap_dir, fmt, sheet=sh)
        if d:
            out[sh] = d
    return out


def audit_book(path, inv, apply=False, detail=False, prune=False, ref_date=None,
               alarm_only=False):
    """`inv` 可以是：
    · {sku: {avail,inb}}            → 整簿共用（旧口径，会串味，仅兼容）
    · {库存sheet名: inv字典}        → 逐 sheet 取自己站点（推荐，见 inv_map_for_book）
    """
    wb = load_workbook(path)
    out = []
    alarms = []
    per_sheet = None
    if inv and all(isinstance(v, dict) and "inb" not in v for v in inv.values()):
        per_sheet = inv
    try:
        for sn in wb.sheetnames:
            if "分析表" not in sn:
                continue
            ws = wb[sn]
            inv_s = inv
            if per_sheet is not None:
                key = C.stock_sheet_for(sn)
                inv_s = per_sheet.get(str(key).strip()) if key else None
                if not inv_s:
                    # 这本簿没有该站点的在途库存报告（美客多 Full / eMAG）→ 跳过
                    continue
            sku_rows = Sh._sku_rows(ws)
            fc, lc = Sh.detect_sheet_cols(ws)
            if not any(ws.cell(1, c).value for c in range(fc, lc + 1)):
                continue
            cov = TM.coverage(ws, inv_s, sku_rows, fc, lc)
            attr = TM.attribute(ws, inv_s, sku_rows, fc, lc, ref_date=ref_date)
            dec = TM.judge(cov, attr)
            mis = TM.missing_shipment_alarm(ws, inv_s, sku_rows, fc, lc)
            if mis:
                alarms.append((sn.strip(), mis))
            for c, d in dec.items():
                d["label"] = ws.cell(5, c).value
                d["sid"] = str(ws.cell(1, c).value or "").strip()
                d["col"] = Sh.col_letter(c)
                if detail:
                    d["present"] = TM.report_skus_present(ws, inv_s, c, sku_rows)
            out.append((sn, ws, dec, attr.get("veto")))
        if apply:
            n_lab = n_part = n_prune = n_keep = 0
            for sn, ws, dec, _veto in out:
                for c, d in dec.items():
                    if d["act"] not in ("arrived", "partial"):
                        continue
                    flag = "arrived" if d["act"] == "arrived" else "partial"
                    old, new = Sh.flag_label(ws, c, flag, on_date=ref_date)
                    if old != new:
                        n_lab += 1
                        print(f"    [改] {sn.strip()} {d['col']} {d['sid']} 行5：{old!r} → {new!r}")
                    if d["act"] == "partial":
                        cut = _apply_partial(ws, c, d)
                        if cut:
                            n_part += 1
                            print(f"        部分到货：数量 {cut[0]} → {cut[1]}（已识别约 {cut[2]} 件）")
                if prune:
                    n, kept = _prune_arrived(ws, ref_date=ref_date)
                    n_prune += n
                    n_keep += len(kept)
                    for col, sid, lab, left in kept:
                        print(f"    [留] {sn.strip()} {col} {sid} {str(lab)[:18]}"
                              f" —— 到货未满 {Sh.KEEP_ARRIVED_DAYS} 天"
                              + (f"，还剩 {left} 天" if left is not None else ""))
            if n_lab or n_part or n_prune:
                wb.calculation.fullCalcOnLoad = True
                wb.save(path)
            print(f"    落盘：改标签 {n_lab} 格，部分到货扣量 {n_part} 票"
                  + (f"，清理已到货列 {n_prune} 个（保留期内 {n_keep} 个）" if prune else ""))
    finally:
        wb.close()
    return out, alarms


def _apply_partial(ws, col, d):
    """把该列数量按「未识别比例」等比缩减。返回 (原总量, 新总量, 已识别件数)。"""
    total = d.get("qty") or 0
    share = d.get("share") or 0.0
    if not total:
        return None
    recognised = int(round(total * share))
    if recognised <= 0:
        return None
    remaining = max(total - recognised, 0)
    ratio = remaining / total if total else 0
    for r in range(Sh.DATA_FIRST_ROW, (ws.max_row or Sh.DATA_FIRST_ROW) + 1):
        v = ws.cell(r, col).value
        if isinstance(v, (int, float)) and v:
            ws.cell(r, col).value = int(round(v * ratio))
    return total, remaining, recognised


def _prune_arrived(ws, ref_date=None, keep_days=Sh.KEEP_ARRIVED_DAYS):
    """清空「已到货」列的货件号/标签/数量 —— 只保留未到货的票。

    **保留期**（2026-11-06 用户：「到货的保留5天吧」）：打完「到货」标签后列还留
    `keep_days` 天，超期才清。到期判定见 `shipments.arrived_expired`
    （优先按行5 里的打标日期 `到货11/6` 算，没记就退化成「发货日 + 5 天」）。

    返回 (清理数, 保留数, [(列字母, 货件号, 标签, 剩余天数)])
    ⚠️ 行2/3/4 是公式（从行5 解析），清掉行5 后它们会自然变空，不需要动。"""
    fc, lc = Sh.detect_sheet_cols(ws)
    n, kept = 0, []
    for c in range(fc, lc + 1):
        lab = ws.cell(5, c).value
        if not Sh.is_arrived_label(lab):
            continue
        if not ws.cell(1, c).value:
            continue
        if ref_date is not None and not Sh.arrived_expired(lab, ref_date, keep_days):
            md = Sh.arrive_md(lab) or Sh._label_md(lab)
            if md:
                # 剩余天数 = keep_days − 已过天数
                from datetime import date as _date
                for yr in (ref_date.year, ref_date.year - 1):
                    try:
                        d = _date(yr, md[0], md[1])
                    except ValueError:
                        continue
                    if d <= ref_date:
                        left = keep_days - (ref_date - d).days
                        break
                else:
                    left = None
            else:
                left = None
            kept.append((Sh.col_letter(c), str(ws.cell(1, c).value).strip(), lab, left))
            continue
        ws.cell(1, c).value = None
        base = Sh.label_base(lab)
        ws.cell(5, c).value = base if base else None
        for r in range(Sh.DATA_FIRST_ROW, (ws.max_row or Sh.DATA_FIRST_ROW) + 1):
            v = ws.cell(r, c).value
            if isinstance(v, (int, float)) and v:
                ws.cell(r, c).value = None
        n += 1
    return n, kept


def _latest_snap():
    base = os.path.join(os.path.dirname(ROOT), "trunk", "data", "数据归属周期")
    cands = sorted(glob.glob(os.path.join(base, "*", "快照日期", "20*")))
    return cands[-1] if cands else None


def site_reconcile(snap_dir, fmt, book_filter=None):
    """逐站 G（库存报告在途）vs H（我方货件列在途）对账表。

    G = 该站库存报告的 afn-inbound-working + shipped + receiving
    H = 该站分析表里「行5 未标到货」的货件列数量之和
    收窄口径（见 inv_map_for_book）：每张分析表只吃自己的报告，不串味。
    返回 [(簿名, 表名, G, H, 列数)]
    """
    rows = []
    for path in sorted(glob.glob(os.path.join(C.WORK_DIR, "*.xlsx"))):
        name = os.path.basename(path)
        if name.startswith("~$"):
            continue
        if book_filter and book_filter not in name:
            continue
        cfg = next((b for b in C.WORKBOOKS if b.get("file") == name), None)
        if cfg is None:
            continue
        inv_map = inv_map_for_book(cfg, snap_dir, fmt)
        if not inv_map:
            continue
        wb = load_workbook(path)
        try:
            for sn in wb.sheetnames:
                if "分析表" not in sn:
                    continue
                key = C.stock_sheet_for(sn)
                inv = inv_map.get(str(key).strip()) if key else None
                if not inv:
                    continue
                ws = wb[sn]
                fc, lc = Sh.detect_sheet_cols(ws)
                sku_rows = Sh._sku_rows(ws)
                G = sum(int(v.get("inb") or 0) for v in inv.values())
                H, ncols = 0, 0
                for c in range(fc, lc + 1):
                    if not ws.cell(1, c).value:
                        continue
                    if Sh.is_arrived_label(ws.cell(5, c).value):
                        continue
                    ncols += 1
                    for _sku, r in sku_rows.items():
                        v = ws.cell(r, c).value
                        if isinstance(v, (int, float)) and v:
                            H += v
                rows.append((name, sn.strip(), G, int(H), ncols))
        finally:
            wb.close()
    return rows


def print_reconcile(rows):
    """打印对账表 + 按簿小计。"""
    if not rows:
        print("无可对账的站点（本簿没有 afn-* 在途库存报告）")
        return
    print(f"\n{'簿':34s} {'表':22s} {'G亚马逊':>8s} {'H我方':>8s} {'差':>7s} {'吻合':>7s} 列")
    print("-" * 100)
    tG = tH = 0
    cur = None
    sub_G = sub_H = 0
    for name, sn, G, H, n in rows:
        if cur and name != cur:
            print(f"  └ {cur[:30]:30s} 小计 {sub_G:8d} {sub_H:8d} {sub_G-sub_H:+7d} "
                  f"{(min(sub_G,sub_H)/max(sub_G,sub_H)*100 if max(sub_G,sub_H) else 100):6.1f}%")
            sub_G = sub_H = 0
        cur = name
        sub_G += G
        sub_H += H
        tG += G
        tH += H
        rate = (min(G, H) / max(G, H) * 100) if max(G, H) else 100.0
        print(f"{name:34s} {sn:22s} {G:8d} {H:8d} {G-H:+7d} {rate:6.1f}% {n}")
    if cur:
        print(f"  └ {cur[:30]:30s} 小计 {sub_G:8d} {sub_H:8d} {sub_G-sub_H:+7d} "
              f"{(min(sub_G,sub_H)/max(sub_G,sub_H)*100 if max(sub_G,sub_H) else 100):6.1f}%")
    print("-" * 100)
    print(f"{'合计':34s} {'':22s} {tG:8d} {tH:8d} {tG-tH:+7d} "
          f"{(min(tG,tH)/max(tG,tH)*100 if max(tG,tH) else 100):6.1f}%")
    print("G−H >0：亚马逊已识别、我方没记（漏货件或数量少登记）；"
          "G−H <0：我方记多（整票未被识别 / 挂错簿 / 已到货没打标）")


def main(argv):
    apply = "--apply" in argv
    backup = "--backup" in argv
    detail = "--detail" in argv
    prune = "--prune" in argv
    alarm = "--alarm" in argv
    recon = "--recon" in argv
    book_filter = argv[argv.index("--book") + 1] if "--book" in argv else None

    snap_day = os.environ.get("TRANSIT_SNAP") or datetime.date.today().strftime("%Y%m%d")
    period = os.environ.get("TRANSIT_PERIOD") or snap_day[:6]
    base = os.path.join(os.path.dirname(ROOT), "trunk", "data", "数据归属周期")
    snap_dir = os.path.join(base, period, "快照日期", snap_day)
    if not os.path.isdir(snap_dir):
        snap_dir = _latest_snap()
        if not snap_dir:
            print("找不到任何快照目录")
            return 2
        snap_day = os.path.basename(snap_dir)
        period = os.path.basename(os.path.dirname(os.path.dirname(snap_dir)))
    fmt = dict(d=snap_day, d4=snap_day[4:], m=period)
    ref_date = datetime.date(int(snap_day[:4]), int(snap_day[4:6]), int(snap_day[6:8]))
    print(f"快照：{snap_day}  参照日：{ref_date}  模式：{'落盘' if apply else '预演'}"
          + ("  +prune" if prune else ""))

    if recon:
        print_reconcile(site_reconcile(snap_dir, fmt, book_filter))
        return 0

    if backup and apply:
        bdir = os.path.join(ROOT, ".workbuddy", f"backup_transit_{snap_day}")
        os.makedirs(bdir, exist_ok=True)
        for p in glob.glob(os.path.join(C.WORK_DIR, "*.xlsx")):
            shutil.copy2(p, bdir)
        print(f"备份 → {bdir}")

    tot = dict(arrived=0, partial=0, pending=0, weak=0)
    all_alarms = []
    for path in sorted(glob.glob(os.path.join(C.WORK_DIR, "*.xlsx"))):
        name = os.path.basename(path)
        if name.startswith("~$"):
            continue
        cfg = next((b for b in C.WORKBOOKS if b.get("file") == name), None)
        if cfg is None:
            continue
        if book_filter and book_filter not in name:
            continue
        inv = inv_map_for_book(cfg, snap_dir, fmt)     # 逐 sheet 各配自己的库存报告
        if not inv:
            continue
        if not alarm:
            n_sku = sum(len(v) for v in inv.values())
            print(f"\n==== {name}（{len(inv)} 份库存报告 / {n_sku} SKU·站点）："
                  + "、".join(f"{k}={len(v)}" for k, v in inv.items()))
        res, alarms = audit_book(path, inv, apply=apply, detail=detail,
                                 prune=prune, ref_date=ref_date, alarm_only=alarm)
        for sn, mis in alarms:
            all_alarms.append((name, sn, mis))
        if alarm:
            continue
        for sn, ws, dec, veto in res:
            acts = {"arrived": [], "partial": [], "pending": [], "weak": [], "empty": []}
            for _c, d in dec.items():
                acts[d["act"]].append(d)
            if veto:
                print(f"  {sn.strip()}: [撤回] {veto}")
            if not any(acts[k] for k in ("arrived", "partial", "weak")):
                if acts["pending"]:
                    print(f"  {sn.strip()}: 全部 {len(acts['pending'])} 票都还在途（无需动作）")
                continue
            print(f"  {sn.strip()}")
            for d in acts["arrived"]:
                tot["arrived"] += 1
                age = f"{d['age']}天前" if d["age"] is not None else "无日期"
                print(f"    OK 已到货        {d['col']:2s} {d['sid']:16s} "
                      f"{str(d['label'])[:14]:16s} 数量 {d['qty']:5d} 发货{age:8s} "
                      f"差额贡献 {d['delta']:+6d}")
                if detail:
                    print(f"       证据 SKU：{d['present'][:12]}")
            for d in acts["partial"]:
                tot["partial"] += 1
                rec = int(round(d["qty"] * d["share"]))
                print(f"    ~  部分到货      {d['col']:2s} {d['sid']:16s} "
                      f"{str(d['label'])[:14]:16s} 数量 {d['qty']:5d}"
                      f"（已识别约 {rec}，留 {d['qty']-rec}）")
            for d in acts["pending"]:
                tot["pending"] += 1
            for d in acts["weak"]:
                tot["weak"] += 1
                print(f"    ?? 待人工        {d['col']:2s} {d['sid']:16s} "
                      f"{str(d['label'])[:14]:16s} 数量 {d['qty']:5d} "
                      f"差额贡献 {d['delta']:+6d} 原因={d['reason']}")

    if all_alarms:
        print("\n" + "=" * 62)
        print("库存 > 在途 = 货件表可能漏更新（亚马逊有在途、我方货件列里没有）")
        print("=" * 62)
        for name, sn, mis in all_alarms:
            total_gap = sum(m[3] for m in mis)
            print(f"\n  {name} / {sn}：{len(mis)} 个 SKU，合计缺 {total_gap} 件")
            for sku, G, H, gap in mis[:8]:
                print(f"      {sku:10s} 亚马逊在途={G:5d}  我方调拨={H:5d}  差 {gap:+5d}")
            if len(mis) > 8:
                print(f"      … 另有 {len(mis)-8} 个 SKU")

    if not alarm:
        print(f"\n汇总：已到货 {tot['arrived']} 票、部分到货 {tot['partial']} 票、"
              f"在途 {tot['pending']} 票、待人工 {tot['weak']} 票")
        if not apply:
            print("未落盘。确认无误后加 --apply（建议同时 --backup）；"
                  f"要清已到货列再加 --prune（到货保留 {Sh.KEEP_ARRIVED_DAYS} 天）")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
