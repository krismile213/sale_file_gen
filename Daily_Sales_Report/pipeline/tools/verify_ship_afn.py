# -*- coding: utf-8 -*-
"""货件台账 × AFN 库存报告 三道闸对账（防手工出错 + 时间错位归因）。

闸1 票级机验：表内每列货件合计 vs TSV 原始报告「已发货」合计（抓抄错数）。
闸2 时间分层：未到货票按 发货日 vs 快照日 分层——
    在网层(发货日<快照日) / 边界层(=快照日, 报告可能未含) / 未来层(>, 报告必不含)。
闸3 AFN 对账：快照库存报告 afn-inbound-working/shipped/receiving 合计 vs 手工未到货，
    差额 = 未入网在途（海运未登记等）或 收货未标「到货」的多计。
差分（可选）：两快照 AFN 差 = 期间净入网，可定位哪些票已入网。

用法: python verify_ship_afn.py [快照日=20261013] [对照日=20261012]
输出: 说明文档/核查-货件台账与AFN对账-{快照日}.md
"""
# --- pipeline 路径引导（2026-10-14 目录收敛自动注入）---
import os as _os, sys as _sys
_bp = _os.path.dirname(_os.path.abspath(__file__))
for _i in range(6):
    if _os.path.isdir(_os.path.join(_bp, "1.每日销售数据表")):
        break
    _bp = _os.path.dirname(_bp)
DSR = _bp
_bcore = _os.path.join(DSR, "pipeline", "core")
if _bcore not in _sys.path:
    _sys.path.insert(0, _bcore)
# --- 引导结束 ---
import os, sys, csv, re, datetime

# [normalize] DSR 由上方引导块赋值
pass  # [normalize] 路径已由引导块设置
import openpyxl
import config as C
import shipments as Sh
from openpyxl.utils import get_column_letter as L

DAY = sys.argv[1] if len(sys.argv) > 1 else "20261013"
DAY2 = sys.argv[2] if len(sys.argv) > 2 else "20261012"
snap = datetime.date(int(DAY[:4]), int(DAY[4:6]), int(DAY[6:8]))
snap_dir = os.path.join(C.SNAP_BASE, C.CURRENT_PERIOD, "快照日期", DAY)
snap_dir2 = os.path.join(C.SNAP_BASE, C.CURRENT_PERIOD, "快照日期", DAY2)
OUT = os.path.join(os.path.dirname(DSR), "说明文档", f"核查-货件台账与AFN对账-{DAY}.md")

_md = []
def emit(s=''):
    _md.append(str(s))

ships = {s['id']: s for s in Sh.load_shipments(C.SHIP_DIR)}

def parse_label(v):
    """'8/18空运9/9到货' -> (8,18, True)；'9/29海运' -> (9,29, False)。"""
    m = re.match(r"^(\d{1,2})/(\d{1,2})", str(v or ""))
    if not m:
        return None, False
    return (int(m.group(1)), int(m.group(2))), ('到货' in str(v))

def d_of(md, ref):
    """(m,d) -> date，年份取参考日附近（m>ref.m 用上一年，反之当年）。"""
    m, d = md
    y = ref.year
    if m - ref.month > 6:
        y -= 1
    elif ref.month - m > 6:
        y += 1
    return datetime.date(y, m, d)

def read_csv_rows(path):
    for enc in ("utf-8-sig", "shift_jis", "gb18030", "latin-1"):
        try:
            with open(path, encoding=enc, newline="") as fh:
                return list(csv.reader(fh, delimiter=","))
        except (UnicodeDecodeError, OSError):
            continue
    return None

def afn_from_csv(day, snap_d, report_tpls):
    """从快照 CSV 读 AFN 入库合计。返回 (合计, 明细dict) 或 (None, 原因)。"""
    if isinstance(report_tpls, str):
        report_tpls = [report_tpls]
    for tpl in report_tpls:
        try:
            nm = tpl.format(m=C.ORDER_MONTH, d=day, d4=day[4:])
        except (KeyError, IndexError):
            nm = tpl
        p = os.path.join(snap_d, nm)
        if not os.path.exists(p):
            continue
        rows = read_csv_rows(p)
        if not rows:
            continue
        hdr = [h.strip().lower() for h in rows[0]]
        need = ['afn-inbound-working-quantity', 'afn-inbound-shipped-quantity',
                'afn-inbound-receiving-quantity']
        if not all(n in hdr for n in need):
            continue
        idx = [hdr.index(n) for n in need]
        tot = 0.0
        det = dict(w=0.0, s=0.0, r=0.0)
        for row in rows[1:]:
            try:
                w, s_, r_ = (float(row[i]) if i < len(row) and row[i] else 0.0 for i in idx)
            except ValueError:
                continue
            det['w'] += w; det['s'] += s_; det['r'] += r_
            tot += w + s_ + r_
        return tot, det
    return None, "报告缺 inbound 列或文件缺失"

def afn_from_sheet(wbv, src_sheet):
    """从工作簿「可售+在途库存」sheet 读 AFN 入库合计（表内 G 公式真源）。"""
    if src_sheet not in wbv.sheetnames:
        return None, f"源sheet[{src_sheet}]不存在"
    ws = wbv[src_sheet]
    hr = None
    hdr = []
    for r, row in enumerate(ws.iter_rows(min_row=1, max_row=6, values_only=True), 1):
        low = [str(v).strip().lower() if v is not None else '' for v in row]
        if 'afn-inbound-shipped-quantity' in low:
            hr, hdr = r, low
            break
    if hr is None:
        return None, "sheet 无 afn 表头"
    def ci(n):
        return hdr.index(n) + 1 if n in hdr else None
    cw, cs_, cr_ = ci('afn-inbound-working-quantity'), ci('afn-inbound-shipped-quantity'), ci('afn-inbound-receiving-quantity')
    det = dict(w=0.0, s=0.0, r=0.0)
    for row in ws.iter_rows(min_row=hr + 1, values_only=True):
        def gv(c):
            if c is None or c > len(row):
                return 0.0
            v = row[c - 1]
            return float(v) if isinstance(v, (int, float)) else 0.0
        det['w'] += gv(cw); det['s'] += gv(cs_); det['r'] += gv(cr_)
    return det['w'] + det['s'] + det['r'], det

# ---------- 主循环 ----------
emit(f"# 货件台账 × AFN 库存报告 三道闸对账（快照 {DAY}）")
emit(f"")
emit(f"> 生成于 2026-10-14。目的：①防手工抄错数（闸1）；②解释「快照是发货前导的」造成的偏差（闸2）；③AFN 系统口径 vs 手工台账的差额归因（闸3）。")
emit(f"")
g1_bad, rows3 = [], []
for wbinfo in C.WORKBOOKS:
    if wbinfo.get('skip'):
        continue
    bk = wbinfo['file']
    p = os.path.join(C.WORK_DIR, bk)
    if not os.path.exists(p):
        continue
    wbv = openpyxl.load_workbook(p, read_only=True, data_only=True)
    wbf = openpyxl.load_workbook(p, read_only=True, data_only=False)
    for sn in [s for s in wbv.sheetnames if '分析' in s]:
        wsv = wbv[sn]
        wsf = wbf[sn]
        # 货件列自动发现：行1 含 FBA 开头编号
        row1 = {}
        for row in wsv.iter_rows(min_row=1, max_row=1, max_col=60):
            for i, c in enumerate(row, 1):
                if isinstance(c.value, str) and c.value.strip().upper().startswith('FBA'):
                    row1[i] = c.value.strip()
        if not row1:
            continue
        # 行5 标签
        row5 = {}
        for row in wsv.iter_rows(min_row=5, max_row=5, max_col=60):
            for i, c in enumerate(row, 1):
                v = c.value
                if v is not None and (i in row1 or re.match(r"^\d{1,2}/\d{1,2}", str(v))):
                    row5[i] = v
        # G7 公式 → 源 sheet 名（优先引号内表名，避免抓到 "A:A,表名"）
        g7v = wsf.cell(7, 7).value
        g7 = str(g7v.text if hasattr(g7v, 'text') else (g7v or ''))
        m = re.search(r"'([^']+)'!", g7)
        if not m:
            m = re.search(r"([一-龥A-Za-z0-9_]+)!\$", g7)
        src = m.group(1).strip("'") if m else None
        # 逐票
        tickets = []
        for col, sid in sorted(row1.items()):
            lab = row5.get(col, '')
            md, arrived = parse_label(lab)
            # 列合计
            csum, nsku = 0.0, 0
            ci = col
            for row in wsv.iter_rows(min_row=7, min_col=ci, max_col=ci, values_only=True):
                v = row[0]
                if isinstance(v, (int, float)):
                    csum += v; nsku += 1
            t = ships.get(sid)
            tsum = sum(t['items'].values()) if t else None
            tmd = t['ship_md'] if t else None
            tickets.append(dict(col=L(col), sid=sid, lab=str(lab), md=md or tmd,
                                arrived=arrived, csum=csum, tsum=tsum, has_tsv=t is not None,
                                t_date=tmd))
        if not tickets:
            continue
        # ---- 闸1 票级机验 ----
        emit(f"## {bk} · {sn}")
        emit("")
        emit("### 闸1 票级机验（表内列合计 vs TSV 已发货）")
        emit("")
        emit("| 列 | 货件号 | 标签 | 表内合计 | TSV合计 | 结论 |")
        emit("|---|---|---|---:|---:|---|")
        for t in tickets:
            if t['has_tsv']:
                ok = abs(t['csum'] - t['tsum']) < 0.5
                note = "✓" if ok else "**✗ 差异!**"
                if not ok:
                    g1_bad.append((bk, sn, t))
            else:
                note = "无TSV(历史票,未机验)"
            emit(f"| {t['col']} | {t['sid']} | {t['lab'][:22]} | {t['csum']:.0f} | "
                 f"{'' if t['tsum'] is None else format(t['tsum'], '.0f')} | {note} |")
        # ---- 闸2 分层 + 闸3 AFN ----
        undeliv = [t for t in tickets if not t['arrived']]
        onweb = [t for t in undeliv if t['md'] and d_of(t['md'], snap) < snap]
        border = [t for t in undeliv if t['md'] and d_of(t['md'], snap) == snap]
        future = [t for t in undeliv if t['md'] and d_of(t['md'], snap) > snap]
        sum_onweb = sum(t['csum'] for t in onweb)
        sum_border = sum(t['csum'] for t in border)
        sum_future = sum(t['csum'] for t in future)
        sum_undeliv = sum(t['csum'] for t in undeliv)
        afn, det = (None, None)
        if src:
            afn, det = afn_from_sheet(wbv, src)
        emit("")
        emit("### 闸2 时间分层（未到货票 × 发货日 vs 快照日）")
        emit("")
        emit(f"- 快照日：{snap}（库存报告为当日导出，早于/等于当日发货的货可能未入网）")
        for t in onweb:
            emit(f"  - 在网层 {t['col']} {t['sid']} 发货 {t['md'][0]}/{t['md'][1]} {t['csum']:.0f}件")
        for t in border:
            emit(f"  - **边界层** {t['col']} {t['sid']} 发货 {t['md'][0]}/{t['md'][1]}（=快照日，报告可能未含）{t['csum']:.0f}件")
        for t in future:
            emit(f"  - **未来层** {t['col']} {t['sid']} 发货 {t['md'][0]}/{t['md'][1]}（>快照日，报告必不含）{t['csum']:.0f}件")
        if not undeliv:
            emit("  - （无未到货票）")
        # 闸3
        emit("")
        emit("### 闸3 AFN 对账")
        emit("")
        if afn is None:
            emit(f"- AFN 读取失败：{det}（G7={g7[:70]}）")
        else:
            gap = afn - sum_onweb
            emit(f"| AFN在途(working+shipped+receiving) | 手工未到货 | 在网层 | 边界层 | 未来层 | AFN−在网层 |")
            emit(f"|---:|---:|---:|---:|---:|---:|")
            emit(f"| {afn:.0f} | {sum_undeliv:.0f} | {sum_onweb:.0f} | {sum_border:.0f} | {sum_future:.0f} | {gap:+.0f} |")
            emit("")
            if abs(gap) < 0.5:
                emit("- **精确对上**：全部在网层票已入网、无多计。")
            elif gap > 0:
                emit(f"- AFN 比在网层多 {gap:.0f} 件：通常=某票已到 FC 正在收货（receiving）但行5 还没标「到货」。")
            else:
                emit(f"- AFN 比在网层少 {-gap:.0f} 件：=未入网在途（海运/卡航货件未在亚马逊系统登记；属正常时间差，不是抄错数）。")
        emit("")
    wbv.close()
    wbf.close()

# ---- 差分 ----
emit(f"## 附：快照差分（{DAY2} → {DAY} 的 AFN 净入网）")
emit("")
emit("| 簿·表 | 报告 | AFN@" + DAY2 + " | AFN@" + DAY + " | 净入网 |")
emit("|---|---|---:|---:|---:|")
for wbinfo in C.WORKBOOKS:
    if wbinfo.get('skip'):
        continue
    for st in wbinfo.get('stocks', []):
        if st.get('kind') != 'inv':
            continue
        a1, _ = afn_from_csv(DAY2, snap_dir2, st['report'])
        a2, _ = afn_from_csv(DAY, snap_dir, st['report'])
        if a1 is None or a2 is None:
            emit(f"| {wbinfo['file'][:14]}·{st['sheet'][:10]} | {st['report'][:30]} | "
                 f"{'N/A' if a1 is None else format(a1, '.0f')} | "
                 f"{'N/A' if a2 is None else format(a2, '.0f')} | 旧源无入库列/缺报告 |")
        else:
            emit(f"| {wbinfo['file'][:14]}·{st['sheet'][:10]} | {st['report'][:30]} | "
                 f"{a1:.0f} | {a2:.0f} | {a2 - a1:+.0f} |")
emit("")
emit(f"**附注**：")
emit(f"- 10-12 快照的美/加/墨三份管理库存报告为**独立数据**（可售 5155/5153/5157 各不同），")
emit(f"  但 inbound 三件套三国完全相同（W=1427 S=758 R=0）——符合北美统一账户(NARF)的")
emit(f"  区域级在途口径：同批货先入美国 FC 再转运加/墨。非导出错误，解读时按区域看待。")
emit(f"- 日本/澳大利亚/interflow澳 的 CSV 为日文界面导出（无英文 afn-* 表头），差分标 N/A；")
emit(f"  工作簿内对应 sheet 已做映射，闸3 用的就是工作簿口径，不受影响。")

emit("")
emit("## 使用方法（以后每天）")
emit("")
emit("```")
emit("python .workbuddy/verify_ship_afn.py 20261014        # 每天快照导入后跑一次")
emit("```")
emit("三道闸的读法：")
emit("1. **闸1 有 ✗** → 手工/程序抄错数，立即修（对照 TSV 原始件数）。")
emit("2. **闸2 边界/未来层** → 这些票不出现在当天 AFN 里是正常的，别当成差异追查；")
emit("   判断某票是否应出现在 AFN：发货日 < 快照日 → 应该在；≥ 快照日 → 不一定/不该在。")
emit("3. **闸3 差额为负**（AFN < 手工在途）= 未入网（海运登记滞后，正常）；")
emit("   **差额为正** = 有票已到 FC 在收货但没标「到货」→ 去行5 补「到货」二字。")
emit("4. **差分行**：净入网 ≈ 某票件数 → 该票已入网；≈ 0 或为负 → 新发货未登记+在收货抵扣。")

os.makedirs(os.path.dirname(OUT), exist_ok=True)
with open(OUT, 'w', encoding='utf-8') as f:
    f.write('\n'.join(_md) + '\n')
print('written', OUT)
print('闸1差异票数:', len(g1_bad))
for bk, sn, t in g1_bad:
    print('  差异:', bk[:14], sn[:10], t['sid'], '表内', t['csum'], 'TSV', t['tsum'])
