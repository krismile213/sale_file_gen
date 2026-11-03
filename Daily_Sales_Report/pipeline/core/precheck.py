# -*- coding: utf-8 -*-
"""跑前体检：确认当天快照「报告是否导完整」+「内容是否与上一快照相同」+「有无重复文件」。

被 pipeline/run.py 在正式跑之前自动调用；也可单独跑：
    python pipeline/core/precheck.py            # 体检最新快照
    python pipeline/core/precheck.py 20261014   # 体检指定快照日
    python pipeline/core/precheck.py --quiet    # 只输出了问题的部分

检查九块：
  ① 完整性   —— 按 config.WORKBOOKS 逐本簿核对所需报告是否都在、格式是否对、有没有空文件
  ② 工作簿占用 —— WPS/Excel 开着 → 跑时该簿会被整份跳过
  ③ 内容对比 —— 与「上一个快照」逐个同名文件比 MD5：相同/已更新/新增/消失
  ④ 重复内容 —— 本快照内内容完全相同的文件组（重复下载 or 同源）。
                     **硬闸**：组内 ≥2 个库存类报告且不是同一份重复下载 ⇒ 必然导错账号/站点
                     （2026-10-30 加拿大实拍：主账号报告被存成 interflow 那份，MD5 相同、
                     SKU 362→83，放行后「清空5430格」，279 个 SKU 可售变空）。
                     订单报告不判硬伤 —— 中东大区一份报告存成沙特/阿联酋两个名字是合法形态。
  ⑤ 货件     —— TSV 票数、重复货件号（同一票下载两次）
  ⑥ 货件列容量 —— 每张分析表还剩几个空列位
  ⑦ 数组公式隐患 —— 分析表里 `t="array"` 且整列引用被当标量用（如 VLOOKUP(A:A,…)）
                     的格子，会静默算成 0（2026-10-16 阿联酋簿实测）
  ⑧ 订单报告日期覆盖 —— 订单报告 purchase-date 的 min/max：月初/月末被本地时间
                     date range 截断、月内单日 0 单（详见 report_coverage.py）
  ⑨ 订单报告渠道错位 —— 订单报告 sales-channel 分布 vs 本簿 expect_channel：
                     整份下错站点（外来渠道 ≥50% 或无一条本店渠道）→ runner 会整份拒收
  ⑩ 销售报告 SKU 归属 —— 订单报告 SKU 集 ∩ 本簿库存 sheet SKU 集 ÷ 报告 SKU 集。
                     <60% ⇒ 几乎肯定是「导成了别的账号的同渠道报告」（2026-10-22 补）。
                     ⑨ 挡不住这种：Amazon.com 同属北美簿与 Interflowd北美簿，渠道全合法。

返回码：0=可跑；1=有缺失报告或闸门硬伤（run.py 默认据此中止，--force 可强行继续）
"""
import os
import re
import sys
import glob
import hashlib

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import config as C


# ---------------- 基础 ----------------
def _md5(path):
    h = hashlib.md5()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def _fmt(tmpl, day):
    try:
        return tmpl.format(m=C.ORDER_MONTH, d=day, d4=day[4:])
    except (KeyError, IndexError):
        return tmpl


def _probe(path):
    """轻量格式探测 → (True/False/None, 说明)。False=明显有问题。"""
    if path is None:
        return False, "缺失"
    size = os.path.getsize(path)
    if size == 0:
        return False, "空文件(0B)"
    if path.lower().endswith((".xlsx", ".xls")):
        return None, f"xlsx({size} B)"
    try:
        with open(path, "rb") as f:
            raw = f.read(300000)
    except OSError as e:
        return False, f"读取失败 {e!r}"
    txt = None
    for enc in ("utf-8-sig", "utf-8", "shift_jis", "gb18030", "latin-1"):
        try:
            txt = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    if txt is None:
        return False, "编码无法解码"
    lines = (txt or "").splitlines()
    if len(lines) <= 1:
        return False, f"疑似只有表头({len(lines)} 行)"
    head = (lines[0] or "").lower()
    if "afn-" in head:
        return True, f"inv(≥{len(lines)} 行)"
    if "amazon-order-id" in head:
        return True, f"订单报告(≥{len(lines)} 行)"
    if "seller-sku" in head:
        return True, "multi(多国库存)"
    if "inv-age" in head or "snapshot-date" in head:
        return True, "age(库龄式)"
    if "reserved_qty" in head:
        return True, "reserved(预留式)"
    return None, "未识别列头: " + head[:50]


def _prev_day(day):
    """比 day 更早的最近一个快照日（跨周期），返回 (day, period) 或 (None, None)。"""
    cands = []
    if os.path.isdir(C.SNAP_BASE):
        for per in os.listdir(C.SNAP_BASE):
            sd = os.path.join(C.SNAP_BASE, per, "快照日期")
            if not os.path.isdir(sd):
                continue
            for d in os.listdir(sd):
                if re.match(r"^\d{8}$", d) and d < day:
                    cands.append((d, per))
    if not cands:
        return None, None
    cands.sort()
    return cands[-1]


def _norm(fname):
    """把报告名里的日期挖成占位符，得到「同一类报告」的归一化键。

    亚马逊报告名带下载日期（…20261014.csv），逐日不同；直接按字面名比对会把
    每天的报告全判成「新增+消失」。挖掉 8 位日 / 6 位月 / 4 位 MMDD 后即可配对：
        亚马逊英国管理库存报告20261014.csv → 亚马逊英国管理库存报告{d}.csv
        亚马逊欧盟站点202610订单报告.txt  → 亚马逊欧盟站点{m}订单报告.txt
        多国库存报告.csv                  → 多国库存报告.csv（无日期，同名即同键）
    """
    base, ext = os.path.splitext(fname)
    base = re.sub(r"(?<!\d)\d{8}(?!\d)", "{d}", base)
    base = re.sub(r"(?<!\d)\d{6}(?!\d)", "{m}", base)
    base = re.sub(r"(?<!\d)\d{4}(?!\d)", "{d4}", base)
    return base + ext


# 库存类报告：同一「账号 × 站点」只会有一份，两份内容相同 = 必然导错（10-30 加拿大事故）。
# 订单报告不在此列 —— 中东大区一份报告存成沙特/阿联酋两个名字是**合法**形态（runner 按渠道切分）。
_STOCK_KEYS = ("管理库存报告", "库存报告", "inventory", "库存-health")


def _is_stock_report(fname):
    n = fname.lower()
    return any(k.lower() in n for k in _STOCK_KEYS)


def _dup_key(fname):
    """重复下载的同一份文件（xxx.csv / xxx (1).csv）算同键；不同来源的报告算不同键。

    ⚠️ 必须**先去 `(1)` 再挖日期** —— `(1)` 在扩展名之前（…20261030 (1).csv），
    对整串跑 `\\(\\d+\\)$` 是匹配不到的（10-30 实测踩过）。
    """
    base, ext = os.path.splitext(fname)
    base = re.sub(r"\s*\(\d+\)$", "", base)
    return _norm(base + ext)


def _index(d):
    """目录内「归一化报告名 → 实际文件名」；同键多个取第一个。"""
    out = {}
    for f in sorted(os.listdir(d)):
        p = os.path.join(d, f)
        if os.path.isfile(p) and not f.startswith("~$"):
            out.setdefault(_norm(f), f)
    return out


def _locked(path):
    """文件是否被其他程序占用（WPS/Excel 开着）。

    两个信号：① 同目录存在 `~$文件名` 锁文件（WPS/Excel 打开时创建）；
    ② 可写句柄试探失败（只开句柄，不改内容）。
    """
    if not os.path.exists(path):
        return False
    d, n = os.path.split(path)
    if os.path.exists(os.path.join(d, "~$" + n)):
        return True
    try:
        with open(path, "r+b"):
            pass
        return False
    except PermissionError:
        return True
    except OSError:
        return False


# ---------------- ⑩ 销售报告 SKU 归属 ----------------
SKU_MIN_RATE = 0.6     # 硬伤线：与 runner.screen_sku_overlap 同阈值 → 会整份拒收
SKU_WARN_RATE = 0.9    # 提示线：合法组合实测 ≥96.5%，掉到 90% 以下要人看一眼

_ENCS = ("utf-8-sig", "utf-8", "shift_jis", "gb18030", "latin-1")


def _resolve(snap_dir, day, tmpl):
    """报告候选模板 → 第一个存在的路径（与 collect / runner.process 同规则）。

    2026-10-24 起支持 `*` / `?` 通配符（取字典序最新、忽略 "(1)" 副本），
    与 runner._resolve_report 完全一致 —— interflow 沙特仓的库存报告文件名
    会滞后几天（内容新、名字旧），死认 {d} 会天天误报缺报告。
    """
    import glob as _glob
    for t in (tmpl if isinstance(tmpl, list) else [tmpl]):
        try:
            c = t.format(m=C.ORDER_MONTH, d=day, d4=day[4:])
        except (KeyError, IndexError):
            c = t
        p = c if os.path.isabs(c) else os.path.join(snap_dir, c)
        if "*" in p or "?" in p:
            hits = [h for h in _glob.glob(p) if "(" not in os.path.basename(h)]
            hits = hits or _glob.glob(p)
            if hits:
                return sorted(hits)[-1]
            continue
        if os.path.exists(p):
            return p
    return None


def _read_table(path):
    """轻量读 tsv/csv/xlsx → (header, body)；读不动返回 (None, None)。

    只给 ⑩ 用：要的是「报告里有哪些 SKU」，不需要完整列映射，所以走最省的路径
    （txt 按行 split，xlsx 只读第一张表）。
    """
    if not path or not os.path.exists(path):
        return None, None
    ext = os.path.splitext(path)[1].lower()
    if ext in (".xlsx", ".xlsm"):
        try:
            import openpyxl
            wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        except Exception:
            return None, None
        try:
            ws = wb[wb.sheetnames[0]]
            it = ws.iter_rows(values_only=True)
            head = next(it, None)
            body = [r for r in it if r and any(v not in (None, "") for v in r)]
            return (list(head) if head else None), body
        except Exception:
            return None, None
        finally:
            wb.close()
    try:
        with open(path, "rb") as f:
            raw = f.read()
    except OSError:
        return None, None
    txt = None
    for enc in _ENCS:
        try:
            txt = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    if txt is None:
        return None, None
    lines = [l for l in txt.splitlines() if l.strip()]
    if len(lines) < 2:
        return None, None
    sep = "\t" if "\t" in lines[0] else ","
    head = [h.strip() for h in lines[0].split(sep)]
    body = [l.split(sep) for l in lines[1:]]
    return head, body


def _order_sku_set(path):
    """订单报告里的 SKU 集（列名 `sku`）；非亚马逊报告（美客多等）无此列 → 空集。"""
    head, body = _read_table(path)
    if not head or not body:
        return set()
    low = [str(h).strip().lower() for h in head]
    try:
        i = low.index("sku")
    except ValueError:
        return set()
    out = set()
    for r in body:
        if len(r) > i:
            v = r[i]
            v = str(v).strip() if v is not None else ""
            if v:
                out.add(v)
    return out


def stock_sku_map(workbooks=None):
    """{簿文件名: 本簿库存 sheet 的 SKU 集}。

    口径与 runner.process() 完全一致：库存 sheet 的 **A 列**、前 3000 行、非空字符串。
    读不到（不存在/被占用/损坏）的簿**不进结果** —— 调用方据此跳过判定，不能瞎判。
    """
    import openpyxl
    workbooks = C.WORKBOOKS if workbooks is None else workbooks
    out = {}
    for bk in workbooks:
        if bk.get("skip"):
            continue
        path = os.path.join(C.WORK_DIR, bk["file"])
        if not os.path.exists(path) or _locked(path):
            continue
        try:
            wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        except Exception:
            continue
        try:
            s = set()
            for st in (bk.get("stocks") or []):
                sn = st.get("sheet")
                if sn not in wb.sheetnames:
                    continue
                for row in wb[sn].iter_rows(min_col=1, max_col=1, max_row=3000,
                                            values_only=True):
                    v = row[0] if row else None
                    if isinstance(v, str) and v.strip():
                        s.add(v.strip())
            out[bk["file"]] = s
        except Exception:
            continue
        finally:
            try:
                wb.close()
            except Exception:
                pass
    return out


def sales_sku_overlap(snap_dir, day, workbooks, min_rate=SKU_MIN_RATE):
    """⑩ 销售报告 SKU 归属 → (rows, n_checked)。

    rows = [(簿, sheet, 报告名, rate, hit, n_rep, [(级别, 文本)])]
    n_checked = 实际判过的报告份数。**必须回报**：库存读不到时判不了，
    若只看 rows 为空就显示 ✅，会变成「假绿灯」（2026-10-22 加）。

    **为什么 ⑨ 渠道闸挡不住**：`Amazon.com` / `Amazon.ca` 同时属于「北美加拿大簿」和
    「Interflowd北美加拿大簿」，两本都在 owner 集合里 → 渠道闸认为合法，照单全收。
    2026-10 实测：interflow 北美订单报告被换成了主账号的（5325 行），渠道全合法，
    `Interflowd北美加拿大库存销售表` 被灌进 4288 行主账号订单。

    唯一可靠的判据是 SKU 归属：报告 SKU 集 ∩ 本簿库存 SKU 集 ÷ 报告 SKU 集。
    实测合法组合 ≥96.5%（欧盟英国 96.5%、其余 100%），错位报告只有 10~15%。
    """
    if not snap_dir or not os.path.isdir(snap_dir):
        return [], 0
    smap = stock_sku_map(workbooks)
    out, n_checked = [], 0
    for bk in workbooks:
        if bk.get("skip"):
            continue
        fname = bk.get("file", "?")
        stock = smap.get(fname) or set()
        if not stock:
            continue                       # 库存读不到 → 不判（不制造假红灯）
        specs = bk.get("sales") or []
        specs = specs if isinstance(specs, list) else [specs]
        for sc in specs:
            rp = _resolve(snap_dir, day, sc.get("report"))
            if not rp:
                continue                   # 缺报告归 ① 管，这里不重复报
            rep = _order_sku_set(rp)
            if not rep:
                continue                   # 无 sku 列（美客多/PDF 类）→ 不适用
            n_checked += 1
            hit = len(rep & stock)
            rate = hit / len(rep)
            # 按簿放宽阈值（2026-10-24）：interflow 沙特簿的库存只覆盖沙特仓 29 个 SKU，
            # 而 Amazon.sa 渠道有 19 个 SKU —— 多数从阿联酋仓发货，本就不在这 29 个里，
            # 天然只有 32% 重合。这不是"导错账号"（同渠道错账号实测 10~15%），
            # 所以按簿下调阈值到 0.2，仍能抓住真正的错账号报告。
            _min = float(sc.get("sku_gate_min", min_rate))
            iss = []
            if rate < _min:
                iss.append(("✗", f"SKU 归属不符：报告 {len(rep)} 个 SKU 只有 {hit} 个"
                                  f"在本簿库存里（{rate:.0%} < {_min:.0%}）"))
                iss.append(("", f"几乎肯定是**导成了别的账号**的同渠道报告"
                                f"（{sc.get('expect_channel')} 两个账号都有这个渠道）"
                                f" —— runner 会整份拒收，本簿当日销售不更新"))
            elif rate < SKU_WARN_RATE:
                iss.append(("△", f"SKU 重合率偏低 {rate:.0%}（{hit}/{len(rep)}）"
                                  f"—— 合法组合实测 ≥96.5%，请核对是否漏导了站点"))
            if iss:
                out.append((fname, sc.get("sheet", "?"), os.path.basename(rp),
                            rate, hit, len(rep), iss))
    return out, n_checked


# ---------------- 收集所需报告 ----------------
def collect(snap_dir, day):
    """按 config.WORKBOOKS 列出每本簿每张 sheet 所需的报告 → 列表。"""
    rows = []
    for wb in C.WORKBOOKS:
        if wb.get("skip"):
            continue
        items = list(wb.get("stocks", []))
        s = wb.get("sales")
        if s:
            items += s if isinstance(s, list) else [s]
        for it in items:
            rp = it.get("report")
            cands = rp if isinstance(rp, list) else [rp]
            # 必须走 _resolve（支持 * 通配符回退），不能只 os.path.exists ——
            # 否则沙特仓这种「文件名滞后、内容已更新」的报告会被误报缺报告，
            # 而 runner 那边其实能读到，两边判断不一致（2026-10-24 踩坑）。
            path = _resolve(snap_dir, day, cands)
            nm = os.path.basename(path) if path else None
            label = nm or " / ".join(os.path.basename(_fmt(x, day)) for x in cands)
            rows.append(dict(book=wb["file"], sheet=it["sheet"], name=label, path=path))
    return rows


# ---------------- 主检查 ----------------
def check(day=None, snap_dir=None):
    if day is None:
        snap_dir, day = C.latest_snapshot()
    if snap_dir is None:
        snap_dir = os.path.join(C.SNAP_BASE, C.CURRENT_PERIOD, "快照日期", day)

    res = dict(day=day, snap_dir=snap_dir, rows=[], missing=[], bad=[], locked=[],
               prev=None, prev_period=None, same=[], changed=[], added=[], gone=[],
               dup_groups=[], ship=dict(tsv=0, tickets=0, dups=[]))

    # ---- ① 完整性 ----
    rows = collect(snap_dir, day)
    res["rows"] = rows
    for r in rows:
        ok, note = _probe(r["path"])
        r["note"] = note
        if r["path"] is None:
            res["missing"].append(r)
        elif ok is False:
            res["bad"].append((r, note))

    # ---- 工作簿占用（WPS/Excel 开着 → 跑时该簿会被整份跳过、少算数）----
    for wb in C.WORKBOOKS:
        if wb.get("skip"):
            continue
        if _locked(os.path.join(C.WORK_DIR, wb["file"])):
            res["locked"].append(wb["file"])

    # ---- ② 内容对比（vs 上一个快照，按归一化报告名配对）----
    prev, prev_per = _prev_day(day)
    res["prev"], res["prev_period"] = prev, prev_per
    if prev:
        pdir = os.path.join(C.SNAP_BASE, prev_per, "快照日期", prev)
        if os.path.isdir(pdir):
            cur, prv = _index(snap_dir), _index(pdir)
            for k in sorted(cur):
                f = cur[k]
                if k in prv:
                    if _md5(os.path.join(snap_dir, f)) == _md5(os.path.join(pdir, prv[k])):
                        res["same"].append(f)
                    else:
                        res["changed"].append(f)
                else:
                    res["added"].append(f)
            res["gone"] = sorted(prv[k] for k in prv if k not in cur)

    # ---- ③ 重复内容（本快照内 MD5 相同）----
    by_hash = {}
    for f in sorted(os.listdir(snap_dir)):
        p = os.path.join(snap_dir, f)
        if os.path.isfile(p) and not f.startswith("~$"):
            by_hash.setdefault(_md5(p), []).append(f)
    res["dup_groups"] = [g for g in by_hash.values() if len(g) > 1]
    # ④-h 硬闸：组内 ≥2 个**库存类**报告、且它们不是同一份的重复下载 → 必然导错账号/站点。
    #    10-30 实拍：亚马逊加拿大管理库存报告被存成 interflow 加拿大那份（MD5 相同、SKU 362→83），
    #    放行后 runner 照跑「清空5430格」，279 个 SKU 可售变空 → 缺货/断货全虚高。
    res["dup_hard"] = [
        g for g in res["dup_groups"]
        if sum(1 for f in g if _is_stock_report(f)) >= 2
        and len({_dup_key(f) for f in g}) > 1
    ]

    # ---- ④ 货件（TSV + PDF：美客多入库准备单 / eMAG 送货单）----
    tsvs, pdfs = [], []
    if os.path.isdir(C.SHIP_DIR):
        for root, _d, files in os.walk(C.SHIP_DIR):
            for f in files:
                if f.lower().endswith(".tsv"):
                    tsvs.append(os.path.join(root, f))
                elif f.lower().endswith(".pdf"):
                    pdfs.append(os.path.join(root, f))
    ids = {}
    for p in tsvs:
        b = os.path.basename(p)
        m = re.match(r"^(FBA[0-9A-Za-z]+)", b)
        ids.setdefault(m.group(1) if m else b, []).append(b)
    for p in pdfs:
        b = os.path.basename(p)
        # 美客多 Inbound-<运单号>-…；eMAG reception_document_<hash>（真单号在 PDF 内，此处按文件名去重）
        m = re.match(r"^Inbound-([0-9A-Za-z\-]+)-", b, re.I)
        ids.setdefault(m.group(1) if m else os.path.splitext(b)[0], []).append(b)
    res["ship"] = dict(tsv=len(tsvs), pdf=len(pdfs), tickets=len(ids),
                       dups=[(k, v) for k, v in sorted(ids.items()) if len(v) > 1])

    # 无对应簿的货件（2026-10-20）：如 interflow 目录里发往沙特 RUH/JED/DMM 的票。
    # 这类票不会落任何分析表（以前会兜底算成"美国"，虚报在途）→ 跑前就暴露。
    try:
        import shipments as Sh
        _all = Sh.load_shipments(C.SHIP_DIR)
        res["ship_orphan"] = [x for x in _all if x.get("book") == Sh.SKIP]
        res["ship_region_warn"] = [x for x in _all if x.get("warn")]
    except Exception as e:
        res["ship_orphan"] = []
        res["ship_region_warn"] = []
        res["ship_orphan_err"] = repr(e)

    # ---- ⑥ 货件列容量 ----
    try:
        res["capacity"] = ship_capacity()
    except Exception as e:
        res["capacity"] = []
        res["capacity_err"] = repr(e)

    # ---- ⑦ 数组公式隐患（会静默算成 0）----
    try:
        res["arr_risk"] = array_formula_risk()
    except Exception as e:
        res["arr_risk"] = []
        res["arr_risk_err"] = repr(e)

    # ---- ⑧ 订单报告日期覆盖（本地时间 date range 截断月初/月末）----
    try:
        import report_coverage as RC
        res["coverage"] = RC.snapshot_coverage(snap_dir, day, C.ORDER_MONTH)
        res["coverage_err"] = sum(1 for _f, iss in res["coverage"]
                                  for lv, _t in iss if lv == "✗")
    except Exception as e:
        res["coverage"] = []
        res["coverage_err"] = 0
        res["coverage_run_err"] = repr(e)
    # ---- ⑨ 订单报告渠道错位（整份下错站点）----
    try:
        import report_coverage as RC
        res["mislabel"] = RC.channel_mislayout(snap_dir, day, C.WORKBOOKS)
        res["mislabel_err"] = sum(1 for r in res["mislabel"]
                                  for lv, _t in r[3] if lv == "✗")
    except Exception as e:
        res["mislabel"] = []
        res["mislabel_err"] = 0
        res["mislabel_run_err"] = repr(e)
    # ---- ⑩ 销售报告 SKU 归属（渠道闸的盲区：同名渠道被两个账号共用）----
    try:
        res["sku_ov"], res["sku_ov_n"] = sales_sku_overlap(snap_dir, day, C.WORKBOOKS)
        res["sku_ov_err"] = sum(1 for r in res["sku_ov"]
                                for lv, _t in r[6] if lv == "✗")
    except Exception as e:
        res["sku_ov"] = []
        res["sku_ov_n"] = 0
        res["sku_ov_err"] = 0
        res["sku_ov_run_err"] = repr(e)
    # ---- ⑪ 分析表 ASIN 公式行号错位（静默显示成别人的 ASIN / 空，不报错）----
    try:
        res["asin_shift"], res["asin_shift_n"] = asin_ref_shift()
    except Exception as e:
        res["asin_shift"] = []
        res["asin_shift_n"] = 0
        res["asin_shift_err"] = repr(e)
    # ---- ⑫ 库存报告格式（下错报表类型，runner 只会静默跳过该项）----
    try:
        res["stock_kind"] = stock_kind_check(snap_dir, day)
    except Exception as e:
        res["stock_kind"] = []
        res["stock_kind_err"] = repr(e)
    # 闸门硬伤合计：⑨ 渠道整份错位 + ⑩ SKU 归属不符 + ④ 库存报告同内容 + ⑫ 报表类型下错
    res["gate_err"] = ((res.get("mislabel_err") or 0) + (res.get("sku_ov_err") or 0)
                       + len(res.get("dup_hard") or []) + len(res.get("stock_kind") or []))
    return res


# ---------------- 渲染 ----------------
def render(res, quiet=False):
    L = []
    a = L.append
    a("=" * 100)
    a(f"跑前体检  |  快照 {res['day']}")
    a(f"  {res['snap_dir']}")
    a("=" * 100)

    # ① 完整性
    n_req = len(res["rows"])
    a(f"① 导完整：共需 {n_req} 份报告")
    if not res["missing"] and not res["bad"]:
        a("   ✅ 齐全，格式正常")
    for r in res["missing"]:
        a(f"   ✗ 缺  [{r['book']}] {r['sheet']}  <- {r['name']}")
    for r, note in res["bad"]:
        a(f"   !! 异常  [{r['book']}] {r['sheet']}  <- {r['name']}  ({note})")

    # ② 工作簿占用
    if res.get("locked"):
        a(f"② ⚠ 工作簿被占用 {len(res['locked'])} 本（WPS/Excel 开着 → 跑时会被整份跳过、少算数，先关掉）：")
        for f in res["locked"]:
            a(f"      ✗ {f}")
    else:
        _nwb = len([b for b in C.WORKBOOKS if not b.get("skip")])
        a(f"② 工作簿：✅ {_nwb} 本均可写（无占用）")

    # ③ 内容对比
    if res["prev"]:
        a(f"③ 内容对比（vs 上一快照 {res['prev']}）")
        a(f"   =  未变 {len(res['same'])} 份    ≠ 已更新 {len(res['changed'])} 份    "
          f"+ 新增 {len(res['added'])} 份    - 消失 {len(res['gone'])} 份")
        if res["same"] and not quiet:
            a("   未变（今日未更新，确认是否正常）：")
            for f in res["same"][:20]:
                a(f"      = {f}")
            if len(res["same"]) > 20:
                a(f"      …… 共 {len(res['same'])} 份")
        if res["added"]:
            a(f"   新增：{', '.join(res['added'][:20])}")
        if res["gone"]:
            a(f"   ⚠ 上一快照有、本快照没有：{', '.join(res['gone'][:20])}")
    else:
        a("③ 内容对比：无更早快照可比")

    # ④ 重复内容
    if res["dup_groups"]:
        hard = {id(g) for g in (res.get("dup_hard") or [])}
        n_hard = len(hard)
        head = f"④ {'⛔' if n_hard else '⚠'} 内容完全相同的文件组（{len(res['dup_groups'])} 组）"
        a(head + ("—— 其中 %d 组是库存报告，必然导错账号/站点：" % n_hard if n_hard
                  else "—— 可能重复下载，也可能本就同源："))
        for g in res["dup_groups"]:
            if id(g) in hard:
                a("      ⛔ " + "  ≡  ".join(g))
                a("         ⇒ 同一账号同一站点只会有一份库存报告，两份相同即导错。"
                  "请重导正确那份（主账号报告 SKU 数应远多于 interflow），覆盖后重跑")
            else:
                a("      " + "  ≡  ".join(g))
        if n_hard:
            a("      说明：订单报告不判硬伤 —— 中东大区一份报告存成沙特/阿联酋两个名字是合法形态")
    else:
        a("④ 无内容重复文件")

    # ⑤ 货件
    s = res["ship"]
    import shipments as Sh          # 仅用于 sid_label（货件号带地区说明）
    extra = f" + {s['pdf']} 个 PDF" if s.get("pdf") else ""
    a(f"⑤ 货件：{s['tsv']} 个 TSV{extra} → {s['tickets']} 票（按货件号）")
    if s["dups"]:
        a(f"   ⚠ 重复货件号 {len(s['dups'])} 个（同一票下载了多次，代码会自动去重）：")
        for k, v in s["dups"]:
            a(f"      {k}: {', '.join(v)}")
    orph = res.get("ship_orphan") or []
    if orph:
        a(f"   ⚠ 无对应簿的货件 {len(orph)} 票（不会落任何分析表，需业务确认）：")
        for o in orph:
            a(f"      {Sh.sid_label(o)} ← {o.get('folder','')}，{int(o.get('total') or 0)} 件")
        a("      ⇒ 已跳过不落列（interflow 账户的沙特站业务暂无统计出口；要统计需建 Interflowd 沙特簿）")
    rw = res.get("ship_region_warn") or []
    if rw:
        a(f"   ⚠ 文件夹与仓库地区不符 {len(rw)} 票（归属可能不对，请核对）：")
        for o in rw:
            a(f"      {Sh.sid_label(o)}：{o['warn']}")

    # ⑥ 货件列容量
    if res.get("capacity_err"):
        a(f"⑥ 货件列容量：!! 检查失败 {res['capacity_err']}")
    else:
        cap = res.get("capacity") or []
        low = [(b, s, f) for b, s, f in cap if f <= 2]
        if not cap:
            a("⑥ 货件列容量：未取得（工作簿被占用或读取失败）")
        elif low:
            a(f"⑥ ⚠ 货件列容量：{len(cap)} 张分析表中 {len(low)} 张余量 ≤2：")
            for b, s, f in low:
                a(f"      ⚠ {b} / {s}  仅剩 {f} 个空列位 —— 满了之后新货件无法落表（需插列扩容）")
        else:
            mn = min(cap, key=lambda x: x[2])
            a(f"⑥ 货件列容量：✅ {len(cap)} 张分析表全部余量充足（最少：{mn[0]} / {mn[1]} 剩 {mn[2]}）")

    # ⑦ 数组公式隐患
    if res.get("arr_risk_err"):
        a(f"⑦ 数组公式：!! 检查失败 {res['arr_risk_err']}")
    else:
        risk = res.get("arr_risk") or []
        if not risk:
            a("⑦ 数组公式：✅ 无「整列引用被数组公式破坏」的隐患")
        else:
            a(f"⑦ ⚠ 数组公式隐患：{len(risk)} 处会静默算成 0（VLOOKUP(A:A,…) 类）——")
            for b, s, col, cnt, sample in risk:
                a(f"      ✗ {b} / {s}  {col} 列 {cnt} 格   例: ={sample[:70]}")
            a("      修法：python pipeline/tools/maintenance/normalize_array_formulas_20261016.py --apply")

    # ⑧ 订单报告日期覆盖
    if res.get("coverage_run_err"):
        a(f"⑧ 订单报告日期覆盖：!! 检查失败 " + res["coverage_run_err"])
    else:
        import report_coverage as RC
        a(RC.render_snapshot(res.get("coverage") or []))

    # ⑨ 订单报告渠道错位
    if res.get("mislabel_run_err"):
        a("⑨ 订单报告渠道错位：!! 检查失败 " + res["mislabel_run_err"])
    else:
        import report_coverage as RC
        a(RC.render_channels(res.get("mislabel") or []))

    # ⑩ 销售报告 SKU 归属
    if res.get("sku_ov_run_err"):
        a("⑩ 销售报告 SKU 归属：!! 检查失败 " + res["sku_ov_run_err"])
    else:
        a(render_sku_overlap(res.get("sku_ov") or [], res.get("sku_ov_n") or 0))

    # ⑪ 分析表 ASIN 公式行号错位
    if res.get("asin_shift_err"):
        a("⑪ 分析表 ASIN 列：!! 检查失败 " + res["asin_shift_err"])
    elif res.get("asin_shift"):
        a(f"⑪ 分析表 ASIN 列：⚠ {len(res['asin_shift'])} 张表公式引用了别的行 "
          f"（会显示成别的 SKU 的 ASIN，或空）")
        for f, s, bad, rows, col, sample in res["asin_shift"]:
            a(f"      ✗ {f} / {s}  {col} 列 {bad}/{rows} 行错位   例: {sample}")
        a("      修法：把该列公式的查找值改回本行（XLOOKUP(A行号,…)），"
          "或 python pipeline/tools/maintenance/fix_sa_asin_ref_20261030.py")
    else:
        a(f"⑪ 分析表 ASIN 列：✓ {res.get('asin_shift_n') or 0} 张表公式行引用正常")

    # ⑫ 库存报告格式
    if res.get("stock_kind_err"):
        a("⑫ 库存报告格式：!! 检查失败 " + res["stock_kind_err"])
    elif res.get("stock_kind"):
        a(f"⑫ 库存报告格式：⚠ {len(res['stock_kind'])} 份下错了报表类型"
          f"（跑批会静默跳过该页，库存不更新）")
        for f, s, name, why in res["stock_kind"]:
            a(f"      ✗ {name}  ← {f} / {s}")
            a(f"          {why}")
        a("      修法：回卖家后台按「管理 FBA 库存」重下（不是库龄/仓储费报告），"
          "覆盖同名文件后重跑体检")
    else:
        a("⑫ 库存报告格式：✓ 表头类型与配置一致")

    a("=" * 100)
    if res["missing"]:
        a(f">>> 结论：缺 {len(res['missing'])} 份必需报告，**不建议现在跑**（缺报告的簿会被整份跳过）。")
    elif res["bad"]:
        a(f">>> 结论：有 {len(res['bad'])} 份报告异常，请先确认。")
    elif res.get("dup_hard"):
        a(f">>> 结论：⛔ 不要跑 —— {len(res['dup_hard'])} 组库存报告内容完全相同，"
          f"必然是导错账号/站点（同一账号同一站点只会有一份）：")
        for grp in res["dup_hard"]:
            a("          " + "  ≡  ".join(grp))
        a("          后果：runner 会把报告里没有的 SKU 库存整片清空 → 缺货/断货全部虚高。")
        a("          已前置拦截，工作簿一个字都没动；请重导正确那份覆盖后重跑"
          "（确要继续请加 --force）。")
        if res.get("sku_ov_err") or res.get("mislabel_err"):
            a(f"          另有 ⑨{res.get('mislabel_err') or 0} / ⑩{res.get('sku_ov_err') or 0} 处硬伤，一并处理。")
    elif res.get("sku_ov_err"):
        a(f">>> 结论：⛔ 不要跑 —— {res['sku_ov_err']} 份销售报告的 SKU 与本簿库存对不上"
          f"（导成了别的账号的同渠道报告）。")
        a("          已前置拦截，工作簿一个字都没动；先按 ⑩ 重导该报告再跑。")
        if res.get("mislabel_err"):
            a(f"          另有 {res['mislabel_err']} 份渠道整份错位（⑨），一并重导。")
        a("          确要继续（只跑没问题的簿）请加 --force。")
    elif res.get("mislabel_err"):
        a(f">>> 结论：⛔ 不要跑 —— {res['mislabel_err']} 份订单报告整份渠道错位（下错站点），"
          f"相关簿当日销售不会更新。")
        a("          已前置拦截，工作簿一个字都没动；先按正确站点重下该报告。"
          "确要继续请加 --force。")
    elif res.get("stock_kind"):
        a(f">>> 结论：⛔ 不要跑 —— {len(res['stock_kind'])} 份库存报告下错了报表类型"
          f"（如库龄报告当成管理库存报告），相关库存页会**静默不更新**，"
          f"缺货/断货会基于旧库存算错。")
        a("          已前置拦截，工作簿一个字都没动；按 ⑫ 重下正确报表后重跑。"
          "确要继续请加 --force。")
    elif res.get("coverage_err"):
        a(f">>> 结论：⚠ 可以跑。有 {res['coverage_err']} 份订单报告缺月初/月末时段"
          f"（date range 按本地时间设的）—— 日销簿已导入的不受影响；"
          f"但**拿该报告重建月度汇总会少这段**，下次导数请按 UTC 设或前后各多导一天。")
    elif res.get("locked"):
        a(f">>> 结论：⚠ 可以跑，但有 {len(res['locked'])} 本工作簿被占用会被跳过 —— 建议先关掉 WPS/Excel 再跑。")
    else:
        a(">>> 结论：✅ 可以跑。")
    a("=" * 100)
    return "\n".join(L)


# ---------------- ⑩ 渲染 ----------------
def render_sku_overlap(rows, n_checked=0):
    """体检第 ⑩ 块的文本。rows = sales_sku_overlap() 的 rows，n_checked = 判过的份数。"""
    if not rows:
        if not n_checked:
            return ("⑩ 销售报告 SKU 归属：⚠ **未能判定**（本簿库存读不到，或报告里没有 sku 列）"
                    " —— 别把它当绿灯，确认 ② 没占用、报告是亚马逊订单报告")
        return (f"⑩ 销售报告 SKU 归属：✅ {n_checked} 份销售报告的 SKU 都落在本簿库存里"
                "（渠道闸挡不住的「同名渠道换账号」已排除）")
    n_err = sum(1 for r in rows for lv, _t in r[6] if lv == "✗")
    if n_err:
        L = [f"⑩ ⛔ 销售报告 SKU 归属不符：{len(rows)} 份（硬伤 {n_err} 处，判过 {n_checked} 份）",
             "      —— 渠道对得上、账号不对（同名 sales-channel 两个卖家都有），runner 会整份拒收"]
    else:
        L = [f"⑩ ⚠ 销售报告 SKU 重合率偏低：{len(rows)} 份（判过 {n_checked} 份）"]
    for fname, sheet, rep, rate, hit, n_rep, iss in rows:
        L.append(f"      {fname} / {sheet} ← {rep}   重合 {hit}/{n_rep} = {rate:.0%}")
        for lv, txt in iss:
            L.append(f"        {lv + ' ' if lv else '  '}{txt}")
    if n_err:
        L.append("      修法：切到**本账号**的卖家后台，按同一站点/日期范围重下该报告，"
                 "覆盖同名文件后重跑（订单号去重，不会重复入库）")
    return "\n".join(L)


# ---------------- ⑥ 货件列容量 ----------------
def ship_capacity():
    """每张分析表的货件列空位余量。

    口径与 shipments.detect_sheet_cols / plan_columns 一致：
    行5 找「在途差额」定界（找不到则跳过该表）；
    空列位 = 行1 无货件号 且 (行5 空 或 行5 是占位公式)；**外加**行5 含「到货」的
    已完成货件列（2026-10-14 方案① 起可回收复用，故计入余量）。
    余量归零后新货件无法落表（runner 打「列位不足」、该票积压）。
    """
    import openpyxl
    out = []
    for wb_cfg in C.WORKBOOKS:
        if wb_cfg.get("skip"):
            continue
        path = os.path.join(C.WORK_DIR, wb_cfg["file"])
        if not os.path.exists(path) or _locked(path):
            continue
        try:
            wb = openpyxl.load_workbook(path, read_only=True)
        except Exception:
            continue
        try:
            for sn in [s for s in wb.sheetnames if "分析" in s]:
                ws = wb[sn]
                row1, row5, diff_col = {}, {}, None
                for ri, row in enumerate(ws.iter_rows(min_row=1, max_row=5, min_col=9, max_col=60), start=1):
                    for ci, c in enumerate(row, start=9):   # read_only 下 EmptyCell 无 row/col 属性，用位置
                        v = getattr(c, "value", None)
                        if ri == 1:
                            row1[ci] = v
                        elif ri == 5:
                            row5[ci] = v
                            if isinstance(v, str) and "在途差额" in v:
                                diff_col = ci
                if diff_col is None:
                    continue
                free = 0
                for cn in range(9, diff_col):
                    v1, v5 = row1.get(cn), row5.get(cn)
                    has_id = v1 is not None and str(v1).strip() != ""   # 货件号可能是 int（美客多）
                    arrived = isinstance(v5, str) and "到货" in v5
                    if arrived:
                        free += 1         # 已到货旧列 → 可回收复用（方案①）
                    elif not has_id and (v5 is None or (isinstance(v5, str) and v5.startswith("="))):
                        free += 1         # 真空白列位（行1 无号 且 行5 空或占位公式）
                out.append((wb_cfg["file"], sn.strip(), free))
        finally:
            wb.close()
    return out


# ---------------- ⑦ 数组公式隐患 ----------------
# 整列引用（列号后不跟数字，排除 $A$3:$A$30999 / I7:R7）
_COL = r"\$?[A-Z]{1,3}:\$?[A-Z]{1,3}(?![0-9A-Z])"
_ARR_HARM = [
    # 整列引用当「查找值」：数组语义下退化成用第 1 行去查 → VLOOKUP 全 0
    re.compile(r"(?:VLOOKUP|HLOOKUP|XLOOKUP|MATCH|LOOKUP)\s*\(\s*" + _COL, re.I),
    # 整列引用当「条件参数」
    re.compile(r"(?:SUMIF|SUMIFS|COUNTIF|COUNTIFS)\s*\(\s*[^,()]+,\s*" + _COL, re.I),
]
_CELL_RE = re.compile(r'<c r="([A-Z]+)(\d+)"([^>]*)>(.*?)</c>', re.S)


def array_formula_risk():
    """扫各簿分析表：返回「数组公式里整列引用被当标量用」的隐患列表。

    2026-10-16 阿联酋簿实测：F/G 被存成 `t="array"` 的
      `IFERROR(VLOOKUP(A:A,阿联酋在途库存!A:Y,COLUMN(...$W$3),0),0)`
    数组语义下 `A:A` 是整列数组，单元格只取第 1 个元素（等于拿 A1 去查）→ #N/A
    → IFERROR → **整列 0**。同一公式在别的簿是普通公式就正常。
    """
    import zipfile
    out = []
    for wb_cfg in C.WORKBOOKS:
        if wb_cfg.get("skip"):
            continue
        path = os.path.join(C.WORK_DIR, wb_cfg["file"])
        if not os.path.exists(path):
            continue
        try:
            with zipfile.ZipFile(path) as z:
                wbx = z.read("xl/workbook.xml").decode("utf-8", "ignore")
                rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8", "ignore")
                rmap = {}
                for m in re.finditer(r"<Relationship\b[^>]*>", rels):
                    t = m.group(0)
                    i, g = re.search(r'Id="([^"]*)"', t), re.search(r'Target="([^"]*)"', t)
                    if i and g:
                        rmap[i.group(1)] = g.group(1)
                for m in re.finditer(r"<sheet\b[^>]*>", wbx):
                    t = m.group(0)
                    nm, ri = re.search(r'name="([^"]*)"', t), re.search(r'r:id="([^"]*)"', t)
                    if not (nm and ri) or "分析" not in nm.group(1):
                        continue
                    tg = rmap.get(ri.group(1), "").lstrip("/")
                    part = tg if tg.startswith("xl/") else "xl/" + tg
                    try:
                        xml = z.read(part).decode("utf-8", "ignore")
                    except KeyError:
                        continue
                    per_col = {}
                    for cm in _CELL_RE.finditer(xml):
                        col, r, attrs, body = cm.group(1), int(cm.group(2)), cm.group(3), cm.group(4)
                        if r < 7:
                            continue
                        fm = re.search(r'<f([^>]*)>(.*?)</f>', body, re.S)
                        if not fm or 't="array"' not in fm.group(1):
                            continue
                        text = fm.group(2)
                        for rx in _ARR_HARM:
                            if rx.search(text):
                                per_col.setdefault(col, [0, text])[0] += 1
                                break
                    for col, (cnt, sample) in sorted(per_col.items()):
                        out.append((wb_cfg["file"], nm.group(1).strip(), col, cnt, sample))
        except Exception:
            continue
    return out


def asin_ref_shift():
    """⑪ 分析表 ASIN 列「公式行号错位」。

    2026-10-30 事故（Interflowd沙特簿）：ASIN 列公式是
      =XLOOKUP(A42, 沙特在途库存!A:A, 沙特在途库存!C:C)   ← 却写在**第 7 行**
    派生建簿时把源簿的公式原样抄来，行引用没重写。后果：
      · 引用到空行 → 查出 0 / 空（表现为「ASIN 大量为空」）
      · 引用到别人 → 显示成**别的 SKU 的 ASIN**（G583 显示成 G587 的）
    两种情况都不报错、其它数字也不异常，肉眼极难发现 —— 所以进体检。

    返回 (out, checked)：out=[(簿文件, 分析表, 错位数, 数据行数, 单元格列, 样例)]，
    checked=真正核过的分析表张数（用于区分「查过没问题」和「压根没找到 ASIN 列」）。
    """
    import zipfile
    RX = re.compile(r"(?:XLOOKUP|VLOOKUP)\(\s*\$?A\$?(\d+)", re.I)
    out = []
    checked = 0
    for wb_cfg in C.WORKBOOKS:
        if wb_cfg.get("skip"):
            continue
        path = os.path.join(C.WORK_DIR, wb_cfg["file"])
        if not os.path.exists(path):
            continue
        try:
            with zipfile.ZipFile(path) as z:
                wbx = z.read("xl/workbook.xml").decode("utf-8", "ignore")
                rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8", "ignore")
                rmap = {}
                for m in re.finditer(r"<Relationship\b[^>]*>", rels):
                    t = m.group(0)
                    i, g = re.search(r'Id="([^"]*)"', t), re.search(r'Target="([^"]*)"', t)
                    if i and g:
                        rmap[i.group(1)] = g.group(1)
                ss = []
                if "xl/sharedStrings.xml" in z.namelist():
                    sx = z.read("xl/sharedStrings.xml").decode("utf-8", "ignore")
                    for m in re.finditer(r"<si>(.*?)</si>", sx, re.S):
                        ss.append("".join(re.findall(r"<t[^>]*>(.*?)</t>", m.group(1), re.S)))
                for m in re.finditer(r"<sheet\b[^>]*>", wbx):
                    t = m.group(0)
                    nm, ri = re.search(r'name="([^"]*)"', t), re.search(r'r:id="([^"]*)"', t)
                    if not (nm and ri) or "分析" not in nm.group(1):
                        continue
                    tg = rmap.get(ri.group(1), "").lstrip("/")
                    part = tg if tg.startswith("xl/") else "xl/" + tg
                    try:
                        xml = z.read(part).decode("utf-8", "ignore")
                    except KeyError:
                        continue

                    # 1) 找 ASIN 表头所在的列（前 8 行内）
                    asin_col = None
                    for cm in _CELL_RE.finditer(xml):
                        col, r, _attrs, body = cm.group(1), int(cm.group(2)), cm.group(3), cm.group(4)
                        if r > 8:
                            continue
                        txt = ""
                        tm = re.search(r"<t[^>]*>(.*?)</t>", body, re.S)
                        vm = re.search(r"<v>(\d+)</v>", body)
                        if tm:
                            txt = tm.group(1)
                        elif vm and 't="s"' in _attrs:
                            k = int(vm.group(1))
                            if k < len(ss):
                                txt = ss[k]
                        if txt.strip().upper() == "ASIN":
                            asin_col = col
                            break
                    if not asin_col:
                        continue
                    checked += 1

                    # 2) 逐行核对该列的查找值引用
                    rows = bad = 0
                    sample = ""
                    for cm in _CELL_RE.finditer(xml):
                        col, r, _attrs, body = cm.group(1), int(cm.group(2)), cm.group(3), cm.group(4)
                        if col != asin_col or r < 7:
                            continue
                        fm = re.search(r"<f[^>]*>(.*?)</f>", body, re.S)
                        if not fm:
                            continue
                        rows += 1
                        q = RX.search(fm.group(1))
                        if q and int(q.group(1)) != r:
                            bad += 1
                            if not sample:
                                sample = "r%d→A%s" % (r, q.group(1))
                    if rows and bad:
                        out.append((wb_cfg["file"], nm.group(1).strip(), bad, rows, asin_col, sample))
        except Exception:
            continue
    return out, checked


# ---------------- ⑫ 库存报告格式（下错报表类型） ----------------
def _read_head(path):
    """读 csv 表头 + 数据行数。按站点首选编码（日文/澳洲是 shift_jis）。

    只看 ASCII 特征（inv-age- 等），编码解成乱码也不影响判断。
    """
    encs = list(getattr(C, "ENC_CANDIDATES", ["utf-8-sig"]))
    for site, enc in (getattr(C, "REPORT_ENC", None) or {}).items():
        if site in os.path.basename(path):
            if enc in encs:
                encs.remove(enc)
            encs.insert(0, enc)
            break
    for enc in encs:
        try:
            with open(path, encoding=enc, newline="") as f:
                head = f.readline()
                nrows = sum(1 for ln in f if ln.strip())
            return head, nrows
        except UnicodeDecodeError:
            continue
    return "", 0


def stock_kind_check(snap_dir, day):
    """库存类报告「下错了报表类型」——体检原本查不出，runner 只会静默跳过该项。

    2026-11-02 实例：阿联酋主账号把「管理库存报告」下成了「库龄报告」（90 列、
    表头含 inv-age-），解析出 0 行 → 阿联酋可售/在途两页都不更新，而体检全绿。

    返回 [(簿文件, sheet, 报告名, 问题说明)]，按报告名去重。
    """
    out, seen = [], set()
    for wb in C.WORKBOOKS:
        for st in (wb.get("stocks") or []):
            rep, kind = st.get("report"), st.get("kind")
            if not rep or not kind:
                continue
            # report 可能是候选列表（如 interflow 中东、美客多），逐个试到命中为止
            reps = rep if isinstance(rep, (list, tuple)) else [rep]
            name = None
            for r in reps:
                cand = _fmt(r, day)
                if "*" in cand or "?" in cand:
                    hits = sorted(glob.glob(os.path.join(snap_dir, cand)))
                    cand = os.path.basename(hits[-1]) if hits else cand
                if os.path.isfile(os.path.join(snap_dir, cand)):
                    name = cand
                    break
            if not name or name in seen:
                continue
            p = os.path.join(snap_dir, name)
            head, nrows = _read_head(p)
            low = head.lower()
            is_age = "inv-age-" in low
            if kind in ("inv", "balance") and is_age:
                seen.add(name)
                out.append((wb["file"], st.get("sheet", ""), name,
                            "下成了库龄报告（表头含 inv-age-，应为管理库存格式）"))
            elif kind == "age" and not is_age:
                seen.add(name)
                out.append((wb["file"], st.get("sheet", ""), name,
                            "应为库龄报告，但表头没有 inv-age-"))
            elif nrows == 0:
                seen.add(name)
                out.append((wb["file"], st.get("sheet", ""), name,
                            "只有表头、没有数据行"))
    return out


def main(argv):
    quiet = "--quiet" in argv
    snap = None
    if "--snap" in argv:
        i = argv.index("--snap")
        if i + 1 < len(argv):
            snap = argv[i + 1]
            argv = argv[:i] + argv[i + 2:]
    argv = [x for x in argv if not x.startswith("--")]
    day = argv[0] if argv else None
    if snap:
        snap = os.path.abspath(snap)
        day = day or os.path.basename(snap.rstrip("\\/"))
    res = check(day, snap_dir=snap)
    print(render(res, quiet=quiet))
    # 1 = 不可跑：缺报告，或闸门硬伤（⑨ 渠道整份错位 / ⑩ SKU 归属不符 / ④ 库存报告同内容）
    return 1 if (res["missing"] or res.get("gate_err")) else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
