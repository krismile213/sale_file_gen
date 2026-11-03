# -*- coding: utf-8 -*-
"""快照齐全性体检：按 config.WORKBOOKS 逐本簿列出所需报告，检查某快照日是否齐全、格式是否正确。

★ 日常「跑前体检」已由 `pipeline/core/precheck.py` 承担（run.py 默认自动调用，含内容比对/
  重复检测/货件票数）。本脚本保留用于**历史快照的深度体检**——它独有「EU 旧源回落探测」
  （2026-10-13 前的快照只有 欧洲库存{d4}.csv 旧源，会标 ~）。

用法：
    python pipeline/tools/check_snapshot.py 20261012 [20261013 ...]
不带参数则体检 config.latest_snapshot() 指向的最新快照。

注意：报告的「应有文件名」随 config 的源定义走。EU 库存源 2026-10-13 已切换
（欧洲可售+在途库存 ← 快照内 亚马逊欧盟管理库存报告{d}.csv），此前的快照里只有旧源
（DSR 根 欧洲库存{d4}.csv，93 列 age 式）。体检旧快照时新报告必缺，本脚本会按旧源
回落探测，命中标 ~ —— 属源约定变更，不算数据缺失。
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
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import config as C

# EU 源切换回落候选：键=新源模板首项，值=旧源候选（先快照目录、再 DSR 根）
LEGACY_FALLBACK = {
    "亚马逊欧盟管理库存报告{d}.csv": ("欧洲库存{d4}.csv", "欧洲库存.csv"),
}


def resolve(snap_dir, files, day):
    """按候选顺序取第一个存在的报告，返回 (文件名, 绝对路径) 或 (None, None)。"""
    if isinstance(files, str):
        files = [files]
    for f in files:
        try:
            nm = f.format(m=C.ORDER_MONTH, d=day, d4=day[4:])
        except (KeyError, IndexError):
            nm = f
        cand = nm if os.path.isabs(nm) else os.path.join(snap_dir, nm)
        if os.path.exists(cand):
            return nm, cand
    return None, None


def probe(path):
    """返回 (ok:True/False/None, 说明)。"""
    if path is None:
        return False, "缺失"
    ext = os.path.splitext(path)[1].lower()
    if ext == ".xlsx":
        return None, f"xlsx(未解析, {os.path.getsize(path)}B)"
    try:
        with open(path, "rb") as fh:
            raw = fh.read(300000)
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
    head = (txt.splitlines() or [""])[0]
    low = head.lower()
    if "afn-" in low:
        return True, "inv(afn-*)"
    if "amazon-order-id" in low:
        return True, f"订单报告(tab, ≥{len(txt.splitlines())}行/首300KB)"
    if "seller-sku" in low:
        return True, "multi(多国库存)"
    if "inv-age" in low or "snapshot-date" in low:
        return True, "age(库龄式)"
    if "reserved_qty" in low:
        return True, "reserved(预留式)"
    if "\t" in head or "," in head:
        return None, "未识别列头: " + head[:60]
    return None, "未识别格式"


def check(day):
    snap_dir = os.path.join(C.SNAP_BASE, C.CURRENT_PERIOD, "快照日期", day)
    print("=" * 104)
    print(f"快照 {day}\n  {snap_dir}")
    if not os.path.isdir(snap_dir):
        print("  ✗ 目录不存在")
        return -1
    print("-" * 104)
    missing = 0
    for wb in C.WORKBOOKS:
        if wb.get("skip"):
            continue
        print(f"[{wb['file']}]")
        items = list(wb.get("stocks", []))
        sales = wb.get("sales")
        if sales:
            items += sales if isinstance(sales, list) else [sales]
        for it in items:
            nm, p = resolve(snap_dir, it.get("report"), day)
            ok, note = probe(p)
            if nm is None:
                # 旧快照回落：新源缺失时按旧源候选探测（先快照目录、再 DSR 根）
                reps = it.get("report")
                first = reps[0] if isinstance(reps, list) else reps
                for fb in LEGACY_FALLBACK.get(first, ()):
                    fnm, fp = resolve(snap_dir, fb, day)
                    if fnm is None:
                        fnm, fp = resolve(C.DSR, fb, day)
                    if fnm is not None:
                        ok2, note2 = probe(fp)
                        nm, p, ok, note = fnm, fp, ok2, "回落旧源: " + note2
                        break
            flag = "OK "
            if nm is None:
                missing += 1
                flag = "缺 "
            elif ok is False:
                flag = "!! "
            elif note.startswith("回落旧源"):
                flag = "~  "
            rep = it.get("report")
            rep = rep if isinstance(rep, str) else " / ".join(rep)
            print(f"    {flag} {it['sheet']:22} <- {rep:56} {note}")
    print(f"  >>> 缺失报告数 = {missing}")
    return missing


if __name__ == "__main__":
    days = sys.argv[1:]
    if not days:
        _, day = C.latest_snapshot()
        days = [day]
    for d in days:
        check(d)
