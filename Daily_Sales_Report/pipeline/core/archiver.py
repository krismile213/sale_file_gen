# -*- coding: utf-8 -*-
"""每日跑批后自动归档：12 本工作簿按「年份/月份/天数」复制留档。

目录结构（与既有月归档同构，月份目录下按「日」建子目录）：
    1.每日销售数据表/2026年每日数据/2026年10月/10月15日/2026年10月15日-XX库存销售表.xlsx

与 month_rollover「月末归档」的分工（两者互不覆盖，路径不同）：
  - 月末归档 = 窗口裁剪 + 锚点定格的**月封存源**，供次月 resolve_month_sales 读取；
  - 本模块   = 跑批当日**原样快照**，供回溯/对账/误删恢复，月内每天一份。

幂等：同日重跑且内容未变 → 跳过；内容有变 → 覆盖为最新机器版。
防误覆盖（2026-10-20 事故后新增）：归档目录是用户在用的**手工台账**
    （标「到货」、删历史货件列）。若归档件在本模块写入后被外部改过，
    先把旧件另存为 `YYYY年M月D日-XX库存销售表-手工版-HHMMSS.xlsx`，
    再写入机器版——两条线都不丢。
    判定依据 = `Daily_Sales_Report/.workbuddy/archive_state/{day}.json` 里登记的
    md5；与归档件当前 md5 不一致（或该历史归档尚无登记）即视为被动过，
    只留存、不静默覆盖。`--force-overwrite` 可强制机器版覆盖一切。
安全：复制前 testzip 校验源簿完整性；跳过 ~$ 锁文件；失败不影响其它簿。
"""
import datetime
import hashlib
import json
import os
import shutil
import sys
import zipfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import config as C  # noqa: E402


def day_dir_name(day):
    """'20261015' → '10月15日'（与既有「2026年10月」月目录同中文体系）。"""
    return f"{int(day[4:6])}月{int(day[6:8])}日"


def _state_dir(work):
    """归档登记簿目录：Daily_Sales_Report/.workbuddy/archive_state/。"""
    return os.path.join(os.path.dirname(os.path.abspath(work)),
                        ".workbuddy", "archive_state")


def _md5(path):
    h = hashlib.md5()
    with open(path, "rb") as fh:
        for chunk in iter(lambda: fh.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def _load_state(work, day):
    """读当日归档登记簿；不存在/损坏一律返回空 dict（按未登记处理）。"""
    p = os.path.join(_state_dir(work), f"{day}.json")
    try:
        with open(p, encoding="utf-8") as fh:
            return json.load(fh)
    except Exception:
        return {}


def _save_state(work, day, state):
    d = _state_dir(work)
    os.makedirs(d, exist_ok=True)
    p = os.path.join(d, f"{day}.json")
    tmp = p + ".tmp"
    with open(tmp, "w", encoding="utf-8") as fh:
        json.dump(state, fh, ensure_ascii=False, indent=1, sort_keys=True)
    os.replace(tmp, p)


def _rec(path):
    return {"md5": _md5(path), "size": os.path.getsize(path),
            "ts": datetime.datetime.now().isoformat(timespec="seconds")}


def archive_day(day=None, work_dir=None, quiet=False, force_overwrite=False):
    """把 WORK_DIR 根目录全部 *.xlsx 复制进「年/月/日」归档目录。

    day：8 位日期字符串（默认取 config 最新快照日）。
    force_overwrite：True 时机器版无条件覆盖，不留存手工版（逃逸阀）。
    返回 (ok_list, fail_list)；ok_list 元素 = (文件名, 字节数)。
    另外把「未变跳过 / 手工版留存」明细写进模块级变量，供 run.py 汇报。
    """
    day = (day or C._LATEST_DAY)
    if not (len(day) == 8 and day.isdigit()):
        raise SystemExit(f"[归档] 日期格式错误: {day!r}（应为 YYYYMMDD）")
    work = work_dir or C.WORK_DIR
    y, m, d = day[:4], int(day[4:6]), int(day[6:8])
    dest = os.path.join(work, f"{y}年每日数据", f"{y}年{m}月", day_dir_name(day))
    os.makedirs(dest, exist_ok=True)

    files = sorted(f for f in os.listdir(work)
                   if f.endswith(".xlsx") and not f.startswith("~$"))
    if not files:
        print(f"  [归档] {work} 下没有 .xlsx，无事可做")
        return [], []
    ok, fail = [], []
    kept, preserved = [], []
    state = _load_state(work, day)
    total = 0
    for f in files:
        src = os.path.join(work, f)
        dst = os.path.join(dest, f"{y}年{m}月{d}日-{f}")
        try:
            with zipfile.ZipFile(src) as z:
                bad = z.testzip()
            if bad:
                raise IOError(f"zip 损坏成员 {bad}")
            if not os.path.exists(dst):
                shutil.copy2(src, dst)
                state[f] = _rec(dst)
                sz = os.path.getsize(dst)
                total += sz
                ok.append((f, sz))
                if not quiet:
                    print(f"    [归档] {f} → {os.path.relpath(dst, work)} ({sz/1048576:.1f}MB)")
                continue
            # ---- 目标已存在：判断是被手工改过还是重跑覆盖 ----
            dst_md5, src_md5 = _md5(dst), _md5(src)
            if dst_md5 == src_md5:
                state.setdefault(f, _rec(dst))
                kept.append(f)
                continue
            rec = state.get(f) or {}
            registered = "md5" in rec
            # 登记簿里没有这本 = 保护升级前的历史归档（可能正是手工维护过的），
            # 一律按「被动过」处理：只留存、不静默覆盖。
            manually_touched = (not registered) or (dst_md5 != rec["md5"])
            if manually_touched and not force_overwrite:
                # 手工改过 → 先留存旧件，再写机器版
                stamp = datetime.datetime.now().strftime("%H%M%S")
                alt_name = f"{y}年{m}月{d}日-{f[:-5]}-手工版-{stamp}.xlsx"
                alt = os.path.join(dest, alt_name)
                shutil.copy2(dst, alt)          # copy2 保留手工件的 mtime
                shutil.copy2(src, dst)
                state[f] = _rec(dst)
                preserved.append(alt_name)
                if registered:
                    print(f"    [归档] !! {f} 归档件被手工改过 → 留存 {alt_name}，"
                          f"机器版照常写入")
                else:
                    print(f"    [归档] !! {f} 归档件无登记（保护升级前历史归档）"
                          f" → 留存 {alt_name}，机器版照常写入")
                ok.append((f, os.path.getsize(dst)))
                continue
            # 登记一致 = 我们写的机器版没人动过 → 直接覆盖
            shutil.copy2(src, dst)
            state[f] = _rec(dst)
            sz = os.path.getsize(dst)
            total += sz
            ok.append((f, sz))
            if not quiet:
                print(f"    [归档] {f} → {os.path.relpath(dst, work)} ({sz/1048576:.1f}MB)")
        except Exception as e:  # 单本失败不阻断其它簿
            fail.append((f, repr(e)))
            print(f"    [归档] !! {f} 失败 {e!r}")
    _save_state(work, day, state)
    rel = os.path.relpath(dest, work)
    msg = f"  [归档] {day}：{len(ok)}/{len(files)} 本入库（共 {total/1048576:.1f}MB）→ {rel}"
    if kept:
        msg += f"；未变跳过 {len(kept)} 本"
    if preserved:
        msg += f"；手工版留存 {len(preserved)} 份"
    if fail:
        msg += f"；失败：{[f for f, _ in fail]}"
    print(msg)
    return ok, fail


def main(argv):
    import argparse
    ap = argparse.ArgumentParser(description="按 年/月/日 归档 12 本日销簿")
    ap.add_argument("--day", default=None, help="YYYYMMDD，默认取最新快照日")
    ap.add_argument("--quiet", action="store_true")
    ap.add_argument("--force-overwrite", action="store_true",
                    help="机器版无条件覆盖归档件（不保留手工版）")
    args = ap.parse_args(argv)
    _ok, fail = archive_day(day=args.day, quiet=args.quiet,
                            force_overwrite=args.force_overwrite)
    return 1 if fail else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
