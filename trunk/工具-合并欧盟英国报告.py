# -*- coding: utf-8 -*-
"""
工具-合并欧盟英国报告.py
====================================================================
亚马逊《欧盟站点》与《英国站点》两份订单报告**都是全欧导出**（同一批订单），
但两次导出的**截止时刻不同**，会各自漏掉一个很小的边界时段：

    欧盟版常见止于 10-31 22:58 UTC  → 漏掉 23:00~23:59 UTC
    英国版常见止于 10-31 23:50 UTC  → 漏掉 09-30 22:00~23:00 UTC

而脚本「欧洲」组是按 **Europe/London** 归月的（见 month_sale_file.py L159，
实测该报告里的时间就是英国时间，不是 CET）。所以：

    UTC 10-31 23:xx  →  伦敦 10-31 23:xx  →  **属于 10 月**
    UTC 09-30 22:xx  →  伦敦 09-30 23:xx  →  属于 9 月（应排除）

⇒ 只读欧盟版会**静默丢掉每月最后 1 小时的英国单**（202610 实测 33 单 / 31 件）。

本工具把「英国版有、欧盟版没有、且归属月 = 本期」的行，按**原始文本行**追加进欧盟版，
不动任何既有行、不用 pandas 重写（避免制表符/编码/尾部空格列名被改写）。

用法：
    python 工具-合并欧盟英国报告.py 202610 20261107          # 执行合并（幂等）
    python 工具-合并欧盟英国报告.py 202610 20261107 --check   # 只体检，不写文件

幂等判据：按 order-item-id 判重。重复执行不会重复追加。
备份：backup/pre_eu_uk_merge_<YYYYMMDD>/亚马逊欧盟站点<期别>订单报告.txt（首次执行时留一份）
====================================================================
"""
import io
import os
import shutil
import sys

TRUNK_DIR = os.path.dirname(os.path.abspath(__file__))
KEY_COL = "order-item-id "      # ★ 亚马逊导出列名尾部自带一个空格，勿删
TZ_EUROPE = "Europe/London"     # ★ 与 month_sale_file.py L159 保持一致


def _read_lines(path):
    with io.open(path, "r", encoding="utf-8-sig", newline="") as f:
        return f.read().split("\n")


def _key_index(header_line):
    cols = header_line.rstrip("\r").split("\t")
    if KEY_COL in cols:
        return cols.index(KEY_COL)
    stripped = [c.strip() for c in cols]
    if KEY_COL.strip() in stripped:
        return stripped.index(KEY_COL.strip())
    raise AssertionError("报告里找不到列 %r，请确认导出格式" % KEY_COL.strip())


def _month_of(purchase_date_str, tzname):
    """UTC 时间串 → 目标时区的 YYYYMM；解析失败返回 None。"""
    try:
        from datetime import datetime
        from zoneinfo import ZoneInfo
        s = purchase_date_str.strip().replace("Z", "+00:00")
        dt = datetime.fromisoformat(s)
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=ZoneInfo("UTC"))
        return dt.astimezone(ZoneInfo(tzname)).strftime("%Y%m")
    except Exception:
        return None


def main(argv):
    if len(argv) < 3:
        print(__doc__)
        return 1
    period, snapshot = argv[1], argv[2]
    check_only = "--check" in argv

    snap_dir = os.path.join(TRUNK_DIR, "data", "数据归属周期", period, "快照日期", snapshot)
    eu_path = os.path.join(snap_dir, "亚马逊欧盟站点%s订单报告.txt" % period)
    uk_path = os.path.join(snap_dir, "亚马逊英国站点%s订单报告.txt" % period)

    for p in (eu_path, uk_path):
        if not os.path.exists(p):
            print("★ 缺文件：%s" % p)
            return 2

    eu_lines = _read_lines(eu_path)
    uk_lines = _read_lines(uk_path)
    hdr = eu_lines[0].rstrip("\r")

    if uk_lines[0].rstrip("\r") != hdr:
        raise AssertionError("两份报告表头不一致，拒绝合并（请确认是同一套导出模板）")

    ki = _key_index(hdr)
    pd_i = [i for i, c in enumerate(hdr.split("\t")) if c.strip() == "purchase-date"]
    pd_i = pd_i[0] if pd_i else None
    status_i = [i for i, c in enumerate(hdr.split("\t")) if c.strip() == "order-status"]
    status_i = status_i[0] if status_i else None

    eu_keys = set()
    for ln in eu_lines[1:]:
        if not ln.strip():
            continue
        parts = ln.rstrip("\r").split("\t")
        if len(parts) > ki:
            eu_keys.add(parts[ki].strip())

    add_lines, added_months, skipped_month, added_cancelled = [], {}, 0, 0
    for ln in uk_lines[1:]:
        if not ln.strip():
            continue
        parts = ln.rstrip("\r").split("\t")
        if len(parts) <= ki:
            continue
        k = parts[ki].strip()
        if k in eu_keys:
            continue
        m = _month_of(parts[pd_i], TZ_EUROPE) if pd_i is not None else None
        if m is not None and m != period:
            skipped_month += 1
            continue
        add_lines.append(ln)
        added_months[m] = added_months.get(m, 0) + 1
        if status_i is not None and parts[status_i].strip() == "Cancelled":
            added_cancelled += 1

    print("期别 %s / 快照 %s（归月基准：%s）" % (period, snapshot, TZ_EUROPE))
    print("  欧盟版数据行 %d（唯一键 %d）" % (len(eu_lines) - 1, len(eu_keys)))
    print("  英国版数据行 %d" % (len(uk_lines) - 1))
    print("  → 待补行数 %d；因不属本期跳过 %d 行（上月尾单，不并入）" % (len(add_lines), skipped_month))
    if add_lines:
        print("  → 补入行的归属月分布 %s（其中 Cancelled %d 行）" % (added_months, added_cancelled))

    if not add_lines:
        print("\n✅ 已是并集状态，无需改动")
        return 0

    if check_only:
        print("\n[--check] 仅体检，未写文件")
        return 0

    bak_dir = os.path.join(TRUNK_DIR, "backup", "pre_eu_uk_merge_%s" % snapshot)
    os.makedirs(bak_dir, exist_ok=True)
    bak = os.path.join(bak_dir, os.path.basename(eu_path))
    if not os.path.exists(bak):
        shutil.copy2(eu_path, bak)
        print("\n已备份原始欧盟报告 → %s" % bak)

    body = "\n".join(eu_lines).rstrip("\r\n")
    out = body + "\n" + "\n".join(add_lines) + "\n"
    with io.open(eu_path, "w", encoding="utf-8", newline="") as f:
        f.write(out)
    print("✅ 已合并：+%d 行 → 现共 %d 数据行" % (len(add_lines), len(eu_lines) - 1 + len(add_lines)))
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
