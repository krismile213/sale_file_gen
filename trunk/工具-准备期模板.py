# -*- coding: utf-8 -*-
"""
工具-准备期模板.py —— 换月第一步：把上一期成品直接复制成这一期的模板

背景（2026-11-02 用户拍板）
    四张模板（月销售总表 / 东南亚 / 欧洲 / 北美）**一律复制上一期成品**，不再另找模板、
    也不用 make_eu_na_templates.py 生成：
      · 复制当月成品 → 下月列块天然是空的、列分组/样式/迷你图/条件格式全部继承；
      · 复制**上上期**会继承 SEA Total 错位级联（历史踩过），所以源期默认锁死上一期。

用法
    python 工具-准备期模板.py 202611 20261201            # 从 202610 成品复制到 202611 快照目录
    python 工具-准备期模板.py 202611 20261201 --dry-run  # 只看会拷什么，不落盘
    python 工具-准备期模板.py 202611 20261201 --force    # 目标已存在且内容不同时强制覆盖
    python 工具-准备期模板.py 202701 20270105 --src 202612   # 跨年时文件名年份自动改写

落点（与四个生成脚本的读取路径严格一致）
    data/数据归属周期/<期>/快照日期/<快照日>/月销售总表.xlsx                     ← result/数据归属周期/<源期>/
    data/数据归属周期/<期>/快照日期/<快照日>/东南亚/2026年东南亚销量表….xlsx      ← result/东南亚销量表/<源期>/
    data/数据归属周期/<期>/快照日期/<快照日>/欧洲/2026年欧洲销量表….xlsx          ← result/欧洲销量表/<源期>/
    data/数据归属周期/<期>/快照日期/<快照日>/北美/2026年北美销量表….xlsx          ← result/北美销量表/<源期>/

行为约定
    · 只新增、只复制，**从不删除**任何文件；
    · 幂等：目标已存在且字节一致 → 跳过；已存在但不同 → 默认不覆盖并告警（--force 才覆盖）；
    · 文件名里的「YYYY年」按**目标期年份**改写（跨年自动处理），其余原样保留；
    · 顺手建好空目录 东南亚/跨境/、<期>/库龄/<快照日>/。

退出码
    0 = 四张模板全部就位（含本次跳过）    2 = 有缺失/有冲突未覆盖
"""
import argparse
import datetime
import hashlib
import os
import re
import shutil
import sys

TRUNK_DIR = os.path.dirname(os.path.abspath(__file__))

# (说明, 目标子目录（相对快照目录，"" = 快照根）, 源 result 子目录, 目标文件名正则/字面量)
TEMPLATES = [
    ("月销售总表", "", ("数据归属周期",), "月销售总表.xlsx"),
    ("东南亚总表", "东南亚", ("东南亚销量表",), r"^\d{4}年东南亚销量表.*\.xlsx$"),
    ("欧洲总表", "欧洲", ("欧洲销量表",), r"^\d{4}年欧洲销量表.*\.xlsx$"),
    ("北美总表", "北美", ("北美销量表",), r"^\d{4}年北美销量表.*\.xlsx$"),
]


def _md5(path, chunk=1 << 20):
    h = hashlib.md5()
    with open(path, "rb") as f:
        while True:
            b = f.read(chunk)
            if not b:
                break
            h.update(b)
    return h.hexdigest()


def _prev_period(period):
    y, m = int(period[:4]), int(period[4:])
    m -= 1
    if m == 0:
        y, m = y - 1, 12
    return f"{y:04d}{m:02d}"


def _rename_year(filename, target_year):
    """把文件名开头/中间的「YYYY年」改成目标年份（只动 4 位年份+年，其余原样）"""
    if re.match(r"^\d{4}年", filename):
        return re.sub(r"^\d{4}年", f"{target_year}年", filename, count=1)
    return filename


def _pick_source(src_dir, pattern):
    """在源目录里按正则/字面量找唯一文件；字面量优先精确匹配"""
    if not os.path.isdir(src_dir):
        return None, f"源目录不存在：{src_dir}"
    if not pattern.startswith("^"):
        p = os.path.join(src_dir, pattern)
        return (p, None) if os.path.isfile(p) else (None, f"源文件不存在：{p}")
    hits = sorted(f for f in os.listdir(src_dir) if re.match(pattern, f))
    hits = [f for f in hits if not f.startswith("~$")]
    if not hits:
        return None, f"源目录里没有匹配 {pattern} 的文件：{src_dir}"
    if len(hits) > 1:
        hits.sort(key=lambda f: os.path.getmtime(os.path.join(src_dir, f)), reverse=True)
        return os.path.join(src_dir, hits[0]), f"源目录有 {len(hits)} 个匹配，取最新：{hits[0]}"
    return os.path.join(src_dir, hits[0]), None


def main():
    ap = argparse.ArgumentParser(description="把上一期成品复制成这一期的四张模板")
    ap.add_argument("period", help="目标数据归属周期 YYYYMM")
    ap.add_argument("snapshot", help="目标快照日期 YYYYMMDD")
    ap.add_argument("--src", default=None, help="源期（默认 = 目标期上一月）")
    ap.add_argument("--force", action="store_true", help="目标已存在且内容不同时覆盖")
    ap.add_argument("--dry-run", action="store_true", help="只打印计划，不落盘")
    args = ap.parse_args()

    period, snapshot = args.period, args.snapshot
    if not re.fullmatch(r"\d{6}", period):
        sys.exit("❌ 期别格式应为 YYYYMM，例如 202611")
    if not re.fullmatch(r"\d{8}", snapshot):
        sys.exit("❌ 快照日期格式应为 YYYYMMDD，例如 20261201")

    src_period = args.src or _prev_period(period)
    prev = _prev_period(period)
    target_year = period[:4]

    snap_dir = os.path.join(TRUNK_DIR, "data", "数据归属周期", period, "快照日期", snapshot)
    result_root = os.path.join(TRUNK_DIR, "result")

    print(f"[目标] {period} / 快照 {snapshot}")
    print(f"[源期] {src_period}（目标期上一月 = {prev}）" + ("  ← 手工指定" if args.src else ""))
    if src_period != prev:
        print(f"⚠️ 源期不是紧邻的上一期！复制上上期可能继承 SEA Total 错位级联，确认无误再继续。")
    if src_period >= period:
        sys.exit(f"❌ 源期 {src_period} 不早于目标期 {period}")
    print(f"[落点] {snap_dir}")
    print("-" * 78)

    ok = miss = conflict = skip = 0
    for label, sub, src_parts, pattern in TEMPLATES:
        src_dir = os.path.join(result_root, *src_parts, src_period)
        src_file, note = _pick_source(src_dir, pattern)
        if not src_file:
            print(f"❌ {label:<8} {note}")
            miss += 1
            continue
        if note:
            print(f"   {label:<8} ℹ️ {note}")

        dst_dir = os.path.join(snap_dir, sub) if sub else snap_dir
        dst_file = os.path.join(dst_dir, _rename_year(os.path.basename(src_file), target_year))

        if os.path.isfile(dst_file):
            if _md5(dst_file) == _md5(src_file):
                print(f"⏭  {label:<8} 已存在且一致，跳过  ({os.path.getsize(dst_file):,}B)")
                skip += 1
                continue
            # 目标比源新 ⇒ 本期已经用过这份模板（主脚本的「得分列表头自愈」会就地改写 KH2，
            # 典型表现是「字节差 2B、KH2 从 N月 变 N+1月」）。这不是冲突，别覆盖。
            if os.path.getmtime(dst_file) > os.path.getmtime(src_file):
                print(f"⏭  {label:<8} 本期已用过（目标比源新，表头自愈已改写），跳过  "
                      f"({os.path.getsize(dst_file):,}B)")
                print(f"      要重置成上期原样：先删目标文件，或加 --force")
                skip += 1
                continue
            if not args.force:
                print(f"⚠️  {label:<8} 目标已存在但内容不同（目标比源旧），未覆盖（加 --force 才覆盖）")
                print(f"      目标 {dst_file}  ({os.path.getsize(dst_file):,}B)")
                print(f"      源   {src_file}  ({os.path.getsize(src_file):,}B)")
                conflict += 1
                continue
            print(f"♻️  {label:<8} 目标已存在且不同 → 覆盖（--force）")

        if args.dry_run:
            print(f"→  {label:<8} [dry-run] {src_file} → {dst_file}")
            ok += 1
            continue

        os.makedirs(dst_dir, exist_ok=True)
        shutil.copy2(src_file, dst_file)
        print(f"✅ {label:<8} {os.path.basename(dst_file)}  ({os.path.getsize(dst_file):,}B)  ← {src_period}")
        ok += 1

    # 顺手建好后续步骤要用的空目录（已存在则无操作）
    extra_dirs = [
        os.path.join(snap_dir, "东南亚", "跨境"),
        os.path.join(TRUNK_DIR, "data", "数据归属周期", period, "库龄", snapshot),
    ]
    if not args.dry_run:
        for d in extra_dirs:
            os.makedirs(d, exist_ok=True)
        print(f"📁 已确保目录存在：东南亚/跨境/、库龄/{snapshot}/")

    print("-" * 78)
    tag = "[dry-run] " if args.dry_run else ""
    print(f"{tag}结果：就位 {ok} / 跳过 {skip} / 冲突 {conflict} / 缺失 {miss}")
    if miss or conflict:
        print("⚠️ 有模板没准备好，先解决上面 ❌ / ⚠️ 再跑生成脚本。")
        return 2

    print()
    print("下一步：")
    print(f"  1) 把本期订单/库存/库龄报告放进 {snap_dir}")
    print(f"     （东南亚的放 东南亚/、跨境的放 东南亚/跨境/）")
    print(f"  2) 体检： python 工具-换月资料体检.py {period} {snapshot}")
    print(f"  3) 得分列： python 工具-更新商品资料得分列.py --apply")
    print(f"  4) 依次跑： month_sale_file.py → southeast_asia_sale_file.py → europe_sale_file.py → north_america_sale_file.py")
    return 0


if __name__ == "__main__":
    sys.exit(main())
