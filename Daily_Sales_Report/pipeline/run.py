# -*- coding: utf-8 -*-
"""日销表管线 —— 统一入口。

日常只记这一条命令：
    python pipeline/run.py

默认流程：**先跑前体检**（报告是否导完整 / 内容是否与上一快照相同 / 有无重复文件），
体检通过再执行 core/runner.py；有缺失报告会中止并提示（--force 可强行继续）。

常用参数（除 --check* / --force / --archive* 外，其余原样转发给 core/runner.py）：
    --dry-run             只打印计划，不写文件
    --only 日本,澳大利亚    只跑文件名含关键字的簿
    --snap <目录>          指定快照目录（默认自动取最新 YYYYMMDD）
    --period <YYYYMM>     指定数据归属周期（默认自动取最新）
    --jobs N              并行度（默认 3）

归档相关（2026-10-15 新增：跑批成功即按「年/月/日」自动留档）：
    --no-archive          跑完不归档
    --archive-only        不跑批，只把当前 12 本簿归档一份（补归档用）
    --force-archive       归档时机器版无条件覆盖，不留存被手工改过的归档件
                          （默认：手工改过的归档件先另存「-手工版-HHMMSS」再覆盖）

缺货表 / 断货明细（随日销表周期输出）：
    --no-shortage         不生成缺货情况表
    --no-stockout         不生成断货明细（详细版）
    --keep-unsold         缺货表不拆「未起量」，仍全量计入三级（回到 2026-10-29 前的口径）
    跑批成功后默认依次自动生成：
      `断货分析/亚马逊/每日缺货/<YYYYMMDD>/亚马逊美客多EMAG缺货情况表.xlsx`（2026-10-22 新增）
        三级标记：一级 可售天数+7<计划备货天数 / 二级 库存0有在途 / 三级 库存0无在途
        （2026-10-29 起：三级里「今年从未有过销量」的 SKU 单列到「未起量」sheet，
          不计入缺货数。为什么：可售为 0 本来就没货可卖，近30天销量 0 是结果不是原因，
          只按近30天销量判会把「断货导致的 0」误判成「弃售」。实测 20261028 的
          326 个无销量 SKU 里 listing 一个都没下架，且美客多墨西哥 78 个里 64 个今年卖过。）
      `断货分析/亚马逊/断货明细/<YYYYMMDD>/断货明细_<YYYYMMDD>.xlsx` + `断货报告_<YYYYMMDD>.md`
        （2026-10-28 接入；原因分析 / 采购方案 / 库龄呆滞，此前需手动跑）
    两个都不影响跑批结果与归档：失败只提示，退出码仍以 runner 为准。

销量趋势月报（2026-10-28 新增，**默认关闭**）：
    --trend               额外生成 `销量趋势/销量趋势月报-<YYYYMMDD>.xlsx` + `.md`
    默认不跑：它要扫 13 本簿 27 张表的月度块（8000+ SKU），比缺货表/断货明细慢不少。
    同样「失败只提示、不影响跑批与归档」。

体检相关：
    --check-only          只做跑前体检，不跑
    --no-check            跳过体检直接跑
    --force               体检红灯也照跑（缺报告 / 渠道错位 / SKU 归属不符的簿会被整份跳过）

体检是「十项」，其中两项是硬闸（2026-10-22 前置）：
    ⑨ 订单报告渠道整份错位（下错站点）
    ⑩ 销售报告 SKU 与本簿库存对不上（导成了别的账号的同渠道报告）
红灯 ⇒ 默认**一步都不写**直接退出（码 2）；以前是跑到写销售列才发现，那时库存列已落盘。

实现说明：runner 以**子进程**方式直接执行 core/runner.py——保证多进程 spawn 语义
与原来「python daily_update/runner.py」完全一致，不做任何导入层面的改动。
"""
import os
import subprocess
import sys

_HERE = os.path.dirname(os.path.abspath(__file__))   # …\Daily_Sales_Report\pipeline
_DSR = os.path.dirname(_HERE)                        # …\Daily_Sales_Report
_CORE = os.path.join(_HERE, "core")

_FLAGS = ("--no-check", "--check-only", "--force", "--no-archive",
          "--archive-only", "--force-archive", "--no-shortage", "--no-stockout",
          "--trend", "--keep-unsold", "--no-history", "--legacy-dos")


def _resolve_day(runner_args):
    """归档用的日期：--snap 指定了取其目录名（YYYYMMDD），否则取最新快照日。"""
    if "--snap" in runner_args:
        i = runner_args.index("--snap")
        if i + 1 < len(runner_args):
            return os.path.basename(runner_args[i + 1].rstrip("\\/"))
    if _CORE not in sys.path:
        sys.path.insert(0, _CORE)
    import config
    return config._LATEST_DAY


def _archive(runner_args, force_overwrite=False):
    """跑批成功后按「年/月/日」归档 12 本簿（archiver 内部容错，失败不炸主流程）。

    归档件若被手工改过，archiver 会先另存「-手工版-HHMMSS」再写机器版；
    确要机器版无条件覆盖（放弃手工版）时 force_overwrite=True（--force-archive）。
    """
    if _CORE not in sys.path:
        sys.path.insert(0, _CORE)
    import archiver
    day = _resolve_day(runner_args)
    print(f"\n[归档] 数据日期 {day}", flush=True)
    try:
        archiver.archive_day(day=day, force_overwrite=force_overwrite)
    except Exception as e:
        print(f"[归档] !! 整体失败 {e!r}（不影响已完成的跑批结果）", flush=True)


def _run_precheck(runner_args):
    """跑前体检；返回 (可跑?, 中止原因)。

    两道闸（⑨ 渠道整份错位 / ⑩ SKU 归属不符）已前置到体检：**跑之前就红，一步都不写**。
    以前是跑到 append_sales 才发现，此时库存列、记录列已经写进去了，处于半截状态。
    """
    if _CORE not in sys.path:
        sys.path.insert(0, _CORE)
    import precheck

    snap = None
    if "--snap" in runner_args:
        i = runner_args.index("--snap")
        if i + 1 < len(runner_args):
            snap = os.path.abspath(runner_args[i + 1])
    if snap:
        day = os.path.basename(snap.rstrip("\\/"))
        res = precheck.check(day, snap_dir=snap)
    else:
        res = precheck.check()

    print(precheck.render(res), flush=True)
    print(flush=True)
    if res["missing"]:
        return False, f"缺 {len(res['missing'])} 份必需报告"
    if res.get("dup_hard"):
        return False, (f"{len(res['dup_hard'])} 组库存报告内容完全相同"
                       f"（同一账号同一站点本只会有一份 → 必然是导错，会清空整片库存）")
    if res.get("sku_ov_err"):
        return False, f"{res['sku_ov_err']} 份销售报告的 SKU 与本簿库存对不上（导错账号）"
    if res.get("mislabel_err"):
        return False, f"{res['mislabel_err']} 份订单报告整份渠道错位（下错站点）"
    if res.get("stock_kind"):
        return False, (f"{len(res['stock_kind'])} 份库存报告下错了报表类型"
                       f"（如把库龄报告当成管理库存报告 → 该页会静默不更新）")
    return True, ""


def _shortage(runner_args):
    """缺货情况表：随日销表周期输出（2026-10-22 需求）。

    产出 `断货分析/亚马逊/每日缺货/<YYYYMMDD>/亚马逊美客多EMAG缺货情况表.xlsx`，三级标记：
      一级 库存≠0 且 可售天数+7 < 计划备货天数
      二级 库存0、在途>0
      三级 库存0、在途0
    依赖快照目录的库存/订单报告（与 runner 同源），失败不影响跑批结果。
    """
    _tools = os.path.join(_HERE, "tools")
    for p in (_CORE, _tools):
        if p not in sys.path:
            sys.path.insert(0, p)
    try:
        import shortage_report
    except Exception as e:
        print(f"[缺货表] !! 导入失败 {e!r}（跳过，不影响跑批）", flush=True)
        return
    args = []
    if "--snap" in runner_args:
        i = runner_args.index("--snap")
        if i + 1 < len(runner_args):
            args += ["--snap", runner_args[i + 1]]
    # 2026-10-29：未起量 SKU 默认拆到独立 sheet（三级、无在途、今年从未有过销量）。
    # 想回到旧口径（全量计入三级）时加 --keep-unsold。
    if "--keep-unsold" in runner_args:
        args.append("--keep-unsold")
    # --no-history：不读月度块（省约 16 秒），不出「上月/近3月/今年累计」三列，也不拆未起量
    if "--no-history" in runner_args:
        args.append("--no-history")
    # --legacy-dos：一级判定回退到「可售÷近30天日均」（2026-10-29 前口径）
    if "--legacy-dos" in runner_args:
        args.append("--legacy-dos")
    print("\n[缺货表] 生成亚马逊/美客多/eMAG 缺货情况表", flush=True)
    try:
        shortage_report.main(args)
    except Exception as e:
        print(f"[缺货表] !! 生成失败 {e!r}（不影响已完成的跑批）", flush=True)


def _stockout(runner_args):
    """断货明细（详细版）：随日销表周期输出（2026-10-28 用户拍板接入）。

    产出 `断货分析/亚马逊/断货明细/<YYYYMMDD>/`：
      · 断货明细_<YYYYMMDD>.xlsx —— 汇总/严重/一般/预备/库龄呆滞/采购建议/无库存记录
      · 断货报告_<YYYYMMDD>.md   —— 文字版，可直接发群
    比「每日缺货情况表」多的东西：断货原因、采购建议件数、库龄呆滞判断、无库存记录清单。

    **为什么用子进程而不是 import**：stockout_report 是脚本式入口（main 里 parse_args +
    末尾直接打印），以子进程跑与手动敲 `python pipeline/tools/stockout_report.py` 是完全
    同一条代码路径 —— 不会出现「自动跑的」和「手动跑的」两套口径。与 _shortage 的区别
    只是这里没有可以复用的 main(args) 形态（它读 sys.argv），故走 subprocess。

    失败/超时只提示，**不影响跑批结果与归档**，也不改退出码。
    """
    script = os.path.join(_HERE, "tools", "stockout_report.py")
    if not os.path.exists(script):
        print(f"[断货明细] !! 找不到 {script}（跳过，不影响跑批）", flush=True)
        return
    cmd = [sys.executable, script]
    if "--snap" in runner_args:
        i = runner_args.index("--snap")
        if i + 1 < len(runner_args):
            cmd += ["--snap", runner_args[i + 1]]
    print("\n[断货明细] 生成断货明细 + 断货报告（原因分析 / 采购方案）", flush=True)
    try:
        r = subprocess.run(cmd, cwd=_DSR, capture_output=True, text=True,
                           encoding="utf-8", errors="replace", timeout=1800)
    except Exception as e:
        print(f"[断货明细] !! 执行失败 {e!r}（不影响已完成的跑批）", flush=True)
        return
    out = (r.stdout or "").strip()
    if r.returncode != 0:
        tail = "\n".join((out.splitlines() + (r.stderr or "").strip().splitlines())[-6:])
        print(f"[断货明细] !! 退出码 {r.returncode}（不影响已完成的跑批）\n{tail}", flush=True)
        return
    # 只回放关键行：在册/断货分级、待补货件数、输出路径（路径行可能不带缩进）
    lines = [x.strip() for x in out.splitlines()]
    i_out = next((k for k, x in enumerate(lines) if "已输出" in x), None)
    for k, s in enumerate(lines):
        if s.startswith(("在册", "待补货")):
            print("  " + s, flush=True)
        elif i_out is not None and k > i_out and s.endswith((".xlsx", ".md")):
            print("  " + s, flush=True)
    print("[断货明细] 完成", flush=True)


def _trend(runner_args):
    """销量趋势月报：默认关闭，加 `--trend` 才跑（2026-10-28 用户拍板）。

    产出 `销量趋势/销量趋势月报-<YYYYMMDD>.xlsx` + `.md`（5 sheet：异动/连跌预警/新品/汇总/趋势明细）。
    为什么默认关：要扫 13 本簿 27 张表的月度块（8000+ SKU），比其他产出慢不少，
    用户选择「需要时再出」，不让它拖慢每天的常规跑批。

    同样走 **subprocess**（`sales_trend_report.main()` 读 `sys.argv`）；
    同样「失败只提示、不改退出码、不阻断归档」。
    """
    script = os.path.join(_HERE, "tools", "sales_trend_report.py")
    if not os.path.exists(script):
        print(f"[趋势月报] !! 找不到 {script}（跳过，不影响跑批）", flush=True)
        return
    cmd = [sys.executable, script]
    if "--snap" in runner_args:
        i = runner_args.index("--snap")
        if i + 1 < len(runner_args):
            # 只取目录名（20261028）：趋势脚本按 2026MMDD 解析，并自己拼快照路径
            cmd += ["--snap", os.path.basename(runner_args[i + 1].rstrip("\\/"))]
    print("\n[趋势月报] 生成销量趋势月报（--trend 指定）", flush=True)
    try:
        r = subprocess.run(cmd, cwd=_DSR, capture_output=True, text=True,
                           encoding="utf-8", errors="replace", timeout=1800)
    except Exception as e:
        print(f"[趋势月报] !! 执行失败 {e!r}（不影响已完成的跑批）", flush=True)
        return
    out = (r.stdout or "").strip()
    if r.returncode != 0:
        tail = "\n".join((out.splitlines() + (r.stderr or "").strip().splitlines())[-6:])
        print(f"[趋势月报] !! 退出码 {r.returncode}（不影响已完成的跑批）\n{tail}", flush=True)
        return
    # 回放：脚本自己的摘要行 + 「已输出:」下面那几行文件路径
    lines = [x.strip() for x in out.splitlines()]
    i_out = next((k for k, x in enumerate(lines) if "已输出" in x), None)
    for k, s in enumerate(lines):
        if s.startswith(("快照", "扫描", "异动")):
            print("  " + s, flush=True)
        elif i_out is not None and k > i_out and s.endswith((".xlsx", ".md")):
            print("  " + s, flush=True)
    print("[趋势月报] 完成", flush=True)


def _ensure_emag(runner_args):
    """EMAG 分段报告自动合并（2026-10-27 新增）。

    背景：eMAG 后台可以按日期段分批导出，用户拿到的是
    `EMAG全站点{m}订单报告1.xlsx` / `...2.xlsx`，而 config 只认
    `EMAG全站点{m}订单报告.xlsx` → 体检直接报「缺必需报告」红灯。
    10-27 首次遇到（1.xlsx 中文表头 10-08~10-26 / 2.xlsx 英文表头 09-30~10-08）。

    这里在体检之前先看一眼：**主文件不存在且有 ≥2 份分段**才合并，
    其余情况（主文件在 / 没有分段 / 只有 1 份）一律不动，旧行为完全不变。
    失败只提示不阻断 —— 让体检照常报缺，不掩盖真实问题。
    """
    if _CORE not in sys.path:
        sys.path.insert(0, _CORE)
    try:
        import config
    except Exception:
        return
    snap = config.SNAP
    if "--snap" in runner_args:
        i = runner_args.index("--snap")
        if i + 1 < len(runner_args):
            snap = os.path.abspath(runner_args[i + 1])
    day = os.path.basename(snap.rstrip("\\/"))
    period = day[:6]
    tools = os.path.join(_HERE, "tools")
    if tools not in sys.path:
        sys.path.insert(0, tools)
    try:
        import merge_emag_orders as ME
        res = ME.merge_day(snap, period, apply=True)
    except Exception as e:
        print(f"[EMAG] 自动合并失败（不影响跑批，体检会照常报缺）：{e!r}", flush=True)
        return
    if res.get("ok"):
        print("[EMAG] 分段报告已合并 → %s（%d 单，原件保留）"
              % (os.path.basename(res["target"]), res["rows"]), flush=True)


def main():
    argv = sys.argv[1:]
    no_check = "--no-check" in argv
    check_only = "--check-only" in argv
    force = "--force" in argv
    no_archive = "--no-archive" in argv
    archive_only = "--archive-only" in argv
    dry_run = "--dry-run" in argv          # dry-run 无副作用：不落盘、也不该覆盖归档
    force_archive = "--force-archive" in argv
    no_shortage = "--no-shortage" in argv
    no_stockout = "--no-stockout" in argv
    trend = "--trend" in argv                # 趋势月报默认关，显式 --trend 才跑
    runner_args = [a for a in argv if a not in _FLAGS]

    if archive_only:                      # 只归档：不体检、不跑批
        _archive(runner_args, force_overwrite=force_archive)
        return 0

    # 型号首库存登记（2026-10-23）：刷新「断货分析/型号首库存登记.xlsx」里的 {SKU: 首现日期}。
    # 必须跑在 runner 之前 —— runner 补行时按「首现日期 >= 当月 1 号」判新品，读的就是这份表。
    # 失败只提示不阻断（补行会退化为「无登记表 → 跳过新品补行」，不会误补全量）。
    if not dry_run:
        _reg = os.path.join(_HERE, "tools", "first_stock_register.py")
        if os.path.exists(_reg):
            try:
                r = subprocess.run([sys.executable, _reg, "--apply"], cwd=_DSR,
                                   capture_output=True, text=True, timeout=300)
                tail = [x for x in (r.stdout or "").strip().splitlines() if x.strip()]
                print(f"[登记] 型号首库存登记 {tail[-1] if tail else '已完成'}", flush=True)
            except Exception as e:
                print(f"[登记] 刷新失败（不影响跑批）：{e!r}", flush=True)

    # EMAG 分段报告合并（2026-10-27）：必须在体检之前，否则会误报「缺必需报告」。
    # dry-run 也执行：它只是把用户给的分段原件合成一份**输入物料**（原件保留、不动任何
    # 工作簿），若 dry-run 跳过就会给出与实跑不一致的红灯，失去预演意义。
    _ensure_emag(runner_args)

    if not no_check:
        ok, why = _run_precheck(runner_args)
        if check_only:
            return 0 if ok else 1
        if not ok and not force:
            print(f"[中止] {why}，未开跑（工作簿一个字都没动）。", flush=True)
            print("       按体检提示重导报告后重跑；确要继续请加 --force"
                  "（出问题的簿会被整份跳过，其余簿正常）。", flush=True)
            return 2

    runner = os.path.join(_CORE, "runner.py")
    if not os.path.exists(runner):
        raise SystemExit(f"找不到入口 {runner}")
    cmd = [sys.executable, runner] + runner_args
    rc = subprocess.call(cmd, cwd=_DSR)
    if rc == 0:
        if not no_shortage and not dry_run:
            _shortage(runner_args)                             # 缺货情况表（随日销周期）
        if not no_stockout and not dry_run:
            _stockout(runner_args)                             # 断货明细（详细版，2026-10-28）
        if trend and not dry_run:
            _trend(runner_args)                                # 趋势月报（默认关，--trend）
        if not no_archive and not dry_run:
            _archive(runner_args, force_overwrite=force_archive)   # 跑批成功才留档
    elif rc != 0:
        print(f"[归档] runner 退出码 {rc}，本次不归档", flush=True)
    elif dry_run:
        print("[归档] --dry-run，跳过归档", flush=True)
    return rc


if __name__ == "__main__":
    sys.exit(main())
