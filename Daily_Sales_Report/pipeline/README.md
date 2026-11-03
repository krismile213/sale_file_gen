# 日销表管线（pipeline）

一条命令跑完 12 本「XX库存销售表」的每日更新。**这里只有代码**；数据表、报告、快照都在别处（见下）。

---

## 一、日常用法

```bash
python pipeline/run.py                  # 全量跑 12 本（先跑前体检，通过再执行）
python pipeline/run.py --dry-run        # 只打印计划，不写文件
python pipeline/run.py --check-only     # 只做跑前体检，不跑
python pipeline/run.py --snap <目录>     # 指定快照目录
python pipeline/run.py --period 202610  # 指定数据归属周期
python pipeline/run.py --only 日本,澳大利亚   # 只跑文件名含关键字的簿
python pipeline/run.py --jobs 4         # 并行度（默认 3）
python pipeline/run.py --no-check       # 跳过体检直接跑
python pipeline/run.py --force          # 体检发现缺报告也照跑（缺报告的簿会被跳过）
python pipeline/run.py --no-archive     # 跑完不归档
python pipeline/run.py --archive-only   # 不跑批，只把当前 12 本簿归档一份（补归档用）
python pipeline/run.py --force-archive  # 归档时机器版无条件覆盖（默认：手工改过的会先留存）
```

**周期/快照全自动**：`config.py` 扫 `数据归属周期/{YYYYMM}/快照日期/{YYYYMMDD}` 自动取
「日期最大」的周期与快照，并据此推导订单报告月（`ORDER_MONTH`）和当月表窗口下界
（`WINDOW_LO` = 上月最后 7 天起点）。**换月、换快照日都不用再改代码**，`--snap` / `--period` 可手工覆盖。

### 跑前体检（默认自动）

`run.py` 开跑前先做一次体检（也可单独 `python pipeline/core/precheck.py [--snap <目录>]`）：

| 检查 | 说明 |
|---|---|
| ① 导完整 | 按 12 簿配置核对当天所需 **39 份报告**是否都在、格式对不对、有没有空文件/只有表头 |
| ② 工作簿占用 | WPS/Excel 开着 → 跑时该簿会被整份跳过 |
| ③ 内容对比 | 与**上一个快照**按「归一化报告名」逐个比 MD5 → 未变 / 已更新 / 新增 / 消失 |
| ④ 重复内容 | 本快照内内容完全相同的文件组（重复下载 or 同源） |
| ⑤ 货件 | TSV 票数、同一票是否重复下载（代码会自动去重） |
| ⑥ 货件列容量 | 每张分析表还剩几个空列位 |
| ⑦ 数组公式隐患 | 整列引用被数组公式当标量用（会静默算成 0） |
| ⑧ 订单报告日期覆盖 | purchase-date 的 min/max：**date range 按站点本地时间设**会截掉月初/月末几小时（北美 = 月初 7 小时），详见 `core/report_coverage.py` |
| ⑨ 订单报告渠道错位 | 报告 sales-channel 分布 vs 本簿站点：**整份下错站点**（外来渠道 ≥50%）→ runner 会整份拒收、相关簿当日销售不更新 |

**有缺失报告时默认中止**（exit 2，缺报告的簿会被整份跳过、会少算数），补齐后重跑；确要继续加 `--force`。

### runner 安全闸与退出码

| 闸 | 触发 | 行为 |
|---|---|---|
| 库存空报告 | 报告解析出 0 行（只有表头/空白模板） | 该 sheet 当缺报告跳过（防清表）**且锚点不滚动** |
| 销售报告错位 | 外来渠道 ≥50% 或无一条本簿渠道 | 整份拒收，当日销售不更新 |
| 下载日期虚报 | 任一库存 sheet 未更新（缺/空/sheet 不存在） | 锚点（=库存数据快照日）保持旧值 |

- 锚点（`config.ANCHORS`）是下载日期格，**只有该簿全部库存报告都入库才滚动**；部分 sheet 还是旧数据时把共享日期滚到新快照日 = 虚报库存新鲜度（2026-10-20 沙特库存报告只有表头，表里还是前一天数据，锚点却被写成本日 —— 已修）。
- 退出码：`0` 正常（允许软提醒）；`2` 体检缺报告中止；`3` runner 有硬伤（报告错位/缺失、库存缺/空、sheet 不存在、跑批异常）——末尾「今日异常汇总」逐条列出。

### 每日自动归档（跑批成功即留档）

`run.py` 跑批退出码为 0 时，自动把 12 本工作簿按**年/月/日**复制一份：

```
1.每日销售数据表/2026年每日数据/2026年10月/10月15日/2026年10月15日-XX库存销售表.xlsx
```

- 与 `month_rollover` 的**月末归档分工不同**：月末归档做窗口裁剪 + 锚点定格，是次月封存的取数源；
  每日归档是**当日原样快照**（约 17MB/天），供回溯、对账、误删恢复。两者路径不同、互不覆盖。
- 归档日期默认取最新快照日（`--snap` 指定时跟随其目录名）；同日重跑且内容未变 → 跳过，
  内容有变 → 覆盖为最新机器版（幂等）。
- **手工维护保护（2026-10-20 事故后新增）**：归档目录是用户在用的台账（标「到货」、删历史货件列）。
  归档件若在 archiver 写入后被外部改过（比对 `.workbuddy/archive_state/{day}.json` 登记的 md5），
  先把旧件另存为 `YYYY年M月D日-XX库存销售表-手工版-HHMMSS.xlsx` 再写机器版，两条线都不丢；
  尚无登记的历史归档同样按「手工版」保护处理。要机器版无条件覆盖加 `--force-archive`。
- 复制前 `testzip` 校验源簿，单本失败不阻断其它簿；`--no-archive` 可关闭，`--archive-only` 可补归档。

人工只做两件事：
1. **喂数**：当天报告放进快照目录；发一批货件就往 `货件/{地区}/` 丢一份 TSV。
2. **点货**：货到了，在分析表**行5 标签**里补「到货」二字（H 调拨在途合计 / AC 差额自动刷新）。

日常运维工具（手动跑）：

```bash
python pipeline/core/precheck.py                       # 跑前体检（同 run.py 自动调用的那份）
python pipeline/tools/check_report_coverage.py --day 20261016   # 订单报告日期覆盖（防本地时间 date range 丢单）
python pipeline/tools/check_report_coverage.py --book <工作簿.xlsx>   # 反向审计：哪天丢了凌晨/傍晚时段
python pipeline/tools/check_snapshot.py                # 快照齐全性深度体检（含 EU 旧源回落探测）
python pipeline/tools/verify_ship_afn.py 20261013      # 货件台账 × AFN 三道闸对账
python pipeline/tools/month_rollover.py                # 月度换月（归档上月 + 开新）
python pipeline/tools/inject_small_base_cf.py          # 红色「小基数档」CF 注入/图例更新（幂等，--dry-run 先看）
python pipeline/tools/audit_hist_formulas.py           # 统计归档簿关键列是公式还是静态值
```

---

## 二、目录结构

```
pipeline/
├─ run.py                 统一入口（以子进程方式调 core/runner.py，保证多进程语义不变）
├─ core/                  ← 每天跑的「活代码」，就这 8 个
│   ├─ config.py          路径 + 12 簿配置 + 各簿所需报告清单 + 参数（周期/快照/窗口全自动推导）
│   ├─ precheck.py        跑前体检九块：导完整 / 占用 / 内容比对 / 重复 / 货件 / 列容量 / 数组公式 / 报告日期覆盖 / 报告渠道错位
│   ├─ report_coverage.py 订单报告 purchase-date 覆盖 + sales-channel 错位检查（本地时间 date range / 下错站点）
│   ├─ runner.py          主流程：备份→库存sheet→销售append→记录补列→货件落列→月汇总→保真回填
│   ├─ analysis.py        分析表维护：追加缺失 SKU 行 / 月销量封存 / 「当月总销量」表头滚动
│   ├─ shipments.py       货件 TSV → 各分析表「货件列」（含地区子文件夹与仓库分站）
│   ├─ record_extend.py   库存记录表按天补列（XML 手术）
│   └─ archiver.py        跑批后按「年/月/日」自动归档 12 本簿（当日原样快照）
├─ tools/                 手动运维工具（不参与每日自动跑）
│   ├─ verify_ship_afn.py      货件台账 × AFN 三道闸对账 → 说明文档/核查-*.md
│   ├─ check_snapshot.py       快照齐全性体检
│   ├─ check_report_coverage.py 订单报告日期覆盖 / 工作簿反向审计（导数规范配套）
│   ├─ month_rollover.py       月度换月
│   ├─ inject_small_base_cf.py 红色小基数档 CF 注入 + B1 图例（XML 手术，幂等）
│   ├─ audit_hist_formulas.py  统计归档簿关键列公式/静态值（2025-2026 全量）
│   ├─ sep_compare.py / compare_daily_vs_monthly.py   日销 vs 月销对撞
│   ├─ build_history_doc.py / extract_history_2025_2026.py  历史总档生成
│   ├─ lean_make2.py / lean_verify.py                  精简版工作簿生成与验收
│   ├─ fix_archive_ids.py / patch_archive_formulas.py / fix_bloat.py
│   │  / backfill_emag_orders.py / verify_formula_columns_20261012.py
│   └─ maintenance/            一次性修复脚本（均已执行，留档备查）
│       ├─ fix_onsale_blocks.py / fix_e_reserved_and_ref.py / backfill_richdata.py
│       └─ fix_int_missing_rows.py / fix_g598_rows.py
└─ docs/                  设计 / 规格 / 结构文档
```

> **路径不敏感**：`config.py` 靠「向上找含 `1.每日销售数据表` 的目录」定位根，所以整个
> `Daily_Sales_Report/` 可以整体搬盘符、改目录名，不用改一行代码。
> 新增工具脚本请直接沿用 `tools/` 里首个 import 前那段**引导块**，不要再写死路径。

---

## 三、数据在哪（不在本目录）

| 内容 | 位置 |
|---|---|
| 12 本工作簿 | `Daily_Sales_Report/1.每日销售数据表/` |
| 月度归档（月末封存源） | `Daily_Sales_Report/1.每日销售数据表/2026年每日数据/2026年{M}月/` |
| 每日归档（跑批快照） | `Daily_Sales_Report/1.每日销售数据表/2026年每日数据/2026年{M}月/{M}月{D}日/` |
| 快照 / 货件 | `sale_file_gen/trunk/data/数据归属周期/{YYYYMM}/快照日期|货件/` |
| 业务文档 | `sale_file_gen/说明文档/` |
| 本次整理前的散文件归档 | `Daily_Sales_Report/_archive/cleanup_20261014.zip`（+ `.manifest.txt` 清单） |

---

## 四、核心口径（改代码前必读）

- **有效订单** = `E<>"Cancelled" 且 MID(G,1,7)="Amazon."`；地区 = `IF(G="Amazon.co.uk","英国","欧洲")`，统计欧洲要**排除 co.uk**。
- **分析表列义**：E=预留 F=可售 G=在途 H=调拨在途合计（Σ未标「到货」的货件列）AC=在途差额(G−H)
  AD=实时速度 AE=上月日均 AF/AG=可售天数 AH=当月总销量 AI+=日销；右侧为 12 个月汇总区。
- **月封存**：上月封存**必须读归档簿**（当前簿换月已裁掉上月早期订单行，直读＝月尾残值）。
  口径由 `analysis.scan_analysis/resolve_month_sales` 在运行时从各表日销 SUMIFS 反解，自动与表内一致。
- **货件**：数量取 TSV 明细第 10 列「已发货」；文件夹/仓库→簿 的映射见 `shipments.FOLDER_RULES`；
  AFN 三件套**不能替代**手工货件列（只覆盖已入网货件），G/H/AC 双轨保留。
  - `interflow/` 目录按仓库分簿：YYZ…=加拿大、XSY…=澳洲、DXB…=阿联酋、其余=美国；
    **RUH/JED/DMM（沙特）= 无对应簿 → 跳过不落列**并告警（Interflowd 没有沙特簿，
    2026-10-20 前会兜底算成"美国"，77 件沙特货件永久挂在途）。跑前体检 ⑤ 也会列出来。
  - **文件夹 × 仓库交叉校验**（`shipments.cross_region_warn`，跑前体检 ⑤）：仓库明确属于
    别国（CA/MX/AU/AE/SA/JP/IN/UK 前缀）但放错文件夹 → 报警（只报警、不改归属）。
    美国/欧洲仓代码无法穷举，未知前缀不判，因此不会误报。`OAK*` 有歧义（澳洲 vs 美国奥克兰）单独提示。
  - **特殊情况货件号带地区说明**（`shipments.sid_label`）：`FBA15MHHH99G（RUH8沙特）`；
    正常票只写货件号。体检 ⑤ 与 runner 告警共用同一格式。
- **保真**：openpyxl 保存会清公式缓存（无害，WPS/Excel 打开自算）；`restore_parts` 负责把
  customXml / 缩略图 / pivot 刷新标志 / richData 等部件按原字节回填。

---

## 五、整理记录

2026-10-14 做过一次结构收敛：原 `daily_update/` + `日销表程序版/daily_update/`（旧副本）
+ 根部散脚本 + `.workbuddy/` 下数百个探针/备份，统一为上面的 `pipeline/`；
被移除的旧文件全部打进 `_archive/cleanup_20261014.zip`，可随时解压找回。

2026-10-14 追加：① `config.py` 的周期 / 快照 / 窗口下界改为**自动推导**（不再手改）；
② 新增**跑前体检** `core/precheck.py`，`run.py` 默认先体检再跑；③ 销售报告 `report` 支持候选列表
（兜住「亚马逊欧洲站点 / 欧盟站点」这类下改名漂移）。

2026-10-15 追加：④ 新增**每日自动归档** `core/archiver.py`——`run.py` 跑批成功（退出码 0）即把
12 本簿按「年/月/日」复制留档（`2026年每日数据/2026年10月/10月15日/2026年10月15日-XX.xlsx`，
约 17MB/天）；与月末归档（月封存源）分工不同、路径不重叠。`--no-archive` 关闭、`--archive-only` 补归档。
⑤ 新增 `tools/audit_hist_formulas.py`：统计 2025-2026 全部归档簿的关键列是公式还是静态值
（首跑结论：上月日均 259/260 为静态值，唯一 1 条公式是 2026-07 Interflow阿联酋簿 T 列公式右移一格的
手工错贴；T/V/W/X 四列全公式）。
⑥ 2026-10-15：「上月日均」列改为**活公式** `=IF(封存列="","",ROUND(封存列/天数,2))`（round 2 位，
救回被 round 1 抹零的 703 行）；新增 `tools/inject_small_base_cf.py` 给 16 张分析表注入**红色小基数档**
CF（`AND(0<上周均<=门槛, 日销>=2, 日销>上周均*4)`，复用现有红色 dxf、与大基数档互斥、不抢色）并同步
B1 图例；openpyxl 往返测试确认规则幸存。详见 `说明文档/分析表颜色判定与小基数档变更-20261015.md`。
另：安装 pypdf 到用户站点（EMAG 货件 PDF 解析依赖，缺失会让美客多/罗马尼亚簿崩在货件步骤）。

2026-10-20 追加：⑦ `core/archiver.py` 加**手工维护保护**。此前归档是无脑 `copy2` 覆盖，
用户在「10月20日」归档目录手工标的 16 个「到货」标记、删的 17 个历史货件列，被 15:44 跑批
整体冲掉（已从 WPS 备份中心抢救并合并回归档件）。现在归档前比对 `.workbuddy/archive_state/{day}.json`
登记的 md5：归档件被动过 → 先另存 `-手工版-HHMMSS.xlsx` 再写机器版；内容未变 → 跳过；
`run.py` 新增 `--force-archive` 强制机器版覆盖。详见 `.workbuddy/memory/2026-10-20.md`。
