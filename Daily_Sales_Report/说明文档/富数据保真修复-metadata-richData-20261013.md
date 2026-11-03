# 富数据保真修复：`xl/metadata.xml` + `xl/richData/*` + 单元格 `vm=` 索引

**日期**：2026-10-13（会话 6）　**范围**：`daily_update/runner.py` 保真层 + 8 本生产簿回填
**触发**：上一轮报告的「顺带发现」——runner 每日 openpyxl 存盘会丢 `xl/metadata.xml`，用户回复「需要」修。

---

## 1. 先纠正上一轮的判断（重要）

| 我上一轮的说法 | 实测结论 |
|---|---|
| 丢的是 `xl/metadata.xml` | ❌ 不完整。还丢 **`xl/richData/rdrichvalue.xml` + `rdrichvaluestructure.xml`** |
| 类型是动态数组（XLDAPR），格上有 `cm=` 索引 | ❌ **类型是 `XLRICHVALUE`（富数据/数据类型）**；**格上根本没有 `cm=`，只有 `vm=`** |
| `可售库存!F1` 的 5 列溢出退化成单格 CSE | ❌ **不成立**。跑批前原件本身就是 `ref="F1"`（单格）且缓存值已是 `#VALUE!` |
| 只加 `_RESTORE_PREFIX` 修不了 | ✅ 对。真修法=**部件 + CT + rels + 原位补 `vm=`** 一起做（本轮已实现） |

**真实影响面**：8 本簿带富数据，另 4 本（日本 / 欧盟英国 / 罗马尼亚 / 美客多）本就没有。`vm="1"` 只落在极少数格：

| 簿 | 带 `vm=` 的格 |
|---|---|
| Interflowd澳 / 印度 / 澳大利亚 | `可售库存!F1` |
| Interflowd阿联酋 / 阿联酋 | `阿联酋可售库存!F1` |
| 沙特 | `沙特阿拉伯可售库存!F1` |
| 北美加拿大 / Interflowd北美加拿大 | `美国!F1`、`美国!A3`、`加拿大!F1`、`加拿大!A3`（+ 北美簿 `墨西哥可售库存!F1`） |

这些格的公式都是 `=CHOOSECOLS(FILTER(HSTACK(Sheet1!…)))`，**缓存值在跑批前就已是 `#VALUE!`**。
所以丢掉富数据的后果 = **只丢「数据类型/富值」这层标记，可见值不变**（openpyxl 本就把公式缓存清成 `<v/>`，
`fullCalcOnLoad="1"` 打开时全量重算）。**这是一个结构保真问题，不是数据/显示问题。**

---

## 2. 修复方案（写进 runner，非一次性脚本）

openpyxl 保存时把「不认识的部件」整段丢掉，并且在重写 sheet XML 时丢掉单元格上的 `vm=`。所以要在**存盘后**做一次包级手术：

| 步骤 | 实现 |
|---|---|
| ① 补回部件 | `_RESTORE_PREFIX` 加 `"xl/metadata.xml"`、`"xl/richData/"` |
| ② 补 `[Content_Types].xml` | `_ct_patch()` 新增 3 处 `Override`（sheetMetadata / rdrichvalue / rdrichvaluestructure），否则 Excel 报「文件损坏」 |
| ③ 补 `xl/_rels/workbook.xml.rels` | **新增** `_wb_rels_patch()`：补 `sheetMetadata` / `rdRichValue` / `rdRichValueStructure` 三条关系 |
| ④ 原位补单元格 `vm=`/`cm=` | **新增** `_sheet_name_map()` / `_collect_vmcm()` / `_inject_vmcm()`：按 **sheet 名** 匹配新旧部件路径（路径会因 openpyxl 重排而不同），把原件里带索引的格原位补回 |
| ⑤ 挂载 | `restore_parts(orig, tmp)` 内联调用（`orig` = 跑批前文件，天然就是富数据的源）；另暴露可复用函数 `restore_richdata(orig, new)` |

**安全设计**：
- **幂等**：`_inject_vmcm` 只在「该格尚未带 `vm`/`cm`」且「该格仍是公式格（紧跟 `<f`）」时注入 ⇒ 二次注入无效、也不会误标普通值格。
- **不炸跑批**：`restore_richdata` 全程 `try/except`，失败仅打印一行日志。
- **零影响**：无富数据的簿（原件没有这些部件）走不到任何分支，输出不变。
- **调用顺序安全**：`restore_parts` 在 `wb.save()` 之后、`record_extend`（只动记录表）之前，互不干扰。

---

## 3. 落地与核验

### A) 8 本生产簿回填（源 = 跑批前原件 `backup_daily_20261013`）
脚本 `.workbuddy/backfill_richdata.py`（**默认干跑**，加 `--apply` 才写；写前备份到
`.workbuddy/backup_metadata_20261013/`）。

| 簿 | 补回部件 | 回填 vm 的 sheet 数 | zip | CT | rels | openpyxl 加载 |
|---|---|---|---|---|---|---|
| Interflowd北美加拿大 | 4 | 2 | OK | ✓ | ✓ | OK(13 sheet) |
| 北美加拿大 | 4 | 3 | OK | ✓ | ✓ | OK(15 sheet) |
| Interflowd澳 / 印度 / 沙特 / 澳大利亚 / 阿联酋 / Interflowd阿联酋 | 各 4 | 各 1 | OK | ✓ | ✓ | OK(6–7 sheet) |

> 4 个部件 = `xl/metadata.xml` + `xl/richData/`(目录) + `rdrichvalue.xml` + `rdrichvaluestructure.xml`。

### B) 端到端 + 幂等（沙箱 `bench_run.py --jobs 4`）
- 第一遍：每本带富数据的簿日志出现 `[保真] 富数据 vm/cm 索引原位回填 N 个 sheet` + `回填4部件`；
  4 本无富数据的簿 `回填0部件`（未被误注入）。
- 第二遍（`--keep`，在已处理过的簿上再跑）：**vm 格前后逐格 0 差异**、meta 仍在、zip 完好、无重复注入 ⇒ **幂等**。

### C) 数据中性
沙箱产物 vs 生产**逐格对比 = 0 差异**（欧盟英国簿含透视/货件列、Interflowd北美加拿大簿含富数据+记录表、
印度簿含富数据、日本簿无富数据）⇒ 本轮改动**只动包结构，不动任何单元格公式/值**。

### D) 顺带确认
`backup()` 有 `if not os.path.exists(dst)` 保护 ⇒ **沙箱跑批不会覆盖生产备份目录**（本轮已复核文件原始 mtime 未变）。

---

## 4. 现在的行为

- 今天起，**日常跑批自动保住富数据**，无需任何人工步骤。
- 以后即便再出现「丢部件」，判据：`zipfile` 比对跑批前备份 vs 现值；或直接 `runner._collect_vmcm(簿路径)` 看当前 vm 格。
- 回填脚本保留可复用：`python .workbuddy/backfill_richdata.py`（干跑）/ `--apply`（写回）。

**改动文件**：`daily_update/runner.py`（保真层）；**新增脚本**：`.workbuddy/backfill_richdata.py`、
`.workbuddy/verify_richdata_bench.py`；**备份**：`.workbuddy/backup_metadata_20261013/`。
