# 代码变更对照：**最原始版 → 当前版**

> 生成时间：2026-09-12  
> 用途：一眼看清 `month_sale_file.py` 从「9.11 最原始那份」到「现在这份」到底改了什么、每处为什么改、对结果有什么影响。
> 📝 **最近一次修订：2026-10-07（202609 期）** —— §⑨「KJ / KK 列代码化」由"列名年份化"修正为**按位置写回 + 硬断言**（见该节标题与正文 ⚠️ 注）。

---

## 0. 对比对象

|       | 最原始版（下称 **A**）                                                                                                                           | 当前版（下称 **B**）                                                 |
| ----- | ---------------------------------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------- |
| 路径    | `D:\Desktop\sale_file_gen\trunk\month_sale_file(old).py`（9.14 由 `month_sale_file.py` 改名而来）  
（也可写作 `C:\Users\dell\Desktop\...`——`C:\Users\dell\Desktop` 是指向 `D:\Desktop` 的符号链接） | `D:\Desktop\work-main\sale_file_gen\trunk\month_sale_file.py` |
| 行数    | **1166** | **1565** |
| 字节    | 60,663 | 84,420 |
| md5   | `b485fa295476421eea0c0051f20b97eb` | `f7dc7b6afb77…` |
| 修改时间  | 2026-09-11 11:11 | 2026-09-15 |
| 能否直接跑 | ❌ **不能**，第 13 行 `import timedelta` 直接 `ModuleNotFoundError`（见下）                                                                          | ✅ 能，9.12 全流程实跑通过（约 38 秒）                                      |

**差异规模**：`diff -u A B` → **19 处 hunk，新增 425 行、删除 26 行**（净 +399，正好 = 1565 − 1166）。

> 🔧 **修正记录（9.15）**：落地 **任务⑩**（列分组自动推进到当月；另含用户自己改的 811 行去重），
> hunk 由 **18 → 19**、规模由 186/25 → **425/26**，行数 1327 → **1565**（+238）。
> ⚠️ **A 的存放位置变了**：原路径 `D://Desktop//sale_file_gen//trunk//month_sale_file(old).py` 已不存在，
> 现只有 `D://Desktop//sale_file_gen.zip`（9.11 整包）里还有原版；
> 本次已解出到 `D://dev-cache//sale_file_gen//zipcheck//trunk//month_sale_file.py`（1166 行 / md5 核对一致）供 diff 使用。
> ⚠️ **本次新增的 `-743,7 +808,7` 那一处不是助手改的**，是**用户自己**把第 811 行 `other_shopee_cross_site`
> 去掉 `"新加坡"`（解决「新加坡 vs 虾皮境内」重复计数 119）——之前的 9.14 版对照用的 B 是未含此改动的副本，故没体现。

> 🔧 **修正记录（9.14 晚）**：9.14 又落地了 **任务⑧**（三个东南亚平台状态过滤）与 **任务⑨**（东南亚汇总列名），hunk 由 **11 → 18**、规模由 146/18 → **186/25**；「当前版」行数 / 字节 / md5 已同步更新。
>
> 🔧 **修正记录（9.14）**：此处原写「新增 147 行、删除 19 行」——那是把 diff 输出的 `+++`／`---` **两个文件名头行**也当成了正文行；重新核对后为 **146 / 18**（该版；现已并入 186/25）。

**怎么复现这份对照**：

```bash
diff -u "/d/Desktop/sale_file_gen/trunk/month_sale_file(old).py" \
        "/d/Desktop/work-main/sale_file_gen/trunk/month_sale_file.py"
```

> ⚠️ 注意 A 是**比修改指南更早**的版本：指南里提到的「1050-1075 行伪代码」「把 1050-1067 整块剪切到 1205 行」等操作，针对的是中间某个版本；**A 里根本没有这些块**，所以下面的对照里它们表现为"纯新增"。



---

## 1. 总览（按主题）

| #  | 主题            | diff 位置             | 一句话                                                         | 影响本月数字？           |
| -- | ------------- | ------------------- | ----------------------------------------------------------- | ----------------- |
| 1  | 删除死导入         | `-10,10 +10,11`     | 删掉 `import timedelta`；同一次改动里补了 `import re`、`from copy import copy`（任务⑦ 要用） | 影响"能不能跑"          |
| 2  | 日期字面量 → 变量    | `-933,7 +968,7`     | 8 处 `"202608","20260907"` → `belonging_date, snapshot_date` | 否（等价）             |
| 3  | 新增得分列变量       | `-898,6 +932,7`     | 新增 `score_col_name`                                         | 否（等价）             |
| 4  | ASIN 反查兜底     | `-958,44 +993,52`   | 反查失败时回退用源表 sku                                              | **是，+298 件**      |
| 5  | 第一层 SKU 清洗    | `-1023,8 +1066,10`  | 销量侧 sku `strip` + `"nan"`→NaN + 丢空 SKU 行                    | 是（行数/新增SKU数）      |
| 6  | 第二层 SKU 清洗    | `-1036,13 +1081,15` | 月份表 SKU `strip` + 三合一过滤                                     | 是（同上）             |
| 7  | 第三层 SKU 清洗    | `-1050,7 +1097,7`   | 补录表 SKU `strip`                                             | 否（防患）             |
| 8  | 得分列改变量 + 断言前置 | `-1064,8 +1111,10`  | 两行 assert 提前拦截列名错                                           | 潜在保护              |
| 9  | KJ / KK 列代码化  | `-1144,6 +1193,20`  | 年度总销量、SKU 数字由代码算                                            | **是，KJ 从过期缓存变正确** |
| 10 | 新增 任务②③ 整段    | `-1160,6 +1223,71`  | 欧洲各国 sheet + 2021-2026 sheet 补当月 + 汇总行三格                    | 新增功能              |
| 11 | **新增行自动套样式**（任务⑦） | `-129,14 +130,47` | `fill_excel_from_df` 给超出模板行数的新增行补齐底纹/边框/字体/行高，并扩自动筛选范围；`wb.save()` 加 `PermissionError` 人话提示 | 否（**只改样式，数值 0 差异**） |

### 1.1 最新对照（9.14 更新：任务⑧⑨ 落地，共 **18** 处）

> ⚠️ 上表是 **9.12 版**的 11 处，它的 `+` 侧（当前版）行号**已被 9.14 的改动整体顶移**，只适合理解改动主题；**要跳代码请用下表**。
>
> 下表为 9.14 重新实测的完整 18 处，**按 B 侧行号从小到大排列**。

| # | 主题 | diff 位置（A → B） | 一句话 | 影响本月数字？ |
| -- | -- | -- | -- | -- |
| 1 | 删除死导入 | `-10,10 +10,11` | 删 `import timedelta`；补 `import re`、`from copy import copy`（任务⑦ 要用） | 影响"能不能跑" |
| 2 | **新增行自动套样式**（任务⑦） | `-129,14 +130,47` | `fill_excel_from_df` 给超出模板行数的新增行补齐底纹/边框/字体/行高，并扩 `auto_filter`；`wb.save()` 加 `PermissionError` 人话提示 | 否（**只改样式，数值 0 差异**） |
| 3 | **任务⑧-1 泰国 `status_map`** | `-225,7 +259,7` | 重复键 `"สำเร็จแล้ว":"Cancelled"`（把「完成」标成「取消」）→ 改为 `"ยกเลิกแล้ว":"Cancelled"` | **是**（前置修复，否则泰国站清零） |
| 4 | **任务⑧-2 退货窗口归一化** | `-304,6 +338,19` | 新增 13 行：5 站 3 语的「已签收、仅处于退货窗口内」状态按**前缀**匹配，归一化为该站「完成」键 | **是**（口径＝算数） |
| 5 | **任务⑧-3 Shopee 过滤赋值** | `-314,8 +361,9` | `order_df[(…)]` → `order_df = order_df[(…)].copy()` + 防呆 `assert` | **是** |
| 6 | **任务⑧-4 Lazada `status_map`** | `-340,7 +388,17` | 空 `{}` → 实测 7 个状态 | **是**（前置修复，否则 Lazada 返回空表） |
| 7 | **任务⑧-5 Lazada 过滤赋值** | `-377,8 +435,9` | 同 #5 | **是** |
| 8 | **任务⑧-6 Tiktok `status_map`** | `-398,7 +457,12` | 空 `{}` → `Completed` / `Shipped` / `Canceled` | **是**（前置修复） |
| 9 | **任务⑧-7 Tiktok 过滤赋值** | `-434,8 +498,9` | 同 #5 | **是** |
| 10 | 新增得分列变量 | `-898,6 +963,7` | 新增 `score_col_name` | 否（等价） |
| 11 | 日期字面量 → 变量 | `-933,7 +999,7` | 8 处 `"202608","20260907"` → `belonging_date, snapshot_date` | 否（等价） |
| 12 | ASIN 反查兜底 **＋ 任务⑨ 列名修复** | `-958,48 +1024,58` | ASIN 失败回退源表 sku；**同一块内**把东南亚汇总 rename 由 `数量dropna` 改为兼容 `{数量, 数量dropna}` + 断言 | 是，**+298 件 / 防整块清零** |
| 13 | 第一层 SKU 清洗 | `-1023,8 +1099,10` | 销量侧 sku `strip` + `"nan"`→NaN + 丢空 SKU 行 | 是（行数/新增 SKU 数） |
| 14 | 第二层 SKU 清洗 | `-1036,13 +1114,15` | 月份表 SKU `strip` + 三合一过滤 | 是（同上） |
| 15 | 第三层 SKU 清洗 | `-1050,7 +1130,7` | 补录表 SKU `strip` | 否（防患） |
| 16 | 得分列改变量 + 断言前置 | `-1064,8 +1144,10` | 两行 assert 提前拦截列名错 | 潜在保护 |
| 17 | KJ / KK 列代码化 | `-1144,6 +1226,20` | 年度总销量、SKU 数字由代码算 | 是，KJ 从过期缓存变正确 |
| 18 | 新增 任务②③ 整段 | `-1160,6 +1256,71` | 欧洲各国 sheet + 2021-2026 sheet 补当月 + 汇总行三格 | 新增功能 |

> **任务⑧ 的净效果**：东南亚 **15122 → 13769**、总计 **73214 → 71861**（−1353 件 ＝ 明确取消/退货），其余分组 0 变化。详见《修改指南》任务⑧ 第七节。

### 1.2 最新对照（9.15 更新：任务⑩ 落地，共 **19** 处）—— **这一张是权威表**

> ⚠️ 上面 1 / 1.1 两张表的 `+B` 行号是**任务⑩ 之前**的（插入 238 行后整体下移），已过期；
> 以本表为准。复现：`diff -u <原版> trunk/month_sale_file.py | grep '^@@'`。

| # | 主题 | diff 位置（A → B） | 一句话 | 影响本月数字？ |
| -- | -- | -- | -- | -- |
| 1 | 删除死导入 | `-10,10 +10,11` | 删 `import timedelta`；补 `import re`、`from copy import copy` | 影响"能不能跑" |
| 2 | 新增行自动套样式（任务⑦） | `-129,14 +130,47` | 超出模板行数的新增行补齐底纹/边框/字体/行高 + 扩 `auto_filter` | 否（只改样式） |
| 3 | 任务⑧-1 泰国 `status_map` | `-225,7 +259,7` | 重复键 `"สำเร็จแล้ว":"Cancelled"` → `"ยกเลิกแล้ว":"Cancelled"` | **是**（否则泰国站清零） |
| 4 | 任务⑧-2 退货窗口归一化 | `-304,6 +338,19` | 5 站 3 语「已签收、仅处于退货窗口内」按**前缀**匹配归一到「完成」键 | **是**（口径＝算数） |
| 5 | 任务⑧-3 Shopee 过滤赋值 | `-314,8 +361,9` | `order_df[(…)]` → `.copy()` + 防呆 `assert` | **是** |
| 6 | 任务⑧-4 Lazada `status_map` | `-340,7 +388,17` | 空 `{}` → 实测 7 个状态 | **是** |
| 7 | 任务⑧-5 Lazada 过滤赋值 | `-377,8 +435,9` | 同 #5 | **是** |
| 8 | 任务⑧-6 Tiktok `status_map` | `-398,7 +457,12` | 空 `{}` → `Completed`/`Shipped`/`Canceled` | **是** |
| 9 | 任务⑧-7 Tiktok 过滤赋值 | `-434,8 +498,9` | 同 #5 | **是** |
| 10 | **虾皮跨境去重（用户自己改的）** | `-743,7 +808,7` | 第 **811** 行 `other_shopee_cross_site` 去掉 `"新加坡"` → `虾皮境内` 503→**384** | **是**（119 重复计数清零） |
| 11 | 新增得分列变量 **＋ 任务⑩ 整段** | `-894,10 +959,241` | `score_col_name`（195 行新增里的 1 行）＋ 任务⑩ 列分组自动推进（4 函数 + `_replace_file` 等，共 ~238 行） | 否（不改任何单元格值） |
| 12 | 日期字面量 → 变量 | `-933,7 +1229,7` | 8 处 `"202608","20260907"` → `belonging_date, snapshot_date` | 否（等价） |
| 13 | ASIN 反查兜底 ＋ 任务⑨ 列名修复 | `-958,48 +1254,58` | ASIN 失败回退源表 sku；东南亚汇总 rename 兼容 `{数量, 数量dropna}` + 断言 | 是，**+298 件 / 防整块清零** |
| 14 | 第一层 SKU 清洗 | `-1023,8 +1329,10` | 销量侧 sku `strip` + `"nan"`→NaN + 丢空 SKU 行 | 是（行数/新增 SKU 数） |
| 15 | 第二层 SKU 清洗 | `-1036,13 +1344,15` | 月份表 SKU `strip` + 三合一过滤 | 是（同上） |
| 16 | 第三层 SKU 清洗 | `-1050,7 +1360,7` | 补录表 SKU `strip` | 否（防患） |
| 17 | 得分列改变量 + 断言前置 | `-1064,8 +1374,10` | 两行 assert 提前拦截列名错 | 潜在保护 |
| 18 | KJ / KK 列代码化 | `-1144,6 +1456,20` | 年度总销量、SKU 数字由代码算 | 是，KJ 从过期缓存变正确 |
| 19 | 新增 任务②③ 整段 | `-1160,6 +1486,79` | 欧洲各国 sheet + 2021-2026 sheet 补当月 + 汇总行三格 | 新增功能 |

> **任务⑩ 的净效果**：**数据 0 变化**——只重写 `xl/worksheets/sheetN.xml` 里 `<cols>` 与 `<sheetViews>` 两段，
> 把模板携带的「上月分组展开态」推进到当月（【{年}年】24 个块 = 1~当月平铺 / 之后折叠；
> 欧洲表 = 当月块展开、之前只留月合计列、之后整块隐藏），并把打开视图横滚到当月。
> 详细改法、实测效果、验收清单见《修改指南》任务⑩；代码意义见《意义说明》任务⑩。
> **顺带修正一处计数**：`+425/−26` 里的 6 行属于**用户自己**的 811 行改动（第 10 行），不是助手贡献（助手 = `+239/−1`）。

> ### 「diff 位置」这一列怎么读
>
> 它是 `diff -u` 的**行头（hunk header）**，完整形态长这样，文档里省掉了前后的 `@@`：
>
> ```
> @@ -958,44 +993,52 @@
> ```
>
> - `-958,44` → **原始版 A**：本块从第 **958** 行开始，共 **44** 行（958~1001）
> - `+993,52` → **当前版 B**：同一块从第 **993** 行开始，共 **52** 行（993~1044）
> - 两侧的**行数相减** = 该处的**净增行数**（52 − 44 = **+8**）
> - 两侧的**起始行号不同**，是因为**行号会漂移**：`B 行号 = A 行号 + 该行上方所有改动的净增行数`（958 + 35 = 993）
> - 所以它**不是"第几处改动"的编号**；而且本表按**主题**排，与行号先后**并不一致**（第 11 行实际在第 129 行，是最早的改动之一）
>
> 🔧 **修正记录（9.14）**：上表 `+` 侧（当前版行号）此前是**手算的近似值**，除第 1、10、11 行外全部有偏差；现已用真实 `diff -u` 逐个核对修正，`-` 侧（原始版行号）本来就正确。

---

## 2. 逐处详解

### ① 删除死导入（A:13 → B 中已无此行）

```diff
 import datetime
-import timedelta
 import dateutil
```

- **为什么**：全脚本从未使用 `timedelta`，但它是**顶层 import**，任何环境下都会让脚本在启动阶段就崩：
  ```
  Traceback (most recent call last):
    File "...\month_sale_file.py", line 13, in <module>
      import timedelta
  ModuleNotFoundError: No module named 'timedelta'
  ```
  实测：直接运行 A 就是这个报错（所以 A **从未在本机跑通过**）。
- **影响**：从"启动即崩"变成"能跑"。

### ② 日期字面量 → 变量（8 处调用点）

```diff
-            last30d_amazon_order_df = get_belonging_amazon_order_df(entity, sale_group,"202608", "20260907")
+            last30d_amazon_order_df = get_belonging_amazon_order_df(entity, sale_group,belonging_date, snapshot_date)
```

同样改法的还有 7 处：`get_belonging_shopee_order_df`(×2)、`get_belonging_lazada_order_df`(×2)、`get_belonging_tkshop_order_df`、`get_belonging_kson_offline_order_df`、`get_belonging_indonesia_warehouse_offline_order_df`。

- **为什么**：A 里 899-900 行**已经定义了** `belonging_date/snapshot_date`，但调用点还在用写死的字符串 → 改日期要改 9 处，漏一处就跑错月份的数据。
- **影响**：等价重构，本月数字不变；但"下月只改 899-901"这条 SOP 才有前提。

### ③ 新增 `score_col_name` 变量

```diff
     belonging_date = "202608"
     snapshot_date = "20260907"
+    score_col_name = "7月销量得分"  # 得分列名以商品资料.xlsx实际列名为准；数据部更新为"8月销量得分"后改这里重跑
```

- **为什么**：得分列名每月由数据部更新（7月→8月），A 里这个列名在代码里出现 3 次，得逐处找。
- **影响**：等价（本月仍是"7月销量得分"）。


### ④ ASIN 反查兜底 ★本轮核心修复

```diff
-    amazon_sale_df["sku"] = amazon_sale_df["asin"].map(asin_sku_map)
+    # 9.12新增：ASIN反查会覆盖源表sku列，映射失败就没退路（欧洲站独立ASIN未登记进商品资料时会整批丢单）
+    # 原逻辑只有一行 amazon_sale_df["sku"] = amazon_sale_df["asin"].map(asin_sku_map)，现改为"ASIN优先 + 源表sku兜底"
+    raw_sku = amazon_sale_df["sku"].astype(str).str.strip()
+    amazon_sale_df["sku"] = amazon_sale_df["asin"].map(asin_sku_map)
+    fallback_sku = raw_sku.str.replace(r"[A-Za-z]{1,3}$", "", regex=True)
+    amazon_sale_df["sku"] = amazon_sale_df["sku"].fillna(fallback_sku[fallback_sku.isin(sku_asin_map.index)])
+    lost_order = amazon_sale_df.loc[amazon_sale_df["sku"].isna(), ["仓库","站点","asin"]].drop_duplicates()
+    if lost_order.shape[0] > 0:
+        print(f"【警告】{lost_order.shape[0]}组订单ASIN无法对应商品资料，将被丢弃：\n{lost_order.to_string()}")
```

- **病灶**：A 那一行**既是"反查"又是"覆盖"**——把源表本来就正确可用的 `sku` 列（欧洲站写作 `G261IT`）直接覆盖掉。一旦 ASIN 查不到（欧洲站独立刊登 `B0BTNC4PN4` 没登记进商品资料），覆盖后的值是 NaN，源表的 `G261IT` 已经没了，下游 `dropna()` 就把这批订单当无效行丢掉。
- **结果**：G261 的欧洲 8 月整批 217 件 + 英国 78 件 + 澳洲 3 件（C005/G136）消失。
- **兜底逻辑**：ASIN 能命中就照旧（占 99.5%）；命不中才用源表 sku 剥掉尾部 1~3 位字母，**且必须剥完能命中商品资料的 SKU 集合**才采用（避免把 `G43J` 误配到 `G43`）；再命不中就保留 NaN 并**打警告日志**，不再静默丢单。
- **影响**：本月 +298 件（J 欧洲 28098→28315、V 英国 7940→8018、BR 澳洲 1948→1951）。


### ⑤⑥⑦ 三层 SKU 清洗

**第一层 —— 销量侧（A 无 → B 新增）**

```diff
-    total_sku_sale_df = pd.concat([amazon_sale_df,se_a_sku_sale_df,other_sku_sale_df]).copy()#连接亚马逊和东南亚
-
+    total_sku_sale_df = pd.concat([amazon_sale_df,se_a_sku_sale_df,other_sku_sale_df]).copy()#连接亚马逊和东南亚销量表
+    # 第一层清洗：销量侧sku规范化（9.12新增，去除空格；"nan"转回NaN防止混入新SKU）
+    total_sku_sale_df["sku"] = total_sku_sale_df["sku"].astype(str).str.strip().replace("nan", np.nan)
+    total_sku_sale_df = total_sku_sale_df[total_sku_sale_df["sku"].notna() & (total_sku_sale_df["sku"] != "")].copy()  # 丢弃空白SKU的订单行（美客多存在此类脏数据）
```

- **为什么**：`astype(str)` 会把 NaN 变成**文本 `"nan"`**，它就躲过了下游的 `dropna()`、被当成一个叫 "nan" 的新 SKU 追加进表；另有美客多来源的空 SKU 订单行需要丢掉。

**第二层 —— 月份表（A 只过滤 "SKU"）**

```diff
-    month_sale_df = month_sale_df[month_sale_df["SKU"] != "SKU"].copy()
+    # 第二层清洗：月份表SKU规范化 + 空行/合计行过滤（9.12修订）
+    month_sale_df["SKU"] = month_sale_df["SKU"].astype(str).str.strip()
+    month_sale_df = month_sale_df[~month_sale_df["SKU"].isin(["", "nan", "SKU"])].copy()
```

- **为什么**：一行过滤三件事——空串、"nan" 文本、模板第 3 行合计行（它的 A 列文本就是 `"SKU"`）。A 只挡住了第三种，且不 `strip`，带空格的 SKU 会匹配不上。

**第三层 —— 补录表**

```diff
-    new_sku_month_sale_df["SKU"] = missed_sku_df["sku"].dropna().drop_duplicates().astype(str)
+    new_sku_month_sale_df["SKU"] = missed_sku_df["sku"].dropna().drop_duplicates().astype(str).str.strip()
```

- **为什么**：补录表最终会按列名 concat 回主表，SKU 带空格就会与新行对不上。属防患。

### ⑧ 得分列改变量 + 断言前置

```diff
-    product_info_map = product_info_df.set_index("SKU")["7月销量得分"]
-    new_sku_month_sale_df["7月销量得分"] = new_sku_month_sale_df["SKU"].map(product_info_map)
+    assert score_col_name in product_info_df.columns, f"商品资料.xlsx 找不到列 {score_col_name}"
+    assert score_col_name in month_sale_df.columns, f"模板【2026年】KH 列表头不是 {score_col_name}，请先同步再运行"
+    product_info_map = product_info_df.set_index("SKU")[score_col_name]
+    new_sku_month_sale_df[score_col_name] = new_sku_month_sale_df["SKU"].map(product_info_map)
```

- **为什么**：`new_sku_month_sale_df` 是**按列名** concat 到模板 df 尾部的。列名一旦不一致（数据部还没把"7月"改成"8月"），会另立一列、得分整块串位，**而且不报错**。断言让它第一时间大声失败。
- 9.12 实际触发过一次（数据部列名未同步），属预期防护。


### ⑨ KJ / KK 列代码化（A 完全没有）〔2026-10-07 修正：KJ 改为按位置写回〕

```diff
+    # 9.12新增：KJ列代码化：当年总销量 = 总计块已存在月列之和（公式原型 =SUM(JS4:KD4)）
+    # 列名必须与模板 KJ 列表头一致（模板里那是公式 =$B$1&"总销量"），故按年份拼；".23"= "总计"在 month_sale_group 中的下标
+    zongji_cols = [f"{m}月.23" for m in range(1, 13)
+                   if f"{m}月.23" in processed_month_sale_df.columns]
+    processed_month_sale_df[f"{cur_acct_cycle_date.year}年总销量"] = (
+        processed_month_sale_df[zongji_cols]
+        .apply(pd.to_numeric, errors="coerce")  # 停售"00"和空值统一转数值，"00"→0
+        .fillna(0)
+        .sum(axis=1)
+    )
+    # KK列代码化：提取SKU中的数字（公式原型 =1*MID(A4,...)数组公式），to_numeric保持与原缓存值一致的数值类型
+    processed_month_sale_df["Unnamed: 296"] = pd.to_numeric(
+        processed_month_sale_df["SKU"].astype(str).str.extract(r"(\d+)", expand=False), errors="coerce")
```

- **为什么**：这两列原本是 Excel 公式（KJ=`=SUM(JS4:KD4)`、KK=`=1*MID(A4,...)` 数组公式）。`fill_excel_from_df` 写回时会**把模板里的旧缓存值也写回**——于是 KJ 停在 7 月的过期数上。
- **修的 bug**：G287（第 4 行）KJ 从 `3128` → **`3591`**（漏了 8 月的 463）。
- **细节**：停售行月列是**文本 `"00"`**（不是数字 0），所以必须 `to_numeric(errors="coerce")` 再 fillna(0)。
- **年份化**：列名写成 `f"{cur_acct_cycle_date.year}年总销量"`。2026 年拼出来和原来的 `"2026年总销量"` 一字不差（零行为变化）；2027 年自动跟到"2027年总销量"，不会另立新列把 KJ/KK 整块挤歪。
- ⚠️ **2026-10-07 修正（202609 期实测，上面 diff 片段已过时）**：年份化仍以"pandas 能读到模板 KJ 表头字符串"为前提。本期模板沿用上期成品、该表头（公式 `=$B$1&"总销量"`）**无缓存值** → pandas 读成 `Unnamed: 295` → **年份化列名照样匹配不上** → 仍另立新列 → **整列右移一格**：
  `【2026年】KJ(296)` 停在 8 月（**532,372**）、正确值落到**表外 `KL(298)`**（595,009）。
  **现写法**（主脚本 **L1588~1613**）：把值**按位置写回 KJ 既有列**（`_cols[_cols.index("状态") + 1]`）+ `assert _kj_letter == "KJ"`；`KK` 仍按 `Unnamed: 296` 写（模板 KK 表头本就为空）。修复后 `KJ` 数据行合计 = **595,009**。


### ⑩ 新增「任务②③」整段（A 完全没有 → B 末尾 +55 行；**2026-10-07 再补 (a′) 「追加缺行 SKU」段**）

A 跑完只生成【2026年】一张 sheet 的当月列，另两张 sheet 的 8 月**是空的**。B 在 `fill_excel_from_df` 之后补写：

**（a）【2026年-欧洲各国每月销售情况】补当月块**

```python
    # ===== 任务②：欧洲各国 sheet 补当月（必须在 fill_excel_from_df 生成 result 之后执行）=====
    sale_col = f"{cur_acct_cycle_date.strftime('%Y%m')}销量"
    europe_df = total_sku_sale_df2[total_sku_sale_df2["仓库"] == "欧洲"]
    pivot = europe_df.groupby(["SKU", "站点"])[sale_col].sum().unstack(fill_value=0)
    wb = load_workbook(result_path)
    ws = wb[f"{cur_acct_cycle_date.year}年-欧洲各国每月销售情况"]
    sku_row = {str(ws.cell(r, 1).value).strip(): r for r in range(4, ws.max_row + 1)}
    blk = 3 + (month_offset - 1) * 11   # 当月块起始列：1月→3(C) 8月→80(CB) 9月→91(CM)；每月块宽11列(10国+合计)
    country_col = {"比利时": blk, "土耳其": blk + 1, "德国": blk + 2, "西班牙": blk + 3, "法国": blk + 4,
                   "意大利": blk + 5, "荷兰": blk + 6, "波兰": blk + 7, "瑞典": blk + 8, "爱尔兰": blk + 9}
    for sku, row in pivot.iterrows():
        r = sku_row.get(str(sku).strip())
        if r:
            for country, c in country_col.items():
                v = row.get(country, 0)
                if v:
                    ws.cell(r, c, int(v))
```

> 【2026年】的欧洲月列是 `SUMIF` 公式，引用本 sheet 的**当月合计列**取数——所以本 sheet 不补，重算后【2026年】的欧洲 8 月也是空的。这就是必须补这张 sheet 的原因。  
> 月份块起始列用 `blk = 3 + (month_offset-1)×11` 推导（8月=80=CB，9月=91=CM），不写死。

**（a′）任务②-b：模板缺行 → 末尾追加新 SKU 行 〔2026-10-07 新增，L1652~1701〕**

上面那段 `if r:` 的问题是：`sku_row` 只认模板预置的 497 行，**缺行的 SKU 整行静默跳过**（欧/北美 sheet 行数常年落后于【2026年】）。
新增段落在同一个 `ws` 上追加行：

```python
    _missing = [s for s in pivot.index if str(s).strip() not in sku_row]
    if _missing:
        _ref_row = max(sku_row.values())          # 参考行 = 最后一个已有数据行
        for _sku in _missing:                     # 逐个追加
            _r += 1
            for _c in range(1, ws.max_column + 1):
                _src, _dst = ws.cell(_ref_row, _c), ws.cell(_r, _c)
                if _src.has_style: _dst._style = copy(_src._style)      # 同簿复制样式
                if isinstance(_v := _src.value, str) and _v.startswith("="):
                    _dst.value = Translator(_v, origin=_src.coordinate).translate_formula(_dst.coordinate)
                elif isinstance(_v, ArrayFormula):                      # ER 列数组公式，按行重建
                    _dst.value = ArrayFormula(_dst.coordinate, Translator(_v.text, origin=...).translate_formula(...))
            ws.cell(_r, 1).value = _sku_s                       # A = SKU
            ws.cell(_r, 2).value = _model_map.get(_sku_s)       # B = 型号（商品资料）
            ws.cell(_r, _status_col).value = _status_map.get(_sku_s)   # 状态
            for _country, _c in country_col.items():            # 当月 10 国值
                ...
        ws.auto_filter.ref = f"{...}{_r}"                       # 筛选范围同步扩
```

> 依赖 `from openpyxl.formula.translate import Translator` 与 `from openpyxl.worksheet.formula import ArrayFormula`（L10~11 新增导入）。
> **不是 A 也没有、也不是 B 原有的**——是 2026-10-07 才补的第三块（见《最新代码修改.md》§5.4）。
> ⚠️ 行 3 的**月列**合计 `=SUM(x4:x4907)` 覆盖新行；但 `EF~EO` 只 `SUM(x4:x470)`（数据部模板老毛病）→ 新行不进这 10 格。

**（b）【2021-2026年销量情况】补 4 个分组的当月 + 同比/环比**

```python
    # ===== 任务③：2021-2026年销量情况 sheet 补当月（行/列全部由归属月份推导，换月不用改代码）=====
    def group_month_total(platform=None, sites=None):
        cond = pd.Series(True, index=total_sku_sale_df2.index)
        if platform is not None:
            cond &= total_sku_sale_df2["平台"] == platform
        if sites is not None:
            cond &= total_sku_sale_df2["站点"].isin(sites)
        return int(total_sku_sale_df2.loc[cond, sale_col].sum())

    r_cur  = month_offset + 1                  # 当月行：第2~13行 = 1~12月 ⇒ 8月→9行、9月→10行、12月→13行
    r_prev = month_offset                      # 上月行（今年环比的分母）
    col_y  = cur_acct_cycle_date.year - 2019   # 当年列：2021→B(2) 2022→C(3) … 2026→G(7) 2027→H(8)

    cur_vals = [
        group_month_total(),                                                                  # 总计 = 全渠道
        group_month_total(platform="亚马逊"),                                                  # 亚马逊
        group_month_total(sites=["菲律宾", "印尼", "泰国", "马来西亚", "新加坡", "越南"]),        # 东南亚
        group_month_total(platform="亚马逊", sites=["美国", "加拿大", "墨西哥"]),                # 北美（口径已验证：7月AK8=8801=美+加+墨）
    ]

    ws2 = wb["2021-2026年销量情况"]
    for k, val in enumerate(cur_vals):
        c = col_y + 10 * k                      # 四个分组块首列：G(7) / Q(17) / AA(27) / AK(37)
        ws2.cell(r_cur, c).value = val                                        # 当月值
        ws2.cell(r_cur, c + 1).value = val / ws2.cell(r_cur, c - 1).value - 1   # 同比增长率 = 当月 / 去年同月
        if month_offset > 1:                    # 1月时 r_prev=1 是表头（字符串），不写环比
            ws2.cell(r_cur, c + 2).value = val / ws2.cell(r_prev, c).value - 1  # 今年环比 = 当月 / 今年上月
```

> 四个分组块的**年列**首列公式：`col_y + 10k`（k=0..3 → G/Q/AA/AK），同比在 +1、环比在 +2。  
> 1 月跑时 `r_prev = 1` 是表头（字符串），所以环比那行加了 `if month_offset > 1` 保护。这是 **A 里完全没有的边界处理**。

**（c）汇总行三格（G14/I14/J14）—— 模板里它们是静态值**

```python
    # 汇总行（第14行）：模板里 G14/I14/J14 都是静态值（不是公式）→ 用表内月值重算，幂等（重跑不叠加）
    y_sum = 0
    for r0 in range(2, 14):                     # 第2~13行 = 1~12月
        cv = ws2.cell(r0, col_y).value
        if isinstance(cv, (int, float)):
            y_sum += cv
    ws2.cell(14, col_y).value = y_sum                                          # 当年累计 = 当年列 1~12月之和
    ws2.cell(14, col_y + 2).value = y_sum / ws2.cell(14, col_y - 1).value - 1  # 汇总同比（模板把它放在 +2 列，+1 列留空）
    ws2.cell(14, col_y + 3).value = sum(                                       # 历年总计 = SUM(首年列..当年列)
        v for v in (ws2.cell(14, c0).value for c0 in range(2, col_y + 1))
        if isinstance(v, (int, float)))
```

- **为什么**：模板里 G14（当年累计）就是硬编码 `460511`（= `SUM(G2:G8)`，1~7 月累计），**不是公式、脚本也从不写它** → 跑完仍显示 460511，而应该是 533725。I14（汇总同比）、J14（历年总计）同理。
- **写法**：用"**读表内月值重算**"而不是写 `=SUM(...)` 公式——openpyxl 写的公式**没有缓存值**（打开是空的、要 Ctrl+Alt+F9），而重算式写法**存盘即有值**，且**幂等**（重跑不会把 533725 加成 1067450）。
- **列位置坑**：第 14 行**不是** 3 列一组。模板实测是 `年值 col_y` / `同比 col_y+2`（`+1` 列 H 留空）/ `历年总计 col_y+3`——照抄模板口径，别照抄 H1 表头位置。

**（d）末尾新增日志**：跑完打印 `【月度参数】…`、`【2021-2026年销量情况】…月：总计/亚马逊/东南亚/北美`、`【汇总行】…`，以及成品绝对路径 `✅ 结果文件（成品，注意不是data目录下的模板）`。每月跑完扫一眼日志就能判断有没有错位。

---

### ⑪ `fill_excel_from_df` 给新增行套样式 + 保存失败提示（A 没有）

- **位置**：`-129,14 +130,47`（函数体内）+ `-10,10 +10,11`（顶部 import）。
- **A 的做法**：函数只做三件事——加载模板、写值、`wb.save()`。**从不处理样式**。
- **问题**：【2026年】的行数每月增长（本月 13 条新增 SKU → 第 536~548 行），而模板的（带样式的）数据区只到第 535 行。openpyxl 在第 536 行及以后新建单元格时用 **Excel 默认格式** → 新增行**无底纹、无边框、字体变宋体 11、行高变默认**，且 `A3:KK535` 的自动筛选范围也不含新行。**每月都要手工用格式刷补一遍。**
- **B 的做法**：写完数据后，把"模板最后一条数据行"的样式**按列复制**到所有新增行——
  - 参考行取法是"**扫首列取最后一个非空行**"，而不是 `ws.max_row`（防止模板尾部有多余的空格式行）；
  - 样式整体走 `cell._style = copy(src._style)`（openpyxl 内部 `WorksheetCopy` 的做法），底纹/边框/字体/对齐/数字格式/保护**一次全带**；
  - 同步复制**行高**，并把 `auto_filter.ref` 的末行改成 `last_row`。
- **顺带修的**：`wb.save()` 包了 `try/except PermissionError`。因为脚本末尾有 `os.startfile(result_path)`，**跑完不关 Excel，下一轮就会因文件被占用而保存失败**，而原报错只有一句 `Permission denied: 'result\...'`。
- **实测（B）**：日志 `🎨 新增行 536~548 已套用第 535 行样式（13 行 × 297 列，含底纹/边框/字体/行高）`；成品 536~548 行 × 297 列与第 535 行比对**不一致 0 / 3861**；行高 `None → 30.0`；筛选范围 `A3:KK535 → A3:KK548`。
- **为什么**：把"每月重复、且完全可由规则推导"的人工动作交给代码。注意它**只影响样式**——修复前后 7 个 sheet 逐格比对（公式按文本比、浮点按 1e-12 容差）**真实差异 0**。

---

## 3. 行为差异汇总

**先说清一件事**：A 本身**跑不起来**（第 13 行死导入），所以下面的基线用的是 **9.11 那版成品**——它是 A 的最近后继版本（已补上任务0 修改①②③④）跑出来的。A 相对 9.11 那版还**额外缺**两块能力，见 3.2。

### 3.1 「9.11 那版成品」→「当前版 B」（9.12 实跑）

| 指标                             | 9.11 那版成品                             | 当前版 B                                      | 差                   |
| ------------------------------ | ------------------------------------- | ------------------------------------------ | ------------------- |
| 【2026年】J 欧洲 8 月                | 28098                                 | **28315**                                  | +217                |
| 【2026年】V 英国 8 月                | 7940                                  | **8018**                                   | +78                 |
| 【2026年】BR 澳洲 8 月               | 1948                                  | **1951**                                   | +3                  |
| 【2026年】DZ 亚马逊 8 月              | 56155                                 | **56453**                                  | +298                |
| 【2026年】HR 东南亚 8 月              | 15122                                 | 15122                                      | 0                   |
| 【2026年】JZ 总计 8 月               | 72916                                 | **73214**                                  | +298                |
| 【2026年】KJ 2026 年总销量            | 533427                                | **533725**                                 | +298                |
| 【2026年】表内数据行 / 新增 SKU          | 544 行 / 12 条                          | **545 行 / 13 条**                           | +1 行（新 SKU `G136`）  |
| 【欧洲各国】8 月块 非空格 / 10 国合计        | 2579 / 27856                          | **2588 / 28073**                           | +9 格 / +217         |
| 【2021-2026】G9 / Q9 / AA9 / AK9 | 72916 / 56155 / 15122 / 9245          | **73214 / 56453 / 15122 / 9245**           | +298 / +298 / 0 / 0 |
| 【2021-2026】汇总行 G14 / I14 / J14 | 460511 / -0.449634711982379 / 4053257 | **533725 / -0.3621352946027246 / 4126471** | 三格全部刷新              |

> 差值自洽性：**217 + 78 + 3 = 298** = DZ3 增量 = JZ3 增量 = KJ3 增量。

### 3.2 A 相对「9.11 那版」还额外缺什么

| 能力                   | 最原始版 A              | 9.11 那版               | 当前版 B   |
| -------------------- | ------------------- | --------------------- | ------- |
| 能否启动                 | ❌ 死导入，直接崩           | ✅                     | ✅       |
| 8 处日期字面量             | ❌ 写死 `"202608"`     | ✅ 已变量化                | ✅       |
| 得分列名                 | ❌ 写死 3 处            | ✅ 变量 + 断言             | ✅       |
| 三层 SKU 清洗            | ❌ 只有 1 层            | ✅                     | ✅       |
| **ASIN 反查兜底**        | ❌ 无（**静默丢单**）       | ❌ 无                   | ✅       |
| KJ / KK 列代码化         | ❌ 无（KJ 停在过期缓存）      | ✅                     | ✅       |
| 【欧洲各国】8 月块           | ❌ **没有这段代码 → 该块空白** | ✅ 但**写死 8 月列 80..89** | ✅ 按月份推导 |
| 【2021-2026】当月值 + 增长率 | ❌ **没有这段代码**        | ✅ 但**写死第 9 行**        | ✅ 按月份推导 |
| 【2021-2026】汇总行三格     | ❌ 无                 | ❌ 无（停在 460511）        | ✅ 重算写入  |

**未受影响的**：其余分组（美国 / 加拿大 / 日本 / 阿联酋 / 沙特 / 墨西哥 / 印度 / 京东 / eMAG / TEMU / 美客多等）的 8 月值，当前版与 9.11 那版**逐格一致**。

---

## 4. 当前版仍未解决的（提醒，非代码 bug）

1. ~~**欧洲各国 sheet 有 13 个新 SKU 没有行位**（实测该 sheet 共 **48 个**主表 SKU 无行位，另 35 个当月为 0）→ 它们的欧洲销量写不进去，10 国块合计 28073 与 J3 28315 差 **242**。需要把新 SKU 追加到该 sheet 尾行。~~
   ✅ **2026-10-07 已修**（主脚本 **L1652~1701** 任务②-b：末尾追加缺行 SKU + `Translator` 平移公式 + `ArrayFormula` 重建 ER + A/B/状态 取自商品资料）。202609 补 **28 行**，9 月合计 **24,517→24,986 = 主表欧洲块**。残留：新行历史月无数据源留空；行 3 `EF~EO` 只 `SUM(x4:x470)`。
2. **商品资料缺 G261 的欧洲 ASIN `B0BTNC4PN4`**（6 个 sheet 均无）——现在靠兜底救回来了，但根因是主数据不全，实测商品资料 **ASIN 空缺 129/625**（61 个在主表），建议数据部补登。
3. **G578/G579/G580 的"7月销量得分"列被误填成 `XE`**（疑似系列名串列），代码如实映射。
4. **G001 缺属性**，需人工补录。
5. **公式无缓存值**：所有 `=SUM()` 类公式经 openpyxl 存盘后没有缓存值，Excel 打开需 Ctrl+Alt+F9 重算才显示（脚本不写第 1~3 行，这是设计如此）。
6. **跨年提醒**：`col_y = 年 - 2019` 假设【2021-2026年销量情况】的首年恒为 2021；若数据部把窗口平移成"2022-2027"，该公式与 sheet 名都要跟着改。

---

## 5. 一句话总结

A 是"**只填【2026年】一张 sheet、且会静默丢单**"的版本；  
B 在 A 基础上补了 **ASIN 兜底**（找回 298 件）、**KJ/KK 代码化**（修过期缓存）、**另两张 sheet 的补写**（含汇总行三格）、**新增行自动套样式**（免每月手工刷格式），  
并把"8 月"这个事实从**坐标**里彻底消掉（日期、月份块列、行号、年份列、KJ 列名全部推导），最后把保存失败这类"看不懂的报错"改成人话——  
于是"下月只改 899-901 三行"才真正成立，且跑完**不需要再手工补任何东西**。

---

## 6. 遗留问题的解决思路（9.14 实测调研；**其中「欧洲各国缺行」已于 2026-10-07 落实**，其余仍未改）

上文「4. 当前版仍未解决的」6 条的**解决思路**已单独成章，见：
- 《新增代码的意义说明-业务层面与代码层面.md》文末「附：5 个遗留问题的解决思路」（含完整论证与优先级）
- 《修改指南-2026年8月补齐.md》文末「附：5 个遗留问题的解决思路」（含落实抓手速查表）

本次复核顺便校正两处事实：欧洲 sheet 无行位的**不是 12 个**（实为 48 个缺行位 / 13 个有销量，合计 242）；
公式"无缓存值"**不等于"打开是空的"**——workbook.xml 里 openpyxl 已写 `fullCalcOnLoad="1"`，
Excel 打开时会自动重算（实测 J3 直读 28315），真正读到空值的是 WPS 老版 / 预览器 / `data_only=True`。
另：本文件涉及成品数值处均为**任务⑧ 之前**口径，任务⑧⑨ 后的权威值见《意义说明》一页速览
（总计 8 月 71861、G14 532372、J14 4125118）。

---

## 2026-10-16 · solo 化两处代码自动化（数据部只有一人后的换月减负）

📝 修订记录 / 最近一次修订：2026-10-16

**背景**：用户告知「数据部就我自己了」——以前等数据部同事配合的环节（更新商品资料得分列、
给本期新版模板）全部落到一人头上。本轮把两个「等别人喂」的点改成脚本自动，换月必改项再减两项。

### 改动 1：得分列自动探测（`trunk/month_sale_file.py`）

- **原逻辑**：`score_col_name = "8月销量得分"` 写死在 `__main__`，商品资料更新得分列后
  必须人工回来改这行（原注释原文：「数据部更新为'9月销量得分'后改这里重跑」）。
- **新逻辑**（2026-10-16 修订）：置 `None` → 商品资料读取后自动探测——
  ① 优先「归属月的上一个月」得分列（业务规律，202609 → 8月销量得分）；
  ② 找不到则退化取月份最大的「N月销量得分」列并打警告；
  ③ 一个都没有 → 人话报错（列出现有列名），不再裸 assert 一句「找不到列」。
  显式写死列名仍然优先，可随时覆盖自动探测。
- **单测**：202609+有 8 月列 → `8月销量得分`（与原写死值一致，**零行为变化**）；
  202611+有 9/10 月列 → `10月销量得分`（上一个月 ✓）；只有旧列 → 退化+警告；
  无得分列 → 报错列出全部列名。
- **位置**：`__main__` 参数区（原 L1443）+ 商品资料读取后的探测块。

### 改动 2：次表表头动态月（`trunk/southeast_asia_sale_file.py`）

- **原逻辑**（2026-09-15 埋的坑）：`f"Aug {year} Total"` / `"Jan-Aug {year} Cumulative"`
  月份写死，换月重跑表头不跟随归属月——202609 期菲/印/泰三次表表头全是 `Aug 2026`，数据却是 9 月
  （10-15 已用探针 `probe_fix_sub_headers.py` 就地改过，脚本未修则换月必复现）。
- **新逻辑**（2026-10-16 修订）：`strftime("%b")` 取归属月缩写；1 月特判
  （当月=累计，累计列直接 `Jan <年> Cumulative`，不产出 `Jan-Jan`）。
- **单测**：202609 → `Sep 2026` / `Jan-Sep 2026`；202612 → `Dec 2026` / `Jan-Dec 2026`；
  202701 → `Jan 2027` / `Jan 2027`。

### 顺带修正一个此前的错误结论

- 旧记录称「菲律宾 sheet 缺行 trunk 不补（自动补行只覆盖欧洲 sheet）」——**不准确**。
  东南亚脚本（L273~284）早就有补行逻辑：`sold_skus` 里不在 `sku_row_map` 的 SKU →
  复制末行样式补行位。202609 期 G569 丢失的真正原因是 **trunk 数据源里没有那 11 件线下单**
  （299 线下不在旧三报告里）→ G569 不在 `sold_skus` → 无从补起。
  ⇒ **补行逻辑不缺，缺的是数据源**；这正是 11 月必须把两段式 + 299 线下口径固化进脚本的原因。

### 仍未自动化 → 已全部闭环（2026-10-17）

- **两段式口径固化**：监控表（段A）+ WMS 新报告（段B）+ G554 补回 + 299 线下，
  目前只活在探针里。11 月若 10 月新报告不满量（10/1~10/13 有缺口），需产品化为
  `trunk/工具-菲律宾两段式重算.py` 一条命令出结果，否则 trunk 重跑会覆盖回旧口径。
  ✅ **2026-10-17 用户拍板：不需要产品化**——「以后 mws 文件是完整一月，就不用单独分了，
  区分好线下单就行」。线下单区分早已在 trunk（`PH_MANUAL_ORDER_TYPES` / KSON 渠道）；
  11 月起满量新报告走 `USE_PH_NEW_ORDER_REPORT=True`（三步见
  `核查-菲律宾新订单报告能否一敌多.md` §八）。
- **模板表头停上期**：模板=复制上月成品时，【2026年】KH 表头仍是上期得分列名，
  L1653 assert 会崩（要求表头与 score_col_name 同名）。solo 后每月首次跑需手工改这一格，
  或后续做「rename + 成品表头重刷」自愈。
  ✅ **2026-10-17 已做**——rename + 成品表头自动重刷，三层防护，测试 28 项全绿。详见下节。

---

## 2026-10-17 · diff：KH 表头自愈（`trunk/month_sale_file.py`）

📝 修订记录 / 最近一次修订：2026-10-17

| # | 位置 | 旧 | 新 |
|---|---|---|---|
| 1 | `get_month_sale_df()` 签名 | `(belonging_date, snapshot_date)` | `(belonging_date, snapshot_date, score_col_name=None)`；函数体内先 `heal_score_col_header_file(file_path, f"{belonging_date[0:4]}年", score_col_name)` 再 pandas 读；读完后若 `score_col_name` 不在列里 → 按 `_SCORE_HEADER_RE` 把旧表头列 rename（内存兜底） |
| 2 | 新增函数 | — | `_SCORE_HEADER_RE` / `_find_score_header_col(ws)` / `rename_score_col_in_ws(ws, score_col_name)`（返回 `(旧, 新, 列号)` 或 None）/ `heal_score_col_header_file(file_path, sheet_name, score_col_name, save=True)`（返回 `skip/ok/renamed/locked/missing`） |
| 3 | `__main__` 调用处 | `get_month_sale_df(belonging_date, snapshot_date)` | `get_month_sale_df(belonging_date, snapshot_date, score_col_name)` |
| 4 | 补充信息处 assert | `模板【2026年】KH 列表头不是 {score_col_name}，请先同步再运行` | `模板【2026年】行2 找不到得分列「{score_col_name}」——表头自愈未命中，请检查 …/月销售总表.xlsx 的列布局` |
| 5 | 任务②（EP 重刷块之后） | — | `rename_score_col_in_ws(wb[f"{cur_acct_cycle_date.year}年"], score_col_name)` 成品侧兜底，随任务②③共用 wb 的唯一一次 save 生效 |
| 6 | `工具-换月资料体检.py` [D] 区 | KH 表头不一致 = ❌致命（「否则主脚本会崩」） | ℹ️提醒（「主脚本会自动 rename 自愈（2026-10-17 起不用手工改）」）；商品资料缺列仍 ❌致命 |
| 7 | 新增测试 | — | `trunk/工具-KH表头自愈测试.py`（7 组 28 项；T3 用真实 202609 模板 sha256 证明零回归） |

> 兼容性：`get_month_sale_df` 第三参默认 `None` → 不传时行为与旧版逐字等价（旧调用/探针不受影响）。

### 2026-10-17 · 任务⑯：商品资料得分列自动生成（新增工具，不动生产脚本）

| # | 位置 | 改前 | 改后 |
|---|---|---|---|
| 1 | `trunk/工具-更新商品资料得分列.py` | 不存在 | 新增约 330 行：读成品月销总表 → 按 A/B/C/D/E + XA~XE 规则打分 → `--verify` 复刻比对 / `--apply` 备份后写回 + 自检 |
| 2 | `商品资料.xlsx` 列数 | 到 P 列 | 追加 **Q 列「9月销量得分」**（431 行有分；H/O 两列 595 个公式未受影响，因只追加不插列） |
| 3 | 得分列维护方式 | 数据部人工在 Excel 里维护 | 工具生成（规则原文 = A1 注释；202609 自验平级 599/638=93.9%） |
| 4 | `trunk/.workbuddy/backup_scorecol/` | 不存在 | `--apply` 自动落备份（写回前快照） |

> 生产脚本**零改动**：得分列内容变了，但列名规律没变，自动探测 + KH 自愈照旧工作。

### 2026-10-30 · diff：报表文件名别名兜底（`trunk/month_sale_file.py`）

| # | 位置 | 改前 | 改后 |
|---|---|---|---|
| 1 | `month_sale_file.py`（`get_belonging_amazon_order_df` 之前） | 无 | 新增 `REPORT_NAME_ALIASES = {"欧洲": ["欧盟"]}` + `_alias_names()` + `resolve_report_path()`，约 35 行 |
| 2 | 订单取文件 L210 | `file_path = os.sep.join([dir_path, file_name])` | `file_path = resolve_report_path(dir_path, file_name)` |
| 3 | 库存取文件 L1000 | 同上 | 同上 |

> 兼容性：原名存在时 `resolve_report_path` 第一分支直接返回，**路径与旧代码逐字相同** ⇒ 202609 行为零变化
> （实测欧洲订单仍 31,286 行）。别名只在原名缺失时启用，并打印 `ℹ️ 报表名别名命中：…` 留痕。

---

#### 4. 运行目录自愈（2026-11-02，SVN 入库加固）

| # | 位置 | 改动前 | 改动后 |
|---|---|---|---|
| 1 | `month_sale_file.py` `if __name__ == "__main__":` 之后 | 无（靠人 `cd trunk`） | 新增 9 行：CWD 不等于脚本所在目录时打印提示并 `os.chdir(__file__ 所在目录)`；已正确时为空操作 |

> 只包在 `__main__` 里 ⇒ `import month_sale_file` 的调用方不受影响。
> 连带把头部任务表行号整体 +9（≥1620 的部分），文件 2189 → 2198 行。

---

#### 5. 202610 期首次跑批修复（2026-11-02）

| # | 文件 / 位置 | 改动前 | 改动后 |
|---|---|---|---|
| 1 | `month_sale_file.py` `result_dir=` 之后 | 无（靠目录已存在） | `os.makedirs(result_dir, exist_ok=True)` + 5 行注释 |
| 2 | `southeast_asia_sale_file.py` `read_shopee_cross_by_country()` | `pd.read_excel(f)` 后直接 `next(...)` 找列名 | 先判空表（0 列/0 行）跳过并打印；数量列别名补 `"Quantity per Item"` |
| 3 | 同上，`files = glob(...)` 之后 | 只排除「新加坡」 | 先按站点白名单 `("马来西亚","越南")` 过滤，被排除的逐个打印 |

---

## 2026-11-02 增补（202610 期）

### 1) 菲律宾：`get_belonging_ph_new_order_df` 由「备用通道」转为「主通道」

- 开关 `USE_PH_NEW_ORDER_REPORT`：`False` → **`True`**（`PH_NEW_REPORT=0` 环境变量可一次性回退，不用改文件）。
- `PH_NEW_REPORT_START_DATE`：`"2026-09-13"` → **`None`**（202610 新报告满量覆盖 10/01~10/31，不再需要起点过滤）。
- 新增「跳过 KSON 线下报告」分支（约 15 行）：开关打开时不再读 `菲律宾KSON仓{月}线下订单报告.xlsx`，
  只做一次 `pd.read_excel().shape[0]` 探针，>0 就大声告警（防止换月时该文件又开始出数导致手工单双计）。
- 净效果：`site = get_belonging_shopee_order_df` 的调用组合变了（本土菲的三份旧报告 → 一份新报告），
  下游 groupby 键不变（`实体/平台/仓库/站点` 与旧路径逐行对齐），所以「菲律宾销量表」6 列结构零改动。

### 2) `get_belonging_shopee_order_df`：新增「台湾（跨境）」分支

```python
elif sale_group == "台湾" and entity == "跨境":
    order_id_col   = "订单编号"
    purchase_date_col = "订单成立时间"
    sku_col        = "商品规格货号"
    quantity_col   = "Quantity per Item"      # 台湾导出唯一用这个列名
    order_status_col = "订单状态"
    status_map = {"已完成":"Completed", "已送達":"Delivered", "已送达":"Delivered",
                  "不成立":"Cancelled"}       # ← 取消单必须显式映射，否则会被当销量
```

### 3) 三处口径集合各加 `"台湾"`

| 函数/变量 | 作用 | 不加会怎样 |
|---|---|---|
| `southeast_asia_cross_shopee_salegroup` | 决定中间产物里有没有 `站点=台湾` 的行 | 台湾 6 件整批丢失 |
| `sum_shopee_cross_site.other_shopee_cross_site` | 决定「虾皮境内」列含哪些站点 | 虾皮境内少 6 件 |
| `sum_total_southeast_asia_site.se_asia_sites` | 决定「东南亚」列含哪些站点 | 东南亚列与 SEA Total 月块和对不上（恒等链断） |
| 任务③ `group_month_total(sites=[...])` | 【2021-2026年销量情况】sheet 的东南亚行 | 汇总 sheet 比正表少 6 件（11,152→11,146） |

### 4) `southeast_asia_sale_file.py`

- `_CROSS_SITES`：`("马来西亚","越南")` → `+ "台湾"`；同时把「台湾=目的地导出应排除」的错误注释改写为纠错记录。
- 取消词表：`"Cancel|取消|ยกเลิก|Đã hủy|Batal"` → `+ "|不成立"`（台湾站取消态）。
- `SE_A_SITES`：`+ "台湾"`；`[2]` 步三站对撞断言 `+ ("台湾","台湾")`；无文件兜底 `sites` 同步 `+ "台湾"`。

---

## 2026-11-02 晚增补

### 1) `europe_sale_file.py` / `north_america_sale_file.py`
- 月度参数：`202609/20261006/9` → **`202610/20261102/10`**。
- `EXPECT` / `EXPECT_TOTAL`：清空跑通后**回填 202610 实测**（欧洲 31,077、北美 10,085）。
  用途不变：换月时断言失败，逼人先核对再改。

### 2) `工具-欧洲北美验收.py`
```python
def _expect_total(script_name):        # 新增
    src = io.open(os.path.join(HERE, script_name), encoding="utf-8").read()
    m = re.search(r"^EXPECT_TOTAL\s*=\s*(\d+)", src, re.M)
    return int(m.group(1)) if m else 0
```
`expect=31837` → `expect=_expect_total("europe_sale_file.py")`；`expect=9443` → 同理北美。
`expect` 为 0 时打印提示并跳过合计校验（不再假红）。

### 3) `month_sale_file.py` 任务总账
新增 ⑱ / ⑲ 两行；②~⑰ 及三条附例行号整体 +2；表头改为「2252 行 + 行号仍是粗略区间」。

### 4) 新增两个独立工具（不改主流程、只读）
- `trunk/工具-跨境订单状态核对.py`：订单号级 + 状态分桶 + 与成品/月销三方对撞，退出码 0/1/2。
- `trunk/工具-期别汇总打包.py`：按期的 data+result 打包 + 数据速览 + MD5 清单，幂等。
