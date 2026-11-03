# 东南亚全站订单状态 × status_map 对照表（202608 / 快照 20260907）

> 扫描工具：`D:\dev-cache\sale_file_gen\scan_se_status.py`（只读，不写任何文件）
> 覆盖：Shopee 6 站（本土 3 + 跨境 3）、Lazada 7 站（本土 2 + 跨境 5）、Tiktok 2 站（本土）
> 共 **15 个订单报告文件**，行数与脚本实际读取一致。

---

## 一、结论速览

| 项 | 结果 |
|---|---|
| 是否所有文件的每个状态都被 status_map 覆盖？ | ✅ **是，零漏网**（无 B 类未覆盖状态） |
| 有没有 map 配了但文件里不存在的？ | 有，但**无害**（只是空配置） |
| 有没有状态被错误保留/剔除？ | 无。剔除项全部是取消/退货，保留项全部是成交/在途/退货窗口内 |
| 最需要注意的一站 | **Shopee 泰国**：status_map 只写了 2 个键（注释「待补全」），靠**退货窗口前缀归一**兜住了那 4 组带日期的状态 |

---

## 二、Shopee 6 站

### 2.1 泰国（本土）— 3005 行
状态列：`สถานะการสั่งซื้อ`

| 行数 | 原文 | map 里？ | 结果 |
|---:|---|---|---|
| 2544 | `สำเร็จแล้ว` | ✔ `→Completed` | 保留 |
| 273 | `ยกเลิกแล้ว` | ✔ `→Cancelled` | **剔除** |
| 126 | `ผู้ซื้อได้รับสินค้าแล้ว … จนถึง 2026-09-09` | ✘ | 保留（前缀归一 → `สำเร็จแล้ว`） |
| 49 | … `2026-09-10` | ✘ | 保留（前缀归一） |
| 7 | … `2026-09-08` | ✘ | 保留（前缀归一） |
| 6 | … `2026-09-11` | ✘ | 保留（前缀归一） |

> ⚠️ **这是唯一一站 status_map 只有 2 个键的**（代码 260-264 行，注释还留着「待补全」）。
> 之所以没出事，是因为 343-349 行的 `return_window_prefix_map` 用
> `ผู้ซื้อได้รับสินค้าแล้ว` 做前缀匹配，把 4 组（188 行）带日期后缀的状态全部归一到 `สำเร็จแล้ว`。
> **【建议】** 把备注改掉、或在 map 里补一句说明，避免以后误删 `return_window_prefix_map` 时泰国整站静默丢 188 行。

### 2.2 印尼（本土）— 1133 行
状态列：`Status Pesanan`

| 行数 | 原文 | 结果 |
|---:|---|---|
| 918 | `Selesai` | 保留 `→Completed` |
| 150 | `Batal` | **剔除** `→Cancelled` |
| 23+12+10+6+2=**53** | `Pesanan diterima, namun Pembeli masih dapat mengajukan pengembalian …` | 保留（前缀归一 → `Selesai`） |
| 7 | `Telah Dikirim` | 保留 `→Delivered` |
| 5 | `Sedang Dikirim` | 保留 `→Shipping` |

> 印尼是唯一状态语义最全的一站：完成/取消/发货中/已发货 **四类齐全**。

### 2.3 菲律宾（本土）— 2269 行
状态列：`Order Status`

| 行数 | 原文 | 结果 |
|---:|---|---|
| 1857 | `Completed` | 保留 |
| 309 | `Cancelled` | **剔除** |
| 54+17+14+9+2+1+1=**98** | `Order Received, but buyer can raise return/refund until <日期>` | 保留（前缀归一 → `Completed`） |
| 3 | `Delivered` | 保留 |
| 2 | `Shipping` | 保留 |

### 2.4 新加坡（跨境）— 132 行
状态列：`Order Status`。`Completed` 77 / `Delivered` 31 / **`Cancelled` 18（剔除）** / 退货窗口 6 行（前缀归一）。
> map 里的 `Shipping` 本文件未出现（无害）。

### 2.5 马来西亚（跨境）— 389 行
状态列：`Order Status`。`Completed` 210 / `Delivered` 99 / **`Cancelled` 32（剔除）** / `Shipping` 3 / 退货窗口 45 行（前缀归一）。

### 2.6 越南（跨境）— 9 行
状态列：`订单状态`。`Đã giao` 5（→Delivered）/ `Hoàn thành` 3（→Completed）/ **`Đã hủy` 1（剔除）**。
> 越南 map 里 `Shipping` 是注释掉的（「暂未找到」），本文件确实也没有该状态。**无退货窗口状态**（不在 `return_window_prefix_map` 里）。

---

## 三、Lazada 7 站

> Lazada 全站状态是**英文原值**，7 站共用一个 `status_map`（代码 393-401 行，7 个键）。
> 但**保留与否不看 map，看 438 行的过滤条件**：`isin(["shipped","delivered","confirmed"])`。
> map 只是把原值透传一遍（`"confirmed":"confirmed"`），**真正的判定在过滤那一步**——这点容易看错。

| 站点 | 实体 | 行数 | 保留 | | 剔除（4 类） |
|---|---|---:|---|---|---|
| 菲律宾 | 本土 | 892 | confirmed 753 / delivered 1 / shipped 1 | 剔除 | canceled 112 + Package Returned 14 + returned 11 = **137** |
| 泰国 | 本土 | 438 | confirmed 375 / delivered 3 / shipped 1 | 剔除 | canceled 54 + Package Returned 4 + returned 1 = **59** |
| 菲律宾 | 跨境 | 164 | confirmed 100 / delivered 21 / shipped 4 | 剔除 | canceled 29 + Package Returned 9 + returned 1 = **39** |
| 泰国 | 跨境 | 82 | confirmed 57 / delivered 7 / shipped 4 | 剔除 | canceled 13 + `In Transit: Returning to seller` 1 = **14** |
| 新加坡 | 跨境 | 62 | confirmed 52 / delivered 5 / shipped 1 | 剔除 | canceled **4** |
| 马来西亚 | 跨境 | 14 | confirmed 10 | 剔除 | canceled **4** |
| 越南 | 跨境 | 17 | confirmed 15 / delivered 1 | 剔除 | canceled **1** |

**Lazada 出现过的状态共 7 种**：`confirmed` / `shipped` / `delivered`（保留）、
`canceled` / `returned` / `Package Returned` / `In Transit: Returning to seller`（剔除）。
7 种**全部在 map 里**，✅ 零漏网。

> 反向：`In Transit: Returning to seller` 只在【跨境泰国】出现过 1 行；`returned` 只在菲本土(11)／泰本土(1)／菲跨境(1) 出现。

---

## 四、Tiktok 2 站

| 站点 | 行数 | 保留 | 剔除 |
|---|---:|---|---|
| 印尼（本土） | 354 | Completed 293 / Shipped 13 | **Canceled 48** |
| 菲律宾（本土） | 1137 | Completed 889 / Shipped 44 | **Canceled 204** |

状态仅 3 种：`Completed` / `Shipped` / `Canceled`，map 全含。✅
过滤条件（501 行）：`isin(["Shipped","Completed"])`。
> 注：Tiktok **没有跨境站**，只有印尼、菲律宾两个本土站；也没有退货窗口状态。

---

## 五、全局 status_map 清单（脚本现状）

> 🔄 **2026-09-16 已按 9.15 三个工作簿的业务状态字典全量补齐**（见第十节）。
> 下表为**补齐后**的现状。

### Shopee（逐站不同，5 套）
| 站点·实体 | map 键数 | 键内容 |
|---|---:|---|
| 泰国·本土 | **7** | `สำเร็จแล้ว`/`จัดส่งสำเร็จแล้ว`/`ผู้ซื้อได้รับสินค้าแล้ว`→Completed、`การจัดส่ง`/`ที่ต้องจัดส่ง`→Shipping、`ยกเลิกแล้ว`/`ยังไม่ชำระ`→Cancelled |
| 印尼·本土 | **8** | `Selesai`/`Pesanan Diterima`→Completed、`Sedang Dikirim`/`Perlu Dikirim`→Shipping、`Telah Dikirim`→Delivered、`Batal`/`Belum Bayar`/`Pembatalan diajukan`→Cancelled |
| 菲律宾·本土 | **8** | `Completed`/`Order Received`→Completed、`Shipping`/`To ship`→Shipping、`Delivered`、`Cancelled`/`Unpaid`/`Cancellation requested`→Cancelled |
| 新加坡·跨境 | **8** | 同菲律宾（三站合并为一个 elif 分支） |
| 马来西亚·跨境 | **8** | 同菲律宾 |
| 越南·跨境 | 3 | `Hoàn thành`、`Đã hủy`、`Đã giao→Delivered` |

**Shopee 保留/剔除判定**：`订单状态 != "Cancelled"` 且非空 → **反向白名单**（只要不是取消就留）。
即：`Shipping` / `Delivered` / `Completed` / 退货窗口归一后的值 **全部保留**。

### Lazada（7 站共用 1 套）
补齐后 **15 个键**，与业务字典 `status.1` 完全对齐：

- ✅ **计入**：`confirmed` / `shipped` / `delivered`
- ❌ **不计入**：`canceled` / `returned` / `Package Returned` / `In Transit: Returning to seller` /
  `pending` / `ready_to_ship` / `packed` / `Unpaid` / `Buyer Delivery Failed` /
  `Seller Delivery Failed` / `Lost by 3PL` / `Package scrapped`

**保留/剔除判定**：**正向白名单** `isin(["shipped","delivered","confirmed"])`。

### Tiktok（2 站共用 1 套）
补齐后 **9 个键**（中英混用已兼容）：
- ✅ **计入**：`Completed` / `已完成` / `Shipped` / `已发货`
- ❌ **不计入**：`Canceled` / `Cancelled` / `已取消` / `To ship` / `待发货`

**保留/剔除判定**：**正向白名单** `isin(["Shipped","Completed"])`。

> ⚠️ **三平台判定方向不一致**：Shopee 是"非取消即留"，Lazada/Tiktok 是"只留白名单"。
> 这意味着 Shopee 若将来冒出任何**新状态**，只要不是 `Cancelled` 就会被**自动计入销量**（可能误算）；
> Lazada/Tiktok 冒出任何新状态则会**自动被丢弃**（可能漏算）。两者风险方向相反，**新增状态出现时都要盯**。
>
> 📌 补齐后三平台均已把**已知全部状态显式登记**，`.map()` 不会再产生"静默 NaN 丢弃"。但仍建议每月
> 用 `工具-东南亚状态扫描.py` 扫一遍当月文件，确认没出现新状态。

---

## 六、15 个文件的行级对账

| 平台 | 文件数 | 总行数 | 保留行 | 剔除行 |
|---|---:|---:|---:|---:|
| Shopee | 6 | 6937 | 6360 | 577 |
| Lazada | 7 | 1669 | 1411 | 258 |
| Tiktok | 2 | 1491 | 1239 | 252 |
| **合计** | **15** | **10097** | **9010** | **1087** |

> 行数 ≠ 件数：Shopee/Tiktok 有「数量」列（一单多件），Lazada 每行固定 1 件。
> **件数口径**剔除合计 = **1353**（见《1353件定位报告》），与此处行数 1087 不同是对的。

---

## 七、顺带登记的其它东南亚数据源（不在三平台主链路）

> ⚠️ **重要口径澄清**：`sum_total_southeast_asia_site`（代码 827-833 行）过滤条件是
> `站点 ∈ {菲律宾,印尼,泰国,马来西亚,新加坡,越南}`，**只按站点、不限平台**。
> 所以"东南亚合计 13769"里其实**混进了两个线下渠道**，它们没有订单状态、不走 status_map：

| 文件 | 行数 | 平台标签 | 8月销量 | 是否走 status_map |
|---|---:|---|---:|---|
| `印尼海外仓202608订单报告.xlsx` | 1721 | **线下** | **2562** | ❌ 不走 |
| `菲律宾KSON仓202608线下订单报告.xlsx` | 81 | **KSON** | **2132** | ❌ 不走 |
| `印尼海外仓202608商品出库报告.xlsx` | 190 | — | — | ❌ 出库报告，按量记 |

**线下渠道小计 = 4694 件**，与三平台修复完全无关。

```
六站 13769 = 三平台 9075 + 线下 4694
```

**风险**：这 4694 件是**无状态直接计入**的，若线下报告出现退货/取消行，现有代码看不见。
建议月复检时人工确认这两份报告有无负数/退货备注。

---

## 八、一句话给维护者

**当前 status_map 对 202608 快照是完备的，没有静默丢单。**
唯一要记的两件事：
1. **`return_window_prefix_map` 的归一化目标必须写该站 status_map 里真实存在的键**
   （泰国 `"สำเร็จแล้ว"` / 印尼 `"Selesai"` / 菲新马 `"Completed"`），写错就静默丢单（见第十节 10.5）；
2. **三平台判定方向不一致**（Shopee 反向白名单 / Lazada+Tiktok 正向白名单），新增状态时两边都要盯。

> 🔄 2026-09-16 已按 9.15 业务字典**全量补齐** status_map（第十节），
> 旧的第 1 条「Shopee 泰国 map 只有 2 个键」已不再成立。

---

## 九、20260915 快照交叉印证（新增）

9.15 新增的 `2026年{印尼智慧物流,泰国易仓,菲律宾仓}销售表1-12月.xlsx` 里的订单 sheet
**带一个平台报告没有的字段 `有效性`（1/0）**，等于**业务方自己给出的"算不算数"判定**。
把它当**第三条独立路径**与现行 status_map 对撞：

| 站点 | 渠道 | 9.15 工作簿 `有效性=1` | 现行 status_map 保留 | 结论 |
|---|---|---:|---:|---|
| 印尼 | 虾皮 | **997** | 1152 − 155(Batal) = **997** | ✅ 一致 |
| 印尼 | 抖音 | **309** | 309（Completed 296 + Shipped 13） | ✅ 一致 |
| 印尼 | 上马线下 | **2562** | 2562（OUTER 全额） | ✅ 一致 |
| 菲律宾 | 虾皮 | **2015** | 2015 | ✅ 一致 |
| 菲律宾 | LAZADA | **755** | 755 | ✅ 一致 |
| 菲律宾 | 抖音 | **945** | 945（Completed 901 + Shipped 44） | ✅ 一致 |
| 菲律宾 | 线下 | **2132** | 2132（KSON 全额） | ✅ 一致 |
| 泰国 | LAZADA | **379** | 379 | ✅ 一致 |

**含义**：三平台 `status_map` + 退货窗口前缀归一 的判定结果，与业务方在 9.15 工作簿里
用 `有效性` 手工标注的结果**逐笔相同**。这是对现行口径的**强佐证**，无需修改 status_map。

> 详细对撞见 `9.15快照复核报告-月销表与东南亚表对撞.md`。
> 另注：9.15 工作簿的线下单**也没有状态过滤**（`08月当地线下订单` 的 `订单状态` 只有 `OUTER(已出库)` 一种），
> 与第七节「线下 4694 件无状态直接计入」的风险提示一致。

---

## 十、2026-09-16 按业务字典全量补齐 status_map（已实施）

### 10.1 起因

用户指出：9.15 三个工作簿（菲律宾仓 / 泰国易仓 / 印尼智慧物流）的
**Lazada、虾皮、抖音订单表里，业务方手工维护了一张"更全的状态字典"**，
要求**更新进代码，以备未来需要**。

### 10.2 找到的字典（在每类订单表尾部，全是「非订单数据行」）

| 工作簿 | 订单表 | 字典列 | 条数 |
|---|---|---|---:|
| 菲律宾 | `8月虾皮订单` | `Order Status.1` | **8** |
| 菲律宾 | `8月LAZADA订单` | `status.1` / `状态` | **15 / 15** |
| 泰国 | `8月易仓虾皮订单` | `订单状态`（中文） | **8** |
| 泰国 | `8月易仓LAZADA订单` | `status.1` / `状态` | **14 / 14** |
| 印尼 | `08月虾皮订单` | `Status Pesanan.1` / `消息状态` | **8 / 8** |

这些列**只在每个 sheet 的前 8~15 行有值**（其余行为空），是**枚举对照表**，不是订单数据。
它们列出了平台**所有可能的状态**，**包括当月从未出现过的**（例如 `Unpaid`、`Package scrapped`、
`Belum Bayar`、`Pembatalan diajukan`）—— 这正是"以备未来需要"的价值。

### 10.3 与旧 status_map 的差集（全年 9 个月扫描）

用 `probe_status_allyear.py` 扫三工作簿全部 9 个月的平台订单表，得到实际出现过的状态，
再与旧代码 status_map 做差：

| 站点·平台 | 旧 map | 新增键 | 新状态（件数） |
|---|---:|---:|---|
| 菲律宾·Shopee | 4 | **+4** | `Order Received`(2064)、`To ship`(109)、`Unpaid`(0)、`Cancellation requested`(0) |
| 菲律宾·Lazada | 7 | **+8** | `pending`(31)、`ready_to_ship`(4)、`Buyer Delivery Failed`(4)、`Lost by 3PL`(1)、`packed`(1)、`Unpaid`、`Seller Delivery Failed`、`Package scrapped` |
| 菲律宾·Tiktok | 3 | **+6** | `已发货`(172)、`已完成`(101)、`已取消`(63)、`待发货`(39)、`Cancelled`、`To ship` |
| 泰国·Shopee | 2 | **+5** | `จัดส่งสำเร็จแล้ว`(411)、`ที่ต้องจัดส่ง`(131)、`การจัดส่ง`(72)、`ยังไม่ชำระ`(3)、`ผู้ซื้อได้รับสินค้าแล้ว`(2836) |
| 泰国·Lazada | 7 | **+7** | `pending`(18)、`Lost by 3PL`(2)、其余 5 个字典项 |
| 印尼·Shopee | 4 | **+4** | `Perlu Dikirim`(31)、`Pesanan Diterima`(2)、`Belum Bayar`、`Pembatalan diajukan` |
| 印尼·Tiktok | 3 | **+6** | `To ship`(7)、`待发货`、中英混用变体 |

### 10.4 新状态的两类语义 & 处置

| 类别 | 例子 | 处置 | 依据 |
|---|---|---|---|
| **未成交态（新发现）** | `To ship` / `待发货` / `ที่ต้องจัดส่ง` / `Perlu Dikirim` / `pending`<br>`ready_to_ship` / `packed` / `Unpaid` / `Belum Bayar` / `Pembatalan diajukan` | ❌ **剔除** | 平台尚未确认成交，计入会虚增。旧代码里这些值 **不在 map 中 → `.map()` 变 NaN → 被静默丢弃**，结果等价但**无迹可查**；现改为**显式登记** |
| **已成交/已签收态的别名（新发现）** | `Order Received`(菲)、`Pesanan Diterima`(印尼)、`ผู้ซื้อได้รับสินค้าแล้ว`(泰)<br>`จัดส่งสำเร็จแล้ว`(泰"成功交付")、`การจัดส่ง`(泰"送货") | ✅ **保留** | 是「已签收」的**裸状态形态**（退货窗口已过期、无日期后缀） |
| **物流异常** | `Buyer Delivery Failed` / `Seller Delivery Failed` / `Lost by 3PL` / `Package scrapped` | ❌ **剔除** | 已在旧 map 或按同义剔除 |

### 10.5 ⚠️ 最重要的一次改错（已修正）

补齐过程中曾把 `return_window_prefix_map` 的归一化目标统一改成英文 `"Completed"`，

```python
# ❌ 错误写法：归一化目标必须是该站 status_map 里真实存在的键
"泰国":("ผู้ซื้อได้รับสินค้าแล้ว","Completed"),   # 泰国 map 里键是 "สำเร็จแล้ว"，不是 "Completed"
"印尼":("Pesanan diterima","Completed"),        # 印尼 map 里键是 "Selesai"
```

回归验证立刻抓到：**泰国 2895→2691（丢 204 件）、印尼 997→944（丢 53 件）**。
原因是归一化后的值 `.map()` 时查不到 → 变 NaN → 被过滤掉。

**教训**：`return_window_prefix_map` 的第二个元素**必须与对应站点 `status_map` 的键字符串完全一致**。
现已在代码里加注释固化这条约束。

### 10.6 ⚠️ Shopee「退货窗口」状态的两种形态（本次查明）

同一个「已签收、仅处于退货窗口内」状态，平台会输出**两种文本**：

| 形态 | 例子 | 件数（8月） | 谁处理 |
|---|---|---:|---|
| **带窗口日期** | `Order Received, but buyer can raise return/refund until 2026-09-09` | 菲 102 / 泰 204 / 印尼 53 | `return_window_prefix_map` 前缀匹配 |
| **裸状态**（窗口已过期） | `Order Received` / `ผู้ซื้อได้รับสินค้าแล้ว` / `Pesanan diterima` | 菲 102(裸98行) / 泰 204 / 印尼 53 | **本次新增**：写进各自 `status_map` |

⚠️ 两者是**互斥**的（一份文件里只会出现其中一种形态），但**都不能漏** ——
旧代码只有前缀匹配、没有裸状态登记，若某月平台把全部长文本改成裸状态，
**整批会静默丢失**（菲律宾会丢 102 件 / 泰国丢 204 件 / 印尼丢 53 件）。

### 10.7 实测回归（9.07 快照，改动后逐项对撞）

| 项 | 期望 | 实测 | 结论 |
|---|---:|---:|---|
| Shopee 本土 菲律宾 | 2015 | **2015** | ✅ |
| Shopee 本土 泰国 | 2895 | **2895** | ✅ |
| Shopee 本土 印尼 | 997 | **997** | ✅ |
| Lazada 本土 菲律宾 | 755 | **755** | ✅ |
| Lazada 本土 泰国 | 379 | **379** | ✅ |
| Tiktok 本土 菲律宾 | 945 | **945** | ✅ |
| Tiktok 本土 印尼 | 309 | **309** | ✅ |
| Lazada 跨境 菲律宾 | 125 | **125** | ✅ |
| Lazada 跨境 泰国 | 68 | **68** | ✅ |
| KSON 线下 菲律宾 | 2132 | **2132** | ✅ |
| 印尼海外仓 线下印尼 | 2562 | **2562** | ✅ |
| **东南亚六站点合计** | 13769 | **13769** | ✅ |
| **总计** | 71861 | **71861** | ✅ |
| **8月.23（总计块当月）** | 71861 | **71861** | ✅ |

**结论：11 项分组 + 3 项总量全部零变化。** 本次补齐是**纯防御性**改动，
只增加"未来月份出现新状态时不会静默丢单"的覆盖，不改变任何既有数字。

### 10.8 维护建议

1. 每月跑 `工具-东南亚状态扫描.py` 扫当月文件，确认无 B 类未覆盖状态；
2. 出现新状态时：**先判定它是"未成交"还是"已成交别名"**，再决定写进 `status_map` 的保留侧还是剔除侧；
3. 改 `return_window_prefix_map` 时，**务必核对目标是该站 status_map 的真实键**（见 10.5 的教训）；
4. 新增平台订单表（如新站点）时，先看它表尾有没有**业务字典列**，有就直接照抄 —— 比逐个跑月份统计快得多。

## 十一、2026-09-16 下午：按「K~N 区 0/1 查找表」二次校准（已实施）

### 11.1 起因（用户原话）

> "三个表的 sheet **8月易仓虾皮订单、8月易仓LAZADA订单、8月线下订单、08月抖音订单 K到N列
> 有单独列出 status_map，1代表计入销量，0代表不计入**，需要完善到现在的 status_map 中，
> 防止后续遇到漏算"

用户点出的是一个**比 §十 所用的"表尾字典列"更明确、更权威的源**：K~N 区那三列
（`状态原文 | 中文译名 | 0/1`），其中 **1 = 计入销量，0 = 剔除**。

### 11.2 ⚠️ 关键结构修正：它是**独立查找表**，不是行对齐的订单标注

第一次用 `read_map_cols.py` 读全表时，看到 K~N 列**前 8~15 行有值、其余为空**，
一度误以为它是「行对齐的订单标注」（即第 r 行的 0/1 对应第 r 行那条订单）。

**实测证伪**：泰国虾皮 `8月易仓虾皮订单` 前 8 行 ——

| 行 | 主状态列[6] | 字典原文[11] | 中文[12] | flag[13] |
|---|---|---|---|---|
| 0 | `สำเร็จแล้ว` | `การจัดส่ง` | 送货 | 1 |
| 1 | `สำเร็จแล้ว` | `คำขอยกเลิก` | 取消请求 | 0 |
| 2 | `สำเร็จแล้ว` | `ที่ต้องจัดส่ง` | 已收货 | 1 |
| 3 | `สำเร็จแล้ว` | `ยกเลิกแล้ว` | 取消 | 0 |
| … | (全为 `สำเร็จแล้ว`) | … | … | … |
| 5 | `ยกเลิกแล้ว` | `สำเร็จแล้ว` | 完成 | 1 |

主状态列 [6] 行 0~4 全是 `สำเร็จแล้ว`，而字典列 [11] 却是 5 个不同状态 ——
**两者毫无行对齐关系**。正确读法：**cols 11/12/13 的前 8~15 行是一张独立的
「状态原文→0/1」查找表**，与订单数据互不相干。

> ⚠️ 教训：看到"前 N 行有值"的列块，先用**主状态列 vs 字典列**做一次逐行比对，
> 确认是否行对齐，再决定按"标注"还是按"查找表"解析。按错会让校验脚本满屏假冲突。

### 11.3 权威字典全表（K~N 区，1=计入 / 0=剔除）

| 站点·平台 | 文件·sheet | 条数 | 0（剔除） | 1（计入） |
|---|---|---:|---|---|
| 泰国·Shopee | 泰国易仓 / `8月易仓虾皮订单` | 8 | `คำขอยกเลิก`(取消请求)、`ยกเลิกแล้ว`(取消)、`ยังไม่ชำระ`(未付款) | `การจัดส่ง`(送货)、`ที่ต้องจัดส่ง`(已收货)、`สำเร็จแล้ว`(完成)、`ผู้ซื้อได้รับสินค้าแล้ว`(买方已收到货物)、`จัดส่งสำเร็จแล้ว`(成功交付) |
| 印尼·Shopee | 印尼智慧物流 / `08月虾皮订单` | 8 | `Batal`、`Belum Bayar`、`Pembatalan diajukan` | `Perlu Dikirim`、`Sedang Dikirim`、`Selesai`、`Pesanan diterima`、`Telah Dikirim` |
| 菲·Shopee | 菲仓 / `8月虾皮订单` | 8 | `Cancelled`、`Unpaid`、`Cancellation requested` | `Completed`、`Shipping`、`To ship`、`Order Received`、`Delivered` |
| 菲·Lazada | 菲仓 / `8月LAZADA订单` | 15 | `canceled`、`returned`、`Package Returned`、`In Transit: Returning to seller`、`Unpaid`、`Buyer Delivery Failed`、`Seller Delivery Failed` | `confirmed`、`shipped`、`delivered`、`packed`、`pending`、`ready_to_ship`、`Lost by 3PL`、`Package scrapped` |
| 泰·Lazada | 泰易仓 / `8月易仓LAZADA订单` | 14 | 同菲（无 `Seller Delivery Failed`） | 同菲 |
| 菲·Tiktok | 菲仓 / `8月抖音订单` | 8 | `Canceled`、`已取消` | `Completed`、`Shipped`、`To ship`、`待发货`、`已发货`、`已完成` |
| 印尼·Tiktok | 印尼智慧物流 / `08月抖音订单` | 10 | `Canceled`、`Dibatalkan`、`Unpaid`、`Belum dibayar` | `Completed`、`Selesai`、`Shipped`、`Dikirim`、`To ship`、`Perlu dikirim` |
| 印尼·线下 | 印尼智慧物流 / `08月当地线下订单` | 8 | `CANCELLED(回收站)`、`MT_REVIEW(商户待审核)` | `PICKING`/`PACKED`/`OUTER`/`WH_REVIEW`/`PICKED`/`REVIEWED` |

### 11.4 ⚠️ 与 §十 的**两处冲突**，本次按 K~N 区口径修正

§十 用的是"表尾字典列"，其中两处与 K~N 区 0/1 **结论相反**，本次以 **K~N 区 0/1 为准**：

| 状态 | §十 的处置 | K~N 区 0/1 | 本次修正 |
|---|---|---|---|
| Lazada `Lost by 3PL` | ❌ 剔除（按"物流异常"归并） | **1（计入）** | ✅ 改为**计入** |
| Lazada `Package scrapped` | ❌ 剔除（按"物流异常"归并） | **1（计入）** | ✅ 改为**计入** |
| Lazada `packed` / `pending` / `ready_to_ship` | ❌ 剔除（按"未成交"归并） | **1（计入）** | ✅ 改为**计入** |

理由：`Lost by 3PL`(第三方物流遗失) 与 `Package scrapped`(包裹报废) 在业务口径下
**已算作成交**（货已发出、款项已确认，赔偿/报废是后续环节）；`packed/pending/ready_to_ship`
同理属"已进入履约流程"。**以业务 0/1 为准**。

> 实测：202608 两份 Lazada 表实际只出现 `confirmed/canceled/Package Returned/returned/
> delivered/shipped` 六种，**放宽前后差额为 0** —— 纯防御性，不改变当月数字。

### 11.5 本次代码改动（4 处）

1. **泰国·Shopee `status_map`**：补 `"คำขอยกเลิก":"Cancelled"`（取消请求，业务标 0）；
   把 `ที่ต้องจัดส่ง` 的中文注释由"待发货"更正为**"已收货"**（对齐 K~N 区中文列）。
2. **印尼·Shopee `status_map`**：`"Pesanan Diterima"` → **`"Pesanan diterima"`**
   （小写 d，与平台真实值逐字符一致，否则 `.map()` 直接查不到）。
3. **Lazada `status_map` + 过滤白名单**：把 `packed` / `pending` / `ready_to_ship` /
   `Lost by 3PL` / `Package scrapped` 从"剔除侧"**移到计入侧**；过滤条件由
   `isin(["shipped","delivered","confirmed"])` 扩为 8 态。
4. **Tiktok `status_map` + 过滤白名单**：补印尼语 4 键（`Dikirim`/`Selesai` →
   计入；`Dibatalkan`/`Belum dibayar` → 剔除）+ `Unpaid`/`Perlu dikirim`；
   过滤条件由 `isin(["Shipped","Completed"])` 扩为含 `"To ship"`（业务标 1）。

另在 **印尼线下**函数加注释，说明表右侧另有"上马系统状态"字典（业务也标了 0/1），
但它与主状态列 `OUTER(已出库)` 口径一致，代码无需引入该列。

### 11.6 双路验证（本次）

**A. 回归验证**（9.07 快照，`verify_status_map_regression.py`）：**13/13 全部 ✅ 零变化**
（11 分组 + 东南亚 13769 + 总计 71861）。

**B. 业务字典一致性验证**（9.15 字典，`verify_vs_biz_01.py`）：把 K~N 区当成**查找表**抽取，
逐条比对"代码判定"与"业务 0/1"：

```
泰国  Shopee   字典 8 条  ✅ 一致
印尼  Shopee   字典 8 条  ✅ 一致
菲律宾 Shopee   字典 8 条  ✅ 一致
菲律宾 Lazada   字典 15 条  ✅ 一致
泰国  Lazada   字典 14 条  ✅ 一致
菲律宾 Tiktok   字典 8 条  ✅ 一致
印尼  Tiktok   字典 10 条  ✅ 一致
总结： ✅ 代码判定与业务 0/1 字典完全一致
```

### 11.7 维护建议（在 §10.8 基础上补充）

5. **K~N 区是"权威口径源"**，与表尾字典列冲突时**以 K~N 区的 0/1 为准**；
6. 每月例行：用查找表口径（`verify_vs_biz_01.py`）跑一遍当月三工作簿，
   若有 `❌ 不一致` 行，即为新出现的状态，按它那行的 0/1 补进 `status_map`；
7. 解析 K~N 区时**必须按"查找表"读**（前 N 行独立成表），**不要**按行对齐的标注读（见 §11.2）。

