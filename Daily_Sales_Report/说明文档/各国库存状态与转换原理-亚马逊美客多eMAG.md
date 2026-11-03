# 各国库存数据与库存状态转换原理（亚马逊 / 美客多 / eMAG）

> 配套文档：《日销表数据原理与计算方式-扫盲.md》（目录结构、销量口径、作业流）。
> 本文回答两个问题：**① 平台报告里的库存状态字段都什么意思？② 一件货的状态怎么在它们之间流转？**
> 亚马逊字段含义为平台公开口径；恒等式与列位均用日销表真实数据验证过（注明例证）。

---

## 0. 先建立地图：三套库存体系

| 体系 | 覆盖 | 库存托管方 | 状态丰富度 |
|---|---|---|---|
| **亚马逊 FBA** | 美/加/墨、英/欧 9 国、日/沙/阿联酋/印/澳（主账号）+ interflow 账号 3 站 | 亚马逊运营中心（FC） | ★★★★★（在途 3 态 + 在库 3 态 + 特殊态） |
| **美客多 Full** | 墨西哥 | MercadoLibre Full 仓 | ★★★★（在途 2 态 + 在库 4 态） |
| **eMAG** | 罗马尼亚 | 卖家自仓（可选 FBE） | ★★（就一个 stock 数 + 处理时效） |

日销表里，亚马逊的库存字段出现在 `在途库存` / `{国家}可售+在途库存` sheet；
美客多出现在 `库存明细` sheet；eMAG 出现在 `库存明细` sheet。

---

## 1. 亚马逊 FBA：库存清单报告的 afn-* 字段全解

下表左列是报告字段名，中列是**日销表里人工加的中文标签**（第 1 行原样摘录），右列是含义。

### 1.1 在库三态 + 总量

| 报告字段 | 表内标签 | 含义 |
|---|---|---|
| `afn-warehouse-quantity`（J 列） | —— | **在仓总量**：此刻物理上在亚马逊运营中心里的货 |
| `afn-fulfillable-quantity`（K 列） | ——（库存记录矩阵粘的就是它） | **可售**：随时能卖、能被下单的库存 |
| `afn-reserved-quantity`（M 列） | `预留` | **预留**：已被内部流程占用、客户暂时买不到（构成见 §4） |
| `afn-unsellable-quantity`（L 列） | `不可售数量` | **不可售**：破损/过期/缺陷/买家退货处理中，需移除或翻新后才能再卖 |
| `afn-total-quantity`（N 列） | `在库在途` | **总量**：仓库里的 + 三种在途的全部（恒等式见 §3） |

### 1.2 在途三态（日销表 P/Q/R 三列，重点）

| 报告字段 | 表内标签 | 英文名 | 含义 |
|---|---|---|---|
| `afn-inbound-working-quantity`（P 列） | `刚发的货` | **Working** | 货件计划已创建、SKU 已分配进某票货件，**还没交给承运商**。此时改数量/取消都灵活 |
| `afn-inbound-shipped-quantity`（Q 列） | `在运输中的货` | **Shipped** | 承运商已揽收，**在路上**（卡航/海运/空运；分析表的"预计时效"就是这段路的时间） |
| `afn-inbound-receiving-quantity`（R 列） | `亚马逊准备接收的货` | **Receiving** | 货已**到 FC，正在扫码接收**、上架前处理。接收完成即转入可售或不可售 |

> 这三态加总 = 日销表的"在途库存"：`Y3 =P3+Q3+R3`。
> 日销表用它干两件事：分析表逐票跟踪（§6）；每天看"在途还有多少没落账"，预判几天后可售会涨多少。

### 1.3 特殊态

| 报告字段 | 含义 |
|---|---|
| `afn-researching-quantity`（S 列） | **调查中**：亚马逊找不到这批货了（错放/丢失），正在 FC 内部查找；找到回归原状态，确认丢失亚马逊赔付 |
| `afn-fc-transfer-quantity`（V 列） | **运营中心之间调拨**：为了平衡库存，货在 FC 之间转运（interflow 表的 `美国预留库存记录` 第 2 行就标着"运营中心之间调拨库存"） |
| `afn-reserved-future-supply`（T 列） / `afn-future-supply-buyable`（U 列） | **未来供货预留 / 未来可买**：参与促销承诺等未来安排的货 |
| `afn-onhand-buyable-quantity`（W 列） | `现货（可售商品 + 运营中心转运商品）` = 可售 + FC 调拨（详见 §3 恒等式） |
| `mfn-fulfillable-quantity`（H 列） | **卖家自配送（MFN）**库存——自己仓里的货，与 FBA 无关，本套表基本为 0 |
| `per-unit-volume`（O 列） | 单件体积（算仓储费用） |

---

## 2. 状态机：一件货的一生

### 2.1 主流程（进货 → 可卖 → 卖出）

```
【卖家发货侧】                                【亚马逊仓内】                  【买家侧】
创建货件计划 ──→ 交承运商 ──→ 货到运营中心 ──→ 接收完成上架
   WORKING       SHIPPED       RECEIVING    ┌─→ FULFILLABLE（可售）──→ 买家下单 ──→ RESERVED ──→ 出库
   (刚发的货)     (运输中)     (接收中)      │        ↑                    (预留/拣货中)
                                            └─→ UNSELLABLE（不可售）
                                               破损/过期/买家退货
```

- **working → shipped**：贴标装箱、交给承运商（分析表里记下"发货日期"和"预计时效"）。
- **shipped → receiving**：货到 FC，开始扫码。
- **receiving → fulfillable / unsellable**：接收完成的分流——**正常货转可售；破损/过期转不可售**。这就是"到货"的本质：在途清零、可售上涨。
- **fulfillable → reserved → 出库**：买家付款下单后货被"预留"（拣货/打包/待发），发出后彻底离开库存。

### 2.2 库内的其余流转

```
FULFILLABLE ⇄ RESERVED(FC调拨) / FC-TRANSFER   运营中心之间调拨（跨区平衡库存）
FULFILLABLE/UNSELLABLE → RESEARCHING           货丢了，亚马逊内部查找
RESEARCHING → 找到 → 回原状态 / 确认丢失 → 亚马逊赔付
UNSELLABLE → 移除订单（Removal Order）→ 退回卖家 / 弃置   ← 不可售的出路
UNSELLABLE → 重新质检上架 → FULFILLABLE                  ← 少数可"复活"
买家退货 → 先进 UNSELLABLE（退货处理）→ 质检OK可回 FULFILLABLE
```

### 2.3 与日销表的衔接（到货落账）

每日库存分析表（`{国家}-每日库存分析表`）就是这张状态机的**人工跟踪器**：

- 每票货件一列：`货件号 | 发货日期 | 预计时效 | 预计到货(=发货+时效)`，货件在途时把数量填进去；
- 货到了，**在列头手工写上"X月X日到货"**，H 列公式 `IF(ISNUMBER(FIND("到货",表头)),0,数量)+…` 自动把该列剔出"调拨在途合计"——**在途清零、可售库存自然由新下载的报告体现**。

---

## 3. 三个恒等式（背下来，对账用）

用日销表真实数据验证过（例证：`美国可售+在途库存` G102 行：J=14, K=13, L=0, M=1, N=21, Q=7, W=13）：

| 恒等式 | 含义 | 例证 |
|---|---|---|
| **在仓 = 可售 + 预留 + 不可售**<br>`afn-warehouse = fulfillable + reserved + unsellable` | FC 里的货只有这三种形态 | 14 = 13+1+0 ✓ |
| **总量 = 在仓 + 在途三态**<br>`afn-total = warehouse + working + shipped + receiving` | "在库在途"名副其实 | 21 = 14 + 0+7+0 ✓ |
| **现货 = 可售 + FC 调拨**<br>`afn-onhand-buyable = fulfillable + fc-transfer` | 调拨中的货也算"现货资产" | 13 = 13+0 ✓ |

> 对账技巧：拿到任何一版库存报告，先跑这三条式子；对不上说明报告导出不完整（截断/少列）。

---

## 4. 预留（Reserved）的三种构成

亚马逊后台 Reserved 页和文档都明确，预留库存由三部分组成：

| 构成 | 场景 |
|---|---|
| **客户订单** | 已下单待拣货/打包/待发货的货 |
| **FC 之间调拨** | 亚马逊在运营中心之间转运（对应 `afn-fc-transfer-quantity`） |
| **FC 内处理** | 上架前处理、测量体积、仓内转移等 |

> 注意区分：**预留 ≠ 不可售**。预留的货"暂时买不到"但没有损坏，绝大多数几天后就变成已发货订单；
> 不可售（unsellable）是物理/质检层面不能卖。

---

## 5. 美客多（MercadoLibre）Full 库存状态（墨西哥）

美客多 `库存明细` sheet 的状态是西语的，逐个对照（列位为该 sheet 实际列）：

### 5.1 状态清单

| 列 | 西语原文 | 中文 | 对应亚马逊概念 |
|---|---|---|---|
| O | `Pendientes de ingreso` | **待入库** | ≈ working + receiving（已发货到 Full、还没上架） |
| P | `En transferencia` | **转运中** | ≈ shipped / FC transfer |
| R | `Aptas para vender` | **可售** | = fulfillable |
| S | `No aptas para vender` | **不可售** | ≈ unsellable |
| Q | `Devueltas por el comprador` | **买家退回** | 退货处理的细分（算不可售族） |
| T | `Temporalmente no aptas para vender` | **暂时不可售** | 细分三类： |
| T | ├ `Extraviadas (en búsqueda)` | 遗失（查找中） | = researching |
| T | ├ `En revisión` | 审核中 | 仓内核查 |
| T | ├ `Ventas canceladas` | 取消销售 | 订单取消退回 |
| W | `Unidades que ocupan espacio en Full` | **占仓库存** | 计仓储费口径（可售+不可售中占位的） |
| X | `Unidades distribuidas por acción recomendada` | 推荐动作分布 | `Buena calidad 好品质 / Para impulsar ventas 促销售 / Para poner en venta 待上架 / Para evitar descarte 避免弃置` |

表头（第 1 行）人工标注的三个汇总口径：`在途库存(O/P) / 预留库存(Q/T族) / 可售库存(R)`。

### 5.2 计划类字段（美客多特色）

| 字段 | 含义 |
|---|---|
| `Stock promedio últimos 30 días (u.)` | 近 30 天平均库存 |
| `Tiempo hasta agotar stock` | 按预测销量还能卖几天（断货倒计时） |
| `Ventas últ. 30 días ($)` / `Unidades vendidos últ. 30 días` | 近 30 天销售额 / 销量 |
| `Urgencia de envío` | 补货紧急度（平台替你算的） |
| `Stock a enviar` | 平台建议的补货量 |
| `Unidades mínimas para ofrecer envíos rápidos` | 提供快速配送的最低库存门槛 |

### 5.3 状态转换流程

```
卖家仓发货
   │
   ▼
Pendientes de ingreso（待入库）──→ En transferencia（转运中，如跨仓分拨）
   │
   ▼  Full 仓接收完成
Aptas para vender（可售）⇄ No aptas para vender（不可售）
   │                        ├─ Devueltas por el comprador（买家退回）
   │                        └─ Temporalmente no aptas：遗失查找中 / 审核中 / 取消销售
   ▼
买家下单 → Full 配送出库（库存减少）
```

**与 FBA 的两点主要差异**：
1. FBA 的"买家退货"混在 unsellable 里，美客多把**退货单独列**（`Devueltas por el comprador`）；
2. 美客多把"遗失查找"归入**暂时不可售**子项，而不是像 FBA 给独立的 researching 字段；另外它有"占仓库存"和"推荐动作分布"这类运营导向的口径。

---

## 6. eMAG（罗马尼亚）库存

eMAG 没有 FBA 式的多状态机（自发货为主，可选 FBE 代管）：

- `库存明细` sheet 是 **offer 级导出**：`stock`（可用库存数）+ `status`（offer 状态 active/inactive）+ `handling_time`（发货处理天数）+ `lead_time`（缺货补齐天数）+ 价格/税率等；
- 库存就是"我（或 FBE）仓里能卖的数"，**在途全靠自己管**——所以罗马尼亚工作簿同样配了 `罗马尼亚-每日库存分析表` 跟踪到货（与亚马逊部分同一套用法）；
- 订单状态见 `订单数据`：`Finished（完成）/ In progress / Canceled`，付款状态"已兑现"等。注意 eMAG 导出**有英文/中文两个版本**（上级管线雷区 17：中文列名曾导致整批静默丢单）。

---

## 7. 各国库存数据差异速查（"不同国家的库存数据"）

| 国家/站点 | 体系 | 日销表库存 sheet | 特殊点 |
|---|---|---|---|
| 美国 / 加拿大 / 墨西哥（主账号） | 亚马逊北美 | `美国/加拿大/墨西哥可售+在途库存` + `美国/加拿大/墨西哥（结余）` + `库存记录×3` + `分析表×3` | 一个账号三市场；墨西哥**另有一份美客多**工作簿 |
| 欧盟 9 国（IT/FR/ES/DE/TR/NL/SE/PL/BE/IE） | 亚马逊 Pan-EU | `欧洲可售+在途库存`（EU 合并） + `多国库存` + 每国一张监控表 | `quantity-for-local-fulfillment` = 各国**本地履约量**；库存放 A 国可卖 B 国（跨国调拨由亚马逊管） |
| 英国 | 亚马逊（脱欧后独立） | `英国可售+在途库存` + `英国库存记录` + `英国-每日库存分析表` | 与欧盟库存**互不相通**，单独入库单独发 |
| 日本 | 亚马逊 | `可售+在途库存` | 表头日文且已乱码（按列位认列）；时间戳 +09:00 |
| 沙特 / 阿联酋 / 印度 / 澳大利亚 | 亚马逊单站点 | `可售库存`（结余）+ `在途库存`（清单）两张 | 结构最简；印度版多一张有问题 SKU 表 |
| interflow 账号（美/加、澳、阿联酋） | 亚马逊第二账号 | 同上结构；另有 `美国/加拿大预留库存记录`（区分"运营中心之间调拨库存"） | 与主账号同站点的库存**完全独立**，对账别混 |
| 墨西哥 | 美客多 Full | `库存明细` | 状态见 §5；报告有中/西语两个版本 |
| 罗马尼亚 | eMAG | `库存明细` | offer 级 stock，无多状态；导出有英/中两个版本 |

---

## 8. 中英状态对照速查表（贴墙用）

**亚马逊 FBA**

| 中文（日销表标签） | 英文 | 报告字段 |
|---|---|---|
| 刚发的货 | Working | afn-inbound-working-quantity |
| 在运输中的货 | Shipped | afn-inbound-shipped-quantity |
| 亚马逊准备接收的货 | Receiving | afn-inbound-receiving-quantity |
| 可售 | Fulfillable | afn-fulfillable-quantity |
| 预留 | Reserved | afn-reserved-quantity |
| 不可售数量 | Unsellable | afn-unsellable-quantity |
| 在库在途 | Total | afn-total-quantity |
| （调查中） | Researching | afn-researching-quantity |
| （FC 调拨） | FC Transfer | afn-fc-transfer-quantity |
| 现货（可售+转运） | Onhand Buyable | afn-onhand-buyable-quantity |

**美客多 Full**

| 中文 | 西语 |
|---|---|
| 待入库 | Pendientes de ingreso |
| 转运中 | En transferencia |
| 可售 | Aptas para vender |
| 不可售 | No aptas para vender |
| 买家退回 | Devueltas por el comprador |
| 遗失（查找中） | Extraviadas (en búsqueda) |
| 审核中 | En revisión |
| 取消销售 | Ventas canceladas |
| 占仓库存 | Unidades que ocupan espacio en Full |

---

## 9. 常见坑（库存侧）

1. **把预留当"卖掉了"**：预留大头是待发货订单，不是库存减少；真正出库要看可售的下降。
2. **把在途当"已有库存"**：working/shipped 没到仓前都卖不了；分析表"预计到货"才是可售上涨日。
3. **到货忘了写"到货"**：列头不补字，H 列在途合计永远不降，双重虚增（可售+在途）。
4. **接收中出现"不可售上涨"**：正常——破损/退货接收完先进 unsellable；长期不处理要创建移除订单。
5. **日本表乱码、结余表 #VALUE!**：见日销表文档 §8，按列位/以清单表为准。
6. **语言漂移**：美客多/ eMAG 导出语言不固定，粘错版本列会错位。

---

*本文基于 2026-10-08 对日销表 12 份工作簿的实际读取；亚马逊字段定义为平台公开口径，三条恒等式已用表内数据验证。*
