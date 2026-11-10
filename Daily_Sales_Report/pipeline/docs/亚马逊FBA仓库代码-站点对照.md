# 亚马逊 FBA 仓库代码 → 站点 / 国家 知识表

> 用途：**货件 TSV 的仓库代码 → 该落哪本库存销售表 / 哪张分析表**。
> 维护方式：遇到白名单里没有的新仓代码，**先在本文件查**，查不到再问用户，确认后补进来。
> 最后核对：2026-11-06（联网比对亚马逊官方 Seller Central 帮助页 + 多家头程物流商 2026 版清单）

---

## 0. 最快判定法（代码前缀 = 最近的大机场 IATA 码）

亚马逊 FBA 仓代码前 3 位基本是「最近大机场的 IATA 三字码」+ 1~2 位序号。
所以**看到陌生代码，先认前缀属于哪个国家**：

| 前缀 | 国家/地区 | 例 |
|---|---|---|
| `BHX` `LBA` `MAN` `LTN` `EMA` `CWL` `NCL` `EUK` `GLA` `EDI` `MME` `DSA` `LCY` `LTN` `XBH` `XUK` | **英国** | BHX4=考文垂, EMA5=东米德兰兹 |
| `DTM` `HAJ` `FRA` `MUC` `DUS` `LEJ` `CGN` `BER` `NUE` `STR` `EDE` `XDE` | **德国** | DTM2=多特蒙德 |
| `CDG` `ORY` `LYS` `MRS` `LIL` `SXB` | **法国** | |
| `MAD` `BCN` `SVQ` `VLC` | **西班牙** | |
| `MXP` `FCO` `TRN` `BLQ` | **意大利** | |
| `AMS` `RTM` `NRN` | **荷兰** | |
| `WRO` `POZ` `KTW` `SZZ` | **波兰** | |
| `PRG` `BRQ` | **捷克** | |
| `BUD` | **匈牙利** | |
| `ARN` `GOT` | **瑞典** | |
| `KIX` `NRT` `HND` `FSZ` `HSG` `NGO` `TPB` `TYO` `QCB` `XHD` `XKX` `XJE` `XJW` `VJW` `VJN` `TPX` `TPF` `HFO` `HFY` `HIY` `SOO` `SOW` `STK` `VJM` | **日本** | XHD4=流山, NRT1=市川 |
| `BLR` `BOM` `DEL` `HYD` `MAA` `PNQ` `ATX` `JPX` `CCX` `LKO` `AMD` `ISK` `SIDA` `FBHB` `FIDA` `CJB` `DEX` `SLDK` `SBOB` `XWAA` `SATB` `XSAJ` `BLX1` `SGAA` | **印度** | ISK3=比万迪(Thane) |
| `YYZ` `YYC` `YVR` `YOW` `YHM` `YEG` `YOO` `YGK` `YUL` `YXX` `YQR` `YWG` `YXU` `YKF` `YXE` `YYJ` `YQT` `YXS` `YXY` | **加拿大** | YXU1=伦敦(安大略) |
| `MEX` `MTY` `GDL` `QRO` | **墨西哥** | |
| `XSY` `BWU` `MEL` `SYD` `PER` `ADL` `BNE` `OAK`(澳) | **澳大利亚** | |
| `DXB` `DXF` `AUH` `XJB` `FJR` | **阿联酋** | |
| `RUH` `JED` `DMM` | **沙特** | |
| `PHX` `ONT` `OAK` `LGB` `SNA` `SJC` `SMF` `LAX` `BFI` `SEA` `PDX` `DFW` `HOU` `SAT` `IND` `CVG` `SDF` `LEX` `BNA` `CHA` `ATL` `CLT` `CAE` `GSP` `TPA` `MCO` `MIA` `JFK` `EWR` `ABE` `PHL` `MDT` `AVP` `BWI` `RIC` `BDL` `BOS` `MSP` `MKE` `STL` `MCI` `RNO` `LAS` `SLC` `XUS*` `MDW` `MEM` `DTW` `CMH` `RFD` `TUL` `OKC` `AFW` `AUS` `DAL` `HOU` `JAX` `LGA` `RDU` `TVC` `BTV` `ORF` `DTW` | **美国** | |

⚠️ **易混**：
- `OAK` 既是美国奥克兰，也是澳洲（OAK=Oakleigh/Melbourne 周边）—— 按货件所在文件夹判定，别只看前缀。
- `YVR1` 美加都用；`VJM*`/`VJW*` 是日本。
- `XHD` 是日本（流山/海老名），**不是**德国。

---

## 1. 已确认的关键仓（含踩坑记录）

| 代码 | 站点 | 位置 | 备注 |
|---|---|---|---|
| **XHD4** | **日本** | 千葉県流山市 DPL 流山 IV 1F/2F（南棟） | 🆕 2026-11-06 用户确认。此前白名单没有 `XHD` 前缀 → 被回落成欧洲表，FBA15GL271G1（2345 件）错簿。 |
| **ISK3** | **印度** | Maharashtra, Thane, Bhiwandi, Village Pise（Royal Warehousing） | 🆕 2026-11-06 用户确认。**不是伊斯坦布尔**（此前误判）。FBA15MTLJSTM（688 件）错落欧洲表。 |
| EMA5 | 英国 | 东米德兰兹（East Midlands） | 2026-11-06 事故：不在 `UK_WAREHOUSES` → interflow 票 FBA15MRX4YLH 被兜底成美国。已加 `EMA` 前缀 + 新建 `Interflowd英国库存销售表.xlsx`。 |
| YXU1 | 加拿大 | 安大略省伦敦市 | 2026-10-30 事故：缺 `YXU` 前缀 → FBA19VFTFF83（545 件）被兜底成美国。 |
| RUH8 | 沙特 | 利雅得 | 2026-10-24 独立成 `Interflowd沙特库存销售表.xlsx`。 |
| DTM2 | 德国 | 多特蒙德 | 欧洲表主力仓，卡航。 |

---

## 2. 日本仓（`XHD4` 所属站点，完整列表）

亚马逊日本 Partnered Carrier 支持的 FC 清单（官方帮助页 G201530080）：

```
**分拣类 / 标准 FC**
FSZ1 小田原      HFO1/HFO2 茨城宮島     HFY1/HFY2 横浜
HND3 川島        HND6/HND9 川崎          HSG1 鳥栖
KIX1 堺          KIX2 大東              KIX3 茨木
KIX4 藤井寺      KIX5 京田辺            KIX6 尼崎
KIX7 武庫川      NGO2 多治見            NGO3 名古屋港
NRT1 市川        NRT5 川越              QCB1 流山
QCB3 狭山ひろせ台 QCB4 千葉港            QCB5 相模湖
QCB6 流山おおたかの森  TPB2 上尾堤崎      TPB3 神戸西
TPB5 大目新町    TPB6 川西              TPB7 ふじみ野
TPB8 東海        TPB9 戸田              TPX2 印西
TYO1 川口        TYO2 久喜              TYO4 青梅
TYO6 坂戸        TYO7 上尾              TYO8 相模原
TYO9 狭山        VJWA 柏                VJWB 住之江
VJGT 柏          VJMF 柏                VJNA 柏
VJNB 東灘        VJNE 平塚              TPN2 名古屋
TPN3 茨城宮島

**转运中心 TC（不是最终 FC，库存会转运）**
XJE1 海老名中幸町   XJE2 厚木            XJW1 西宮北
XHD3 海老名        **XHD4 流山**        XKX3 茨木
XKX4 鳴尾浜        NPB1 海老名中幸町
```

> ⚠️ `XHD3` / `XHD4` / `XKX3` / `XKX4` 是 **TC（转运中心）**，库存到货后会调拨到 FC。
> 「预计到货」比 FC 慢几天是正常的，别当成「未到货」无限期挂在 H 里。

---

## 3. 印度仓（`ISK3` 所属站点）

2026 年亚马逊印度共 33 个 FC，横跨 12–13 个邦：

| 邦 | 仓代码 |
|---|---|
| Maharashtra | **BOM4 / BOM5 / BOM7**（比万迪 Thane）、**ISK3**（比万迪 Pise）、PNQ3（浦那）、SBOB、XWAA |
| Karnataka | BLR4 / BLR5 / BLR7 / BLR8（班加罗尔）、BLX1、XSAJ |
| Haryana | DEL2 / DEL4 / DEL5 / DEL8、DED3 / DED5、SDEB、XNRW |
| Delhi | DEX3 / DEX8、PNQ2（Mathura Road） |
| Telangana | HYD3 / HYD8（海得拉巴 Shamshabad） |
| Tamil Nadu | MAA4（Thiruvallur）、CJB1（哥印拜陀） |
| West Bengal | CCX1 / CCX2（Howrah Panchla） |
| Madhya Pradesh | SIDA（Indore）、FBHB / FIDA（Bhopal） |
| Punjab | ATX1（Ludhiana）、SATB |
| Rajasthan | JPX1 / JPX2（斋浦尔） |
| Gujarat | AMD2（艾哈迈达巴德） |
| Uttar Pradesh | LKO1（勒克瑙）、SLDK（Bijnour） |
| Assam | SGAA（Kamrup） |

> ⚠️ `ISK3` 是 Bhiwandi（塔那县，孟买都会区）的 **IXD 仓**。`ISK` 不是伊斯坦布尔机场码（伊斯坦布尔是 IST/ISL）。
> 这是本项目踩过的坑：**别按直觉把 `ISK` 读成 Istanbul**。

---

## 4. 英国仓（`UK_WAREHOUSES` / `_UK_PREFIX` 依据）

```
ALT1 Milton Keynes      BHX1 Rugeley(已关)     BHX2 Coalville
BHX3 Daventry           BHX4 Coventry          BHX5 Rugby
BHX7 Hinckley           BHX8 Redditch          BRS1/BRS2 Bristol
CWL1 Swansea(威尔士)     EDI4 Dunfermline       EMA1 Kegworth
EMA2 Sutton-in-Ashfield EMA3 Eastwood          EUK5 Peterborough
EUKA/EUKB Peterborough  EUKC Widnes            GLA1 Gourock(已关)
LBA1 Doncaster(已关)     LBA2/LBA3/LBA4 Doncaster  LBA8 Leeds
LCY2 Tilbury            LCY3 Dartford          LCY8 Rochester
LPL2 Liverpool          LTN1 Ridgmont(已关)     LTN2 Hemel Hempstead
LTN4/LTN9 Dunstable     LTN7 Bedford           MAN1 曼彻斯特机场
MAN2 Warrington         MAN3 Bolton            MAN4 Chesterfield
MAN8 Haydock            MME1 Darlington        MME2 Bowburn
NCL1 Gateshead          STN8 Milton Keynes      XUKA Runcorn
```

> ⚠️ 英国仓**关闭频繁**（BHX1 2024 关、LTN1 2026 关、LBA1/GLA1 2023 关）。
> 关闭只是不再收新货，历史货件仍可能出现在代码表里 —— 所以**前缀判定优先于精确表**。

---

## 5. 代码维护规则（写进 shipments.py 的边界）

1. **前缀优先，精确表兜底**：代码是手工穷举的，新仓一出现就漏。
   `_UK_PREFIX` / `_CA_PREFIX` / `_AU_PREFIX` / `_AE_PREFIX` / `_SA_PREFIX` / `_MX_PREFIX` 就是为此存在。
2. **新前缀必先查本文件**，查不到→按 IATA 机场码推断国家→**仍不确定就问用户**，别静默兜底。
3. **未识别仓默认回落欧洲表**是历史约定，但会静默错簿。
   现已加告警：`resolve_target`/`int_target_of` 遇到未登记前缀要打印「未知仓代码」提示。
4. **同一个仓代码在不同站点可能重名**（如 `OAK`）。落簿以「货件所在文件夹」为准，仓库代码只做交叉校验。

---

## 6. 数据来源

- 日本：`sellercentral.amazon.co.jp/help/hub/reference/external/G201530080`（官方 Partnered Carrier 支持 FC 列表）、`G200141510`（运输路线，含 TC 地址）
- 印度：ormguru / myhq / thegstco 2026 版印度 FC 全表（源自亚马逊官方 FC codes）
- 英国：eurofiscalis 2026 UK 全表、多家头程 2026 版英国仓清单
- 美国：`sellercentral.amazon.com/gp/help/external/G201811680`（官方 US FC 按州清单）
- 其他：`forestshipping`、`guigangbj` 2026 版全球仓代码清单

> 本文件是**辅助判定**，不能替代亚马逊后台。真实发货以 Seller Central 生成的 FBA 发货计划 PDF 上的 FC code + 地址为准。
