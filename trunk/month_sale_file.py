#库存分析
import pandas as pd
import numpy as np
import datetime
from zoneinfo import ZoneInfo
import file_chardetect
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.formula.translate import Translator   # 10.07新增：追加行时按行平移公式
from openpyxl.worksheet.formula import ArrayFormula  # 10.07新增：搬运 ER 列数组公式
import dateutil.parser
import os
import datetime
import dateutil
from zoneinfo import ZoneInfo
import file_chardetect
import re
from copy import copy
'''
数据分层
DWS 汇总层
DWD 明细层
ODS 原始数据层
'''

# ============================================================================
# 月销表【任务总账】—— 2026-11-02 全量重登（上一版 2026-10-29 只登了 7 条）
# ----------------------------------------------------------------------------
# 本文件是「月销售总表」的唯一生成入口，一次跑批按下面固定顺序执行。
# 规矩：新增任务必须 ① 在本表登记（序号 + 加入日期 + 行号 + grep 锚 + 失败影响）
#                  ② 在代码处写 `# ===== 任务N：<做什么> =====`，这样 `grep 任务N` 能直接定位。
#
# ⚠️ 行号是 **2026-11-02 晚 / 本文件 2254 行** 这一时刻的快照（本次新增⑱⑲⑳ 后已整体校准）
#    （这张表自己就占用 30 余行）。**行号仍是粗略区间**（历史插入改过几次，个别行仍有 ±40 行漂移），
#    定位代码请优先用「grep 锚」列，行号只用来快速跳；
#    跑批前复核一次：`grep -n "任务N" month_sale_file.py`。
#
# ⚠️ 编号沿用历史，空号/重号都不重排 —— 别处文档、记忆日志全按这些编号引用：
#    · 0/①②③④⑤⑥⑦ = 2026-09-12「8 月补齐」那一轮；⑧⑨ = 09-14；⑩ = 09-15；
#      ⑪⑫⑬⑭ = 09-16；⑮⑯ = 10-17（solo 化）；⑰ = 10-22（库龄）；
#      ⑱⑲ = 11-02（跨境订单状态核对 / 期别汇总打包，两个独立只读工具）；
#      ⑳ = 11-02（准备期模板：上一期成品→当期模板，换月第一条命令）。
#    · ★ **⑪ 在《修改指南》里指的是「印尼海外仓口径」（09-16）**，本文件 10-22 曾把
#      库龄也登记成 ⑪（撞号）。2026-11-02 起库龄统一叫 **⑰**，代码注释已同步。
#
#  序号  任务                                          加入日期     行号(11-02)       grep 锚                        失败影响
#  ----  -------------------------------------------   ----------   --------------    ----------------------------   --------------------
#  —     主流程：读单→汇总→回填→落盘（骨架，无编号）      2026-09      156~227 / 2010    def fill_excel_from_df         是，核心
#  0     前置：商品资料 + 得分列名自动探测 + 备货计划       2026-09-12   1656~1690         get_product_info_df            是，核心
#  ①    【2026年】KJ 当年总销量＝总计块月列之和            2026-09-12   2006~2030         assert _kj_letter              是，KJ 停在上月
#  ②    【欧洲各国】sheet 补当月各国值                     2026-09-12   2015~2032         任务②                          是
#  ②-b  欧洲各国模板缺行 → 末尾追加新 SKU 行              2026-10-07   2033~2083         任务②-b                        是，新 SKU 无行可写
#  ②补  【欧洲各国】得分列(EP)重刷 + 表头跟当月           2026-10-0x   2084~2110         任务②补                        否，自愈
#  ③    【2021-2026年销量情况】补当月 + 增长率 + 汇总行     2026-09-12   2120~2162         任务③                          是
#  ④    KK 列代码化（从 SKU 里提数字）                      2026-09-12   1990~1992         Unnamed: 296                    否，只影响 KK 列
#  ⑤    ASIN 反查兜底（映射不到就用源表 sku）              2026-09-12   1732~1740         lost_order                      是，会整批丢单
#  ⑥    月度复用加固：月份/列号参数化 + KJ 列名年份化      2026-09-12   1896 / 1981       month_offset / _kj_col          是
#  ⑦    新增行自动套用模板行样式（免手工刷格式）            2026-09-12   201~228           template_last_row               否，只影响格式
#  ⑧    东南亚三平台状态过滤赋值（退货窗口期＝算数）        2026-09-14   335~706           status_map                      是，取消单被计销
#  ⑨    东南亚汇总列名兼容（数量 / 数量dropna）            2026-09-14   1795~1798         数量dropna                      是，会整块清零
#  ⑩    列分组（分级显示）自动推进到当月                    2026-09-15   1392~1605/2166   arrange_month_outline           否，只影响折叠态
#  ⑪    印尼海外仓「线下单」口径（YFDZ 前缀 + 已出库）      2026-09-16   864~910           YFDZ                            是
#  ⑫    status_map 按 9.15 业务状态字典全量补齐             2026-09-16   356~426 / 518 / 612  status_map                      是
#  ⑬    K~N 查找表二次校准 status_map（含两处口径反转）     2026-09-16   同上              status_map                      是
#  ⑭    东南亚表头中英双语化（在下游脚本里，不在本文件）    2026-09-16   —                 southeast_asia_sale_file.py     否
#  ⑮    KH 得分列表头自愈（模板侧 rename + 成品兜底）       2026-10-17   1116~1191/2111   heal_score_col_header_file      否，自愈
#  ⑯    商品资料「N月销量得分」列自动生成                   2026-10-17   —                 工具-更新商品资料得分列.py      否，--verify/--apply
#  ⑰    库龄情况表（随月销表周期输出）                      2026-10-22   2221~2249         任务⑰ / import age_report     否，只打印告警
#  ⑱    跨境订单状态核对（订单号+状态 → 件数，防算多/算少）  2026-11-02   —                 工具-跨境订单状态核对.py        否，独立只读工具
#  ⑲    期别汇总打包（data+result 收一个文件夹 + 速览）     2026-11-02   —                 工具-期别汇总打包.py            否，独立只读工具
#  ⑳    换月第一步：上一期成品直接复制成当期模板            2026-11-02   —                 工具-准备期模板.py              是，模板错则全表错
#  ㉑    本地月份过滤：报告按 last-updated 导出混上月尾单，   2026-11-03   326~341           任务㉑                          是，尾单算错月（10月33件/9月54件）
#        只保留 purchase-date-local 落在归属月的订单（get_belonging_amazon_order_df 内）
#  ㉑b   汇总行四块重算（原只算全球块，Q/AA/AK 块带模板陈旧值） 2026-11-03  2243~2260        任务㉑b / 汇总行                  是，亚马逊/东南亚/北美年汇总缺当月
#  ㉔    日销表日列窗口整体平移（9/24~11/30→10/25~12/31，行5日期+行1星期平移，  2026-11-03   —                 tools/maintenance/shift_daycols_20261103.py  否，公式/CF/结构零改动
#        SUMIFS引本列$5零改动；收编换月阶段2.6 shift_window，折叠边界改快照日次日）
#  ㉒    日销表上月封存源改月销总表（runner.apply_month_summary，非本文件）：  2026-11-03   —             pipeline/core/runner.py 任务㉒    否，回落归档簿重算
#        分析表上月列原来=归档簿明细 resolve 重算（UTC 口径），与上报的月销表（本地月）
#        永远差跨月单+事后修正件；改为优先直取本表「{Y}年」sheet 对应站点块逐 SKU 值，
#        月销总表缺站点/未生成时自动回落老路径。当前簿一次性对齐由
#        pipeline/tools/maintenance/apply_msx_to_current_20261103.py 执行。
#  ㉕    一键跑全链（5 步 + 23 项断言 + 定案数字落盘）        2026-11-10   —                 工具-一键跑全链.py               是，跑批入口（漏步=错数）
#  ㉖    源数据层补充订单 + 非亚马逊渠道报告区间体检           2026-11-10   380~470           merge_supplement_orders /        是，跨月单丢失 / 导错月份静默
#        （详见 audit_channel_spans + <快照>/东南亚/补充订单/README-补充订单.md；替代 apply_yfdz_*.py）      audit_channel_spans
#  附   菲律宾新订单报告「以一敌多」接入（**默认开**）        2026-10-09   707~842 / 1669~1673  USE_PH_NEW_ORDER_REPORT         是，菲整列口径
#       ★2026-11-02 起默认 True（`PH_NEW_REPORT=0` 可回退）；同时**跳过 KSON 线下报告**（1795~1812）避免手工单双计
#  附   Shopee 跨境**台湾站**并入「虾皮境内」（含取消单剔除）  2026-11-02   433~443 / 1254/1273/1826  台湾 / other_shopee_cross_site / se_asia_sites   是，东南亚少 6 件
#  附   亚马逊「欧洲」站改名「欧盟」导出的文件名别名         2026-10-30   230~267           resolve_report_path             是，找不到报表
#
# ※ ②补「得分列(EP)重刷」没有独立日期记录：注释里提到 202609 期实测（表头停在上月），
#   且排在 ②-b(10-07) 与 ⑮(10-17) 之间 → 推定 10-07~10-09 那轮跑批现场补的。
#
# 下游（不在本文件里，但依赖本文件产出的 sale_file_tmp/*.csv，不能并行、不能复用上期）：
#   southeast_asia_sale_file.py / europe_sale_file.py / north_america_sale_file.py
# ============================================================================

entites = ["巨川","interflow"]
amazon_country_inv_map = {
    "法国":"欧洲",
    "德国":"欧洲",
    "意大利":"欧洲",
    "西班牙":"欧洲",
    "土耳其":"欧洲",
    "比利时":"欧洲",
    "瑞典":"欧洲",
    "荷兰":"欧洲",
    "波兰":"欧洲",
    "爱尔兰":"欧洲",
    "英国":"英国",
    "日本":"日本",
    "澳大利亚":"澳大利亚",
    "美国":"美国",
    "加拿大":"加拿大",
    "墨西哥":"墨西哥",
    "沙特阿拉伯":"沙特阿拉伯",
    "阿拉伯联合酋长国":"阿拉伯联合酋长国",
    "印度":"印度",
}
channel_country_map = {
    "Amazon.co.uk":"英国",
    "Amazon.fr":"法国",
    "Amazon.de":"德国",
    "Amazon.it":"意大利",
    "Amazon.es":"西班牙",
    "Amazon.com.tr":"土耳其",
    "Amazon.com.be":"比利时",
    "Amazon.se":"瑞典",
    "Amazon.nl":"荷兰",
    "Amazon.pl":"波兰",
    "Amazon.ie":"爱尔兰",
    "Amazon.com.au":"澳大利亚",
    "Amazon.co.jp":"日本",
    "Amazon.com":"美国",
    "Amazon.com.mx":"墨西哥",
    "Amazon.ca":"加拿大",
    "Amazon.sa":"沙特阿拉伯",
    "Amazon.ae":"阿拉伯联合酋长国",
    "Amazon.in":"印度",
}
amazon_sale_group = {
    "法国":"欧洲",
    "德国":"欧洲",
    "意大利":"欧洲",
    "西班牙":"欧洲",
    "土耳其":"欧洲",
    "比利时":"欧洲",
    "瑞典":"欧洲",
    "荷兰":"欧洲",
    "波兰":"欧洲",
    "爱尔兰":"欧洲",
    "英国":"欧洲",
    "美国":"北美",
    "加拿大":"北美",
    "墨西哥":"北美",
    "阿拉伯联合酋长国":"阿拉伯联合酋长国",
    "沙特阿拉伯":"沙特阿拉伯",
    "印度":"印度",
}
amazon_order_tz = {
    # "欧洲":"Europe/Berlin",
    "欧洲":"Europe/London",#实际上欧洲订单报告用的就是英国时间，并不是所说的CET时间
    "英国":"Europe/London",
    "日本":"Asia/Tokyo",
    "澳大利亚":"Australia/Sydney",#Australia/Melbourne
    "阿拉伯联合酋长国":"Asia/Dubai",
    "沙特阿拉伯":"Asia/Riyadh",
    "北美":"America/Los_Angeles",#亚马逊订单目前使用ptd时区，但是不一定是IANA时区，只是20260824的时候是一致的
    "印度":"Asia/Kolkata",
}
emag_sale_group = {
    "保加利亚":"全站点",
    "匈牙利":"全站点",
    "罗马尼亚":"全站点",
}
tmp_dir = "sale_file_tmp"
def fill_excel_from_df(df, template_path, output_path, 
                       start_row=1, start_col=1, 
                       sheet_name=None):
    """
    将 DataFrame 数据（不含表头）填入 Excel 模板的指定起始位置
    
    参数:
        df: pandas.DataFrame，要填入的数据
        template_path: str，模板文件路径
        output_path: str，输出文件路径
        start_row: int，起始行号（从1开始）
        start_col: int，起始列号（从1开始）
        sheet_name: str，工作表名称，None表示使用第一个工作表
    """
    # 加载模板
    wb = load_workbook(template_path, data_only=False)
    
    # 选择工作表
    if sheet_name is None:
        ws = wb.active
    else:
        ws = wb[sheet_name]
    
    # 处理 NaN，转为 None
    df_clean = df.replace({np.nan: None})
    
    # 只取数据，不要表头
    data_values = df_clean.values.tolist()
    
    # 记录模板最后一条数据行：模板只给这个范围设了样式（底纹/边框/字体/对齐/行高），
    # 超出部分 openpyxl 会建成默认格式 → 每月新增的 SKU 行会"掉样式"，不得不手工刷格式
    template_last_row = start_row - 1
    for _r in range(start_row, ws.max_row + 1):
        if ws.cell(row=_r, column=start_col).value not in (None, ""):
            template_last_row = _r
    max_col = ws.max_column

    # 写入数据
    for row_idx, row_data in enumerate(data_values, start=start_row):
        for col_idx, value in enumerate(row_data, start=start_col):
            if value is not None:
                ws.cell(row=row_idx, column=col_idx, value=value)

    # 9.12新增：给超出模板范围的新增行套用模板行样式（按列复制第 template_last_row 行）
    # 模板数据区（第4行~末行）逐列样式完全一致，故整行套用即可与新行"看起来无缝"
    last_row = start_row + len(data_values) - 1
    if last_row > template_last_row >= start_row:
        for r in range(template_last_row + 1, last_row + 1):
            for c in range(1, max_col + 1):
                src_cell = ws.cell(row=template_last_row, column=c)
                if src_cell.has_style:
                    ws.cell(row=r, column=c)._style = copy(src_cell._style)
        ref_height = ws.row_dimensions[template_last_row].height
        if ref_height is not None:
            for r in range(template_last_row + 1, last_row + 1):
                ws.row_dimensions[r].height = ref_height
        if ws.auto_filter.ref:  # 自动筛选范围同步扩到新行（模板里是 A3:KK535）
            m = re.match(r"^([A-Z]+\d+:[A-Z]*)", ws.auto_filter.ref)
            if m:
                ws.auto_filter.ref = f"{m.group(1)}{last_row}"
        print(f"🎨 新增行 {template_last_row + 1}~{last_row} 已套用第 {template_last_row} 行样式"
              f"（{last_row - template_last_row} 行 × {max_col} 列，含底纹/边框/字体/行高）")

    # 保存
    try:
        wb.save(output_path)
    except PermissionError:
        raise PermissionError(
            f"无法写入【{output_path}】——该文件正被 Excel（或其他程序）打开。\n"
            f"提示：脚本跑完会自动用 Excel 打开成品，若跑完不关掉它，下一轮保存就会失败。请关闭后重跑。") from None
    print(f"✅ 成功写入 {len(data_values)} 行数据到 {output_path}")
    return output_path
# ==================== 报表文件名别名兜底（2026-10-30） ====================
# 亚马逊后台中文界面会把 "Europe" 译成「欧盟」：
#   20261030 实测导出的是 `亚马逊欧盟站点202610订单报告.txt` / `亚马逊欧盟管理库存报告20261030.csv`，
#   而脚本按 amazon_sale_group 里的分组名「欧洲」拼文件名 → 直接 FileNotFoundError（跑批中断）。
# 策略：**原名优先**（存在就用原名，行为零变化）→ 不在才按别名替换重试 → 命中打印提示。
# 后台文案以后还会变，往 REPORT_NAME_ALIASES 里加一行即可，不用改调用处。
REPORT_NAME_ALIASES = {
    "欧洲": ["欧盟"],
}


def _alias_names(name: str) -> list:
    out = []
    for base, alts in REPORT_NAME_ALIASES.items():
        if base in name:
            for alt in alts:
                out.append(name.replace(base, alt))
    return out


def resolve_report_path(dir_path: str, file_name: str) -> str:
    """原名优先、别名兜底地定位一份报表；都找不到时抛出带目录线索的人话异常。"""
    path = os.sep.join([dir_path, file_name])
    if os.path.exists(path):
        return path
    for alt_name in _alias_names(file_name):
        alt_path = os.sep.join([dir_path, alt_name])
        if os.path.exists(alt_path):
            print(f"ℹ️ 报表名别名命中：{file_name} → {alt_name}")
            return alt_path
    siblings = []
    if os.path.isdir(dir_path):
        key = re.sub(r"[0-9]{6,8}|报告|站点|管理库存|订单", "", file_name).strip()
        if key:
            siblings = sorted(f for f in os.listdir(dir_path) if key in f)[:5]
    hint = f"\n  目录里相近文件：{siblings}" if siblings else f"\n  目录：{os.path.abspath(dir_path)}"
    raise FileNotFoundError(f"找不到报表【{file_name}】{hint}")


# 2026-11-10：亚马逊报告「月份区间」体检的告警收集器（见 get_belonging_amazon_order_df 末尾）
AMZ_SANITY_WARNINGS = []

# 2026-11-10：**非亚马逊**渠道（Shopee/Lazada/Tiktok/京东/eMAG/TEMU/美客多/线下）报告区间体检
REPORT_SANITY_WARNINGS = []


def audit_channel_spans(order_df:pd.DataFrame, tag:str, belonging_date:str,
                        min_rows:int=30)->list:
    """按 (实体, 平台, 站点) 体检一份渠道订单数据的日期覆盖，返回告警文本列表。

    为什么要单独做一遍：亚马逊有「本地月份过滤」（任务㉑）兜底，导错月份会在那里被拦下；
    而 Shopee / Lazada / Tiktok / 京东 / eMAG / TEMU / 美客多 / 线下 这些渠道**完全没有月份过滤**
    ⇒ 导错月份时的表现是：日志一句提示都没有，数字静静地变多或变少（interflow 北美那次
    就是同类事故，只是恰好落在亚马逊分支上）。
    判据（只看形状，不改任何数字）：
      · 归属月占比 = 0 或 < 50%   → 【严重】整份报告看起来不是本期的
      · 50% ≤ 占比 < 90%          → 【提醒】报告里混着上/下月单（滚动窗口导出的渠道属正常）
      · 覆盖天数 < 当月天数 − 3   → 【提醒】报告可能截断（只导了半个月）
    """
    out = []
    if order_df is None or len(order_df) == 0:
        return out
    if "下单时间" not in order_df.columns:
        return out
    start = pd.Timestamp(int(belonging_date[:4]), int(belonging_date[4:6]), 1)
    end = start + pd.DateOffset(months=1)
    days = start.days_in_month
    keys = [c for c in ("实体", "平台", "站点") if c in order_df.columns]
    if not keys:
        keys = ["平台"] if "平台" in order_df.columns else None
    grouped = order_df.groupby(keys, dropna=False) if keys else [("全部", order_df)]
    for key, g in grouped:
        if len(g) < min_rows:
            continue                     # 小渠道（如台湾 5 单）噪声大，不体检
        d = pd.to_datetime(g["下单时间"], errors="coerce")
        ok_d = d.notna()
        if not ok_d.any():
            continue
        n = int(ok_d.sum())
        share = float(((d >= start) & (d < end)).sum()) / n
        span = int((d.max() - d.min()).days) + 1
        name = "/".join(str(k) for k in (key if isinstance(key, tuple) else (key,)))
        dist = d.dt.strftime("%Y%m").value_counts().to_dict()
        if share < 0.5:
            out.append(f"【严重·报告区间】{tag} {name}：{n} 行里只有 {share:.1%} 属于 "
                       f"{belonging_date}（实际月份分布 {dist}）—— 这份报告很可能导错月份区间")
        elif share < 0.9:
            out.append(f"【提醒·报告区间】{tag} {name}：{n} 行里 {1 - share:.1%} 不属于 "
                       f"{belonging_date}（分布 {dist}）—— 滚动窗口导出的渠道属正常，否则请核对")
        if span < days - 3:
            out.append(f"【提醒·报告区间】{tag} {name}：日期只覆盖 {span} 天（{d.min():%Y-%m-%d}~"
                       f"{d.max():%Y-%m-%d}，当月 {days} 天）—— 可能截断")
    for w in out:
        print(w)
        REPORT_SANITY_WARNINGS.append(w)
    return out



def get_belonging_amazon_order_df(entity:str, sale_group:str, belonging_date:str, snapshot_date:str)-> pd.DataFrame:
    dir_path = os.path.join("data","数据归属周期",belonging_date,"快照日期",snapshot_date)
    if entity == "巨川":
        prefix = "亚马逊"
    elif entity == "interflow":
        prefix = "interflow"
    else:
        raise Exception(f"不支持的主体{entity}")
    file_name = f"{prefix}{sale_group}站点{belonging_date}订单报告.txt"
    file_path = resolve_report_path(dir_path, file_name)
    encoding = file_chardetect.detect_encoding(file_path)[0]
    order_df = pd.read_csv(file_path, sep="\t", parse_dates=["purchase-date"], encoding=encoding)
    order_df["实体"] = entity
    order_df["订单组"] = sale_group
    order_df["平台"] = "亚马逊"
    order_df["站点"] = order_df["sales-channel"].map(channel_country_map)
    order_df["仓库"] = order_df["站点"].map(amazon_country_inv_map)
    order_df = order_df[order_df["order-status"] != "Cancelled"]
    order_df = order_df[order_df["站点"].notna()]
    order_df = order_df[~order_df["sku"].str.startswith("Amazon.Found.")]#去除其他用途sku号
    order_df = order_df[~order_df["sku"].str.startswith("amzn.gr.")]#去除其他用途sku号
    #去除快照日期当天的订单，因为不完整
    # order_df = order_df[order_df["日期"] != pd.to_datetime(snapshot_date, format="%Y%m%d", utc=True)]#utc时间和普通时间戳不能比较，必须转换相同
    order_df = order_df.copy().reset_index()

    site_list = order_df["站点"].drop_duplicates().tolist()
    print(site_list)
    inv_group = ""
    for site in site_list:
        order_tz = None
        tz_condition = order_df["站点"] == site
        inv_group = amazon_country_inv_map[site]
        if site in amazon_order_tz:
            order_tz = amazon_order_tz[site]
        elif inv_group in amazon_order_tz:
            order_tz = amazon_order_tz[inv_group]
            tz_condition = order_df["仓库"] == inv_group
        elif sale_group in amazon_order_tz:
            order_tz = amazon_order_tz[sale_group]
            tz_condition = order_df["订单组"] == sale_group
        else:
            raise Exception(f"不支持时区的站点{site}和仓库组{inv_group}")

        # 将无时区时间赋予特定时区属性
        dt_tz = datetime.datetime.now().replace(tzinfo=ZoneInfo(order_tz))
        # 获取相对 UTC 的偏移量 (utcoffset)
        offset = dt_tz.utcoffset()
        print(f"站点{site}和仓库组{inv_group}")
        print(f"{order_tz:<18} : UTC {offset}")
        order_df.loc[tz_condition, "purchase-date-local"] = order_df[tz_condition]["purchase-date"].dt.tz_convert(order_tz).dt.tz_localize(None)#最终把时区抹掉，因为pandas不支持一列有多个时区，会产生时区污染最终只有一个时区
        
    order_df["日期"] = order_df["purchase-date-local"].dt.normalize()
    # ===== 任务㉑：本地月份过滤（2026-11-03）=====
    # 月度订单报告按 last-updated 导出，报告里混着「上个月末尾单」（本地下单月=上月）：
    #   202610 实测 欧盟24+英国9 件本地9/30深夜单曾混入「10月」列（50788→修复后 50755）；
    #   202609 同样混 8 月尾单 54 件（全在欧盟报告）。月销表口径=站点本地下单月 →
    #   这里只保留 purchase-date-local 落在归属月内的订单。
    # 本函数仅月销表聚合一个调用方；日销表走 UTC 归月不受影响。
    _bel_start = pd.Timestamp(int(belonging_date[:4]), int(belonging_date[4:6]), 1)
    _bel_end_excl = _bel_start + pd.DateOffset(months=1)
    _rows_before_filter = len(order_df)
    _pre_month_dist = order_df["purchase-date-local"].dt.strftime("%Y%m").value_counts()
    order_df = order_df[(order_df["purchase-date-local"] >= _bel_start)
                        & (order_df["purchase-date-local"] < _bel_end_excl)].copy()
    _rows_dropped = _rows_before_filter - len(order_df)
    if _rows_dropped > 0:
        print(f"【任务㉑·本地月份过滤】{entity}/{sale_group}：剔除本地下单月≠{belonging_date} 的订单 {_rows_dropped} 行（上月尾单，不进本期月销表）")

    # ===== 2026-11-10 新增：报告「月份区间」体检（防静默丢单）=====
    # 事故：20261107 快照的 interflow北美站点202610订单报告.txt 被误导出成 2026-11-01~11-07
    #   （213 行），月份过滤把 197 行当「上月尾单」剔掉、只留 16 行（保留率 7.5%），
    #   日志看着一切正常 ⇒ 静默丢 809 件、总计从 64,007 掉到 63,198。
    # ⇒ 两道体检：①过滤后保留率 ②归属月订单在报告中的占比。
    #   超阈值打【严重】，极端情况（保留率<30% 或 归属月占比<10%）直接中断，避免带病出数。
    #   紧急放行：环境变量 MSF_ALLOW_LOW_COVER=1。
    _allow_low = os.environ.get("MSF_ALLOW_LOW_COVER", "") == "1"
    if _rows_before_filter >= 20:
        _keep_ratio = len(order_df) / _rows_before_filter
        _bel_share = _pre_month_dist.get(belonging_date, 0) / _rows_before_filter
        _dist_txt = ", ".join(f"{k}:{v}" for k, v in sorted(_pre_month_dist.items()))
        _severe = []
        if _keep_ratio < 0.5:
            _severe.append(f"保留率仅 {_keep_ratio:.1%}（{_rows_before_filter} 行只剩 {len(order_df)} 行）")
        if _bel_share < 0.5:
            _severe.append(f"归属 {belonging_date} 的订单只占 {_bel_share:.1%}（实际月份分布 {_dist_txt}）")
        if _severe:
            _msg = (f"【严重·任务㉑】{entity}/{sale_group} 报告疑似导错月份区间："
                    + "；".join(_severe) + f" ——请核对 {file_name} 的导出起止日期")
            print(_msg)
            AMZ_SANITY_WARNINGS.append(_msg)
            if not _allow_low and (_keep_ratio < 0.30 or _bel_share < 0.10):
                raise Exception(_msg + f"（如确认无误，可设 MSF_ALLOW_LOW_COVER=1 放行）")
    order_df.sort_values(by=["purchase-date-local"],ascending=[True]).to_csv(os.path.join(tmp_dir,"order_df.csv"))
    return order_df


def se_a_dir(dir_path:str, *sub:str)-> str:
    """东南亚资料目录：旧版快照放在 <快照日>/东南亚/ 下；新版（20260915 起）直接平铺在 <快照日>/ 下。
    自动按实际存在与否选择，调用方不必关心快照布局。"""
    nested = os.sep.join([dir_path, "东南亚"] + list(sub))
    if os.path.isdir(os.path.join(dir_path, "东南亚")):
        return nested
    flat = os.sep.join([dir_path] + list(sub))
    return flat if os.path.exists(flat) or not sub else nested


# ===== 2026-11-10 新增：补充订单（源数据层补录，替代「跑完再改成品」）=====
# 背景：印尼线下单 YFDZ-20260924-001（8 SKU / 38 件）源报告日期是 9-24，但业务要求计进 10 月。
#   旧做法 = 跑完月销表后跑 apply_yfdz_*.py 改写成品，再在东南亚表里二次注入 —— 重跑必丢、
#   且三处（月销表 / 东南亚表 / 订单明细）容易不同步。
# 新做法 = 把该单的**原始报告行**原样放进快照目录的「补充订单」文件夹，读报告时直接 concat：
#   月销表、东南亚销量表、订单明细 sheet 全部同源，重跑天然幂等。
# 约定
#   · 位置：<快照日>/[东南亚/]补充订单/<渠道标签>补充订单<期>.xlsx
#   · 列结构：与被补的渠道报告**完全一致**（直接复制原始行，不改名、不派生新列）
#   · 幂等：按 key_col（订单号）去重 —— 主报告里已有的订单号不再补，重复跑不会叠加
#   · 停用：文件名里带 `_disabled` 即忽略
SUPPLEMENT_LOG = []


def merge_supplement_orders(order_df:pd.DataFrame, folder:str, tag:str, key_col:str)-> pd.DataFrame:
    """把 folder 下匹配 tag 的补录文件按 key_col 去重后并入 order_df（并入前后列结构一致）"""
    if not folder or not os.path.isdir(folder):
        return order_df
    files = sorted(f for f in os.listdir(folder)
                   if f.lower().endswith(".xlsx") and not f.startswith("~$")
                   and "_disabled" not in f and tag in f)
    if not files:
        return order_df
    if key_col not in order_df.columns:
        raise KeyError(f"补充订单并入失败：主报告里没有订单号列「{key_col}」"
                       f"（现有列：{list(order_df.columns)[:8]}…）")
    base_keys = set(order_df[key_col].astype(str).str.strip())
    add_frames = []
    for f in files:
        p = os.path.join(folder, f)
        sup = pd.read_excel(p)
        if key_col not in sup.columns:
            raise KeyError(f"补充订单文件 {f} 缺少订单号列「{key_col}」"
                           f"（现有列：{list(sup.columns)[:8]}…）")
        keys = sup[key_col].astype(str).str.strip()
        dup = keys.isin(base_keys)
        new = sup[~dup].reindex(columns=order_df.columns)
        print(f"【补充订单】{tag}：{f} 共 {len(sup)} 行，主报告已有 {int(dup.sum())} 行"
              f" → 实补 {len(new)} 行" + ("（幂等跳过，不会重复计入）" if dup.any() else ""))
        if len(new):
            add_frames.append(new)
            base_keys |= set(keys[~dup])
            SUPPLEMENT_LOG.append((tag, f, int(len(new))))
    if not add_frames:
        return order_df
    return pd.concat([order_df] + add_frames, ignore_index=True)


def get_belonging_shopee_order_df(entity:str, sale_group:str, belonging_date:str, snapshot_date:str)-> pd.DataFrame:
    '''
    获取虾皮订单数据
    '''
    dir_path = os.path.join("data","数据归属周期",belonging_date,"快照日期",snapshot_date)
    if entity == "本土":
        prefix = "Shopee"
    elif entity == "跨境":
        prefix = "Shopee跨境"
    else:
        raise Exception(f"不支持的主体{entity}")
    file_name = f"{prefix}{sale_group}站点{belonging_date}订单报告.xlsx"
    if entity == "本土":
        file_path = se_a_dir(dir_path, file_name)
    elif entity == "跨境":
        file_path = se_a_dir(dir_path, "跨境", file_name)
    order_id_col=""
    purchase_date_col=""
    sku_col=""
    quantity_col=""
    order_status_col = ""
    status_map = {}
    if sale_group == "泰国" and entity == "本土":
        order_id_col = "หมายเลขคำสั่งซื้อ"#泰语订单号码
        purchase_date_col = "วันที่ทำการสั่งซื้อ"#泰语下单日期
        sku_col = "เลขอ้างอิง SKU (SKU Reference No.)"#泰语SKU号
        quantity_col = "จำนวน"#泰语数量
        order_status_col = "สถานะการสั่งซื้อ"
        # 完整词表来自 9.15《2026年泰国易仓销售表》8月易仓虾皮订单表内的业务状态字典
        #（「订单状态」列 8 个：การจัดส่ง送货=1 / คำขอยกเลิก取消请求=0 / ที่ต้องจัดส่ง已收货=1 /
        #   ยกเลิกแล้ว取消=0 / ยังไม่ชำระ未付款=0 / สำเร็จแล้ว完成=1 /
        #   ผู้ซื้อได้รับสินค้าแล้ว买方已收到货物=1 / จัดส่งสำเร็จแล้ว成功交付=1）
        status_map = {
            "สำเร็จแล้ว":"Completed",            #完成
            "จัดส่งสำเร็จแล้ว":"Completed",      #成功交付
            "ผู้ซื้อได้รับสินค้าแล้ว":"Completed",#买方已收到货物（退货窗口期，见 return_window_prefix_map）
            "การจัดส่ง":"Shipping",              #送货
            "ที่ต้องจัดส่ง":"Shipping",          #已收货/待发货
            "ยกเลิกแล้ว":"Cancelled",            #取消
            "คำขอยกเลิก":"Cancelled",            #取消请求（业务标 0，须显式剔除）
            "ยังไม่ชำระ":"Cancelled",            #未付款
            #其他的状态就是退款
            }
    elif sale_group == "印尼" and entity == "本土":
        order_id_col = "No. Pesanan"#印尼语订单号码
        purchase_date_col = "Waktu Pesanan Dibuat"#印尼语下单日期
        sku_col = "Nomor Referensi SKU"#印尼语SKU号
        quantity_col = "Jumlah"#印尼语数量
        order_status_col = "Status Pesanan"
        # 完整词表来自 9.15《2026年印尼智慧物流销售表》08月虾皮订单表内的业务状态字典
        # （「Status Pesanan.1」+「消息状态」各 8 个：Batal=0/Belum Bayar=0/Pembatalan diajukan=0/
        #   Perlu Dikirim=1/Sedang Dikirim=1/Selesai=1/Pesanan diterima=1/Telah Dikirim=1）
        # ⚠️ 大小写：平台真实值是 'Pesanan diterima'（小写 d），下面必须与之**逐字符一致**，
        # 否则 .map() 查不到 → 变 NaN → 走 return_window_prefix_map 归一化时才被救回。
        status_map = {
            "Selesai":"Completed",               #完毕
            "Pesanan diterima":"Completed",      #已收到订单（无日期后缀的裸状态）
            "Sedang Dikirim":"Shipping",         #投递中
            "Perlu Dikirim":"Shipping",          #需要发送/待发货
            "Telah Dikirim":"Delivered",         #发送
            "Batal":"Cancelled",                 #取消
            "Belum Bayar":"Cancelled",           #未付款
            "Pembatalan diajukan":"Cancelled",   #取消申请
            #其他的状态就是退款
            }
    elif sale_group in ("菲律宾","新加坡","马来西亚"):
        order_id_col = "Order ID"
        purchase_date_col = "Order Creation Date"
        sku_col = "SKU Reference No."
        quantity_col = "Quantity"
        order_status_col = "Order Status"
        # 完整词表来自 9.15《2026年菲律宾仓销售表》8月虾皮订单表内的业务状态字典
        # （「Order Status.1」列 8 个：Cancelled/Completed/Shipping/To ship/Unpaid/
        #   Cancellation requested/Order Received/Delivered）
        status_map = {
            "Completed":"Completed",
            "Order Received":"Completed",        #已收货（无日期后缀的裸状态）
            "Shipping":"Shipping",
            "To ship":"Shipping",                #待发货
            "Delivered":"Delivered",
            "Cancelled":"Cancelled",
            "Unpaid":"Cancelled",                #未付款
            "Cancellation requested":"Cancelled",#取消申请
            #其他的状态就是退款
            }
    elif sale_group == "越南" and entity == "跨境":
        order_id_col = "订单编号"
        purchase_date_col = "订单成立时间"
        sku_col = "商品规格货号"
        quantity_col = "数量"
        order_status_col = "订单状态"
        status_map = {
            "Hoàn thành":"Completed",
            "Đã hủy":"Cancelled",
            # "Shipping":"Shipping",#暂未找到
            "Đã giao":"Delivered",
            #其他的状态就是退款
            }
    elif sale_group == "台湾" and entity == "跨境":
        # ★ 2026-11-02 新增：Shopee 跨境**台湾站**（中文表头，数量列名是 `Quantity per Item`）。
        #   状态只有三种：已完成 / 已送達 / 不成立。
        #   ⚠️「不成立」= 订单未成立（买家取消、付款失败等，导出里「取消原因」列有值）⇒ 必须剔除，
        #     否则 202610 会把 11 件取消单当成销量（台湾当月真销量是 6 件，不是 17 件）。
        order_id_col = "订单编号"
        purchase_date_col = "订单成立时间"
        sku_col = "商品规格货号"
        quantity_col = "Quantity per Item"
        order_status_col = "订单状态"
        status_map = {
            "已完成":"Completed",
            "已送達":"Delivered",
            "已送达":"Delivered",
            "不成立":"Cancelled",
            # 其他的状态就是退款
            }
    else:
        raise Exception(f"不支持的Shopee销售组{sale_group}")

    order_df = pd.read_excel(file_path, parse_dates=[purchase_date_col])
    order_df = order_df.rename(columns={
        order_id_col:"订单号",
        purchase_date_col:"下单时间",
        sku_col:"sku",
        quantity_col:"数量",
        order_status_col:"订单状态",
        })
    # 退货窗口期状态（已签收、仅处于退货窗口内）按口径计入销量；其原始文本带日期后缀
    #（如 "... until 2026-09-08"），故按前缀匹配归一化为该站的「完成」状态。
    # ⚠️ 归一化目标必须写**该站 status_map 里真实存在的键**（泰国="สำเร็จแล้ว"、印尼="Selesai"），
    # 否则 .map() 查不到 → 变 NaN → 该批订单被静默丢弃。
    # ⚠️ 9.15 实测补充：该状态**有两种形态** —— 带窗口日期的长文本，以及**窗口已过期的裸状态**
    #（泰国 204 件 = 'ผู้ซื้อได้รับสินค้าแล้ว'；印尼 53 件 = 'Pesanan diterima'；菲律宾 102 件 =
    #  'Order Received'）。裸状态已在各自 status_map 里显式登记为完成态，这里的前缀匹配负责
    # 带后缀的长文本，两者合起来才是完整的「已签收」口径。
    return_window_prefix_map = {
        "泰国":("ผู้ซื้อได้รับสินค้าแล้ว","สำเร็จแล้ว"),#泰语：买方已收到货物
        "印尼":("Pesanan diterima","Selesai"),        #印尼语：订单已收到
        "菲律宾":("Order Received, but buyer can raise return/refund","Completed"),
        "马来西亚":("Order Received, but buyer can raise return/refund","Completed"),
        "新加坡":("Order Received, but buyer can raise return/refund","Completed"),
        # 202609 实测新增：越南语「买家已确认收货、但仍在退货窗口内」的长文本
        #   Người mua xác nhận đã nhận được hàng, tuy nhiên ... Trả hàng/Hoàn tiền tới ngày YYYY-MM-DD.
        #   → 归一化成该站 status_map 里真实存在的完成键 "Hoàn thành"（写成别的键会 .map() 查不到 → 静默丢单）。
        "越南":("Người mua xác nhận đã nhận được hàng","Hoàn thành"),
        }
    if sale_group in return_window_prefix_map:
        rw_prefix, rw_completed_key = return_window_prefix_map[sale_group]
        rw_mask = order_df["订单状态"].astype(str).str.startswith(rw_prefix)
        order_df.loc[rw_mask, "订单状态"] = rw_completed_key
    order_df["订单状态"] = order_df["订单状态"].map(status_map)

    order_df["实体"] = entity
    order_df["订单组"] = sale_group
    order_df["平台"] = "Shopee"
    order_df["站点"] = sale_group
    if entity == "本土":
        order_df["仓库"] = sale_group
    elif entity == "跨境":
        order_df["仓库"] = "旭艺仓"
    order_df = order_df[(order_df["订单状态"] != "Cancelled") & (order_df["订单状态"].notna())].copy()
    order_df = order_df[order_df["站点"].notna()]
    # 202609 起：源报告可能整份为空（本期该渠道无销量 / 导出不全）——空源放行、按 0 计；
    # 源有数据、过滤后却全空 才说明 status_map 与状态取值不匹配（保留防呆）。
    if order_df.shape[0] == 0:
        print(f"  ⚠️ 【{sale_group}】该渠道源报告无数据行 → 按 0 计（若不应为 0，请检查导出是否完整）")
        return order_df
    assert order_df.shape[0] > 0, f"{sale_group}订单过滤后为空，请检查status_map与订单状态取值是否匹配"

    return order_df

def get_belonging_lazada_order_df(entity:str, sale_group:str, belonging_date:str, snapshot_date:str)-> pd.DataFrame:
    '''
    获取Lazada订单数据
    '''
    dir_path = os.path.join("data","数据归属周期",belonging_date,"快照日期",snapshot_date)
    if entity == "本土":
        prefix = "Lazada"
    elif entity == "跨境":
        prefix = "Lazada跨境"
    else:
        raise Exception(f"不支持的主体{entity}")
    file_name = f"{prefix}{sale_group}站点{belonging_date}订单报告.xlsx"
    if entity == "本土":
        file_path = se_a_dir(dir_path, file_name)
    elif entity == "跨境":
        file_path = se_a_dir(dir_path, "跨境", file_name)
    order_id_col=""
    purchase_date_col=""
    sku_col=""
    quantity_col=""
    order_status_col = ""
    # 业务口径（9.15 三工作簿「status.1」字典 15 个值，1=计入 / 0=不计入）：
    #   ✅ 计入：confirmed / shipped / delivered / packed / pending / ready_to_ship /
    #            Lost by 3PL / Package scrapped
    #   ❌ 剔除：canceled / returned / Package Returned / In Transit: Returning to seller /
    #            Unpaid / Buyer Delivery Failed / Seller Delivery Failed
    # ⚠️ 202608 实测两份 Lazada 表只出现 confirmed/canceled/Package Returned/returned/
    #    delivered/shipped 六种，业务口径与旧「三态白名单」结果**完全一致（差额 0）**；
    #    此处按业务口径放宽，是为了将来出现 packed/pending/ready_to_ship/Lost by 3PL/
    #    Package scrapped 时不被漏算。
    status_map = {
        # ✅ 计入销量（下方过滤条件只保留这些）
        "confirmed":"confirmed",
        "shipped":"shipped",
        "delivered":"delivered",
        "packed":"packed",
        "pending":"pending",
        "ready_to_ship":"ready_to_ship",
        "Lost by 3PL":"Lost by 3PL",              #第三方物流遗失（业务标 1）
        "Package scrapped":"Package scrapped",    #包裹报废（业务标 1）
        # ❌ 不计入销量（取消 / 退货 / 未成交 / 物流异常）
        "canceled":"canceled",
        "returned":"returned",
        "Package Returned":"Package Returned",
        "In Transit: Returning to seller":"In Transit: Returning to seller",
        "Unpaid":"Unpaid",
        "Buyer Delivery Failed":"Buyer Delivery Failed",
        "Seller Delivery Failed":"Seller Delivery Failed",
        }
    
    # ⚠️ 下面这段「原始版就存在的注释代码」，业务 2026-09-16 明确确认：**保留，不删**。
    #    它已被紧随其后的 4 行通用赋值完全取代（那 4 行对所有 sale_group 一视同仁），属等价冗余；
    #    留着是为了与原始版保持最小 diff、便于回溯对照（原始版基线见 HANDOFF 第一节）。
    # if sale_group == "泰国":
    #     order_id_col = "orderNumber"
    #     purchase_date_col = "createTime"
    #     sku_col = "sellerSku"
    #     order_status_col = "status"
    # elif sale_group == "菲律宾":
    #     order_id_col = "orderNumber"
    #     purchase_date_col = "createTime"
    #     sku_col = "sellerSku"
    #     order_status_col = "status"
    # else:
    #     raise Exception(f"不支持的{prefix}销售组{sale_group}")
    order_id_col = "orderNumber"
    purchase_date_col = "createTime"
    sku_col = "sellerSku"
    order_status_col = "status"

    order_df = pd.read_excel(file_path, parse_dates=[purchase_date_col])
    order_df = order_df.rename(columns={
        order_id_col:"订单号",
        purchase_date_col:"下单时间",
        sku_col:"sku",
        order_status_col:"订单状态",
        })
    order_df["数量"] = 1#lazada每条数据表示一个数量产品的订单
    order_df["订单状态"] = order_df["订单状态"].map(status_map)

    order_df["实体"] = entity
    order_df["订单组"] = sale_group
    order_df["平台"] = "Lazada"
    order_df["站点"] = sale_group
    if entity == "本土":
        order_df["仓库"] = sale_group
    elif entity == "跨境":
        order_df["仓库"] = "旭艺仓"
    order_df = order_df[order_df["订单状态"].isin(
        ["shipped", "delivered", "confirmed", "packed", "pending", "ready_to_ship",
         "Lost by 3PL", "Package scrapped"])].copy()
    order_df = order_df[order_df["站点"].notna()]
    # 202609 起：源报告可能整份为空（本期该渠道无销量 / 导出不全）——空源放行、按 0 计；
    # 源有数据、过滤后却全空 才说明 status_map 与状态取值不匹配（保留防呆）。
    if order_df.shape[0] == 0:
        print(f"  ⚠️ 【{sale_group}】该渠道源报告无数据行 → 按 0 计（若不应为 0，请检查导出是否完整）")
        return order_df
    assert order_df.shape[0] > 0, f"{sale_group}订单过滤后为空，请检查status_map与订单状态取值是否匹配"

    return order_df

def get_belonging_tkshop_order_df(entity:str, sale_group:str, belonging_date:str, snapshot_date:str)-> pd.DataFrame:
    '''
    获取Lazada订单数据
    '''
    dir_path = os.path.join("data","数据归属周期",belonging_date,"快照日期",snapshot_date)
    if entity == "本土":
        prefix = "Tiktok"
    else:
        raise Exception(f"不支持的主体{entity}")
    file_name = f"{prefix}{sale_group}站点{belonging_date}订单报告.csv"
    file_path = se_a_dir(dir_path, file_name)
    order_id_col=""
    purchase_date_col=""
    sku_col=""
    quantity_col=""
    order_status_col = ""
    # Tiktok状态为英文原值，保留 Shipped/Completed（见下方过滤条件），其余（Canceled）被剔除。
    # ⚠️ 9.15 三工作簿实测：抖音表存在**中英+印尼语三语混用**的状态值（同一列既有 'Completed'
    # 也有 '已完成'/'Selesai'），且出现未成交态 'To ship' / '待发货' / 'Perlu dikirim'。
    # 完整词表来自 9.15 三工作簿抖音订单表内的业务字典（「Order Status.1」+ 中文/印尼语译名 + 0/1）：
    #   菲：Canceled=0 / Completed=1 / To ship=1 / Shipped=1 / 待发货=1 / 已发货=1 / 已完成=1 / 已取消=0
    #   印尼：Canceled=0 / Completed=1 / To ship=1 / Shipped=1 / Unpaid=0 / Dikirim=1 /
    #        Dibatalkan=0 / Selesai=1 / Belum dibayar=0 / Perlu dikirim=1
    status_map = {
        # ✅ 计入销量（下方过滤条件只保留 Shipped/Completed 这两类）
        "Completed":"Completed",
        "已完成":"Completed",
        "Selesai":"Completed",      #印尼语：完成
        "Shipped":"Shipped",
        "已发货":"Shipped",
        "Dikirim":"Shipped",        #印尼语：已发货/运输中
        "To ship":"To ship",        #业务标 1（未发货，但口径计入）——见下方过滤条件
        "待发货":"To ship",
        "Perlu dikirim":"To ship",  #印尼语：需要发送
        # ❌ 不计入销量
        "Canceled":"Canceled",
        "Cancelled":"Canceled",
        "已取消":"Canceled",
        "Dibatalkan":"Canceled",    #印尼语：已取消
        "Unpaid":"Canceled",        #未付款
        "Belum dibayar":"Canceled", #印尼语：尚未付款
        }
    
    if sale_group == "印尼":
        order_id_col = "Order ID"
        purchase_date_col = "Created Time"
        sku_col = "Seller SKU"
        order_status_col = "Order Status"
        quantity_col = "Quantity"
    elif sale_group == "菲律宾":
        order_id_col = "Order ID"
        purchase_date_col = "Created Time"
        sku_col = "Seller SKU"
        order_status_col = "Order Status"
        quantity_col = "Quantity"
    else:
        raise Exception(f"不支持的{prefix}销售组{sale_group}")

    # order_df = pd.read_excel(file_path, parse_dates=[purchase_date_col])#tiktok导出的excel pandas读出来只有第一列orderid，不知道为什么
    encode = file_chardetect.detect_encoding(file_path)[0]
    order_df = pd.read_csv(file_path, encoding=encode, parse_dates=[purchase_date_col])
    order_df.info()

    order_df = order_df.rename(columns={
        order_id_col:"订单号",
        purchase_date_col:"下单时间",
        sku_col:"sku",
        order_status_col:"订单状态",
        })
    order_df = order_df.rename(columns={quantity_col:"数量"})
    order_df["订单状态"] = order_df["订单状态"].map(status_map)

    order_df["实体"] = entity
    order_df["订单组"] = sale_group
    order_df["平台"] = "Tiktok"
    order_df["站点"] = sale_group
    order_df["仓库"] = sale_group
    # 业务口径（9.15 三工作簿字典）：Shipped/Completed/**To ship**（含待发货/Perlu dikirim）均计入。
    # ⚠️ 202608 实测抖音只有 Completed/Canceled/Shipped，故把 To ship 纳入不改变当月数字；
    # 但若将来出现待发货订单，旧写法会漏算，故按业务口径一并保留。
    order_df = order_df[order_df["订单状态"].isin(["Shipped","Completed","To ship"])].copy()
    order_df = order_df[order_df["站点"].notna()]
    # 202609 起：源报告可能整份为空（本期该渠道无销量 / 导出不全）——空源放行、按 0 计；
    # 源有数据、过滤后却全空 才说明 status_map 与状态取值不匹配（保留防呆）。
    if order_df.shape[0] == 0:
        print(f"  ⚠️ 【{sale_group}】该渠道源报告无数据行 → 按 0 计（若不应为 0，请检查导出是否完整）")
        return order_df
    assert order_df.shape[0] > 0, f"{sale_group}订单过滤后为空，请检查status_map与订单状态取值是否匹配"

    return order_df

# ---- 菲律宾新订单报告：「订单类型」口径（2026-10-16 业务确认）----
# 2026-10-16 用户确认：2026-10 期起菲律宾销量统一在这一份新订单报告里，手工单统计方式与
# 202609 的 145 件相同——「订单类型」列中**除 Api订单 外的标签**（如 RESELLER / 线下 / 补单等）
# 一律按手工单处理，计入「线下订单」渠道（KSON），不随平台继承进 Shopee/Lazada/TikTok 列。
# 202609 期实测该列只有两个值：Api订单（2096 件）/ 手工订单（145 件）。
PH_MANUAL_ORDER_TYPES = {"手工订单"}                        # 已知手工单标签
PH_API_ORDER_TYPES = {"Api订单", "API订单", "api订单"}       # 已知平台 API 单标签
# 未知标签的处理：按「非 Api 即手工单」计入线下渠道，并在跑批日志打印警告（⚠️），
# 便于当月人工确认；若某标签不应计入，把它加进 PH_API_ORDER_TYPES 或改为显式排除。

def get_belonging_ph_new_order_df(entity:str, sale_group:str, belonging_date:str, snapshot_date:str,
                                  start_date:str=None)-> pd.DataFrame:
    '''
    菲律宾【新订单报告】读取（2026-10-09 新增，默认不启用，见 __main__ 顶部 USE_PH_NEW_ORDER_REPORT）。

    新工具导出的《菲律宾{YYYYMM}订单报告.xlsx》单份即覆盖 Lazada + Shopee + Tiktok + 手工订单，
    目标是「以一敌多」替换上面 get_belonging_{shopee,lazada,tkshop}_order_df 三份分渠道报告。

    ⚠️ 启用条件（证据见《核查-菲律宾新订单报告能否一敌多.md》）：
      · 202609 实测该报告只覆盖 **9/13 起**（9/1~9/12 整段缺失，全月净缺 967 件）→ 本期**不能替换**；
      · 覆盖窗口内（9/13~9/30）新报告 ≥ 旧三报告（+188 件，其中 145 件为旧报告完全没有的手工订单）；
      · 待 2026-11 用户确认 10 月新报告自 10/1 起满量后，把 __main__ 顶部开关置 True 正式切换。

    读表四要点（踩坑记录）：
      ① 这是**合并单元格导出**：同一订单的多行明细只在首行写父级字段，后续行留空 →
         必须先对「除行级列外的全部列」ffill，否则平台/状态/单号会读成空。
      ② SKU 带内部前缀 "66128-"（旧报告是裸 SKU，如 G376）→ 去前缀后才能与商品资料/旧报告对齐。
      ③ 订单状态只有 已交运 / 已发货 / 取消 三态 → 计销量取「已交运 + 已发货」（取消不计）。
      ④ 「订单类型」列中除 Api订单 外的行**归「线下订单」渠道**（平台/仓库改写成 KSON，
         与 KSON 线下报告同键），不随 ffill 继承的平台（Shopee）进 Shopee 列——
         2026-10-16 业务确认（与印尼 YFDZ 手工单归线下同口径；202609 该列即 手工订单 145 件）。
         口径常量见函数上方 PH_MANUAL_ORDER_TYPES / PH_API_ORDER_TYPES；未知标签按「非 Api
         即手工单」计入并打 ⚠️ 警告，跑批时需人工复核。
    '''
    if sale_group != "菲律宾":
        raise Exception("新订单报告仅菲律宾")
    if entity != "本土":
        raise Exception("新订单报告当前仅本土（跨境 Lazada 菲律宾仍走旧报告）")
    dir_path = os.path.join("data","数据归属周期",belonging_date,"快照日期",snapshot_date)
    file_name = f"{sale_group}{belonging_date}订单报告.xlsx"
    file_path = se_a_dir(dir_path, file_name)
    # 2026-11 起后台改了导出文件名：加了「ESTP易速菲-」前缀和「(MWS)」后缀，
    # 精确拼名找不到时按「菲律宾{YYYYMM}订单报告*」兜底，避免直接 FileNotFoundError。
    if not os.path.exists(file_path):
        _root = os.path.dirname(file_path)
        _key = f"{sale_group}{belonging_date}订单报告"
        _hits = sorted(
            os.path.join(_root, f) for f in os.listdir(_root)
            if _key in f and f.lower().endswith((".xlsx", ".xls")) and not f.startswith("~$")
        ) if os.path.isdir(_root) else []
        if _hits:
            file_path = _hits[0]
            print(f"    ℹ️ 【菲律宾·新订单报告】文件名兜底命中：{os.path.basename(file_path)}")

    raw = pd.read_excel(file_path, dtype=str)
    # 坑①：merged-cell 导出 → 除行级列外全部向下填充
    _line_level_cols = ["订单ID","Sku","数量","多属性Id","平台数量","sku图片","sku中文名称"]
    _parent_cols = [c for c in raw.columns if c not in _line_level_cols]
    raw[_parent_cols] = raw[_parent_cols].ffill()

    # 计销量口径：已交运 + 已发货（取消不计）
    order_df = raw[raw["订单状态"].isin(["已交运","已发货"])].copy()
    # ⚠️ 新报告同时有「订单号」（内部号 R66128…，不可 join）与「交易号」（平台单号，可 join）。
    #    下游统一用「订单号」，故先把内部号让位为「内部订单号」，再把「交易号」提到「订单号」。
    order_df = order_df.rename(columns={"订单号":"内部订单号","交易号":"订单号","创建时间":"下单时间","Sku":"sku"})

    # 坑②：SKU 带**客户编码前缀**（本期 = "66128-"，与该报告「客户」列的值完全一致，2067 行全是 66128）
    #   → 去掉前缀才是商品 SKU（G376 这种）。
    #   ⚠️ **不写死 66128**：用「前导数字串 + '-'」剥离——商品 SKU 一律「字母+数字」，绝不会以数字开头，
    #   所以这条规则既精确又自适应（客户编码变了、或以后不带前缀，都能正确处理），并把实际剥掉的域名打印留痕。
    _sku_raw = order_df["sku"].astype(str).str.strip()
    _prefix_seen = _sku_raw.str.extract(r"^(\d+)-")[0]
    if int(_prefix_seen.notna().sum()):
        print(f"    ℹ️ 【菲律宾·新订单报告】剥除 SKU 客户编码前缀：{_prefix_seen.dropna().value_counts().to_dict()}")
    order_df["sku"] = _sku_raw.str.replace(r"^\d+-", "", regex=True).replace("nan", np.nan)
    # 剥完仍不像「字母+数字」→ 安静放过会污染下游关联，必须显式告警（不抛错，交给人工核）
    _weird = order_df["sku"].notna() & ~order_df["sku"].astype(str).str.match(r"^[A-Za-z]+[0-9]+$")
    if int(_weird.sum()):
        print(f"    ⚠️ 【菲律宾·新订单报告】{int(_weird.sum())} 行 SKU 形态异常（非「字母+数字」），请人工确认："
              f"{order_df.loc[_weird, 'sku'].unique()[:10].tolist()}")
    # 单号去空白（旧 TikTok 单号带尾随 Tab、Shopee 有长短格式）
    order_df["订单号"] = order_df["订单号"].astype(str).str.replace(r"\s+","",regex=True)
    order_df["下单时间"] = pd.to_datetime(order_df["下单时间"], errors="coerce")
    order_df["数量"] = pd.to_numeric(order_df["数量"], errors="coerce")

    # 仅在「新报告完整覆盖」的日期之后取值：早于 start_date 的订单新报告缺失，
    # 不能当作 0（否则会静默少算）——这一段是否需要补由调用方决定（本期=启用时不补）。
    if start_date:
        _start = pd.to_datetime(start_date, errors="coerce")
        _n_before = int((order_df["下单时间"] < _start).sum())
        order_df = order_df[order_df["下单时间"] >= _start].copy()
        print(f"    ℹ️ 【菲律宾·新订单报告】按完整覆盖起点 {start_date} 过滤，剔除 {_n_before} 行"
              f"（该段新报告缺失，如需凑齐当月须另用旧报告补齐）")

    # 平台映射：脚本内部统一写法为 "Tiktok"（小写 k），与旧 Tiktok 报告产物一致。
    _plat_map = {"Lazada":"Lazada","Shopee":"Shopee","TikTok":"Tiktok"}
    order_df["平台"] = order_df["平台类型"].map(_plat_map)
    order_df["实体"] = entity
    order_df["订单组"] = sale_group
    order_df["站点"] = sale_group
    # 仓库：沿用旧三份报告的口径（本土 → 站点名），保证切换后 groupby 出的行能与旧路径逐行对齐；
    # 新报告原始仓库（PH-17）保留到附属列便于追溯。
    order_df["源仓库"] = order_df["仓库"]
    order_df["仓库"] = sale_group

    # 2026-10-16：手工/线下订单 → 归「线下订单」渠道（分组键 ("本土","KSON","菲律宾")，与 KSON 线下报告一致）。
    # ⚠️ 不能随平台继承：合并单元格 ffill 会把手工单的平台继承成 Shopee（其上一单的平台），
    #    导致手工单混进「菲仓shopee」列。业务口径 = 菲律宾手工单属线下订单
    #    （与印尼 YFDZ 手工单归「印尼线下订单」同口径，202609 印尼 200 件即此处理）。
    # 2026-10-16 用户补充：10 月起菲律宾销量统一在本报告，与 145 件同口径——「订单类型」除
    #    Api订单 外的标签（RESELLER/线下/补单…）都按手工单计入线下渠道；未知标签打 ⚠️ 警告。
    _type_s = order_df["订单类型"].astype(str).str.strip()
    _manual_mask = _type_s.isin(PH_MANUAL_ORDER_TYPES)
    _unknown_mask = (~_type_s.isin(PH_MANUAL_ORDER_TYPES | PH_API_ORDER_TYPES)
                     & ~_type_s.isin(["", "nan", "None"]))
    if int(_unknown_mask.sum()):
        for _t, _grp in order_df.loc[_unknown_mask].groupby("订单类型"):
            print(f"    ⚠️ 【菲律宾·新订单报告】未知订单类型 {_t!r}：{len(_grp)} 行 / "
                  f"{int(_grp['数量'].sum())} 件 → 按「非 Api 即手工单」计入「线下订单」渠道"
                  f"（如不应计入，请改 PH_API_ORDER_TYPES / 显式排除）")
        _manual_mask = _manual_mask | _unknown_mask
    order_df["_is_manual"] = _manual_mask
    if int(_manual_mask.sum()):
        order_df.loc[_manual_mask, "平台"] = "KSON"
        order_df.loc[_manual_mask, "仓库"] = "KSON"
        print(f"    ℹ️ 【菲律宾·新订单报告】手工/线下订单 {int(_manual_mask.sum())} 行 / "
              f"{int(order_df.loc[_manual_mask, '数量'].sum())} 件 → 归「线下订单」渠道（KSON），不进平台列")

    # 清洗：平台识别不出 / SKU 为空 / 数量为空 的明细丢弃，避免污染汇总
    _n_bad_plat = int(order_df["平台"].isna().sum())
    if _n_bad_plat:
        print(f"    ⚠️ 【菲律宾·新订单报告】{_n_bad_plat} 行平台类型无法识别，已丢弃")
    order_df = order_df[order_df["平台"].notna() & order_df["sku"].notna() & order_df["数量"].notna()].copy()

    if order_df.shape[0] == 0:
        print("    ⚠️ 【菲律宾·新订单报告】过滤后无数据行 → 按 0 计（若不应为 0，请检查导出是否完整）")
    else:
        _manual = order_df[order_df["_is_manual"]]
        if _manual.shape[0]:
            print(f"    ℹ️ 【菲律宾·新订单报告】手工/线下订单 {_manual['订单号'].nunique()} 单 / "
                  f"{int(_manual['数量'].sum())} 件（旧三报告完全没有，属净增量；已归「线下订单」渠道）")

    return order_df


def get_last30d_emag_order_df(sale_group:str, belonging_date:str, snapshot_date:str):
    dir_path = os.path.join("data","数据归属周期",belonging_date,"快照日期",snapshot_date)
    file_name = f"EMAG{sale_group}订单报告{snapshot_date}.txt"
    file_path = os.sep.join([dir_path, file_name])
    order_df = pd.read_csv(file_path, sep="\t", parse_dates=["purchase-date"])
    order_df["站点"] = order_df["sales-channel"].map(channel_country_map)
    order_df["订单组"] = order_df["站点"].map(amazon_sale_group)
    order_df["仓库"] = order_df["站点"].map(amazon_country_inv_map)
    order_df["日期"] = order_df["purchase-date"].dt.normalize()
    order_df = order_df[order_df["order-status"] != "Cancelled"]
    order_df = order_df[order_df["站点"].notna()]
    order_df = order_df[~order_df["sku"].str.startswith("Amazon.Found.")]#去除其他用途sku号
    order_df = order_df[~order_df["sku"].str.startswith("amzn.gr.")]#去除其他用途sku号
    #去除快照日期当天的订单，因为不完整
    # order_df = order_df[order_df["日期"] != pd.to_datetime(snapshot_date, format="%Y%m%d", utc=True)]#utc时间和普通时间戳不能比较，必须转换相同
    order_df = order_df.copy().reset_index()
    return order_df

def get_belonging_kson_offline_order_df(entity:str, sale_group:str, belonging_date:str, snapshot_date:str)-> pd.DataFrame:
    '''
    kson的线下订单
    '''
    if entity != "本土":
        raise Exception("kson是本土销售")
    if sale_group != "菲律宾":
        raise Exception("kson是菲律宾销售")
    dir_path = os.path.join("data","数据归属周期",belonging_date,"快照日期",snapshot_date)
    file_name = f"{sale_group}KSON仓{belonging_date}线下订单报告.xlsx"
    file_path = se_a_dir(dir_path, file_name)
    order_df = pd.read_excel(file_path)
    order_df["实体"] = "本土"
    order_df["平台"] = "KSON"
    order_df["仓库"] = "KSON"
    order_df["站点"] = "菲律宾"
    order_df["订单组"] = sale_group
    order_df["日期"] = datetime.datetime.strptime(belonging_date, "%Y%m")
    order_df = order_df.rename(columns={"SKU":"sku"})
    return order_df

def get_belonging_indonesia_warehouse_offline_order_df(entity:str, sale_group:str, belonging_date:str, snapshot_date:str)-> pd.DataFrame:
    '''
    印尼海外仓
    '''
    if entity != "本土":
        raise Exception("这是印尼海外仓")
    if sale_group != "印尼":
        raise Exception("这是印尼海外仓")
    dir_path = os.path.join("data","数据归属周期",belonging_date,"快照日期",snapshot_date)
    file_name = f"{sale_group}海外仓{belonging_date}订单报告.xlsx"
    file_path = se_a_dir(dir_path, file_name)
    order_df = pd.read_excel(file_path)
    # ★ 2026-11-10：源数据层补录（如 YFDZ-20260924-001 跨月单）——必须在改名/过滤之前并入，
    #   这样后面的 rename / YFDZ+OUTER 过滤 / 渠道键赋值全部一视同仁，不会漏字段。
    order_df = merge_supplement_orders(order_df, se_a_dir(dir_path, "补充订单"),
                                       "印尼海外仓", "订单号\nOrder No.")
    order_df["实体"] = "本土"
    order_df["平台"] = "线下"
    order_df["仓库"] = "印尼海外仓"
    order_df["站点"] = "印尼"
    order_df["订单组"] = sale_group
    
    order_id_col = "订单号\nOrder No."
    purchase_date_col = "订单时间\nOrder Time"
    sku_col = "商品编码\nGoods No"
    order_status_col = "订单状态\nStatus"
    quantity_col = "商品数量\nQuantity"
    order_source_col = "订单来源\nOrder Source"
    order_df = order_df.rename(columns={
        order_id_col:"订单号",
        purchase_date_col:"下单时间",
        sku_col:"sku",
        order_status_col:"订单状态",
        order_source_col:"订单来源",
        })
    order_df = order_df.rename(columns={quantity_col:"数量"})

    order_df = order_df.rename(columns={"SKU":"sku"})
    # ★ 2026-11-10：把「订单类型」也带出来 —— 只服务于「国家订单明细」sheet 的展示列，
    #   不参与任何聚合。不带的话线下渠道那一列在明细里会是空的（业务靠它区分手工单/普通单）。
    if "订单类型\nOrder Type" in order_df.columns and "订单类型" not in order_df.columns:
        order_df = order_df.rename(columns={"订单类型\nOrder Type": "订单类型"})
    #只获取线下订单
    #线下单 = 手动导入的订单，单号以 YFDZ 开头（等价于原「订单来源 == MANUAL(手动导入)」，但不受来源分类改名影响）
    # ⚠️ 9.15 该表右侧另有业务维护的「上马系统状态 + 0/1」字典（PICKING/PACKED/OUTER/WH_REVIEW/
    #    PICKED/REVIEWED = 1，CANCELLED(回收站)/MT_REVIEW(商户待审核) = 0）。那是**上马系统**这一
    #    独立进程的状态，本表主状态列恒为 OUTER(已出库)（业务标 1），两者口径一致；
    #    故这里仍只按「订单状态 == OUTER(已出库)」过滤即可，无需引入上马状态列。
    order_df = order_df[
        (order_df["订单号"].astype(str).str.strip().str.upper().str.startswith("YFDZ"))
        & (order_df["订单状态"] == "OUTER(已出库)")
    ].copy()
    assert len(order_df) > 0, "印尼海外仓：订单号 YFDZ 开头的已出库订单为空，请检查订单号列或订单来源口径"
    return order_df

def get_belonging_jingdong_order_df(entity:str, sale_group:str, belonging_date:str, snapshot_date:str)-> pd.DataFrame:
    '''
    京东国内站
    '''
    if entity != "巨川":
        raise Exception("这是京东国内站")
    if sale_group != "国内":
        raise Exception("这是京东国内站")
    dir_path = os.path.join("data","数据归属周期",belonging_date,"快照日期",snapshot_date)
    file_name = f"京东{sale_group}{belonging_date}订单报告.xlsx"
    file_path = os.sep.join([dir_path, file_name])
    order_df = pd.read_excel(file_path)
    order_df["实体"] = "巨川"
    order_df["平台"] = "京东"
    order_df["仓库"] = "旭艺仓"
    order_df["站点"] = sale_group
    order_df["订单组"] = sale_group
    
    order_id_col = "订单号"
    purchase_date_col = "下单时间"
    sku_col = "商家SKUID"
    order_status_col = "订单状态"
    quantity_col = "订购数量"
    # order_source_col = "订单来源\nOrder Source"
    order_df = order_df.rename(columns={
        order_id_col:"订单号",
        purchase_date_col:"下单时间",
        sku_col:"sku",
        order_status_col:"订单状态",
        # order_source_col:"订单来源",
        })
    order_df = order_df.rename(columns={quantity_col:"数量"})
    order_df = order_df[(order_df["订单状态"] == "完成")].copy()
    return order_df

def get_belonging_emag_order_df(entity:str, sale_group:str, belonging_date:str, snapshot_date:str)-> pd.DataFrame:
    '''
    EMAG
    '''
    if entity != "巨川":
        raise Exception("这是EMAG")
    if sale_group != "eMag":
        raise Exception("这是EMAG")
    prefix = "EMAG"
    dir_path = os.path.join("data","数据归属周期",belonging_date,"快照日期",snapshot_date)
    file_name = f"{prefix}全站点{belonging_date}订单报告.xlsx"
    file_path = os.sep.join([dir_path, file_name])
    order_df = pd.read_excel(file_path)
    order_df["实体"] = "巨川"
    order_df["平台"] = "EMAG"
    order_df["仓库"] = "EMAG"
    order_df["站点"] = "EMAG"
    order_df["订单组"] = sale_group

    # 202609 实测：eMAG 报告有**英文列名 / 中文列名**两种导出（202608 是英文、202609 是中文界面导出），
    # 两版列一一对应。先把中文列名对齐成英文列名，后面归一名逻辑完全复用 → 两种版本都能跑。
    # ⚠️ 不做这一步的后果是"静默丢单"：中文表没有 sku 列 → 汇总后 SKU 为空 → 被后面的
    #    "丢弃空白SKU" 清洗整批丢掉，既不报错也不计销量（202609 期 eMAG 曾因此计 0）。
    _emag_zh_alias = {"订单编号": "Order number", "订单日期": "Order date",
                      "产品代码": "Product code", "数量": "Quantity",
                      "订单状态": "Order status"}
    for _zh, _en in _emag_zh_alias.items():
        if _en not in order_df.columns and _zh in order_df.columns:
            order_df = order_df.rename(columns={_zh: _en})

    order_id_col = "Order number"
    purchase_date_col = "Order date"
    sku_col = "Product code"
    order_status_col = "Order status"
    quantity_col = "Quantity"
    # order_source_col = "订单来源\nOrder Source"
    order_df = order_df.rename(columns={
        order_id_col:"订单号",
        purchase_date_col:"下单时间",
        sku_col:"sku",
        order_status_col:"订单状态",
        # order_source_col:"订单来源",
        })
    order_df = order_df.rename(columns={quantity_col:"数量"})
    order_df = order_df[~(order_df["订单状态"].isin(["Storno","Canceled by customer"]))].copy()
    return order_df

def get_belonging_temu_order_df(entity:str, sale_group:str, belonging_date:str, snapshot_date:str)-> pd.DataFrame:
    '''
    Temu
    '''
    if entity != "巨川":
        raise Exception("这是Temu")
    if sale_group != "TEMU平台":
        raise Exception("这是Temu")
    prefix = "Temu"
    dir_path = os.path.join("data","数据归属周期",belonging_date,"快照日期",snapshot_date)
    file_name = f"{prefix}全球{belonging_date}订单报告.csv"
    file_path = os.sep.join([dir_path, file_name])
    encoding = file_chardetect.detect_encoding(file_path)[0]
    order_df = pd.read_csv(file_path, encoding=encoding)
    order_df["实体"] = "巨川"
    order_df["平台"] = "Temu"
    order_df["仓库"] = "亚马逊美国仓"
    # order_df["站点"] = "Temu"#temu报告有站点字段，这个字段会指示订单国家
    order_df["订单组"] = sale_group

    order_df.drop(["订单号","订单状态"], axis=1, inplace=True)

    order_id_col = "子订单号"
    purchase_date_col = "订单创建时间"
    sku_col = "SKU货号"
    order_status_col = "子订单状态"
    quantity_col = "应履约件数"
    # order_source_col = "订单来源\nOrder Source"#海外仓订单会有这个字段，表示支持了哪个平台的履约
    order_df = order_df.rename(columns={
        order_id_col:"订单号",
        purchase_date_col:"下单时间",
        sku_col:"sku",
        order_status_col:"订单状态",
        # order_source_col:"订单来源",
        })
    order_df = order_df.rename(columns={quantity_col:"数量"})
    order_df = order_df[~(order_df["订单状态"].isin(["已取消"]))].copy()
    return order_df

def get_belonging_mercado_order_df(entity:str, sale_group:str, belonging_date:str, snapshot_date:str)-> pd.DataFrame:
    '''
    Temu
    '''
    if entity != "巨川":
        raise Exception("这是美客多")
    if sale_group != "美客多":
        raise Exception("这是美客多")
    prefix = "美客多墨西哥"
    dir_path = os.path.join("data","数据归属周期",belonging_date,"快照日期",snapshot_date)
    file_name = f"{prefix}站点{belonging_date}订单报告.xlsx"
    file_path = os.sep.join([dir_path, file_name])
    order_df = pd.read_excel(file_path, skiprows=5)
    order_df["实体"] = "巨川"
    order_df["平台"] = "美客多"
    order_df["仓库"] = "美客多墨西哥仓"
    order_df["站点"] = "墨西哥"
    order_df["订单组"] = sale_group
    # 美客多报告有**中文版 / 西班牙语原版**两种（202608 给的是中文版、202609 给的是西语原版），
    # 两版列一一对应。这里先按别名把西语列名统一成中文列名，后面逻辑完全不变 → 两种版本都能跑。
    _mercado_es_alias = {"订单号": "# de venta", "销售日期": "Fecha de venta",
                         "状态": "Estado", "单位": "Unidades"}
    for _zh, _es in _mercado_es_alias.items():
        if _zh not in order_df.columns and _es in order_df.columns:
            order_df = order_df.rename(columns={_es: _zh})
    order_id_col = "订单号"
    purchase_date_col = "销售日期"
    sku_col = "SKU"
    order_status_col = "状态"
    quantity_col = "单位"
    # order_source_col = "订单来源\nOrder Source"#海外仓订单会有这个字段，表示支持了哪个平台的履约
    order_df = order_df.rename(columns={
        order_id_col:"订单号",
        purchase_date_col:"下单时间",
        sku_col:"sku",
        order_status_col:"订单状态",
        # order_source_col:"订单来源",
        })
    order_df = order_df.rename(columns={quantity_col:"数量"})
    # 剔除"未成交"状态：中文版 退款/退货/取消；西语版 reembols*/Devoluci*/Cancelad*/Anulaci*
    not_available_keyword = ["退款", "退货", "取消", "reembols", "Devoluci", "Cancelad", "Anulaci"]
    for keyword in not_available_keyword:
        order_df = order_df[~(order_df["订单状态"].str.contains(keyword, case=False, na=False))].copy()
    return order_df

def get_manage_inv_df(inv_group:str,belonging_date:str,snapshot_date:str):
    dir_path = os.path.join("data","数据归属周期",belonging_date,"快照日期",snapshot_date)
    file_name = f"亚马逊{inv_group}管理库存报告{snapshot_date}.csv"
    file_path = resolve_report_path(dir_path, file_name)
    inv_df = pd.read_csv(file_path)
    inv_df["平台"] = "亚马逊"
    inv_df["仓库"] = inv_group

    inv_df["日期"] = pd.to_datetime(snapshot_date, format='%Y%m%d')
    inv_df["日期"] = inv_df["日期"]# - pd.Timedelta(days=1)
    inv_df["日期"] = inv_df["日期"].dt.date

    inv_df = inv_df[~inv_df["sku"].str.startswith("Amazon.Found.")]
    inv_df = inv_df[~inv_df["sku"].str.startswith("amzn.gr.")]

    return inv_df.copy().reset_index()

def get_product_info_df():
    file_name = "商品资料.xlsx"
    product_info_df = pd.read_excel(file_name, sheet_name="商品资料", skiprows=1)
    return product_info_df

def get_manual_purchase_plan(date:str):
    dir_path = os.path.join("compare",f"{date}")
    file_path = os.path.join(dir_path, f"采购计划表{date}.xlsx")
    manual_purchase_plan_df = pd.read_excel(file_path, sheet_name="常规采购与备货", skiprows=2)
    manual_purchase_plan_df.to_csv(os.path.join("tmp", "manual_purchase_plan_df.csv"))
    return manual_purchase_plan_df

# ===== 2026-10-17 新增：得分列表头自愈（模板=复制上月成品 → KH/EP 表头停上期）=====
# 月销总表模板沿用上月成品时，【{year}年】KH2（294列）表头会停在上期得分列名
# （202610 实测：staged 模板停在「8月销量得分」，而 10 月期望「9月销量得分」）——
# 过去每月首次跑要手工改一格，否则补充信息处的 assert 直接崩。
# 现按当月 score_col_name 自动 rename，三层防护：
#   ① 模板侧：get_month_sale_df → heal_score_col_header_file（改名并落盘，一劳永逸）；
#   ② 成品侧：任务② 共用的 wb 里 rename_score_col_in_ws 兜底（模板被占用没能落盘时）；
#   ③ 内存侧：get_month_sale_df 对 pandas df 列名兜底改名（模板被占用时 assert 也不崩）。
_SCORE_HEADER_RE = re.compile(r"^\d{1,2}月销量得分$")

def _find_score_header_col(ws):
    """在 worksheet 行2 找「N月销量得分」表头列号；找不到返回 None。"""
    for _c in range(2, ws.max_column + 1):
        if _SCORE_HEADER_RE.match(str(ws.cell(2, _c).value or "").strip()):
            return _c
    return None

def rename_score_col_in_ws(ws, score_col_name:str):
    """把 ws 行2 的「N月销量得分」表头改成 score_col_name（只改内存，落盘由调用方保存）。
    返回 (旧表头, 新表头, 列号)；无需改（已是当月 / 行2 没有得分表头 / 未传列名）时返回 None。"""
    if not score_col_name:
        return None
    _col = _find_score_header_col(ws)
    if _col is None:
        return None
    _old = str(ws.cell(2, _col).value).strip()
    if _old == score_col_name:
        return None
    ws.cell(2, _col).value = score_col_name
    return (_old, score_col_name, _col)

def heal_score_col_header_file(file_path:str, sheet_name:str, score_col_name:str, save:bool=True)->str:
    """文件级得分列表头自愈：把 file_path 的 sheet_name 行2「N月销量得分」表头 rename 成
    当月 score_col_name 并落盘（模板=上月成品时表头停上期，过去要手工改一格）。
    返回状态：'skip' 未传列名｜'ok' 已是当月，无需改｜'renamed' 已改名并落盘｜
    'locked' 改了但文件被其他程序占用、没能落盘（成品侧还有兜底）｜'missing' 行2 没有得分表头。"""
    if not score_col_name:
        return "skip"
    _wb = load_workbook(file_path, data_only=False)
    try:
        if sheet_name not in _wb.sheetnames:
            print(f"⚠️ 【得分列表头】{os.path.basename(file_path)} 没有「{sheet_name}」sheet，跳过自愈")
            return "missing"
        _ws = _wb[sheet_name]
        _ren = rename_score_col_in_ws(_ws, score_col_name)
        if _ren is None:
            _col = _find_score_header_col(_ws)
            if _col is None:
                print(f"⚠️ 【得分列表头】{os.path.basename(file_path)}「{sheet_name}」行2 找不到"
                      f"「N月销量得分」表头 → 无法自愈，请检查模板列布局")
                return "missing"
            print(f"【得分列表头】{os.path.basename(file_path)}「{sheet_name}」"
                  f"{get_column_letter(_col)}2 已是「{score_col_name}」✔")
            return "ok"
        if save:
            try:
                _wb.save(file_path)
            except PermissionError:
                print(f"⚠️ 【得分列表头】{os.path.basename(file_path)} 被其他程序占用，表头只在内存改名"
                      f"（{get_column_letter(_ren[2])}2：「{_ren[0]}」→「{_ren[1]}」），未能落盘")
                return "locked"
        print(f"【得分列表头】{os.path.basename(file_path)}「{sheet_name}」{get_column_letter(_ren[2])}2："
              f"「{_ren[0]}」→「{_ren[1]}」（上月成品带过来的旧表头，已自动 rename）")
        return "renamed"
    finally:
        _wb.close()

def get_month_sale_df(belonging_date:str,snapshot_date:str,score_col_name:str=None):
    dir_path = os.path.join("data","数据归属周期",belonging_date,"快照日期",snapshot_date)
    file_name = f"月销售总表.xlsx"
    file_path = os.sep.join([dir_path, file_name])
    # ===== 2026-10-17 新增：KH 表头自愈 =====
    # 模板=复制上月成品 → 【{year}年】KH2 表头停上期；过去每月首次跑要手工改一格、
    # 否则补充信息处 assert 硬崩。这里按当月 score_col_name 自动 rename 落盘，一劳永逸。
    heal_score_col_header_file(file_path, f"{belonging_date[0:4]}年", score_col_name)
    month_sale_df = pd.read_excel(file_path, sheet_name=f"{belonging_date[0:4]}年", skiprows=1)
    # 内存保险：模板被占用没能落盘改名时，把 df 里的旧表头名改成新名（列位不变，
    # 后续 assert / KH 刷新逻辑不受影响；成品表头由结果侧兜底再补）。
    if score_col_name and score_col_name not in month_sale_df.columns:
        for _c in list(month_sale_df.columns):
            if _SCORE_HEADER_RE.match(str(_c).strip()):
                month_sale_df = month_sale_df.rename(columns={_c: score_col_name})
                print(f"【得分列表头】⚠️ 模板未能落盘改名，df 列名「{_c}」→「{score_col_name}」"
                      f"（成品表头由结果侧自愈兜底）")
                break
    return month_sale_df

def fill_missing(df:pd.DataFrame):
    """
    填充缺失的 SKU 或 ASIN，生成唯一流水号
    
    规则：
    - SKU 有值但 ASIN 为空 → 生成 ASIN 流水号
    - ASIN 有值但 SKU 为空 → 生成 SKU 流水号
    - 两者都有值 → 保持不变
    - 两者都为空 → 保持不变
    """
    # 复制一份，避免修改原数据
    df = df.copy()
    
    # 获取当前最大的 SKU 和 ASIN 流水号
    def get_max(prefix):
        all_vals = pd.concat([df['SKU'], df['ASIN']]).dropna()
        matched = all_vals[all_vals.astype(str).str.startswith(prefix, na=False)]
        if len(matched) == 0:
            return 0
        nums = matched.astype(str).str.replace(prefix, '', regex=False)
        nums = nums[nums.str.isdigit()]
        return nums.astype(int).max() if len(nums) > 0 else 0
    
    max_sku = get_max('SKU')
    max_asin = get_max('ASIN')
    
    # 情况1：SKU 有值，ASIN 为空 → 生成 ASIN
    mask1 = df['SKU'].notna() & df['ASIN'].isna()
    count1 = mask1.sum()
    if count1 > 0:
        df.loc[mask1, 'ASIN'] = [f'ASIN{i:04d}' for i in range(max_asin + 1, max_asin + count1 + 1)]
        max_asin += count1
    
    # 情况2：ASIN 有值，SKU 为空 → 生成 SKU
    mask2 = df['ASIN'].notna() & df['SKU'].isna()
    count2 = mask2.sum()
    if count2 > 0:
        df.loc[mask2, 'SKU'] = [f'SKU{i:04d}' for i in range(max_sku + 1, max_sku + count2 + 1)]
    print(df)
    return df

def sum_shopee_cross_site(cur_acct_cycle_date:datetime.date, processed_month_sale_df:pd.DataFrame, i:int, group:str, month_offset:str, total_sku_sale_df2:pd.DataFrame):
    # ★ 2026-11-02：加入「台湾」——台湾是 Shopee 跨境站点，月销总表没有独立列，
    #   历来并入「虾皮境内」（1~2 月历史值可验：527马+16台=543）。用户 2026-11-02 拍板沿用此口径。
    other_shopee_cross_site = ["马来西亚","越南","台湾"]
    group_select_condition = (total_sku_sale_df2["平台"] == "Shopee") & (total_sku_sale_df2["站点"].isin(other_shopee_cross_site)) & ((total_sku_sale_df2["实体"] == "跨境"))
    cur_group_month_sale_df = total_sku_sale_df2[group_select_condition].copy()
    cur_group_month_sale_df = cur_group_month_sale_df.groupby(by=["SKU"])[f"{cur_acct_cycle_date.strftime('%Y%m')}销量"].sum().reset_index()
    if cur_group_month_sale_df.shape[0] > 0:
        print(group)
        print(cur_group_month_sale_df)
        processed_month_sale_df = pd.merge(left=processed_month_sale_df, right=cur_group_month_sale_df[["SKU",f"{cur_acct_cycle_date.strftime('%Y%m')}销量"]], how="left", left_on=["SKU"], right_on=["SKU"]).copy()
        suffix = f".{i}"
        if i == 0:
            suffix = ""
        located_col_name = f"{month_offset}月{suffix}"
        processed_month_sale_df[located_col_name] = processed_month_sale_df[f"{cur_acct_cycle_date.strftime('%Y%m')}销量"]
        processed_month_sale_df[located_col_name] = processed_month_sale_df[located_col_name].fillna(0)  # ★2026-11-07 补 0：merge(how=left) 只给「本月有销量」的 SKU 赋值，归零 SKU 会残留模板旧值
        processed_month_sale_df.drop(f"{cur_acct_cycle_date.strftime('%Y%m')}销量", axis=1, inplace=True)
    return processed_month_sale_df

def sum_total_southeast_asia_site(cur_acct_cycle_date:datetime.date, processed_month_sale_df:pd.DataFrame, i:int, group:str, month_offset:str, total_sku_sale_df2:pd.DataFrame):
    # ★ 2026-11-02：`台湾` 必须一起进「东南亚」列——恒等链要求 月销东南亚 == SEA Total 月块和，
    #   而 SEA Total 的「虾皮跨境」列含台湾（东南亚销量表《虾皮跨境》sheet 本就有「中国台湾销量」列）。
    se_asia_sites = ["菲律宾","印尼","泰国","马来西亚","新加坡","越南","台湾"]
    group_select_condition = (total_sku_sale_df2["站点"].isin(se_asia_sites))
    cur_group_month_sale_df = total_sku_sale_df2[group_select_condition].copy()

    cur_group_month_sale_df = cur_group_month_sale_df.groupby(by=["SKU"])[f"{cur_acct_cycle_date.strftime('%Y%m')}销量"].sum().reset_index()
    cur_group_month_sale_df.to_csv(os.path.join(tmp_dir, "total_southeast_asia_summary.csv"))
    print(group)
    print(cur_group_month_sale_df)
    if cur_group_month_sale_df.shape[0] > 0:
        processed_month_sale_df = pd.merge(left=processed_month_sale_df, right=cur_group_month_sale_df[["SKU",f"{cur_acct_cycle_date.strftime('%Y%m')}销量"]], how="left", left_on=["SKU"], right_on=["SKU"]).copy()
        suffix = f".{i}"
        if i == 0:
            suffix = ""
        located_col_name = f"{month_offset}月{suffix}"
        processed_month_sale_df[located_col_name] = processed_month_sale_df[f"{cur_acct_cycle_date.strftime('%Y%m')}销量"]
        processed_month_sale_df[located_col_name] = processed_month_sale_df[located_col_name].fillna(0)  # ★2026-11-07 补 0：merge(how=left) 只给「本月有销量」的 SKU 赋值，归零 SKU 会残留模板旧值
        processed_month_sale_df.drop(f"{cur_acct_cycle_date.strftime('%Y%m')}销量", axis=1, inplace=True)
    return processed_month_sale_df

def sum_se_single_site(cur_acct_cycle_date:datetime.date, processed_month_sale_df:pd.DataFrame, i:int, group:str, month_offset:str, total_sku_sale_df2:pd.DataFrame, month_sale_group_lazada_shopee_site_mapping:dict[str,str]):
    if group == "新加坡":
        group_select_condition = total_sku_sale_df2["站点"] == month_sale_group_lazada_shopee_site_mapping[group]
    elif group == "马来":
        group_select_condition = (total_sku_sale_df2["站点"] == month_sale_group_lazada_shopee_site_mapping[group]) & (total_sku_sale_df2["平台"] == "Lazada")
    else:
        group_select_condition = (total_sku_sale_df2["站点"] == month_sale_group_lazada_shopee_site_mapping[group]) & (total_sku_sale_df2["平台"] == "Lazada")
    cur_group_month_sale_df = total_sku_sale_df2[group_select_condition].copy()
    cur_group_month_sale_df = cur_group_month_sale_df.groupby(by=["SKU"])[f"{cur_acct_cycle_date.strftime('%Y%m')}销量"].sum().reset_index()
    if cur_group_month_sale_df.shape[0] > 0:
        print(group)
        print(cur_group_month_sale_df)
        processed_month_sale_df = pd.merge(left=processed_month_sale_df, right=cur_group_month_sale_df[["SKU",f"{cur_acct_cycle_date.strftime('%Y%m')}销量"]], how="left", left_on=["SKU"], right_on=["SKU"]).copy()
        suffix = f".{i}"
        if i == 0:
            suffix = ""
        located_col_name = f"{month_offset}月{suffix}"
        processed_month_sale_df[located_col_name] = processed_month_sale_df[f"{cur_acct_cycle_date.strftime('%Y%m')}销量"]
        processed_month_sale_df[located_col_name] = processed_month_sale_df[located_col_name].fillna(0)  # ★2026-11-07 补 0：merge(how=left) 只给「本月有销量」的 SKU 赋值，归零 SKU 会残留模板旧值
        processed_month_sale_df.drop(f"{cur_acct_cycle_date.strftime('%Y%m')}销量", axis=1, inplace=True)
    return processed_month_sale_df

def sum_se_local_site(cur_acct_cycle_date:datetime.date, processed_month_sale_df:pd.DataFrame, i:int, group:str, month_offset:str, total_sku_sale_df2:pd.DataFrame):
    #三个本土销售市场，都要统计跨境数量（口径 A：本土 + 跨境）
    # ⚠️ 业务 2026-09-16 明确确认：本列口径**保持「本土 + 跨境」，不要改成「仅本土」**。
    #    原因：跨境订单的「站点」也被写成 "菲律宾"/"泰国"（见 Lazada 分支 order_df["站点"]=sale_group），
    #    所以"只按站点过滤"必然带上跨境。若改成仅本土，跨境 Lazada（菲 125 / 泰 68）会失去落点，
    #    而「虾皮境内」只覆盖 Shopee → 东南亚 7 个子列之和会对不上总计。
    #    实测（202608）：菲 5972 / 泰 3342 / 印 3868，三者之和与总计 71861 自洽。
    cur_group_month_sale_df = total_sku_sale_df2[total_sku_sale_df2["站点"] == group].copy()
    cur_group_month_sale_df = cur_group_month_sale_df.groupby(by=["SKU"])[f"{cur_acct_cycle_date.strftime('%Y%m')}销量"].sum().reset_index()
    if cur_group_month_sale_df.shape[0] > 0:
        print(group)
        print(cur_group_month_sale_df)
        processed_month_sale_df = pd.merge(left=processed_month_sale_df, right=cur_group_month_sale_df[["SKU",f"{cur_acct_cycle_date.strftime('%Y%m')}销量"]], how="left", left_on=["SKU"], right_on=["SKU"]).copy()
        suffix = f".{i}"
        if i == 0:
            suffix = ""
        located_col_name = f"{month_offset}月{suffix}"
        processed_month_sale_df[located_col_name] = processed_month_sale_df[f"{cur_acct_cycle_date.strftime('%Y%m')}销量"]
        processed_month_sale_df[located_col_name] = processed_month_sale_df[located_col_name].fillna(0)  # ★2026-11-07 补 0：merge(how=left) 只给「本月有销量」的 SKU 赋值，归零 SKU 会残留模板旧值
        processed_month_sale_df.drop(f"{cur_acct_cycle_date.strftime('%Y%m')}销量", axis=1, inplace=True)
    return processed_month_sale_df

def sum_jingdong_site(cur_acct_cycle_date:datetime.date, processed_month_sale_df:pd.DataFrame, i:int, group:str, month_offset:str, total_sku_sale_df2:pd.DataFrame):
    group_select_condition = (total_sku_sale_df2["平台"] == "京东") & (total_sku_sale_df2["站点"] == "国内")
    cur_group_month_sale_df = total_sku_sale_df2[group_select_condition].copy()
    cur_group_month_sale_df = cur_group_month_sale_df.groupby(by=["SKU"])[f"{cur_acct_cycle_date.strftime('%Y%m')}销量"].sum().reset_index()
    if cur_group_month_sale_df.shape[0] > 0:
        print(group)
        print(cur_group_month_sale_df)
        processed_month_sale_df = pd.merge(left=processed_month_sale_df, right=cur_group_month_sale_df[["SKU",f"{cur_acct_cycle_date.strftime('%Y%m')}销量"]], how="left", left_on=["SKU"], right_on=["SKU"]).copy()
        suffix = f".{i}"
        if i == 0:
            suffix = ""
        located_col_name = f"{month_offset}月{suffix}"
        processed_month_sale_df[located_col_name] = processed_month_sale_df[f"{cur_acct_cycle_date.strftime('%Y%m')}销量"]
        processed_month_sale_df[located_col_name] = processed_month_sale_df[located_col_name].fillna(0)  # ★2026-11-07 补 0：merge(how=left) 只给「本月有销量」的 SKU 赋值，归零 SKU 会残留模板旧值
        processed_month_sale_df.drop(f"{cur_acct_cycle_date.strftime('%Y%m')}销量", axis=1, inplace=True)
    return processed_month_sale_df

def sum_emag_site(cur_acct_cycle_date:datetime.date, processed_month_sale_df:pd.DataFrame, i:int, group:str, month_offset:str, total_sku_sale_df2:pd.DataFrame):
    group_select_condition = (total_sku_sale_df2["平台"] == "EMAG")
    cur_group_month_sale_df = total_sku_sale_df2[group_select_condition].copy()
    cur_group_month_sale_df = cur_group_month_sale_df.groupby(by=["SKU"])[f"{cur_acct_cycle_date.strftime('%Y%m')}销量"].sum().reset_index()
    if cur_group_month_sale_df.shape[0] > 0:
        print(group)
        print(cur_group_month_sale_df)
        processed_month_sale_df = pd.merge(left=processed_month_sale_df, right=cur_group_month_sale_df[["SKU",f"{cur_acct_cycle_date.strftime('%Y%m')}销量"]], how="left", left_on=["SKU"], right_on=["SKU"]).copy()
        suffix = f".{i}"
        if i == 0:
            suffix = ""
        located_col_name = f"{month_offset}月{suffix}"
        processed_month_sale_df[located_col_name] = processed_month_sale_df[f"{cur_acct_cycle_date.strftime('%Y%m')}销量"]
        processed_month_sale_df[located_col_name] = processed_month_sale_df[located_col_name].fillna(0)  # ★2026-11-07 补 0：merge(how=left) 只给「本月有销量」的 SKU 赋值，归零 SKU 会残留模板旧值
        processed_month_sale_df.drop(f"{cur_acct_cycle_date.strftime('%Y%m')}销量", axis=1, inplace=True)
    return processed_month_sale_df

def sum_temu_site(cur_acct_cycle_date:datetime.date, processed_month_sale_df:pd.DataFrame, i:int, group:str, month_offset:str, total_sku_sale_df2:pd.DataFrame):
    group_select_condition = (total_sku_sale_df2["平台"] == "Temu")
    cur_group_month_sale_df = total_sku_sale_df2[group_select_condition].copy()
    cur_group_month_sale_df = cur_group_month_sale_df.groupby(by=["SKU"])[f"{cur_acct_cycle_date.strftime('%Y%m')}销量"].sum().reset_index()
    if cur_group_month_sale_df.shape[0] > 0:
        print(group)
        print(cur_group_month_sale_df)
        processed_month_sale_df = pd.merge(left=processed_month_sale_df, right=cur_group_month_sale_df[["SKU",f"{cur_acct_cycle_date.strftime('%Y%m')}销量"]], how="left", left_on=["SKU"], right_on=["SKU"]).copy()
        suffix = f".{i}"
        if i == 0:
            suffix = ""
        located_col_name = f"{month_offset}月{suffix}"
        processed_month_sale_df[located_col_name] = processed_month_sale_df[f"{cur_acct_cycle_date.strftime('%Y%m')}销量"]
        processed_month_sale_df[located_col_name] = processed_month_sale_df[located_col_name].fillna(0)  # ★2026-11-07 补 0：merge(how=left) 只给「本月有销量」的 SKU 赋值，归零 SKU 会残留模板旧值
        processed_month_sale_df.drop(f"{cur_acct_cycle_date.strftime('%Y%m')}销量", axis=1, inplace=True)
    return processed_month_sale_df

def sum_mercado_site(cur_acct_cycle_date:datetime.date, processed_month_sale_df:pd.DataFrame, i:int, group:str, month_offset:str, total_sku_sale_df2:pd.DataFrame):
    group_select_condition = (total_sku_sale_df2["平台"] == "美客多")
    cur_group_month_sale_df = total_sku_sale_df2[group_select_condition].copy()
    cur_group_month_sale_df = cur_group_month_sale_df.groupby(by=["SKU"])[f"{cur_acct_cycle_date.strftime('%Y%m')}销量"].sum().reset_index()
    if cur_group_month_sale_df.shape[0] > 0:
        print(group)
        print(cur_group_month_sale_df)
        processed_month_sale_df = pd.merge(left=processed_month_sale_df, right=cur_group_month_sale_df[["SKU",f"{cur_acct_cycle_date.strftime('%Y%m')}销量"]], how="left", left_on=["SKU"], right_on=["SKU"]).copy()
        suffix = f".{i}"
        if i == 0:
            suffix = ""
        located_col_name = f"{month_offset}月{suffix}"
        processed_month_sale_df[located_col_name] = processed_month_sale_df[f"{cur_acct_cycle_date.strftime('%Y%m')}销量"]
        processed_month_sale_df[located_col_name] = processed_month_sale_df[located_col_name].fillna(0)  # ★2026-11-07 补 0：merge(how=left) 只给「本月有销量」的 SKU 赋值，归零 SKU 会残留模板旧值
        processed_month_sale_df.drop(f"{cur_acct_cycle_date.strftime('%Y%m')}销量", axis=1, inplace=True)
    return processed_month_sale_df

def sum_total_site(cur_acct_cycle_date:datetime.date, processed_month_sale_df:pd.DataFrame, i:int, group:str, month_offset:str, total_sku_sale_df2:pd.DataFrame):
    cur_group_month_sale_df = total_sku_sale_df2.copy()
    cur_group_month_sale_df = cur_group_month_sale_df.groupby(by=["SKU"])[f"{cur_acct_cycle_date.strftime('%Y%m')}销量"].sum().reset_index()
    if cur_group_month_sale_df.shape[0] > 0:
        print(group)
        print(cur_group_month_sale_df)
        processed_month_sale_df = pd.merge(left=processed_month_sale_df, right=cur_group_month_sale_df[["SKU",f"{cur_acct_cycle_date.strftime('%Y%m')}销量"]], how="left", left_on=["SKU"], right_on=["SKU"]).copy()
        suffix = f".{i}"
        if i == 0:
            suffix = ""
        located_col_name = f"{month_offset}月{suffix}"
        processed_month_sale_df[located_col_name] = processed_month_sale_df[f"{cur_acct_cycle_date.strftime('%Y%m')}销量"]
        processed_month_sale_df[located_col_name] = processed_month_sale_df[located_col_name].fillna(0)  # ★2026-11-07 补 0：merge(how=left) 只给「本月有销量」的 SKU 赋值，归零 SKU 会残留模板旧值
        processed_month_sale_df.drop(f"{cur_acct_cycle_date.strftime('%Y%m')}销量", axis=1, inplace=True)
    return processed_month_sale_df


# ==================== 任务⑩：列分组（分级显示）自动推进到当月 ====================
# 背景：模板的「列分组」永远停在上一个月的展开态，每月都要人工「取消再重新排列」一遍：
#   · 【{年}年】sheet：24 个销售分组块（欧洲/英国/美国/…/总计），每块 = 12 个月列。
#     模板（7 月做的）= 每块 1~7 月平铺、8~12 月整段分组并折叠隐藏。
#   · 【{年}年-欧洲各国每月销售情况】sheet：12 个月块，每块 = 10 国明细列 + 1 个月合计列。
#     模板 = 1~6 月只留月合计列（明细折叠）、7 月全展开、8~12 月整段隐藏。
# 本步把两张表的展开窗口统一「往后推到当月」，并把打开视图横滚到当月；只改 <cols> 与
# <sheetView>，不碰任何单元格值。（列宽/样式/公式全部原样保留；重复执行结果一致。）
MONTH_SHEET_SUFFIX = "年"
EU_SHEET_SUFFIX = "年-欧洲各国每月销售情况"
COL_ATTR_ORDER = ["style", "width", "customWidth", "hidden", "outlineLevel", "collapsed"]


def _parse_cols(xml):
    """把 <cols> 摊平成 {列号: {属性: 值}}。
    分组三属性（hidden/outlineLevel/collapsed）作为"原值"一并带上：这样只改我们在 state 里
    显式指定的列，其它列（如欧洲表尾部的「比利时2026年总销量」块）保持模板原样。"""
    m = re.search(r"<cols>.*?</cols>", xml, re.S)
    if not m:
        return {}
    attrs = {}
    for tag in re.findall(r"<col\b[^>]*?/>", m.group(0)):
        a = dict(re.findall(r'([\w:]+)="([^"]*)"', tag))
        lo, hi = int(a.pop("min")), int(a.pop("max"))
        for c in range(lo, hi + 1):
            attrs[c] = dict(a)
    return attrs


def _build_cols(attrs, state):
    """按 state 覆盖后重新拼 <cols>；不在 state 里的列沿用模板原值，相邻同属性列自动合并"""
    def attr_of(c):
        a = dict(attrs.get(c, {}))
        if c in state:
            st = state[c]
            for k in ("hidden", "collapsed"):
                if st.get(k):
                    a[k] = "1"
                else:
                    a.pop(k, None)
            if st.get("outlineLevel"):
                a["outlineLevel"] = str(st["outlineLevel"])
            else:
                a.pop("outlineLevel", None)
        return a

    def tag(c0, c1, a):
        parts = [f'min="{c0}"', f'max="{c1}"'] + [f'{k}="{a[k]}"' for k in COL_ATTR_ORDER if k in a]
        parts += [f'{k}="{a[k]}"' for k in sorted(a) if k not in COL_ATTR_ORDER]
        return "<col " + " ".join(parts) + "/>"

    out, run0, run_attr, prev = [], None, None, None
    for c in sorted(attrs):
        a = attr_of(c)
        if run0 is None:
            run0, run_attr = c, a
        elif a != run_attr or c != prev + 1:
            out.append(tag(run0, prev, run_attr))
            run0, run_attr = c, a
        prev = c
    if run0 is not None:
        out.append(tag(run0, prev, run_attr))
    return "<cols>" + "".join(out) + "</cols>"


def _patch_view(xml, left_col, active_col):
    """把冻结窗格的横滚起点移到 left_col，活动单元格落在 active_col 第 2 行（替换旧选区，不追加）"""
    pm = re.search(r"<pane\b[^>]*/>", xml)
    if not pm:
        return xml, None
    pane = pm.group(0)
    rm = re.search(r'topLeftCell="[A-Z]+(\d+)"', pane)
    row = rm.group(1) if rm else "1"
    new_pane = re.sub(r'topLeftCell="[^"]*"', f'topLeftCell="{get_column_letter(left_col)}{row}"', pane)
    sel = (f'<selection pane="bottomRight" activeCell="{get_column_letter(active_col)}2" '
           f'sqref="{get_column_letter(active_col)}2"/>')
    if re.search(r'<selection pane="bottomRight"[^>]*/>', xml):
        xml = re.sub(r'<selection pane="bottomRight"[^>]*/>', sel, xml, count=1)
    else:
        xml = xml.replace(pane, pane + sel, 1)
    return xml.replace(pane, new_pane, 1), row


def _month_blocks(ws, maxc):
    """【{年}年】sheet：第 2 行是「1月…12月」逐列标签，且各月块首尾相连 → 按 12 列切块"""
    blocks, run = [], []
    for c in range(1, maxc + 1):
        if not re.fullmatch(r"\d{1,2}月", str(ws.cell(2, c).value or "").strip()):
            continue
        if run and c != run[-1] + 1:
            blocks.append(run)
            run = []
        run.append(c)
        if len(run) == 12:
            blocks.append(run)
            run = []
    if run:
        blocks.append(run)
    return blocks


def _eu_month_blocks(ws, maxc):
    """【{年}年-欧洲各国…】sheet：第 2 行只有「月合计列」带月份标签（形如 ='2026年'!$B$1&"8月"）
    → 用它的列号反推月块（块 = [合计列-width+1 .. 合计列]）"""
    totals = []
    for c in range(1, maxc + 1):
        s = str(ws.cell(2, c).value or "")
        m = re.search(r'"(\d{1,2})月"', s) or re.fullmatch(r"(\d{1,2})月", s.strip())
        if m:
            totals.append((c, int(m.group(1))))
    if len(totals) < 2:
        return totals, 11
    return totals, totals[1][0] - totals[0][0]


def _sheet_xml_map(blobs):
    """sheet 名 → xl/worksheets/sheetN.xml（openpyxl 写的 rels 属性顺序不固定，必须与顺序无关地解析）"""
    relmap = {}
    for tag in re.findall(r"<Relationship\b[^>]*/>", blobs["xl/_rels/workbook.xml.rels"].decode("utf-8")):
        a = dict(re.findall(r'([\w:]+)="([^"]*)"', tag))
        if "Id" in a and "Target" in a:
            relmap[a["Id"]] = a["Target"]
    xml_of = {}
    for tag in re.findall(r"<sheet\b[^>]*/>", blobs["xl/workbook.xml"].decode("utf-8")):
        a = dict(re.findall(r'([\w:]+)="([^"]*)"', tag))
        m = re.search(r"worksheets/(sheet\d+\.xml)", relmap.get(a.get("r:id", ""), "")) if a.get("name") else None
        if m:
            xml_of[a["name"]] = "xl/worksheets/" + m.group(1)
    return xml_of


def arrange_month_outline(result_path, wb, cur_year, month_offset):
    """把成品里两张表的列分组从模板的「上月态」推进到当月（幂等，重复跑结果一致）"""
    import zipfile

    sheet_cur = f"{cur_year}{MONTH_SHEET_SUFFIX}"
    sheet_eu = f"{cur_year}{EU_SHEET_SUFFIX}"
    reports, warnings = [], []

    with zipfile.ZipFile(result_path) as z:
        blobs = {n: z.read(n) for n in z.namelist()}
    xml_of = _sheet_xml_map(blobs)

    for sheet_name in (sheet_cur, sheet_eu):
        if sheet_name not in xml_of or sheet_name not in wb.sheetnames:
            warnings.append(f"【{sheet_name}】不在成品里，跳过")
            continue
        key = xml_of[sheet_name]
        xml = blobs[key].decode("utf-8")
        ws = wb[sheet_name]
        attrs = _parse_cols(xml)
        if not attrs:
            warnings.append(f"【{sheet_name}】没有 <cols>，跳过")
            continue
        state, maxc = {}, ws.max_column

        if sheet_name == sheet_cur:
            # 每块：≤当月平铺；>当月 整段分组折叠。折叠按钮落在块后一列（= 下一块的 1 月列 / 尾部属性列）
            blocks = [b for b in _month_blocks(ws, maxc)
                      if len(b) == 12 and [str(ws.cell(2, c).value).strip() for c in b]
                      == [f"{m}月" for m in range(1, 13)]]
            for i, b in enumerate(blocks):
                for c, m in zip(b, range(1, 13)):
                    state[c] = {"outlineLevel": 1, "hidden": True} if m > month_offset else {}
                if i and month_offset < 12:
                    state[b[0]]["collapsed"] = True
            if blocks and month_offset < 12 and (blocks[-1][-1] + 1) in attrs:
                state.setdefault(blocks[-1][-1] + 1, {})["collapsed"] = True
            left_col, active_col = blocks[-1][0], blocks[-1][month_offset - 1]  # 视图：总计块 1 月列 / 当月列
            reports.append(f"【{sheet_name}】{len(blocks)} 个分组块：1~{month_offset}月平铺、"
                           f"{month_offset + 1}~12月分组折叠" + ("" if month_offset < 12 else "（已到 12 月，无折叠）"))
        else:
            # 月块 = [10 国明细列][1 月合计列]：之前的月只留合计列；当月展开明细；之后的月整块隐藏
            totals, width = _eu_month_blocks(ws, maxc)
            if not totals:
                warnings.append(f"【{sheet_name}】没识别到月合计列，跳过")
                continue
            for c, m in totals:
                first = c - width + 1
                if m < month_offset:
                    for cc in range(first, c):
                        state[cc] = {"outlineLevel": 1, "hidden": True}
                    state[c] = {"collapsed": True}
                elif m == month_offset:
                    for cc in range(first, c):
                        state[cc] = {"outlineLevel": 1}
                    state[c] = {}
                else:
                    for cc in range(first, c):
                        state[cc] = {"outlineLevel": 1, "hidden": True}
                    state[c] = {"outlineLevel": 1, "hidden": True, "collapsed": True}
            future = [c for c, m in totals if m > month_offset]
            if future:
                # 9~12 月整块隐藏后，在其后第一个可见列上留一个分级显示按钮，方便再次展开
                state.setdefault(max(future) + 1, {})["collapsed"] = True
            cur_col = next((c for c, m in totals if m == month_offset), totals[0][0] if totals else 3)
            left_col, active_col = cur_col - width + 1, cur_col                  # 视图：当月块首列 / 当月合计列
            reports.append(f"【{sheet_name}】{len(totals)} 个月块：1~{month_offset - 1}月只留月合计列、"
                           f"{month_offset}月展开明细、{month_offset + 1}~12月隐藏")

        xml = xml.replace(re.search(r"<cols>.*?</cols>", xml, re.S).group(0), _build_cols(attrs, state), 1)
        xml, row = _patch_view(xml, left_col, active_col)
        reports.append(f"    打开视图：横滚到 {get_column_letter(left_col)}{row or '?'}"
                       f"（活动单元格 {get_column_letter(active_col)}2）")
        blobs[key] = xml.encode("utf-8")

    tmp = result_path + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zw:
        for n, data in blobs.items():
            zw.writestr(n, data)
    _replace_file(tmp, result_path)
    return reports, warnings


def _replace_file(tmp, dst):
    """优先原子替换；目标被 WPS / 预览器打开时 Windows 拒绝"删除+改名" → 退化为原地覆写"""
    try:
        os.replace(tmp, dst)
        return "replace"
    except PermissionError:
        with open(tmp, "rb") as f:
            data = f.read()
        with open(dst, "wb") as f:
            f.write(data)
        os.remove(tmp)
        print("    ⚠️ 成品被其他程序占用，已退化为原地覆写（建议关掉 WPS/预览后重跑）")
        return "overwrite"


if __name__ == "__main__":
    # ===== 运行目录自愈（2026-11-02 新增，随 SVN 入库的加固）=====
    # 本文件内部路径全是相对写法（"data/数据归属周期/..."、"sale_file_tmp"），
    # 必须以 trunk/ 为工作目录。以前靠人工 cd，换个入口（仓库根、IDE 默认目录、
    # SVN 检出后直接双击/计划任务）就会因为找不到 data/ 而 FileNotFoundError。
    # 这里统一钉到脚本自身所在目录；CWD 已经正确时是空操作，不改变任何既有行为。
    _TRUNK_DIR = os.path.dirname(os.path.abspath(__file__))
    if os.path.abspath(os.getcwd()) != os.path.abspath(_TRUNK_DIR):
        print(f"[cwd] 工作目录切到 trunk/：{os.getcwd()} -> {_TRUNK_DIR}")
        os.chdir(_TRUNK_DIR)
    need_type_check = True
    # 2026-11-03：支持环境变量覆盖期别/快照（补算历史期用，如 MSF_PERIOD=202609 MSF_SNAPSHOT=20261006）；
    # 不设置时走下面的默认值。★ 2026-11-10：默认快照已跟到本期 20261107（原来停在 20261102，
    # 忘了传 env 就会静默用旧快照，和欧洲/北美脚本的 L35 常量对不上）。
    belonging_date = os.environ.get("MSF_PERIOD", "202610")
    snapshot_date = os.environ.get("MSF_SNAPSHOT", "20261107")
    # 得分列名。None = 自动探测（2026-10-16 新增：数据部只有一人后，不再"等别人改这行"）：
    # 规律 = 归属月的上一个月（202609 用「8月销量得分」），探测块见下方商品资料读取后。
    # 仍可写死列名（如 "8月销量得分"）覆盖自动探测。
    score_col_name = None

    # ===== 菲律宾新订单报告（以一敌多）接入开关 —— 2026-10-09 新增 =====
    # 新工具导出的《菲律宾{YYYYMM}订单报告.xlsx》单份覆盖 Lazada+Shopee+Tiktok+手工订单。
    # 202609 实测仅覆盖 9/13 起（9/1~9/12 整段缺失，全月净缺 967 件）⇒ 202609 那期仍用旧三份报告。
    # ★ 2026-11-02 用户拍板：**202610 起按新报告走**（开关置 True）。
    #   202610 实测新报告 创建时间 10/01 00:05 ~ 10/31 23:57 全覆盖、无跨月行
    #   ⇒ PH_NEW_REPORT_START_DATE 设为 None（不过滤）；若以后某月不满量，再填当月起始日。
    #   口径核对：已交运+已发货 = 3,408 件（其中手工订单 465 件 / 2 单，走 KSON 线下渠道），
    #   加跨境 Lazada 166 件 ⇒ 菲 = 3,574 件（旧三渠道口径 3,092 件，差 +482）。
    #   证据与判据见《核查-菲律宾新订单报告能否一敌多.md》。
    USE_PH_NEW_ORDER_REPORT = True
    if os.environ.get("PH_NEW_REPORT", "1") == "0":   # 一次性回退：设 PH_NEW_REPORT=0 即走旧三报告
        USE_PH_NEW_ORDER_REPORT = False
    PH_NEW_REPORT_START_DATE = None  # None = 不过滤（该月新报告满量）；不满量的月份填 "YYYY-MM-DD"

    cur_acct_cycle_date = pd.to_datetime(belonging_date, format="%Y%m")
    last_6m_acct_cycle_date = cur_acct_cycle_date - pd.DateOffset(months=5)#获取包括本月在内的最近六个月日期
    print(cur_acct_cycle_date)
    print(last_6m_acct_cycle_date)
    # manual_purchase_plan_df = get_manual_purchase_plan("20260826")
    # 商品资料
    product_info_df = get_product_info_df()
    # ===== 得分列自动探测（2026-10-16 新增）=====
    # 数据部只有一人：商品资料.xlsx 更新得分列后，不用再回来改 score_col_name。
    # 探测顺序：①「归属月上一个月」得分列（业务规律，202609 → 8月销量得分）；
    #          ② 退化取月份最大的「N月销量得分」列并打警告；
    #          ③ 一个都没有 → 人话报错（列出现有列名），不再裸 assert 一句"找不到列"。
    if score_col_name is None:
        _prev_m = (cur_acct_cycle_date - pd.DateOffset(months=1)).month
        _want = f"{_prev_m}月销量得分"
        _avail = [c for c in product_info_df.columns if re.match(r"^\d+月销量得分$", str(c))]
        if _want in product_info_df.columns:
            score_col_name = _want
            print(f"【得分列】自动探测 → {score_col_name}（归属月上一个月）")
        elif _avail:
            score_col_name = max(_avail, key=lambda c: int(re.match(r"^(\d+)月", str(c)).group(1)))
            print(f"【得分列】⚠️ 商品资料没有「{_want}」，退化取最新列 {score_col_name}"
                  f"（候选：{_avail}）——不符合预期请在商品资料补列")
        else:
            raise AssertionError(
                f"商品资料.xlsx 里找不到任何「N月销量得分」列（现有列：{list(product_info_df.columns)}）"
                f"——请先在商品资料维护得分列再跑")
    sku_asin_full_map = fill_missing(product_info_df[["SKU","ASIN"]])
    sku_asin_full_map.to_csv(os.path.join(tmp_dir, "sku_asin_full_map.csv"))
    sku_asin_map = sku_asin_full_map.set_index("SKU")["ASIN"]
    asin_sku_map = sku_asin_full_map.set_index("ASIN")["SKU"]
    # 月销售总表（2026-10-17：传入 score_col_name，模板 KH 表头若停上期会自动 rename 自愈）
    month_sale_df = get_month_sale_df(belonging_date, snapshot_date, score_col_name)
    #管理库存报告原始数据
    manage_inv_df = get_manage_inv_df("欧洲", belonging_date, snapshot_date)
    manage_inv_df = manage_inv_df.groupby(by=["仓库","sku","asin"])[["afn-fulfillable-quantity","afn-fc-transfer-quantity","afn-inbound-working-quantity","afn-inbound-shipped-quantity","afn-inbound-receiving-quantity"]].sum().reset_index()
    current_sale_df = manage_inv_df[["仓库","sku","asin"]].copy()#备货计划df
    current_sale_df[["sku_prefix","sku_sequence"]] = current_sale_df['sku'].str.extract(r'^([A-Za-z]+)(\d+)')
    current_sale_df["sku_sequence"] = current_sale_df["sku_sequence"].astype(int)
    current_sale_df = current_sale_df.sort_values(by=["仓库","sku_prefix", "sku_sequence"], ascending=[True, True,True])
    current_sale_df = current_sale_df[["仓库","sku","asin"]].copy().reset_index()
    #订单原始数据
    juchuan_sale_groups = ["欧洲","日本","澳大利亚","北美","沙特阿拉伯","阿拉伯联合酋长国","印度"]
    interflow_sale_groups = ["北美","澳大利亚","阿拉伯联合酋长国"]
    total_last30d_amazon_order_df = pd.DataFrame({})
    for entity in entites:
        target_sale_groups = []
        if entity == "巨川":
            target_sale_groups =juchuan_sale_groups
        elif entity == "interflow":
            target_sale_groups = interflow_sale_groups
        else:
            raise Exception("不支持的实体")
        for sale_group in target_sale_groups:
            last30d_amazon_order_df = get_belonging_amazon_order_df(entity, sale_group,belonging_date, snapshot_date)
            order_date_col = "purchase-date-local"
            max_sale_date = last30d_amazon_order_df[order_date_col].max()
            min_sale_date = last30d_amazon_order_df[order_date_col].min()
            if not need_type_check:
                max_sale_date = datetime.datetime.now()
                min_sale_date = datetime.datetime.now()
            sale_date_delta = max_sale_date - min_sale_date
            # 覆盖天数 = 差值 + 1；按月天数比对（9/4/6/11/2 月差值天然 <30，写死 30 会误报）
            _cover_days = sale_date_delta.days + 1
            _days_in_month = pd.Timestamp(f"{belonging_date[:4]}-{belonging_date[4:6]}-01").days_in_month
            if _cover_days < _days_in_month:
                print(f"【警告】{entity}{sale_group}订单日期跨度小于{_days_in_month}天（实际覆盖 {_cover_days} 天），请检查是否导出数据完整")
            if _cover_days > _days_in_month + 1:
                print(f"【警告】{entity}{sale_group}订单日期跨度大于{_days_in_month + 1}天（实际覆盖 {_cover_days} 天），请检查原始数据是否有误")
            check_order_days_completion = last30d_amazon_order_df[["amazon-order-id",order_date_col,"sku","asin","quantity"]].copy()
            check_order_days_completion[order_date_col] = check_order_days_completion[order_date_col].dt.normalize()
            check_order_days_completion = check_order_days_completion.groupby(by=[order_date_col])["quantity"].sum().reset_index()
            if check_order_days_completion.shape[0] < _days_in_month:
                print(f"【信息】{entity}{sale_group}有些日期没有订单（{check_order_days_completion.shape[0]}/{_days_in_month} 天有单），请注意")
            if check_order_days_completion.shape[0] > 31:
                print(f"【警告】{entity}{sale_group}订单日期含有不止31天数据，请检查。【{min_sale_date}】-【{min_sale_date}】跨度【{sale_date_delta}】")
            total_last30d_amazon_order_df = pd.concat([total_last30d_amazon_order_df,last30d_amazon_order_df])
    
    total_last30d_amazon_order_df.to_csv(os.path.join(tmp_dir,"total_last30d_amazon_order_df.csv"))#把所有亚马逊订单综合起来
    #汇总亚马逊销量表
    amazon_sale_df = total_last30d_amazon_order_df.groupby(by=["实体","平台","仓库", "站点","sku","asin"])["quantity"].sum().reset_index()
    amazon_sale_df = amazon_sale_df.rename(columns={"quantity":f"{cur_acct_cycle_date.strftime('%Y%m')}销量"})
    # 9.12新增：ASIN反查会覆盖源表sku列，映射失败就没退路（欧洲站独立ASIN未登记进商品资料时会整批丢单）
    # 原逻辑只有一行 amazon_sale_df["sku"] = amazon_sale_df["asin"].map(asin_sku_map)，现改为"ASIN优先 + 源表sku兜底"
    raw_sku = amazon_sale_df["sku"].astype(str).str.strip()
    amazon_sale_df["sku"] = amazon_sale_df["asin"].map(asin_sku_map)
    fallback_sku = raw_sku.str.replace(r"[A-Za-z]{1,3}$", "", regex=True)
    amazon_sale_df["sku"] = amazon_sale_df["sku"].fillna(fallback_sku[fallback_sku.isin(sku_asin_map.index)])
    lost_order = amazon_sale_df.loc[amazon_sale_df["sku"].isna(), ["仓库","站点","asin"]].drop_duplicates()
    if lost_order.shape[0] > 0:
        print(f"【警告】{lost_order.shape[0]}组订单ASIN无法对应商品资料，将被丢弃：\n{lost_order.to_string()}")

    total_last30d_se_a_order_df = pd.DataFrame({})
    entity = "本土"
    # 2026-10-09：菲律宾三渠道可整体由「新订单报告」替代（开关 USE_PH_NEW_ORDER_REPORT）。
    # 开关关闭（默认）时行为与历史完全一致：仍读旧 Shopee/Lazada/Tiktok 三份报告。
    southeast_asia_local_shopee_salegroup = ["印尼","菲律宾","泰国"]
    if USE_PH_NEW_ORDER_REPORT:
        southeast_asia_local_shopee_salegroup = ["印尼","泰国"]  # 菲律宾改由新订单报告统一提供
    for sale_group in southeast_asia_local_shopee_salegroup:
        shopee_se_a_order_df = get_belonging_shopee_order_df(entity, sale_group,belonging_date, snapshot_date)
        total_last30d_se_a_order_df = pd.concat([total_last30d_se_a_order_df, shopee_se_a_order_df])
    
    southeast_asia_local_lazada_salegroup = ["菲律宾","泰国"]
    if USE_PH_NEW_ORDER_REPORT:
        southeast_asia_local_lazada_salegroup = ["泰国"]  # 菲律宾改由新订单报告统一提供
    for sale_group in southeast_asia_local_lazada_salegroup:
        lazada_se_a_order_df = get_belonging_lazada_order_df(entity, sale_group,belonging_date, snapshot_date)
        total_last30d_se_a_order_df = pd.concat([total_last30d_se_a_order_df, lazada_se_a_order_df])

    southeast_asia_local_tkshop_salegroup = ["菲律宾","印尼"]
    if USE_PH_NEW_ORDER_REPORT:
        southeast_asia_local_tkshop_salegroup = ["印尼"]  # 菲律宾改由新订单报告统一提供
    for sale_group in southeast_asia_local_tkshop_salegroup:
        tkshop_se_a_order_df = get_belonging_tkshop_order_df(entity, sale_group,belonging_date, snapshot_date)
        total_last30d_se_a_order_df = pd.concat([total_last30d_se_a_order_df, tkshop_se_a_order_df])

    if USE_PH_NEW_ORDER_REPORT:
        # 菲律宾：一份新订单报告替代上面三份（含 Shopee 手工订单）
        ph_new_order_df = get_belonging_ph_new_order_df(entity, "菲律宾", belonging_date, snapshot_date,
                                                        start_date=PH_NEW_REPORT_START_DATE)
        total_last30d_se_a_order_df = pd.concat([total_last30d_se_a_order_df, ph_new_order_df])

    southeast_asia_offline_salegroup = ["菲律宾"]
    if USE_PH_NEW_ORDER_REPORT:
        # ★ 2026-11-02：开关打开后**不能再读 KSON 线下报告**——新报告的「手工订单」已按同一分组键
        #   ("本土","KSON","菲律宾") 计入线下渠道，两边同源会双计（202610 该文件本身已是 0 行）。
        #   保留一次只读探针：文件仍有数据就大声警告，避免"换月时它又开始出数"却无人察觉。
        _kson_path = se_a_dir(os.path.join("data","数据归属周期",belonging_date,"快照日期",snapshot_date),
                              f"菲律宾KSON仓{belonging_date}线下订单报告.xlsx")
        try:
            _kson_rows = int(pd.read_excel(_kson_path).shape[0])
        except Exception as _e:
            _kson_rows = -1
            print(f"    ℹ️ 【菲律宾】KSON 线下报告不可读（{_e}）→ 跳过读取（开关已开，不影响口径）")
        if _kson_rows > 0:
            print(f"    ⚠️ 【菲律宾】KSON 线下报告仍有 {_kson_rows} 行，但开关已开 → **已跳过**"
                  f"（手工订单以新报告为准）。若这批单不在新报告内，请人工核对后决定是否回退开关")
        elif _kson_rows == 0:
            print("    ℹ️ 【菲律宾】KSON 线下报告为 0 行 → 手工订单已全部并入新报告，跳过读取")
        southeast_asia_offline_salegroup = []
    for sale_group in southeast_asia_offline_salegroup:
        kson_offline_order_df = get_belonging_kson_offline_order_df(entity, sale_group,belonging_date, snapshot_date)
        total_last30d_se_a_order_df = pd.concat([total_last30d_se_a_order_df, kson_offline_order_df])

    indonesia_warehouse_salegroup = ["印尼"]
    for sale_group in indonesia_warehouse_salegroup:
        belonging_indonesia_warehouse_offline_order_df = get_belonging_indonesia_warehouse_offline_order_df(entity, sale_group,belonging_date, snapshot_date)
        total_last30d_se_a_order_df = pd.concat([total_last30d_se_a_order_df, belonging_indonesia_warehouse_offline_order_df])

    entity = "跨境"
    # ★ 2026-11-02：「台湾」并入跨境 Shopee（→ 月销并进「虾皮境内」列，见 sum_shopee_cross_site）。
    southeast_asia_cross_shopee_salegroup = ["马来西亚","新加坡","越南","台湾"]
    for sale_group in southeast_asia_cross_shopee_salegroup:
        # ★ 2026-11-03（任务㉑配套）：「台湾」跨境报告 202610 起才存在（11-02 新增任务）。
        #   补算更早历史期（202609 等）时快照里没有这份报告 → 跳过并告警（与当时首跑口径一致，台湾=0）；
        #   202610 起缺报告仍然报错中断，不静默。
        if sale_group == "台湾" and belonging_date < "202610":
            try:
                shopee_se_a_order_df = get_belonging_shopee_order_df(entity, sale_group, belonging_date, snapshot_date)
            except FileNotFoundError as _e:
                print(f"【警告】{belonging_date} 快照缺 Shopee 跨境台湾报告（该报告 202610 起才有），本期按 0 处理")
                continue
        else:
            shopee_se_a_order_df = get_belonging_shopee_order_df(entity, sale_group,belonging_date, snapshot_date)
        total_last30d_se_a_order_df = pd.concat([total_last30d_se_a_order_df, shopee_se_a_order_df])
    
    southeast_asia_cross_lazada_salegroup = ["菲律宾","泰国","马来西亚","新加坡","越南"]
    for sale_group in southeast_asia_cross_lazada_salegroup:
        lazada_se_a_order_df = get_belonging_lazada_order_df(entity, sale_group,belonging_date, snapshot_date)
        total_last30d_se_a_order_df = pd.concat([total_last30d_se_a_order_df, lazada_se_a_order_df])
    # 2026-11-10：非亚马逊渠道报告区间体检（只看形状，不改数字；命中会在末尾汇总再打一次）
    audit_channel_spans(total_last30d_se_a_order_df, "东南亚各渠道", belonging_date)
    #汇总东南亚销量表
    se_a_sku_sale_df = total_last30d_se_a_order_df.groupby(by=["实体", "平台", "仓库", "站点", "sku"], dropna=False)["数量"].sum().reset_index()
    # groupby聚合后的销量列名随pandas版本而异（旧版为"数量dropna"，新版为"数量"），两种都兼容
    se_a_sku_sale_df = se_a_sku_sale_df.rename(columns={"数量":f"{cur_acct_cycle_date.strftime('%Y%m')}销量","数量dropna":f"{cur_acct_cycle_date.strftime('%Y%m')}销量"})
    assert f"{cur_acct_cycle_date.strftime('%Y%m')}销量" in se_a_sku_sale_df.columns, "东南亚销量表列名重命名失败（groupby聚合列名与预期不符）"
    se_a_sku_sale_df["asin"] = se_a_sku_sale_df["sku"].map(sku_asin_map)
    se_a_sku_sale_df.to_csv(os.path.join(tmp_dir,"se_a_sku_sale_df.csv"))
    # ===== 2026-11-07 新增：东南亚「逐笔订单明细」中间产物（供国家表订单明细 sheet 使用）=====
    # 业务需求：具体国家的表要能看到「订单号 / SKU / 订单状态」。
    # 这里把聚合前的逐笔明细落盘（只保留展示与核对需要的列），东南亚脚本据此生成明细 sheet。
    # ⚠️ 只做落盘、不改变任何既有口径或产物 —— 销量数字仍以 se_a_sku_sale_df.csv 为准。
    _detail_cols = ["订单号", "实体", "平台", "仓库", "站点", "sku", "数量", "下单时间", "订单状态", "订单类型"]
    _detail_df = total_last30d_se_a_order_df.copy()
    for _c in _detail_cols:
        if _c not in _detail_df.columns:
            _detail_df[_c] = None
    _detail_df = _detail_df[_detail_cols].copy()
    _detail_df["数量"] = pd.to_numeric(_detail_df["数量"], errors="coerce")
    _detail_df = _detail_df[_detail_df["数量"].notna()]
    _detail_df.to_csv(os.path.join(tmp_dir, "se_a_order_detail_df.csv"), index=False)
    print(f"【订单明细中间产物】{len(_detail_df)} 行 → se_a_order_detail_df.csv")
    #其他渠道
    total_last30d_other_order_df = pd.DataFrame({})
    entity = "巨川"

    jingdong_order_df = get_belonging_jingdong_order_df(entity, "国内", belonging_date, snapshot_date)
    total_last30d_other_order_df = pd.concat([total_last30d_other_order_df, jingdong_order_df])

    emag_order_df = get_belonging_emag_order_df(entity, "eMag", belonging_date, snapshot_date)
    total_last30d_other_order_df = pd.concat([total_last30d_other_order_df, emag_order_df])

    temu_order_df = get_belonging_temu_order_df(entity, "TEMU平台", belonging_date, snapshot_date)
    total_last30d_other_order_df = pd.concat([total_last30d_other_order_df, temu_order_df])

    mercado_order_df = get_belonging_mercado_order_df(entity, "美客多", belonging_date, snapshot_date)
    total_last30d_other_order_df = pd.concat([total_last30d_other_order_df, mercado_order_df])
    # 2026-11-10：非亚马逊渠道报告区间体检（京东 / eMAG / TEMU / 美客多）
    audit_channel_spans(total_last30d_other_order_df, "其他渠道", belonging_date)
    #汇总其他销量表
    other_sku_sale_df = total_last30d_other_order_df.groupby(by=["实体", "平台", "仓库", "站点", "sku"], dropna=False)["数量"].sum().reset_index()
    other_sku_sale_df = other_sku_sale_df.rename(columns={"数量":f"{cur_acct_cycle_date.strftime('%Y%m')}销量"})
    other_sku_sale_df.to_csv(os.path.join(tmp_dir,"other_sku_sale_df.csv"))

    #汇总所有销量表
    total_sku_sale_df = pd.concat([amazon_sale_df,se_a_sku_sale_df,other_sku_sale_df]).copy()#连接亚马逊和东南亚销量表
    # 第一层清洗：销量侧sku规范化（9.12新增，去除空格；"nan"转回NaN防止混入新SKU）
    total_sku_sale_df["sku"] = total_sku_sale_df["sku"].astype(str).str.strip().replace("nan", np.nan)
    total_sku_sale_df = total_sku_sale_df[total_sku_sale_df["sku"].notna() & (total_sku_sale_df["sku"] != "")].copy()  # 丢弃空白SKU的订单行（美客多存在此类脏数据）
    total_sku_sale_df = total_sku_sale_df.sort_values(by=["实体","平台","仓库", "站点", "sku","asin"], ascending=[True, True, True, True, True, True])
    total_sku_sale_df.to_csv(os.path.join(tmp_dir,"total_sku_sale_df.csv"))
    
    month_sale_group = [
        "欧洲","英国","美国","加拿大","日本","澳洲","阿联酋","沙特","墨西哥","印度","亚马逊",
        "菲律宾","泰国","印尼","新加坡","马来","越南","虾皮境内","东南亚",
        "京东","eMAG","TEMU平台","美客多",
        "总计"]
    
    init_pos = 1
    month_sale_df = month_sale_df.rename(columns={"Unnamed: 0":"SKU","Unnamed: 1":"型号"})
    # 第二层清洗：月份表SKU规范化 + 空行/合计行过滤（9.12修订）
    month_sale_df["SKU"] = month_sale_df["SKU"].astype(str).str.strip()
    month_sale_df = month_sale_df[~month_sale_df["SKU"].isin(["", "nan", "SKU"])].copy()
    month_sale_df = month_sale_df.reset_index(drop=True)
    month_sale_df.to_csv(os.path.join(tmp_dir,"month_sale_df.csv"))
    group_to_amazon_inv_map = {
        "欧洲":"欧洲","英国":"英国","美国":"美国","加拿大":"加拿大","日本":"日本","澳洲":"澳大利亚","阿联酋":"阿拉伯联合酋长国","沙特":"沙特阿拉伯","墨西哥":"墨西哥","印度":"印度",
    }

    missed_sku_df = pd.merge(left=total_sku_sale_df, right=month_sale_df, how="left", left_on=["sku"], right_on=["SKU"]).copy().reset_index()
    missed_sku_df = missed_sku_df[missed_sku_df["类型"].isna()]
    missed_sku_df.to_csv(os.path.join(tmp_dir,"missed_sku_df.csv"))

    #补充信息
    #类型,品牌,系列,7月销量得分,状态,
    new_sku_month_sale_df = pd.DataFrame({})
    new_sku_month_sale_df["SKU"] = missed_sku_df["sku"].dropna().drop_duplicates().astype(str).str.strip()

    product_info_map = product_info_df.set_index("SKU")["型号"]
    new_sku_month_sale_df["型号"] = new_sku_month_sale_df["SKU"].map(product_info_map)

    product_info_map = product_info_df.set_index("SKU")["类型"]
    new_sku_month_sale_df["类型"] = new_sku_month_sale_df["SKU"].map(product_info_map)

    product_info_map = product_info_df.set_index("SKU")["品牌"]
    new_sku_month_sale_df["品牌"] = new_sku_month_sale_df["SKU"].map(product_info_map)

    product_info_map = product_info_df.set_index("SKU")["系列"]
    new_sku_month_sale_df["系列"] = new_sku_month_sale_df["SKU"].map(product_info_map)

    assert score_col_name in product_info_df.columns, f"商品资料.xlsx 找不到列 {score_col_name}"
    # 2026-10-17：模板 KH 表头已在 get_month_sale_df 自动 rename 自愈；这里保留 assert 兜底
    # （模板列布局变了 / 行2 根本没有得分表头时才会触发）。
    assert score_col_name in month_sale_df.columns, (
        f"模板【2026年】行2 找不到得分列「{score_col_name}」——表头自愈未命中，"
        f"请检查 data/数据归属周期/{belonging_date}/快照日期/{snapshot_date}/月销售总表.xlsx 的列布局")
    product_info_map = product_info_df.set_index("SKU")[score_col_name]
    new_sku_month_sale_df[score_col_name] = new_sku_month_sale_df["SKU"].map(product_info_map)

    # 得分列以「商品资料.xlsx」为权威：模板里的得分列可能只是上期成品带过来的静态值
    # （换月若模板沿用上月成品，会出现"表头改了、值还是上月"——202609 期实测：模板 KH/EP
    #  值仍是 7 月得分，而商品资料已给到 8 月）。这里按当月得分列重刷整个 KH 列，
    #  只覆盖商品资料里有值的 SKU；商品资料无该 SKU / 未评级（NaN）→ 保留模板原值。
    _kh_mapped = month_sale_df["SKU"].map(product_info_map)
    _kh_valid = _kh_mapped.notna()
    _kh_changed = int((_kh_valid & (month_sale_df[score_col_name].astype(object)
                                    != _kh_mapped.astype(object))).sum())
    month_sale_df.loc[_kh_valid, score_col_name] = _kh_mapped[_kh_valid]
    print(f"【得分列】按商品资料「{score_col_name}」刷新【2026年】KH：命中 {int(_kh_valid.sum())} 行，"
          f"其中变更 {_kh_changed} 行")

    product_info_map = product_info_df.set_index("SKU")["状态"]
    new_sku_month_sale_df["状态"] = new_sku_month_sale_df["SKU"].map(product_info_map)

    processed_month_sale_df = month_sale_df.copy()
    processed_month_sale_df = pd.concat([processed_month_sale_df, new_sku_month_sale_df]).copy().reset_index(drop=True)
    processed_month_sale_df.to_csv(os.path.join(tmp_dir,"processed_month_sale_df_with_new_sku.csv"))
    for i in range(0, len(month_sale_group)):
        group = month_sale_group[i]
        sale_group_offset = i * 12
        month_offset = cur_acct_cycle_date.month
        # for month in range(1, month_offset+1):
        #     print(group, month)

        # 先连接SKU，将销量对齐，再补充缺失的SKU
        # 处理亚马逊单库
        total_sku_sale_df = total_sku_sale_df.rename(columns={"sku":"SKU"})#换成大写，利用merge特性删除重复列
        total_sku_sale_df2 = total_sku_sale_df.groupby(by=["实体", "平台", "仓库", "站点", "SKU"])[f"{cur_acct_cycle_date.strftime('%Y%m')}销量"].sum().reset_index()
        if group in group_to_amazon_inv_map:
            cur_group_month_sale_df = total_sku_sale_df2[total_sku_sale_df2["仓库"] == group_to_amazon_inv_map[group]].copy()
            cur_group_month_sale_df = cur_group_month_sale_df.groupby(by=["SKU"])[f"{cur_acct_cycle_date.strftime('%Y%m')}销量"].sum().reset_index()
            if cur_group_month_sale_df.shape[0] > 0:
                print(group)
                print(cur_group_month_sale_df)
                # raise Exception(processed_month_sale_df[processed_month_sale_df["SKU"].isin(new_sku_month_sale_df["SKU"])],cur_group_month_sale_df[cur_group_month_sale_df["SKU"].isin(new_sku_month_sale_df["SKU"])])
                processed_month_sale_df = pd.merge(left=processed_month_sale_df, right=cur_group_month_sale_df[["SKU",f"{cur_acct_cycle_date.strftime('%Y%m')}销量"]], how="left", left_on=["SKU"], right_on=["SKU"]).copy()
                suffix = f".{i}"
                if i == 0:
                    suffix = ""
                located_col_name = f"{month_offset}月{suffix}"
                processed_month_sale_df[located_col_name] = processed_month_sale_df[f"{cur_acct_cycle_date.strftime('%Y%m')}销量"]
                processed_month_sale_df[located_col_name] = processed_month_sale_df[located_col_name].fillna(0)  # ★2026-11-07 补 0：merge(how=left) 只给「本月有销量」的 SKU 赋值，归零 SKU 会残留模板旧值
                processed_month_sale_df.drop(f"{cur_acct_cycle_date.strftime('%Y%m')}销量", axis=1, inplace=True)
        #处理亚马逊合计
        if group == "亚马逊":
            cur_group_month_sale_df = total_sku_sale_df2[total_sku_sale_df2["平台"] == "亚马逊"].copy()
            cur_group_month_sale_df = cur_group_month_sale_df.groupby(by=["SKU"])[f"{cur_acct_cycle_date.strftime('%Y%m')}销量"].sum().reset_index()
            if cur_group_month_sale_df.shape[0] > 0:
                print(group)
                print(cur_group_month_sale_df)
                processed_month_sale_df = pd.merge(left=processed_month_sale_df, right=cur_group_month_sale_df[["SKU",f"{cur_acct_cycle_date.strftime('%Y%m')}销量"]], how="left", left_on=["SKU"], right_on=["SKU"]).copy()
                suffix = f".{i}"
                if i == 0:
                    suffix = ""
                located_col_name = f"{month_offset}月{suffix}"
                processed_month_sale_df[located_col_name] = processed_month_sale_df[f"{cur_acct_cycle_date.strftime('%Y%m')}销量"]
                processed_month_sale_df[located_col_name] = processed_month_sale_df[located_col_name].fillna(0)  # ★2026-11-07 补 0：merge(how=left) 只给「本月有销量」的 SKU 赋值，归零 SKU 会残留模板旧值
                processed_month_sale_df.drop(f"{cur_acct_cycle_date.strftime('%Y%m')}销量", axis=1, inplace=True)
        #包含了本土和跨境
        se_a_local_group = ["菲律宾","印尼","泰国"]
        if group in se_a_local_group:
            processed_month_sale_df = sum_se_local_site(cur_acct_cycle_date, processed_month_sale_df, i, group, month_offset, total_sku_sale_df2)

        se_cross_singlesite_group = ["新加坡","马来","越南"]#月销表的马来指的是lazada跨境站
        month_sale_group_lazada_shopee_site_mapping = {"马来":"马来西亚","新加坡":"新加坡","越南":"越南"}
        #新加坡包含Lazada和Shopee
        #马来和越南都只包含Lazada
        if group in se_cross_singlesite_group:
            processed_month_sale_df = sum_se_single_site(cur_acct_cycle_date, processed_month_sale_df, i, group, month_offset, total_sku_sale_df2, month_sale_group_lazada_shopee_site_mapping)

        #"虾皮境内" 所有其他非本土三国的shopee店铺
        # ,"东南亚",
        if group == "虾皮境内":
            processed_month_sale_df = sum_shopee_cross_site(cur_acct_cycle_date, processed_month_sale_df, i, group, month_offset, total_sku_sale_df2)
        if group == "东南亚":
            processed_month_sale_df = sum_total_southeast_asia_site(cur_acct_cycle_date, processed_month_sale_df, i, group, month_offset, total_sku_sale_df2)
        #"京东","eMAG","TEMU平台","美客多",
        if group == "京东":
            processed_month_sale_df = sum_jingdong_site(cur_acct_cycle_date, processed_month_sale_df, i, group, month_offset, total_sku_sale_df2)
        if group == "eMAG":
            processed_month_sale_df = sum_emag_site(cur_acct_cycle_date, processed_month_sale_df, i, group, month_offset, total_sku_sale_df2)
        if group == "TEMU平台":
            processed_month_sale_df = sum_temu_site(cur_acct_cycle_date, processed_month_sale_df, i, group, month_offset, total_sku_sale_df2)
        if group == "美客多":
            processed_month_sale_df = sum_mercado_site(cur_acct_cycle_date, processed_month_sale_df, i, group, month_offset, total_sku_sale_df2)
        if group == "总计":
            processed_month_sale_df = sum_total_site(cur_acct_cycle_date, processed_month_sale_df, i, group, month_offset, total_sku_sale_df2)
    processed_month_sale_df = processed_month_sale_df.reset_index(drop=True)
    processed_month_sale_df.to_csv(os.path.join(tmp_dir,"processed_month_sale_df.csv"))

    # 9.12新增：KJ列代码化：当年总销量 = 总计块已存在月列之和（公式原型 =SUM(JS4:KD4)）
    # ".23" = "总计"在 month_sale_group 中的下标（pandas 对重名列「1月」自动加序号后缀，总计是第 24 个 → .23）
    zongji_cols = [f"{m}月.23" for m in range(1, 13)
                   if f"{m}月.23" in processed_month_sale_df.columns]
    _kj_series = (
        processed_month_sale_df[zongji_cols]
        .apply(pd.to_numeric, errors="coerce")  # 停售"00"和空值统一转数值，"00"→0
        .fillna(0)
        .sum(axis=1)
    )
    # ⚠️ 202609 修：必须写回 KJ **既有列**，绝不能新建列！
    #   模板 KJ 表头是公式 =$B$1&"总销量"，openpyxl 生成的成品该表头**无缓存值** → pandas 读成 'Unnamed: 295'。
    #   若用新列名（如 "2026年总销量"）赋值，pandas 会把它**追加到 df 末尾**；而 fill_excel_from_df 是
    #   **按列位**写入（df.values.tolist()）→ 整列右移一位：KJ 留模板陈旧值、算好的总销量落到表外 KL(298) 列。
    #   （202609 期实测：KJ=532,372 停在 8 月、KL=595,009 才是 1~9 月；202608 未复现——其模板 KJ 表头
    #     带缓存值，pandas 恰好读到 '2026年总销量' 而原地覆盖；换月复用上期成品才首次暴露。）
    _cols = list(processed_month_sale_df.columns)
    _kj_col = _cols[_cols.index("状态") + 1] if "状态" in _cols else f"{cur_acct_cycle_date.year}年总销量"
    processed_month_sale_df[_kj_col] = _kj_series
    _kj_letter = get_column_letter(_cols.index(_kj_col) + 1)  # df 第 i 列(0基) → 写入 Excel 第 i+1 列
    # 硬断言：当年总销量必须落在 Excel 的 KJ 列。若模板布局变了导致落到别列，宁可当场失败，
    # 也不要再产出"KJ 停在上月、正确值写到表外"这种不报错的坏成品。
    assert _kj_letter == "KJ", (
        f"当年总销量落在 {_kj_letter} 列，不是 KJ——模板列布局可能已变，请核对 df 列序与 KJ 位（df 第{_cols.index(_kj_col)+1}列）")
    print(f"【当年总销量】写回既有列「{_kj_col}」= Excel KJ 列 ✔")
    # KK列代码化：提取SKU中的数字（公式原型 =1*MID(A4,...)数组公式），to_numeric保持与原缓存值一致的数值类型
    processed_month_sale_df["Unnamed: 296"] = pd.to_numeric(
        processed_month_sale_df["SKU"].astype(str).str.extract(r"(\d+)", expand=False), errors="coerce")

    data_dir = os.path.join("data", "数据归属周期", f"{belonging_date}", "快照日期", f"{snapshot_date}")
    template_path = os.path.join(data_dir, "月销售总表.xlsx")
    result_dir = os.path.join("result", "数据归属周期", f"{belonging_date}")
    result_path = os.path.join(result_dir, "月销售总表.xlsx")
    # 2026-11-02 新增：换月首次跑批时 result/数据归属周期/<新月>/ 还不存在，
    # openpyxl 的 wb.save() 不会建目录 → 会在算完全部数据、即将落盘时抛
    # FileNotFoundError 而整轮白跑（202610 期实测踩到）。三个下游脚本本来就有这行，
    # 只有主脚本漏了。放在这里（真正要写文件前）而不是文件开头，才与下游口径一致。
    os.makedirs(result_dir, exist_ok=True)
    
    for col in processed_month_sale_df.columns:
        processed_month_sale_df[col] = processed_month_sale_df[col].astype(object)#转为复杂类型用于标记
        months = [1,2,3,4,5,6,7,8,9,10,11,12]
        for month in months:
            if col == f"{month}月" or col.startswith(f"{month}月."):
                label_condition = (processed_month_sale_df["状态"] == "停售") & ((processed_month_sale_df[col] == 0) | processed_month_sale_df[col].isna())
                processed_month_sale_df.loc[label_condition, col] = "00"#停售标记
                processed_month_sale_df.to_csv(os.path.join(tmp_dir,f"labeled_processed_month_sale_df.csv"))
    
    fill_excel_from_df(df= processed_month_sale_df, start_row=4, start_col= 1, template_path=template_path, output_path=result_path, sheet_name=f"{cur_acct_cycle_date.year}年")
    print(f"{cur_acct_cycle_date}新增SKU{new_sku_month_sale_df.shape[0]}条")

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

    # ===== 任务②-b（10.07新增）：模板缺行 → 末尾追加新 SKU 行 =====
    # 数据部模板的【欧洲各国】sheet 行数常年落后于【2026年】（202609 实测：497 行 vs 545 SKU）；
    # 当月新增/晚出现的 SKU 在这张 sheet 里"无行可写"，旧逻辑 `if r:` 直接整行跳过 →
    # **静默漏销量**（202609 漏 28 个 SKU / 469 件；202608 漏 13 个 / 242 件，逐期累积）。
    # 修法照【2026年】的做法：在末尾追加行，套用参考行样式 + 按行平移公式（月合计/年度总销/ER），
    # 基础列 A/B/状态 从商品资料补，当月各国销量按 pivot 写入。行3 月列合计公式范围到 4907，覆盖新行。
    # ⚠️ 历史月（1~当月-1）无数据源（total_sku_sale_df 只有当月）→ 留空；新行年度总销=仅已填月之和。
    _missing = [s for s in pivot.index if str(s).strip() not in sku_row]
    if _missing:
        _ref_row = max(sku_row.values())          # 参考行 = 最后一个已有数据行
        _max_col = ws.max_column
        _model_map = product_info_df.set_index("SKU")["型号"]
        _status_map = product_info_df.set_index("SKU")["状态"]
        _status_col = None
        for _c in range(1, _max_col + 1):
            if str(ws.cell(2, _c).value or "").strip() == "状态":
                _status_col = _c
                break
        _r = _ref_row
        for _sku in _missing:
            _r += 1
            _sku_s = str(_sku).strip()
            for _c in range(1, _max_col + 1):
                _src = ws.cell(_ref_row, _c)
                _dst = ws.cell(_r, _c)
                if _src.has_style:
                    _dst._style = copy(_src._style)   # 同 workbook 内复制样式，不触发跨簿 IndexError
                _v = _src.value
                if isinstance(_v, str) and _v.startswith("="):
                    _dst.value = Translator(_v, origin=_src.coordinate).translate_formula(_dst.coordinate)
                elif isinstance(_v, ArrayFormula):
                    _dst.value = ArrayFormula(_dst.coordinate,
                        Translator(_v.text, origin=_src.coordinate).translate_formula(_dst.coordinate))
            ws.cell(_r, 1).value = _sku_s
            ws.cell(_r, 2).value = _model_map.get(_sku_s)
            if _status_col:
                ws.cell(_r, _status_col).value = _status_map.get(_sku_s)
            for _country, _c in country_col.items():
                _vv = pivot.loc[_sku].get(_country, 0)
                if _vv:
                    ws.cell(_r, _c, int(_vv))
            if ws.row_dimensions[_ref_row].height is not None:
                ws.row_dimensions[_r].height = ws.row_dimensions[_ref_row].height
        if ws.auto_filter.ref:
            _m = re.match(r"^([A-Z]+\d+:[A-Z]*)", ws.auto_filter.ref)
            if _m:
                ws.auto_filter.ref = f"{_m.group(1)}{_r}"
        print(f"➕ 【欧洲各国】模板缺行 → 追加 {len(_missing)} 个新 SKU 行"
              f"（行 {_ref_row + 1}~{_r}），当月各国值写入 "
              f"{get_column_letter(blk)}~{get_column_letter(blk + 9)}")

    # ===== 任务②补：得分列（EP）重刷 =====
    # 该 sheet 的得分列是模板里的**静态列**，不像年度总销列(EF~EO)那样带公式 → 换月不会自动跟。
    # 若模板来自上期成品，成品里就会出现"表头还是上个月得分列名、值还是上个月得分"（202609 期实测：
    # 表头停在「7月销量得分」）。这里按当月 score_col_name 重刷表头 + 从商品资料回填得分。
    score_col = None
    for _c in range(2, ws.max_column + 1):
        _h = str(ws.cell(2, _c).value or "")
        if "销量得分" in _h:            # 表头形如「8月销量得分」（不会误匹配 EF~EO 的「总销量」）
            score_col = _c
            break
    if score_col:
        _score_map = product_info_df.set_index("SKU")[score_col_name]
        _filled = 0
        for r in range(4, ws.max_row + 1):
            _sku = ws.cell(r, 1).value
            if _sku in (None, ""):
                continue
            _v = _score_map.get(str(_sku).strip())
            if _v is not None and not (isinstance(_v, float) and pd.isna(_v)):
                ws.cell(r, score_col).value = _v
                _filled += 1
        ws.cell(2, score_col).value = score_col_name   # 表头跟当月得分列名
        print(f"✅ 【欧洲各国】得分列 {get_column_letter(score_col)}({score_col}) "
              f"已刷新为「{score_col_name}」，回填 {_filled} 行")
    else:
        print("⚠️ 【欧洲各国】未找到「销量得分」列 → 得分列未刷新，请检查模板列布局")

    # ===== 2026-10-17 新增：成品【{year}年】KH 表头兜底自愈 =====
    # 模板=复制上月成品 → KH2 表头停上期；模板侧已在 get_month_sale_df 里自动 rename 落盘，
    # 成品经 fill_excel_from_df 复制模板带出正确表头。但模板若被占用没能落盘，成品会带旧表头出来
    # → 这里在任务②③共用的 wb 里补一次 rename，随下方唯一一次 wb.save(result_path) 生效，不额外读写文件。
    _kh_ren = rename_score_col_in_ws(wb[f"{cur_acct_cycle_date.year}年"], score_col_name)
    if _kh_ren:
        print(f"【得分列表头】成品【{cur_acct_cycle_date.year}年】{get_column_letter(_kh_ren[2])}2："
              f"「{_kh_ren[0]}」→「{_kh_ren[1]}」（模板侧改名未落盘时的兜底）")

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
        # 2026-11-02：补「台湾」——与上面「东南亚」列口径一致（台湾 Shopee 跨境并入「虾皮境内」）。
        # 不补这张汇总表会少 6 件，与 2026年 sheet 的「东南亚」列对不上。
        group_month_total(sites=["菲律宾", "印尼", "泰国", "马来西亚", "新加坡", "越南", "台湾"]),        # 东南亚
        group_month_total(platform="亚马逊", sites=["美国", "加拿大", "墨西哥"]),                # 北美（口径已验证：7月AK8=8801=美+加+墨）
    ]

    ws2 = wb["2021-2026年销量情况"]
    for k, val in enumerate(cur_vals):
        c = col_y + 10 * k                      # 四个分组块首列：G(7) / Q(17) / AA(27) / AK(37)
        ws2.cell(r_cur, c).value = val                                        # 当月值
        ws2.cell(r_cur, c + 1).value = val / ws2.cell(r_cur, c - 1).value - 1   # 同比增长率 = 当月 / 去年同月
        if month_offset > 1:                    # 1月时 r_prev=1 是表头（字符串），不写环比
            ws2.cell(r_cur, c + 2).value = val / ws2.cell(r_prev, c).value - 1  # 今年环比 = 当月 / 今年上月

    # 汇总行（第14行）：模板里各块汇总都是静态值（不是公式）、脚本原本一格都不写 → 用表内月值重算，幂等。
    # ⚠️ 2026-11-03 修（任务㉑附带发现）：原代码只重算「全球」块（G/I/J），亚马逊(Q/S/T)/
    #    东南亚(AA/AC/AD)/北美(AK/AM/AN) 三块的汇总行一直带模板陈旧值——
    #    202610 交付表实测 Q14=355387(停在1~9月旧口径) vs 月值和513012、AA14=117024 vs 128176、
    #    AK14=58755 vs 84417，整月 10 月都没进汇总。现按同口径重算全部四块：
    #    当年累计=当年列2~13行之和；汇总同比=当年累计/上一年累计-1（写 +2 列，+1 列留空）；
    #    历年总计=SUM(该块 2021 列..当年列的汇总)。
    for _k in range(4):
        _c0 = col_y + 10 * _k                      # 四个分组块当年列：G(7) / Q(17) / AA(27) / AK(37)
        y_sum = 0
        for r0 in range(2, 14):                     # 第2~13行 = 1~12月
            cv = ws2.cell(r0, _c0).value
            if isinstance(cv, (int, float)):
                y_sum += cv
        ws2.cell(14, _c0).value = y_sum                                    # 当年累计 = 当年列 1~12月之和
        ws2.cell(14, _c0 + 2).value = y_sum / ws2.cell(14, _c0 - 1).value - 1  # 汇总同比（模板放在 +2 列，+1 列留空）
        ws2.cell(14, _c0 + 3).value = sum(                                 # 历年总计 = SUM(该块 2021 列..当年列)
            v for v in (ws2.cell(14, c0).value for c0 in range(_c0 - 5, _c0 + 1))
            if isinstance(v, (int, float)))

    print(f"【月度参数】当月行={r_cur} 上月行={r_prev} 当年列={col_y} 欧洲块起始列={blk}")
    print(f"【2021-2026年销量情况】{month_offset}月：总计={cur_vals[0]} 亚马逊={cur_vals[1]} 东南亚={cur_vals[2]} 北美={cur_vals[3]}")
    print(f"【汇总行】当年累计={ws2.cell(14, col_y).value} 同比={ws2.cell(14, col_y + 2).value} 历年总计={ws2.cell(14, col_y + 3).value}")
    wb.save(result_path)  # 任务②③共用一个 wb，统一保存一次

    # 任务⑩：把两张表的列分组（分级显示）从模板的「上月态」自动推进到当月，免去每月手工「取消和排列」
    outline_reports, outline_warnings = arrange_month_outline(
        result_path, wb, cur_acct_cycle_date.year, cur_acct_cycle_date.month)
    for line in outline_reports:
        print(line)
    for line in outline_warnings:
        print(f"⚠️ {line}")

    # 2026-11-10：把亚马逊报告区间问题在末尾再汇总一遍，避免埋在前面几千行日志里被忽略
    if AMZ_SANITY_WARNINGS:
        print("\n" + "=" * 70)
        print(f"🚨 亚马逊报告疑似导错月份区间，共 {len(AMZ_SANITY_WARNINGS)} 份，本期数字可能偏小：")
        for _w in AMZ_SANITY_WARNINGS:
            print("   " + _w)
        print("=" * 70)
    if REPORT_SANITY_WARNINGS:
        _sev = [w for w in REPORT_SANITY_WARNINGS if w.startswith("【严重")]
        print("\n" + "=" * 70)
        print(f"{'🚨' if _sev else '⚠️'} 非亚马逊渠道报告区间体检：{len(REPORT_SANITY_WARNINGS)} 条"
              f"（其中严重 {len(_sev)} 条）")
        for _w in REPORT_SANITY_WARNINGS:
            print("   " + _w)
        print("=" * 70)
    if SUPPLEMENT_LOG:
        print("\n【补充订单】本次共并入 "
              + "、".join(f"{t}({f} → {n} 行)" for t, f, n in SUPPLEMENT_LOG))

    print(f"✅ 结果文件（成品，注意不是data目录下的模板）: {os.path.abspath(result_path)}")
    # os.startfile(result_path)  # 202609 期：无人值守跑，先注释掉（跑完自动弹窗会占住成品，导致下轮保存失败）
    # print(manage_inv_df)

    # ===== 任务⑰：库龄情况表（2026-10-22 需求，随月销表周期输出）=====
    # ⚠️ 编号：这条 10-22 曾登记成「任务⑪」，与《修改指南》里 09-16 的
    #    「任务⑪＝印尼海外仓口径」（本文件 L860）撞号 → 2026-11-02 起改叫 ⑰。
    # 输入  trunk/data/数据归属周期/<期>/库龄/**.csv|.xlsx —— 按**内容**识别，不认文件名；
    #       亚马逊「亚马逊物流库存」（英文/中文版）+ EMAG「FBE Inventory」+ 美客多「长期仓储费」
    # 输出  trunk/data/库龄报表/亚马逊EMAG库龄情况总表.xlsx（**总表逐期累积**，不是每期一个新文件）
    #       4 个 sheet：本期明细 / 站点汇总 / 历史汇总 / 取数清单
    # 期号  优先用本次月销表的归属期；该期 库龄/ 下没有报告时，退到「最新有报告的一期」并告警。
    #       （10-29 加：belonging_date 是手改的硬编码，11 月初跑 202610 时若把 11 月导的库龄
    #         报告放进 202611/库龄，按老逻辑任务⑰ 只认 202610 → 新报告一份都读不到。
    #         反过来，如果归属期目录里已经有报告，就**只**用归属期的，绝不串期。）
    # 失败  只打印 ⚠️，不阻断、不改退出码 —— 库龄是附带产出，不能因为它把月销表成品废掉
    # 详见  trunk/data/数据归属周期/202611/库龄/README-库龄报告下载教学.md
    try:
        import sys as _sys
        _sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
        import age_report
        print("\n【库龄情况表】")
        _age_period = str(belonging_date)
        if not age_report.find_age_files(_age_period):
            _auto = age_report.latest_age_period()
            if _auto and _auto != _age_period:
                print(f"⚠️ {_age_period}/库龄 下没有库龄报告，改用最新有报告的一期：{_auto}")
                _age_period = _auto
        age_report.main(["--period", _age_period])
    except Exception as _e:
        print(f"⚠️ 库龄情况表生成失败 {_e!r}（不影响月销表成品）")


