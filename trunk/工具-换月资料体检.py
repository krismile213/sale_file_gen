# -*- coding: utf-8 -*-
"""
换月资料体检（**只读**，不改任何文件）

用途：每月换月**跑脚本之前**先跑一次，检查
  ① 新一期快照目录在不在、四份模板就位没；
  ② 源报告文件齐不齐（按上一期跑通的实际文件名清单逐条核对）；
  ③ 月销总表模板的「得分列」表头名 与 商品资料.xlsx 是否一致
     （不一致 → 主脚本第 1452/1453 行的 assert 会直接崩）；
  ④ 欧洲/北美模板里「当月块」表头是否预置（决定模板能不能直接复用）。

用法（在 trunk 目录下执行）：
  PYTHONPATH="C:/Users/dell/AppData/Roaming/Python/Python314/site-packages" \
    /c/Python314/python.exe 工具-换月资料体检.py 202609 20261007
参数：<归属周期 YYYYMM> <快照日期 YYYYMMDD>；不传则用下方默认值。
退出码：0 = 无致命缺失；1 = 有致命缺失（红叉）。
"""
import os
import sys
import warnings

import openpyxl
from openpyxl.utils import get_column_letter as _gl

warnings.filterwarnings("ignore")

TRUNK = os.path.dirname(os.path.abspath(__file__))

OK, BAD, WARN = [], [], []


def ck(cond, ok_msg, bad_msg, hard=True):
    if cond:
        OK.append(ok_msg)
    else:
        (BAD if hard else WARN).append(bad_msg)
    return cond


def sec(title):
    print("\n" + "─" * 78)
    print(title)
    print("─" * 78)


def main(belonging="202609", snap="20261007"):
    year = belonging[:4]
    m = int(belonging[4:6])
    # 得分列名 = 上一个月（8 月归属期用「7月销量得分」）
    score_expect = f"{(m - 1) if m > 1 else 12}月销量得分"

    snap_dir = os.path.join(TRUNK, "data", "数据归属周期", belonging, "快照日期", snap)
    print("=" * 78)
    print(f"换月资料体检 · 归属周期 {belonging} · 快照日期 {snap}")
    print(f"目录：{snap_dir}")
    print(f"得分列期望值：{score_expect}（= 归属月 {m} 的上一个月）")
    print("=" * 78)

    # ---------- [A] 目录骨架 ----------
    sec("[A] 目录骨架")
    ck(os.path.isdir(snap_dir), f"快照目录存在",
       f"❌ 快照目录不存在：{snap_dir}\n      → 先建目录，再把本期原始报告放进去", hard=True)
    if not os.path.isdir(snap_dir):
        _report()
        return 1
    for sub in ("东南亚", os.path.join("东南亚", "跨境"), "欧洲", "北美"):
        p = os.path.join(snap_dir, sub)
        ck(os.path.isdir(p), f"{sub} 目录存在", f"❌ 缺子目录：{sub}", hard=True)

    # ---------- [B] 四份模板 ----------
    sec("[B] 四份模板（缺一不可，欧洲/北美可从上月成品复制）")
    tpl_main = os.path.join(snap_dir, "月销售总表.xlsx")
    tpl_se = os.path.join(snap_dir, "东南亚", f"{year}年东南亚销量表Southeast Asia Sales table.xlsx")
    tpl_eu = os.path.join(snap_dir, "欧洲", f"{year}年欧洲销量表Europe Sales table.xlsx")
    tpl_na = os.path.join(snap_dir, "北美", f"{year}年北美销量表North America Sales table.xlsx")
    ck(os.path.isfile(tpl_main), "月销总表模板存在", "❌ 缺 月销售总表.xlsx（数据部给）", True)
    ck(os.path.isfile(tpl_se), "东南亚模板存在", "❌ 缺 东南亚/2026年东南亚销量表…xlsx（数据部给）", True)
    if ck(os.path.isfile(tpl_eu), "欧洲模板存在",
          "❌ 缺 欧洲/{年}年欧洲销量表…xlsx → 用上月成品复制（见下）", True):
        pass
    if not os.path.isfile(tpl_eu):
        prev_eu = os.path.join(TRUNK, "result", "欧洲销量表", _prev_month(belonging),
                               f"{year}年欧洲销量表Europe Sales table.xlsx")
        print(f"      参考命令：copy \"{prev_eu}\" \"{tpl_eu}\"")
    if not os.path.isfile(tpl_na):
        prev_na = os.path.join(TRUNK, "result", "北美销量表", _prev_month(belonging),
                               f"{year}年北美销量表North America Sales table.xlsx")
        print(f"      参考命令：copy \"{prev_na}\" \"{tpl_na}\"")
    ck(os.path.isfile(tpl_na), "北美模板存在",
       "❌ 缺 北美/{年}年北美销量表…xlsx → 用上月成品复制", True)

    # ---------- [C] 源报告清单 ----------
    sec("[C] 源报告清单（按 202608 期跑通的实际文件名逐条核对）")

    def _pick(*names):
        '''同名报表多种写法：取第一个真实存在的；都不在就返回第一个（报缺）。'''
        for _n in names:
            if os.path.isfile(os.path.join(snap_dir, _n)):
                return _n
        return names[0]

    # ★ 2026-11 实测：亚马逊后台把「欧洲」站点改叫「欧盟」导出（主脚本已有
    #   REPORT_NAME_ALIASES 别名，体检这边以前只认「欧洲」→ 每天假报缺 2 项）。
    _eu_tag = "欧盟" if os.path.isfile(
        os.path.join(snap_dir, f"亚马逊欧盟站点{belonging}订单报告.txt")) else "欧洲"
    spec = {
        "亚马逊订单报告(巨川)": [f"亚马逊{g}站点{belonging}订单报告.txt"
                                 for g in (_eu_tag, "日本", "澳大利亚", "北美", "沙特阿拉伯",
                                           "阿拉伯联合酋长国", "印度")],
        "亚马逊订单报告(interflow)": [f"interflow{g}站点{belonging}订单报告.txt"
                                      for g in ("北美", "澳大利亚", "阿拉伯联合酋长国")],
        "亚马逊管理库存(主脚本要读)": [_pick(f"亚马逊欧盟管理库存报告{snap}.csv",
                                             f"亚马逊欧洲管理库存报告{snap}.csv")],
        "EMAG": [f"EMAG全站点{belonging}订单报告.xlsx"],
        "Temu": [f"Temu全球{belonging}订单报告.csv"],
        "美客多": [f"美客多墨西哥站点{belonging}订单报告.xlsx"],
        "京东": [f"京东国内{belonging}订单报告.xlsx"],
        "东南亚 Shopee": [f"东南亚/Shopee{c}站点{belonging}订单报告.xlsx"
                          for c in ("菲律宾", "泰国", "印尼")],
        "东南亚 Lazada": [f"东南亚/Lazada{c}站点{belonging}订单报告.xlsx"
                          for c in ("菲律宾", "泰国")],
        "东南亚 Tiktok": [f"东南亚/Tiktok{c}站点{belonging}订单报告.csv"
                          for c in ("菲律宾", "印尼")],
        # ★ 2026-11 实测：主脚本读的名是「…订单报告.xlsx」，「商品出库报告」是老写法，二选一。
        "东南亚 海外仓/线下": [_pick(f"东南亚/印尼海外仓{belonging}订单报告.xlsx",
                                     f"东南亚/印尼海外仓{belonging}商品出库报告.xlsx"),
                              f"东南亚/菲律宾KSON仓{belonging}线下订单报告.xlsx"],
        "跨境 Shopee": [f"东南亚/跨境/Shopee跨境{c}站点{belonging}订单报告.xlsx"
                        for c in ("马来西亚", "越南", "新加坡")],
        "跨境 Lazada": [f"东南亚/跨境/Lazada跨境{c}站点{belonging}订单报告.xlsx"
                        for c in ("新加坡", "泰国", "菲律宾", "越南", "马来西亚")],
    }
    for cat, files in spec.items():
        miss = [f for f in files if not os.path.isfile(os.path.join(snap_dir, f))]
        if miss:
            BAD.append(f"❌ {cat}：缺 {len(miss)}/{len(files)} → " + "、".join(miss))
        else:
            OK.append(f"{cat}：{len(files)}/{len(files)} 齐")

    # ---------- [D] 表头/列名一致性 ----------
    sec("[D] 表头一致性（商品资料缺列才崩；模板 KH 表头停上期主脚本会自动 rename 自愈）")
    prod = os.path.join(TRUNK, "商品资料.xlsx")
    prod_cols = []
    if os.path.isfile(prod):
        wb = openpyxl.load_workbook(prod, read_only=True)
        ws = wb["商品资料"]
        for i, row in enumerate(ws.iter_rows(min_row=2, max_row=2, values_only=True)):
            prod_cols = [str(v) for v in row if v]
        wb.close()
        got = [c for c in prod_cols if c.endswith("月销量得分")]
        ck(bool(got) and got[0] == score_expect,
           f"商品资料.xlsx 得分列 = {got[0] if got else '无'}（与预期一致）",
           f"⚠️ 商品资料.xlsx 得分列 = {got[0] if got else '找不到'}，但本期期望 {score_expect}"
           f" → 主脚本会退化取最新月列并打警告（得分可能停上期），建议先在商品资料补列", hard=True)
    else:
        BAD.append("❌ 找不到 trunk/商品资料.xlsx")

    if os.path.isfile(tpl_main):
        wb = openpyxl.load_workbook(tpl_main, data_only=True)
        sheets = wb.sheetnames
        cur = f"{year}年"
        ck(cur in sheets, f"月销总表有【{cur}】sheet", f"❌ 月销总表缺【{cur}】sheet")
        if cur in sheets:
            ws = wb[cur]
            v = ws.cell(2, 294).value  # KH = 得分列
            # 2026-10-17：模板 KH 表头停上期（模板=复制上月成品）已由主脚本自动 rename 自愈
            #（get_month_sale_df → heal_score_col_header_file）→ 不再是 fatal，降级为提醒。
            ck(v == score_expect, f"月销总表模板 KH 列 = {v}（与预期一致）",
               f"ℹ️ 月销总表模板 KH 列 = {v}，期望 {score_expect}"
               f" → 模板沿用上月成品；主脚本会自动 rename 自愈（2026-10-17 起不用手工改）",
               hard=False)

            # ★ 202609 期新增：得分列的「值」是否真跟到当月。
            #   得分列在模板里是静态值（不像年度总销列 EF~EO 带公式）→ 若模板沿用上期成品，
            #   会出现"表头改成当月、值还是上期"。这里拿商品资料逐 SKU 对一遍，提前预警。
            if os.path.isfile(prod):
                try:
                    wbp = openpyxl.load_workbook(prod, read_only=True)
                    wsp = wbp["商品资料"]
                    hdr = next(wsp.iter_rows(min_row=2, max_row=2, values_only=True))
                    hi = next((i for i, x in enumerate(hdr) if str(x) == score_expect), None)
                    smap = {}
                    if hi is not None:
                        for r in wsp.iter_rows(min_row=3, values_only=True):
                            k = r[0] if len(r) else None
                            if k in (None, ""):
                                continue
                            smap[str(k).strip()] = r[hi] if hi < len(r) else None
                    wbp.close()
                    n = diff = 0
                    for rr in range(4, ws.max_row + 1):
                        k = ws.cell(rr, 1).value
                        if k in (None, ""):
                            continue
                        b = smap.get(str(k).strip())
                        if b is None:
                            continue
                        n += 1
                        if ws.cell(rr, 294).value != b:
                            diff += 1
                    ck(n == 0 or diff == 0,
                       f"月销总表模板 KH 值与商品资料「{score_expect}」一致（可比 {n} 行）",
                       f"⚠️ 模板 KH 值与商品资料「{score_expect}」有 {diff}/{n} 行不一致"
                       f" → 模板很可能沿用上期成品、得分列没刷新（主脚本已自愈，仍建议索取本期新版模板）",
                       hard=False)
                except Exception as ex:
                    WARN.append(f"⚠️ 得分列值核对跳过（{type(ex).__name__}: {ex}）")

            # ★ 202609 期新增：欧洲各国 sheet 的得分列表头也要跟当月（同样是模板静态列）
            esheet = f"{cur}-欧洲各国每月销售情况"
            if esheet in sheets:
                we = wb[esheet]
                ep = None
                for c in range(2, we.max_column + 1):
                    if "销量得分" in str(we.cell(2, c).value or ""):
                        ep = c
                        break
                if ep:
                    hv = we.cell(2, ep).value
                    ck(hv == score_expect,
                       f"欧洲各国得分列 {_gl(ep)}({ep}) 表头 = {hv}（与预期一致）",
                       f"⚠️ 欧洲各国得分列表头 = {hv}，期望 {score_expect}"
                       f" → 模板沿用上期成品（只改了 KH 表头）。主脚本已自愈，仍建议索取本期新版模板",
                       hard=False)

                # ★ 2026-10-07 新增 [D-4]：欧洲各国 sheet 的**行数**是否跟得上【{cur}】sheet。
                #   数据部模板两张 sheet 常年不同步（202609 实测：欧洲各国 497 行 vs【2026年】545 SKU）
                #   → 当月新增/晚出现的 SKU 在这张 sheet 里"无行可写"（旧脚本 if r: 静默跳过、漏销量：
                #   202609 漏 28 个/469 件、202608 漏 13 个/242 件，逐期累积）。
                #   主脚本已自愈（末尾追加行 + 平移公式），这里提前预警，便于索取本期新版模板。
                n_main = sum(1 for r in range(4, ws.max_row + 1)
                             if ws.cell(r, 1).value not in (None, ""))
                n_eu = sum(1 for r in range(4, we.max_row + 1)
                           if we.cell(r, 1).value not in (None, ""))
                ck(n_eu >= n_main,
                   f"欧洲各国 sheet 行数 {n_eu} ≥ 【{cur}】SKU 数 {n_main}（行数跟得上）",
                   f"⚠️ 欧洲各国 sheet 只有 {n_eu} 行，而【{cur}】有 {n_main} 个 SKU（少 {n_main - n_eu}）"
                   f" → 模板两张 sheet 不同步，晚出现的 SKU 无行可写"
                   f"（主脚本已自愈：末尾追加行 + 平移公式；仍建议向数据部索取新版模板）",
                   hard=False)
            # 当月列必须为空（模板语义）
            col_cur = 3 + (m - 1)  # 欧洲块：1月→C(3)，m月→3+(m-1)
            cl = _gl(col_cur)
            sample = [ws.cell(r, col_cur).value for r in range(4, 12)]
            ck(all(v in (None, 0) for v in sample),
               f"月销总表模板「{m}月」列（欧洲块 {cl} 列）为空 = 是模板不是成品",
               f"⚠️ 月销总表模板「{m}月」列（{cl}）有值 → 这可能是**成品**不是模板，会重复累加",
               hard=False)
        wb.close()

    # ---------- [E] 东南亚模板结构 ----------
    sec("[E] 东南亚模板（10 个 sheet + 当月块表头）")
    if os.path.isfile(tpl_se):
        wb = openpyxl.load_workbook(tpl_se, read_only=True)
        need = ["SEA Total 东南亚总销量表", "Philippines 菲律宾销量表", "Thailand 泰国销量表",
                "Indonesia 印尼销量表", "Singapore 新加坡销量表", "Malaysia Shopee 虾皮跨境销量表"]
        got = wb.sheetnames
        miss = [s for s in need if s not in got]
        ck(not miss, f"关键 sheet 齐（共 {len(got)} 个）",
           f"❌ 东南亚模板缺 sheet：{miss}（完整清单见 {got}）")
        wb.close()

    # ---------- [F] 欧洲/北美模板：当月块表头 ----------
    sec("[F] 欧洲/北美模板「当月块」表头是否预置（决定能否直接复用上月成品）")
    for tag, path, base, n, bw in (("欧洲", tpl_eu, 6, 10, 11), ("北美", tpl_na, 6, 3, 4)):
        if not os.path.isfile(path):
            continue
        wb = openpyxl.load_workbook(path, read_only=True)
        tsheet = f"EU Total 欧洲总销量表" if tag == "欧洲" else "NA Total 北美总销量表"
        if tsheet in wb.sheetnames:
            ws = wb[tsheet]
            c0 = base + n + 1 + (m - 1) * bw
            row1 = []
            for r in ws.iter_rows(min_row=1, max_row=1, min_col=c0, max_col=c0 + bw - 1, values_only=True):
                row1 = list(r)
            ck(any(v for v in row1), f"{tag} Total 表 {m} 月块表头已预置",
               f"❌ {tag} Total 表 {m} 月块表头为空 → 该模板不支持 {m} 月，必须重建")
        wb.close()

    # ---------- [G] 上月成品（当模板 / 交叉验证都要用） ----------
    sec("[G] 上月成品（交叉验证 & 复制模板用）")
    prev = _prev_month(belonging)
    for tag, sub in (("月销总表", "数据归属周期"), ("东南亚", "东南亚销量表"),
                     ("欧洲", "欧洲销量表"), ("北美", "北美销量表")):
        p = os.path.join(TRUNK, "result", sub, prev)
        ck(os.path.isdir(p), f"{tag} 上月产物目录存在（{prev}）",
           f"⚠️ 没有 {prev} 的{tag}产物 → 交叉验证会缺基准（不影响本期出数）", hard=False)

    _report()
    return 1 if BAD else 0


def _prev_month(belonging):
    y, mm = int(belonging[:4]), int(belonging[4:6])
    if mm == 1:
        return f"{y - 1}12"
    return f"{y}{mm - 1:02d}"


def _report():
    print("\n" + "=" * 78)
    print(f"✅ 通过 {len(OK)} 项    ❌ 致命 {len(BAD)} 项    ⚠️ 提醒 {len(WARN)} 项")
    print("=" * 78)
    if BAD:
        print("\n❌ 必须先解决：")
        for b in BAD:
            print("   " + b)
    if WARN:
        print("\n⚠️ 提醒（不阻塞）：")
        for w in WARN:
            print("   " + w)
    if not BAD:
        print("\n通过项摘要：")
        for o in OK:
            print("   ✓ " + o)
    if not BAD and not WARN:
        print("\n🎉 资料齐了，可以按《换月操作手册》开跑。")


if __name__ == "__main__":
    args = sys.argv[1:3]
    sys.exit(main(*(args if len(args) == 2 else ())))
