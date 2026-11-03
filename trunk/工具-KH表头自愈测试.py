# -*- coding: utf-8 -*-
"""
KH 得分列表头自愈测试（2026-10-17 新增，配套 month_sale_file.py 的 heal_score_col_header_file）

背景：月销总表模板 = 复制上月成品 → 【{year}年】KH2（294列）表头停在上期得分列名
（202610 已 staged 模板实测停在「8月销量得分」，而 10 月期望「9月销量得分」）。
过去每月首次跑主脚本要先手工改一格，否则补充信息处 assert 硬崩。现已三层自动自愈：
  ① 模板侧  heal_score_col_header_file：改名并落盘（一劳永逸）
  ② 成品侧  rename_score_col_in_ws：任务②共用 wb 里兜底（模板被占用没能落盘时）
  ③ 内存侧  get_month_sale_df：pandas df 列名兜底改名（模板被占用时 assert 也不崩）

本测试用 202609 真实模板伪造「202611 期 + 十月成品模板」场景逐条验证，并验证 202609 零回归。
用法（在 trunk 目录下执行）：
  PYTHONPATH="C:/Users/dell/AppData/Roaming/Python/Python314/site-packages" \
    /c/Python314/python.exe 工具-KH表头自愈测试.py
"""
import os
import shutil
import stat
import hashlib

import pandas as pd
from openpyxl import load_workbook

import month_sale_file as msf

TRUNK = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(TRUNK, "data", "数据归属周期")

# 真实模板（只读参照 + 零回归对象）
REAL_BELONG, REAL_SNAP, REAL_SCORE = "202609", "20261006", "8月销量得分"
# 伪造场景：202611 期跑批，模板 = 复制 10 月成品 → 表头停在「9月销量得分」，期望「10月销量得分」
FAKE_BELONG, FAKE_SNAP = "202611", "20261117"
FAKE_SCORE, STALE_SCORE = "10月销量得分", "9月销量得分"
SHEET = "2026年"
KH_COL = 294
EU_SHEET = "2026年-欧洲各国每月销售情况"
EP_COL, EP_HEADER = 146, "7月销量得分"

PASS, FAIL = [], []


def check(name, cond, detail=""):
    (PASS if cond else FAIL).append(name)
    print(f"  {'✅' if cond else '❌'} {name}" + (f" —— {detail}" if detail and not cond else ""))


def sha256(p):
    h = hashlib.sha256()
    with open(p, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def kh2(p):
    wb = load_workbook(p, data_only=True)
    v = wb[SHEET].cell(2, KH_COL).value
    wb.close()
    return v


def make_fake(header=STALE_SCORE, blank=False):
    """复制 202609 真实模板到伪造快照目录；可选把 KH2 改成指定表头 / 置空。"""
    src = os.path.join(DATA, REAL_BELONG, "快照日期", REAL_SNAP, "月销售总表.xlsx")
    dst_dir = os.path.join(DATA, FAKE_BELONG, "快照日期", FAKE_SNAP)
    os.makedirs(dst_dir, exist_ok=True)
    dst = os.path.join(dst_dir, "月销售总表.xlsx")
    shutil.copy2(src, dst)
    if header != REAL_SCORE or blank:
        wb = load_workbook(dst)
        wb[SHEET].cell(2, KH_COL).value = None if blank else header
        wb.save(dst)
        wb.close()
    return dst


def _series_equal(a, b):
    """NaN 感知的逐值比较（直接 == 会把 NaN!=NaN 误判成不一致）。"""
    a, b = list(a), list(b)
    return len(a) == len(b) and all(
        (x == y) or (pd.isna(x) and pd.isna(y)) for x, y in zip(a, b))


def t1_stale_rename():
    print("\n[T1] 停上期表头 → 自动 rename 落盘（模板侧①）")
    p = make_fake()
    ref = pd.read_excel(p, sheet_name=SHEET, skiprows=1)      # 改名前的参照
    df = msf.get_month_sale_df(FAKE_BELONG, FAKE_SNAP, FAKE_SCORE)
    check("df 列名变成当月得分列", FAKE_SCORE in df.columns)
    check("df 旧列名已消失", STALE_SCORE not in df.columns)
    check("模板文件落盘改名", kh2(p) == FAKE_SCORE, f"KH2={kh2(p)!r}")
    check("得分列数值不变（只改名不改值）", _series_equal(df[FAKE_SCORE], ref[STALE_SCORE]))
    wb = load_workbook(p, data_only=True)
    ws = wb[SHEET]
    check("相邻表头未被动（291~293 / 295）",
          (ws.cell(2, 291).value, ws.cell(2, 292).value,
           ws.cell(2, 293).value, ws.cell(2, 295).value) == ("类型", "品牌", "系列", "状态"))
    check("欧洲各国 EP 不在本函数管辖范围（仍停 7 月，由既有成品侧逻辑管）",
          wb[EU_SHEET].cell(2, EP_COL).value == EP_HEADER)
    wb.close()


def t2_idempotent():
    print("\n[T2] 幂等：已是当月 → 不改、不多落盘")
    p = make_fake()
    msf.get_month_sale_df(FAKE_BELONG, FAKE_SNAP, FAKE_SCORE)   # 第一次：rename + 落盘
    h0 = sha256(p)
    st = msf.heal_score_col_header_file(p, SHEET, FAKE_SCORE)   # 第二次：应 ok 且不保存
    check("heal 返回 ok", st == "ok", st)
    check("文件字节不变（没有多余落盘）", sha256(p) == h0)
    df = msf.get_month_sale_df(FAKE_BELONG, FAKE_SNAP, FAKE_SCORE)
    check("df 列名仍正确", FAKE_SCORE in df.columns)


def t3_real_zero_regression():
    print("\n[T3] 真实 202609 零回归（已是当月 → 不落盘，字节不变）")
    p = os.path.join(DATA, REAL_BELONG, "快照日期", REAL_SNAP, "月销售总表.xlsx")
    h0 = sha256(p)
    df = msf.get_month_sale_df(REAL_BELONG, REAL_SNAP, REAL_SCORE)
    check("真实模板字节不变", sha256(p) == h0)
    check("df 得分列名不变", REAL_SCORE in df.columns)
    ref = pd.read_excel(p, sheet_name=SHEET, skiprows=1)
    check("df 列集合与直接 pandas 读一致", df.columns.tolist() == ref.columns.tolist())
    check("得分列数值与直接读一致", _series_equal(df[REAL_SCORE], ref[REAL_SCORE]))


def t4_locked_fallback():
    print("\n[T4] 模板被占用（只读文件）→ locked 降级 + 内存改名 + 成品侧兜底")
    p = make_fake()
    os.chmod(p, stat.S_IREAD)                    # Windows 上 = 只读属性，openpyxl 保存抛 PermissionError
    try:
        st = msf.heal_score_col_header_file(p, SHEET, FAKE_SCORE)
        if st == "locked":
            check("heal 返回 locked（不崩）", True)
            check("文件未被改动（表头仍停上期）", kh2(p) == STALE_SCORE, f"KH2={kh2(p)!r}")
            df = msf.get_month_sale_df(FAKE_BELONG, FAKE_SNAP, FAKE_SCORE)
            check("内存兜底：df 列名已改成当月（assert 不会崩）", FAKE_SCORE in df.columns)
            # 成品侧兜底（= 任务②的做法）：load → rename in ws → 存到成品路径
            out = p + ".result.xlsx"
            wb = load_workbook(p)
            ren = msf.rename_score_col_in_ws(wb[SHEET], FAKE_SCORE)
            check("成品侧 rename 返回正确三元组",
                  ren == (STALE_SCORE, FAKE_SCORE, KH_COL), repr(ren))
            wb.save(out)
            wb.close()
            check("成品表头被兜底刷正", kh2(out) == FAKE_SCORE, f"KH2={kh2(out)!r}")
            os.remove(out)
        else:
            print(f"  ⚠️ 本机对只读文件竟可直接写盘（st={st}）——locked 分支跳过（非 Windows 语义，不计失败）")
    finally:
        os.chmod(p, stat.S_IWRITE)


def t5_missing_header():
    print("\n[T5] 行2 无得分表头 → 清晰失败（不静默、不瞎改名）")
    p = make_fake(blank=True)
    check("KH2 已置空", kh2(p) is None)
    st = msf.heal_score_col_header_file(p, SHEET, FAKE_SCORE)
    check("heal 返回 missing", st == "missing", st)
    df = msf.get_month_sale_df(FAKE_BELONG, FAKE_SNAP, FAKE_SCORE)
    check("df 无得分列", FAKE_SCORE not in df.columns)
    # 模拟主脚本补充信息处的 assert：应带清晰信息失败，而不是 KeyError / 静默放过
    try:
        assert FAKE_SCORE in df.columns, (
            f"模板【2026年】行2 找不到得分列「{FAKE_SCORE}」——表头自愈未命中，"
            f"请检查 data/数据归属周期/{FAKE_BELONG}/快照日期/{FAKE_SNAP}/月销售总表.xlsx 的列布局")
        check("assert 应该失败", False)
    except AssertionError as e:
        check("assert 带清晰信息失败", "找不到得分列" in str(e))


def t6_helper_units():
    print("\n[T6] helper 单测（rename_score_col_in_ws / _find_score_header_col）")
    p = make_fake()
    wb = load_workbook(p)
    ws = wb[SHEET]
    check("_find_score_header_col 命中 294", msf._find_score_header_col(ws) == KH_COL)
    ren = msf.rename_score_col_in_ws(ws, FAKE_SCORE)
    check("第一次改名返回三元组", ren == (STALE_SCORE, FAKE_SCORE, KH_COL), repr(ren))
    check("改名已生效（内存）", ws.cell(2, KH_COL).value == FAKE_SCORE)
    check("第二次改名返回 None（幂等）", msf.rename_score_col_in_ws(ws, FAKE_SCORE) is None)
    check("未传列名返回 None（skip）", msf.rename_score_col_in_ws(ws, None) is None)
    wb.close()


def t7_fill_carries_header():
    print("\n[T7] fill_excel_from_df 复制已自愈模板 → 成品带出正确表头（真实产出链路）")
    p = make_fake()
    msf.get_month_sale_df(FAKE_BELONG, FAKE_SNAP, FAKE_SCORE)     # 模板侧 rename + 落盘
    out = p + ".filled.xlsx"
    msf.fill_excel_from_df(df=pd.DataFrame([[1]]), template_path=p, output_path=out,
                           start_row=4, start_col=1, sheet_name=SHEET)
    check("成品 KH2 = 当月得分列（模板自愈被带出）", kh2(out) == FAKE_SCORE, f"KH2={kh2(out)!r}")
    os.remove(out)


def main():
    fake_root = os.path.join(DATA, FAKE_BELONG)
    if os.path.exists(fake_root):
        print(f"❌ 伪造目录已存在：{fake_root}\n   为避免误删真实数据，请先手动处理后再跑本测试。")
        raise SystemExit(2)
    try:
        t1_stale_rename()
        t2_idempotent()
        t3_real_zero_regression()
        t4_locked_fallback()
        t5_missing_header()
        t6_helper_units()
        t7_fill_carries_header()
    finally:
        if os.path.exists(fake_root):
            shutil.rmtree(fake_root)
            print(f"\n🧹 已清理伪造目录 {fake_root}")
    print("\n" + "=" * 60)
    print(f"结果：{len(PASS)} 通过 / {len(FAIL)} 失败")
    if FAIL:
        for f in FAIL:
            print(f"  ❌ {f}")
        raise SystemExit(1)
    print("全部通过 ✅")


if __name__ == "__main__":
    main()
