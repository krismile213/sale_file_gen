# -*- coding: utf-8 -*-
"""merge_day() 全场景测试：月底分页导出（多份/重叠/主文件不完整/断档）。

用临时目录 + 造小 xlsx，不碰真实快照。
"""
import os
import shutil
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.abspath(os.path.join(HERE, "..", "core")))
sys.path.insert(0, HERE)
import merge_emag_orders as ME
import openpyxl

PERIOD = "202610"
CN = ["订单编号", "订单日期", "产品代码", "数量"]
EN = ["Order number", "Order date", "Product code", "Quantity"]
ok = bad = 0


def chk(c, msg):
    global ok, bad
    if c:
        ok += 1
        print("  [OK] " + msg)
    else:
        bad += 1
        print("  [!!] " + msg)


def make(path, hdr, rows):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "纸张 1"
    ws.append(hdr)
    for r in rows:
        ws.append(r)
    wb.save(path)


def rows_of(path):
    wb = openpyxl.load_workbook(path, read_only=True)
    ws = wb[wb.sheetnames[0]]
    out = [tuple(r) for r in ws.iter_rows(values_only=True)]
    wb.close()
    return out


def T(name):
    return os.path.join(tmp, name)


tmp = tempfile.mkdtemp(prefix="emagtest_")
MAIN = "EMAG全站点%s订单报告.xlsx" % PERIOD
try:
    # ── 1 空目录 ──
    print("1) 空目录")
    r = ME.merge_day(tmp, PERIOD, apply=True, quiet=True)
    chk(not r["ok"] and r["reason"] == "无分段文件", "无分段文件")

    # ── 2 只有主文件 ──
    print("2) 只有主文件")
    make(T(MAIN), CN, [["A1", "2026-10-01 10:00:00", "G1", 1]])
    r = ME.merge_day(tmp, PERIOD, apply=True, quiet=True)
    chk(not r["ok"] and r["reason"] == "无分段文件", "仍报无分段文件")
    chk(len(rows_of(T(MAIN))) == 2, "主文件未被碰")

    # ── 3 只有 1 份分段 ──
    print("3) 只有 1 份分段（不该自动改名）")
    os.remove(T(MAIN))
    make(T("EMAG全站点%s订单报告1.xlsx" % PERIOD), CN,
         [["A1", "2026-10-01 10:00:00", "G1", 1]])
    r = ME.merge_day(tmp, PERIOD, apply=True, quiet=True)
    chk(not r["ok"] and "只有 1 份" in r["reason"], "只有 1 份，不处理")
    chk(not os.path.exists(T(MAIN)), "未生成主文件")

    # ── 4 月底分页：3 份，含分页重叠 ──
    print("4) 月底分页 3 份 + 分页重叠")
    os.remove(T("EMAG全站点%s订单报告1.xlsx" % PERIOD))
    make(T("EMAG全站点%s订单报告1.xlsx" % PERIOD), CN,
         [["A1", "2026-10-01 10:00:00", "G1", 1],
          ["A2", "2026-10-02 10:00:00", "G2", 2]])
    make(T("EMAG全站点%s订单报告2.xlsx" % PERIOD), CN,
         [["A2", "2026-10-02 10:00:00", "G2", 2],        # 与 1 重叠
          ["A3", "2026-10-03 10:00:00", "G3", 3]])
    make(T("EMAG全站点%s订单报告3.xlsx" % PERIOD), EN,
         [["A3", "2026-10-03 10:00:00", "G3", 3],        # 与 2 重叠
          ["A4", "2026-10-04 09:00:00", "G4", 4]])
    r = ME.merge_day(tmp, PERIOD, apply=True, quiet=True)
    chk(r["ok"] and r["rows"] == 4, "合并 4 单（3 份共 6 行，去重 2）实际 %d" % r["rows"])
    chk(len(r["parts"]) == 3, "识别到 3 份分段")
    got = rows_of(T(MAIN))
    chk(got[0] == tuple(CN), "表头统一为中文")
    chk([str(x[1])[:10] for x in got[1:]] ==
        ["2026-10-01", "2026-10-02", "2026-10-03", "2026-10-04"], "按日期升序")
    chk(not r["gaps"], "无日期断档")
    d = {str(x[0]): x for x in got[1:]}
    chk(d["A4"][2] == "G4" and d["A4"][3] == 4, "英文表头那份映射正确")

    # ── 5 幂等：分段已被主文件覆盖 ──
    print("5) 幂等（分段数据已被主文件覆盖）")
    before = os.path.getmtime(T(MAIN))
    r = ME.merge_day(tmp, PERIOD, apply=True, quiet=True)
    chk(not r["ok"] and "已包含全部分段" in r["reason"], "判定已覆盖，不动")
    chk(os.path.getmtime(T(MAIN)) == before, "主文件未被重写")

    # ── 6 主文件不完整 + 来了新分段（浏览器重名 (1) 场景）──
    print("6) 主文件不完整 + 新分段 (1)（浏览器重名）")
    make(T("EMAG全站点%s订单报告(1).xlsx" % PERIOD), CN,
         [["A5", "2026-10-05 10:00:00", "G5", 5]])
    r = ME.merge_day(tmp, PERIOD, apply=True, quiet=True)
    chk(r["ok"], "触发重新合并（主文件缺 1 单）：%s" % r["reason"])
    got = rows_of(T(MAIN))
    chk(len(got) == 6, "1 表头 + 5 数据行（实际 %d）" % len(got))
    chk("A5" in {str(x[0]) for x in got[1:]}, "新单 A5 已并入")
    bakdir = os.path.join(ME.DSR, ".workbuddy", "backup_emag_merge")
    baks = [f for f in os.listdir(bakdir)
            if f.startswith(os.path.splitext(MAIN)[0])] if os.path.isdir(bakdir) else []
    chk(bool(baks), "旧主文件已备份（%d 份）" % len(baks))

    # ── 7 中文括号 / 横线 / 空格命名 ──
    print("7) 命名变体识别")
    for nm, seq in [("EMAG全站点%s订单报告（7）.xlsx" % PERIOD, 7),
                    ("EMAG全站点%s订单报告-8.xlsx" % PERIOD, 8),
                    ("EMAG全站点%s订单报告 9.xlsx" % PERIOD, 9)]:
        make(T(nm), CN, [["B%d" % seq, "2026-10-10 10:00:00", "G", 1]])
    found = dict(ME.find_parts(tmp, PERIOD))
    chk(7 in found and 8 in found and 9 in found,
        "识别中文括号/横线/空格：%s" % sorted(found))
    chk(MAIN not in [os.path.basename(p) for p in found.values()],
        "主文件不被当成分段")

    # ── 8 日期断档告警 ──
    print("8) 日期断档告警")
    shutil.rmtree(tmp)
    os.makedirs(tmp)
    make(T("EMAG全站点%s订单报告1.xlsx" % PERIOD), CN,
         [["C1", "2026-10-01 10:00:00", "G1", 1],
          ["C2", "2026-10-02 10:00:00", "G1", 1]])
    make(T("EMAG全站点%s订单报告2.xlsx" % PERIOD), CN,
         [["C3", "2026-10-06 10:00:00", "G1", 1]])   # 缺 10-03~10-05
    r = ME.merge_day(tmp, PERIOD, apply=True, quiet=True)
    chk(r["ok"] and len(r["gaps"]) == 1, "检测到 1 处断档（实际 %d）" % len(r["gaps"]))
    if r["gaps"]:
        a, b, n = r["gaps"][0]
        chk((a, b, n) == ("2026-10-02", "2026-10-06", 3),
            "断档内容正确：%s 之后缺 %d 天 → %s" % (a, n, b))
finally:
    shutil.rmtree(tmp, ignore_errors=True)

print()
print("通过 %d / %d" % (ok, ok + bad))
sys.exit(1 if bad else 0)
