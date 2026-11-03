# -*- coding: utf-8 -*-
"""补丁：归档件 DAY(TODAY格±算术尾巴) → DAY(锚点格±尾巴)
   修复 retarget 括号内尾巴漏匹配（澳/印写法 DAY(...!$F$1-1)），对全部模拟透视归档件幂等重跑。"""
# --- pipeline 路径引导（2026-10-14 目录收敛自动注入）---
import os as _os, sys as _sys
_bp = _os.path.dirname(_os.path.abspath(__file__))
for _i in range(6):
    if _os.path.isdir(_os.path.join(_bp, "1.每日销售数据表")):
        break
    _bp = _os.path.dirname(_bp)
DSR = _bp
_bcore = _os.path.join(DSR, "pipeline", "core")
if _bcore not in _sys.path:
    _sys.path.insert(0, _bcore)
# --- 引导结束 ---
import os, sys
pass  # [normalize] 路径已由引导块设置
import config as C
import month_rollover as M
from openpyxl import load_workbook

for f, anchors in C.ANCHORS.items():
    p = os.path.join(M.ARCHIVE_DIR, M.ARCH_PREFIX + f)
    if not os.path.exists(p):
        print(f"{f}: 缺归档件!")
        continue
    wb = load_workbook(p, data_only=False)
    sn, coord = anchors[0]
    n = M.retarget_today_to_anchor(wb, sn, coord)
    if n:
        tmp = p + ".patch.xlsx"
        wb.save(tmp)
        wb.close()
        M.restore_parts(p, tmp)
        os.replace(tmp, p)
        print(f"{f}: 补重定向 {n} 处 → 已保存")
    else:
        wb.close()
        print(f"{f}: 0（已定向或本就引用锚点格）")

# 抽查：澳/印 归档件透视图公式现在的样子
for f, probe_cells in (("澳大利亚库存销售表.xlsx", [("新品销售情况数据透视图", "E2")]),
                       ("印度库存销售表.xlsx", [("新品销售情况数据透视图", "B2")])):
    p = os.path.join(M.ARCHIVE_DIR, M.ARCH_PREFIX + f)
    wb = load_workbook(p, read_only=True, data_only=False)
    for sn, coord in probe_cells:
        v = wb[sn][coord].value
        print(f"[抽查] {f} {sn}!{coord}: {v}")
    wb.close()
print("PATCH DONE")
