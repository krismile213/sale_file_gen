# -*- coding: utf-8 -*-
"""清理 iter_rows 空单元格膨胀（历史遗留：销售 sheet dimension 虚标到 60000 行）。

原理：openpyxl 普通模式 iter_rows 越界会把空 Cell 实例化进 ws._cells，
保存时写出空 <c/> 并把 dimension 撑到 60000。清理 = 删除"无值+无样式+无批注+无超链接"
的实例化空格，dimension 自动回归真实范围。有值/有样式/合并格一律不动 → 数据零影响。
"""
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
import os, sys, shutil, zipfile, re

HERE = os.path.dirname(os.path.abspath(__file__))
# [normalize] DSR 由上方引导块赋值
sys.path.insert(0, HERE)
import config as C
from runner import restore_parts
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell

BASE = C.WORK_DIR
BAK = os.path.join(DSR, ".workbuddy", "backup_bloat_fix_20261010")

TARGETS = [
    "欧盟英国库存销售表.xlsx",
    "北美加拿大库存销售表.xlsx",
    "日本库存销售表.xlsx",
    "沙特库存销售表.xlsx",
    "澳大利亚库存销售表.xlsx",
    "印度库存销售表.xlsx",
    "阿联酋库存销售表.xlsx",
    "Interflowd北美加拿大库存销售表.xlsx",
    "Interflowd澳大利亚库存销售表.xlsx",
    "Interflowd阿联酋库存销售表.xlsx",
]


def purge_sheet(ws):
    n = 0
    for key in list(ws._cells.keys()):
        c = ws._cells[key]
        if (c.value is None and not c.has_style and c.comment is None
                and c.hyperlink is None and not isinstance(c, MergedCell)):
            del ws._cells[key]
            n += 1
    return n


def max_dim_row(p):
    z = zipfile.ZipFile(p)
    mx = 0
    for n in z.namelist():
        if re.match(r"xl/worksheets/sheet\d+\.xml$", n):
            head = z.read(n)[:2000].decode("utf-8", errors="replace")
            m = re.search(r'<dimension ref="([^"]+)"', head)
            if m:
                mx = max(mx, int(re.findall(r"\d+", m.group(1))[-1]))
    z.close()
    return mx


def main():
    os.makedirs(BAK, exist_ok=True)
    for f in TARGETS:
        p = os.path.join(BASE, f)
        print(f"\n==== {f} ====", flush=True)
        try:
            fd = os.open(p, os.O_RDWR); os.close(fd)
        except OSError as e:
            print(f"  [LOCK] 文件被占用(errno {e.errno})，跳过——关闭后重跑本脚本", flush=True)
            continue
        before = max_dim_row(p)
        if before < 30000:
            print(f"  [skip] dimension 最大行={before}，无膨胀", flush=True)
            continue
        dst = os.path.join(BAK, f)
        if not os.path.exists(dst):
            shutil.copy2(p, dst)
            print(f"  备份 → {dst}", flush=True)
        tmp = p + ".fix.xlsx"
        wb = load_workbook(p)
        total, hit = 0, []
        for ws in wb.worksheets:
            n = purge_sheet(ws)
            if n > 300:
                total += n
                hit.append(f"{ws.title}({n})")
        if total == 0:
            wb.close()
            print(f"  [skip] 无空格实体（仅 dimension 虚标），重写修复", flush=True)
            wb = load_workbook(p)
            wb.save(tmp)
            wb.close()
        else:
            print(f"  清除 {total} 空格: {', '.join(hit)}", flush=True)
            wb.save(tmp)
            wb.close()
        restored, piv = restore_parts(p, tmp)
        shutil.move(tmp, p)
        after = max_dim_row(p)
        ok = "OK" if zipfile.ZipFile(p).testzip() is None else "ZIP-BAD!"
        print(f"  [保真] 回填{len(restored)}部件 pivot刷新{piv}处 | dimension {before}→{after} | zip {ok}", flush=True)
    print("\nFIX BLOAT DONE", flush=True)


if __name__ == "__main__":
    main()
