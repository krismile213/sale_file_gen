# -*- coding: utf-8 -*-
"""订单报告日期覆盖检查（入口）—— 防「date range 按站点本地时间设、purchase-date 是 UTC」丢单。

用法：
    # 体检某个快照目录里所有订单报告（= 跑前体检第 ⑧ 块）
    python pipeline/tools/check_report_coverage.py --snap <快照目录> --day 20261016
    python pipeline/tools/check_report_coverage.py --day 20261016        # 自动定位最新快照

    # 单份报告（归属月默认从文件名的 YYYYMM 挖）
    python pipeline/tools/check_report_coverage.py --report <订单报告.txt> [--month 202610] [--day 20261016]

    # 反向审计：已导入工作簿里哪些 UTC 日丢了凌晨/傍晚时段（定位历史损伤）
    python pipeline/tools/check_report_coverage.py --book <工作簿.xlsx>

原理与 2026-09 实测数据见 core/report_coverage.py 顶部 docstring。
"""
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "core"))
import report_coverage as RC  # noqa: E402

if __name__ == "__main__":
    sys.exit(RC.main(sys.argv[1:]))
