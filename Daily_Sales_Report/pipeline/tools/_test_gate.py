# -*- coding: utf-8 -*-
"""补行证据闸逻辑单测（不碰文件）。

验证三条证据各自能放行、以及「无库存无销量」被正确挡住。
直接复刻 runner 里的判据表达式，避免为了测试去改生产代码。
"""
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
PIPE = os.path.dirname(HERE)
sys.path.insert(0, PIPE)
sys.path.insert(0, os.path.join(PIPE, "core"))

import config as C                                     # noqa: E402
import analysis as An                                  # noqa: E402


def gate(stock_rec, sold_q):
    """复刻 runner.py 补行段的证据闸。返回 (是否放行, 原因)。"""
    if C.TRISTATE_ENABLED and C.TRISTATE_REQUIRE_EVIDENCE \
            and not C.TRISTATE_KEEP_MONTH_GATE:
        av = (stock_rec or {}).get("avail") or 0
        ib = (stock_rec or {}).get("inb") or 0
        if av > 0:
            return True, f"本站可售 {av:.0f}"
        if ib > 0:
            return True, f"本站在途 {ib:.0f}"
        if sold_q > 0:
            return True, f"本月销量 {sold_q:.0f}"
        return False, "无库存无销量"
    return False, "旧口径（不在本测试范围）"


CASES = [
    ("有可售", {"avail": 31, "inb": 0, "listing": True}, 0),
    ("有在途", {"avail": 0, "inb": 350, "listing": True}, 0),
    ("有销量", {"avail": 0, "inb": 0, "listing": True}, 12),
    ("有销量+listing不通", {"avail": 0, "inb": 0, "listing": False}, 3),
    ("三无（listing通）", {"avail": 0, "inb": 0, "listing": True}, 0),
    ("三无（listing不通）", {"avail": 0, "inb": 0, "listing": False}, 0),
    ("源里没有", None, 0),
    ("只有销量1件", {"avail": 0, "inb": 0, "listing": True}, 1),
]

print(f"闸门配置：ENABLED={C.TRISTATE_ENABLED} "
      f"REQUIRE_EVIDENCE={C.TRISTATE_REQUIRE_EVIDENCE} "
      f"KEEP_MONTH_GATE={C.TRISTATE_KEEP_MONTH_GATE}\n")
print(f"{'场景':<22}{'放行':<6}{'原因'}")
print("-" * 52)
for name, rec, q in CASES:
    ok, why = gate(rec, q)
    print(f"{name:<22}{'✅' if ok else '❌':<6}{why}")

print("\n--- 三态判定（同输入）---")
for name, rec, q in CASES:
    print(f"  {name:<22}{An.tri_state(rec)}")

print("\n--- 正则边界 ---")
for t in ("G598", "K013-018", "G136R", "G590J", "SKU139", "Info", "Product code"):
    print(f"  {t:<16}{'✅' if An._SKU_RE.match(t) else '❌'}")
