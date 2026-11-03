# -*- coding: utf-8 -*-
"""零证据补行闸的单元验证（不加载任何工作簿）。

复刻 runner.py 里那段闸的判定顺序，喂构造样本，确认：
  · G599（listing=True、库存全 0）→ 放行
  · amzn.gr.* 临时 SKU       → 挡住
  · listing=False            → 挡住
  · 库存源里根本没有这条记录   → 挡住
  · 老闸（有货/有在途/有销量）→ 照旧放行，与本改动无关
"""
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.abspath(os.path.join(HERE, "..", "core")))

import config as C
import analysis as An


def gate(stock, sold, sku):
    """返回 (是否补行, 原因)。顺序与 runner.py L1746~1770 一致。"""
    rec = stock.get(sku) or {}
    av = rec.get("avail") or 0
    ib = rec.get("inb") or 0
    q = (sold or {}).get(sku, 0)
    st = An.tri_state_of_sku(stock, sku)
    if av > 0 or ib > 0 or q > 0:
        return True, "老闸放行(有货/在途/销量)"
    if (C.ZERO_EVIDENCE_SKU_RE
            and re.match(C.ZERO_EVIDENCE_SKU_RE, sku)
            and st != An.STATE_NOT_LISTED
            and rec):
        return True, "零证据白名单放行(listing 已建、货未到)"
    return False, "挡住"


CASES = [
    # (说明, stock, sold, sku, 期望)
    ("G599 listing=True 全 0",
     {"G599": {"name": "vivo X500 Pro", "avail": 0, "inb": 0, "listing": True}},
     {}, "G599", True),
    ("G602 listing=True 全 0",
     {"G602": {"name": "小米18 Pro", "avail": 0, "inb": 0, "listing": True}},
     {}, "G602", True),
    ("G599 listing=None（未知）",
     {"G599": {"name": "vivo", "avail": 0, "inb": 0, "listing": None}},
     {}, "G599", True),
    ("G599 listing=False（明确未上架）",
     {"G599": {"name": "vivo", "avail": 0, "inb": 0, "listing": False}},
     {}, "G599", False),
    ("amzn.gr.* 临时 SKU",
     {"amzn.gr.G117-UOFBnm4I1": {"name": "x", "avail": 0, "inb": 0, "listing": True}},
     {}, "amzn.gr.G117-UOFBnm4I1", False),
    ("库存源里没有这条记录",
     {}, {}, "G599", False),
    ("G520 有可售（老闸）",
     {"G520": {"name": "x", "avail": 12, "inb": 0, "listing": True}},
     {}, "G520", True),
    ("G700 有在途（老闸）",
     {"G700": {"name": "x", "avail": 0, "inb": 30, "listing": True}},
     {}, "G700", True),
    ("G701 有销量（老闸）",
     {"G701": {"name": "x", "avail": 0, "inb": 0, "listing": True}},
     {"G701": 3}, "G701", True),
]

print("ZERO_EVIDENCE_SKU_RE = %r" % C.ZERO_EVIDENCE_SKU_RE)
print()
ok = bad = 0
for desc, stock, sold, sku, want in CASES:
    got, why = gate(stock, sold, sku)
    flag = "OK " if got == want else "!! "
    if got == want:
        ok += 1
    else:
        bad += 1
    print("%s %-34s → %-6s 期望 %-6s  [%s]"
          % (flag, desc, "补行" if got else "不补", "补行" if want else "不补", why))

print()
print("通过 %d / %d" % (ok, ok + bad))

# 关闭开关应完全回退
print()
C.ZERO_EVIDENCE_SKU_RE = ""
got, why = gate({"G599": {"name": "v", "avail": 0, "inb": 0, "listing": True}},
                {}, "G599")
print("开关置空后 G599 → %s [%s]（应为「不补」，即完全回退老行为）"
      % ("补行" if got else "不补", why))
if got:
    bad += 1

sys.exit(0 if bad == 0 else 1)
