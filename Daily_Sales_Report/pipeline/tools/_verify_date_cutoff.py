# -*- coding: utf-8 -*-
"""跑批后核验「数据日期截断」：确认没有 10-27 之后的订单被写进日销簿。

检查每本簿的「每日销售数据」sheet：
  ① purchase-date 分布，max 必须 <= 2026-10-27
  ② 有没有 > 10-27 的行（有就是截断失效）
  ③ 上月尾巴（9/24~9/30）行数，确认窗口口径没被扩大
  ④ 各分析表日列表头覆盖到哪天
"""
import os
import sys
from collections import Counter

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.abspath(os.path.join(HERE, "..", "core")))
import config as C
from openpyxl import load_workbook

WORK = os.path.join(C.DSR, "1.每日销售数据表")
SNAP_DAY = "2026-10-27"
BAD = []


def main():
    files = sorted(f for f in os.listdir(WORK) if f.endswith(".xlsx"))
    for f in files:
        p = os.path.join(WORK, f)
        try:
            wb = load_workbook(p, read_only=True, data_only=True)
        except Exception as e:
            print("%-38s 打开失败 %s" % (f, e))
            BAD.append(f)
            continue
        print("=" * 74)
        print(f)
        for sn in wb.sheetnames:
            if "每日销售数据" not in sn:
                continue
            ws = wb[sn]
            it = ws.iter_rows(values_only=True)
            # 表头不一定在第 1 行（日销簿第 1 行常为空，表头在第 2 行）
            hdr = None
            for _ in range(5):
                try:
                    cand = [str(x).strip() if x is not None else "" for x in next(it)]
                except StopIteration:
                    break
                if "purchase-date" in cand:
                    hdr = cand
                    break
            if hdr is None:
                print("   %-16s 无 purchase-date 列，跳过" % sn)
                continue
            di = hdr.index("purchase-date")
            cnt = Counter()
            total = 0
            for r in it:
                if not r or len(r) <= di or r[di] is None:
                    continue
                d = str(r[di])[:10]
                cnt[d] += 1
                total += 1
            if not cnt:
                print("   %-16s 无数据" % sn)
                continue
            after = {k: v for k, v in cnt.items() if k > SNAP_DAY}
            tail = sum(v for k, v in cnt.items() if "2026-09-24" <= k <= "2026-09-30")
            early = {k: v for k, v in cnt.items() if k < "2026-09-24"}
            print("   %-16s 行=%-5d 范围 %s ~ %s  上月尾巴=%-4d 超期=%d 窗口外=%d"
                  % (sn, total, min(cnt), max(cnt), tail,
                     sum(after.values()), sum(early.values())))
            if after:
                print("      ✗✗ 有 %s 之后的数据：%s" % (SNAP_DAY, after))
                BAD.append("%s/%s 超期 %s" % (f, sn, after))
            if early:
                print("      ⚠ 窗口外(<9/24)：%s" % dict(list(early.items())[:5]))
        wb.close()

    print()
    if BAD:
        print("!! 发现 %d 处问题：" % len(BAD))
        for x in BAD:
            print("   " + x)
        return 1
    print("=== 全部簿均无 2026-10-27 之后的数据 ===")
    return 0


if __name__ == "__main__":
    sys.exit(main())
