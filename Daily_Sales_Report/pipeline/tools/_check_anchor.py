# -*- coding: utf-8 -*-
"""扫全部「每日库存分析表」的粉色规则锚点日期，判断是否有设成未来的。"""
import sys, os, re, glob, datetime
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _cf_eval as C
import openpyxl

TODAY = datetime.date(2026, 10, 23)
PAT = re.compile(r"([^\s!(),<>=]+)!\$?([A-Z]{1,3})\$?(\d+)")


def main():
    books = sorted(glob.glob(os.path.join(C.WORK, "*库存销售表.xlsx")))
    print(f"{'工作簿':<30}{'sheet':<26}{'锚点':<13}{'判定'}")
    print("-" * 82)
    bad = []
    for p in books:
        base = os.path.basename(p)
        if base.startswith("~$"):
            continue
        try:
            wb = openpyxl.load_workbook(p, data_only=True)
        except Exception as e:
            print(f"  !! {base}: {e}")
            continue
        for sn in wb.sheetnames:
            if "每日库存分析表" not in sn:
                continue
            rules = [r for r in C.load_cf(p, sn) if "LOOKUP" in r["formula"]]
            anc = None
            for r in rules:
                m = PAT.search(r["formula"])
                if m and m.group(1) in wb.sheetnames:
                    anc = wb[m.group(1)].cell(
                        int(m.group(3)),
                        C.column_index_from_string(m.group(2))).value
                    break
            a = anc.date() if hasattr(anc, "date") else anc
            if not isinstance(a, datetime.date):
                flag = "未识别"
            elif a > TODAY:
                flag = "⚠ 未来"
                bad.append((base, sn, a))
            elif a == TODAY:
                flag = "= 今天"
            else:
                flag = "过去"
            print(f"{base[:28]:<30}{sn[:24]:<26}{str(a):<13}{flag}")
        wb.close()
    print("-" * 82)
    if bad:
        print(f"⚠ 有 {len(bad)} 张表锚点在未来：")
        for b in bad:
            print("   ", b)
    else:
        print("✅ 全部锚点 <= 今天，粉色规则不会标到未来日期")


if __name__ == "__main__":
    main()
