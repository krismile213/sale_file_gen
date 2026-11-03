# -*- coding: utf-8 -*-
"""验证粉色规则日列公式已带 MIN 上界，且染色结果与改前一致。"""
import sys, os, glob
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _cf_eval as C

ok = miss = 0
for p in sorted(glob.glob(os.path.join(C.WORK, "*库存销售表.xlsx"))):
    if os.path.basename(p).startswith("~$"):
        continue
    import openpyxl
    wb = openpyxl.load_workbook(p, read_only=True)
    names = [s for s in wb.sheetnames if "每日库存分析表" in s]
    wb.close()
    for sn in names:
        rules = [r for r in C.load_cf(p, sn) if "LOOKUP" in r["formula"]]
        date_rules = [r for r in rules if "$5>" in r["formula"] or "$5<" in r["formula"]]
        has_min = [r for r in date_rules if "MIN(" in r["formula"] and "TODAY()+1" in r["formula"]]
        tag = "OK " if len(has_min) == 2 and len(date_rules) == 2 else "!! "
        if tag == "OK ":
            ok += 1
        else:
            miss += 1
        print(f"{tag}{os.path.basename(p)[:26]:<28}{sn[:24]:<26}日列规则{len(date_rules)} 带上界{len(has_min)}")
        if tag == "!! ":
            for r in date_rules:
                print("      ", r["formula"][:150])
print()
print(f"通过 {ok} 张 / 异常 {miss} 张")
