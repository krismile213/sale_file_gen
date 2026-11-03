# -*- coding: utf-8 -*-
"""把 xl/metadata.xml + xl/richData/* 与单元格 vm=/cm= 索引补回生产簿。
源 = .workbuddy/backup_daily_20261013（跑批前原件，含富数据部件）。
默认干跑（在临时副本上验证，不碰生产）；加 --apply 才真正写回，且先备份到
.workbuddy/backup_metadata_20261013/。
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

# [normalize] DSR 由上方引导块赋值
sys.path.insert(0, DSR)
import runner as R   # noqa: E402

SRC = os.path.join(DSR, ".workbuddy", "backup_daily_20261013")
DST = os.path.join(DSR, "1.每日销售数据表")
BK = os.path.join(DSR, ".workbuddy", "backup_metadata_20261013")
TMPDIR = os.path.join(DSR, ".workbuddy", "_richdata_dry")


def parts_of(p):
    with zipfile.ZipFile(p) as z:
        return set(z.namelist())


def sheet_vmcm(p):
    """{sheet名: {格: 'vm=..'}}"""
    out = {}
    with zipfile.ZipFile(p) as z:
        try:
            sm = R._sheet_name_map(z.read("xl/workbook.xml"), z.read("xl/_rels/workbook.xml.rels"))
        except KeyError:
            return out
        for nm, path in sm.items():
            try:
                raw = z.read(path)
            except KeyError:
                continue
            if b'vm="' not in raw and b'cm="' not in raw:
                continue
            d = {}
            for m in re.finditer(r"<c\b[^>]*/?>", raw.decode("utf-8", "ignore")):
                a = dict(R._XML_ATTR_RE.findall(m.group(0)))
                got = [f'{k}="{a[k]}"' for k in ("cm", "vm") if k in a]
                if "r" in a and got:
                    d[a["r"]] = " ".join(got)
            if d:
                out[nm] = d
    return out


def verify(src, dst, tag):
    ok = True
    ps, pd = parts_of(src), parts_of(dst)
    need = [n for n in ps if n == "xl/metadata.xml" or n.startswith("xl/richData/")]
    miss = [n for n in need if n not in pd]
    with zipfile.ZipFile(dst) as z:
        ct = z.read("[Content_Types].xml").decode("utf-8", "ignore")
        wbr = z.read("xl/_rels/workbook.xml.rels").decode("utf-8", "ignore")
    ct_miss = [p for p in need if p.endswith(".xml") and f'PartName="/{p}"' not in ct]
    rel_miss = []
    if "xl/metadata.xml" in need and "relationships/sheetMetadata" not in wbr:
        rel_miss.append("sheetMetadata")
    if "xl/richData/rdrichvalue.xml" in need and "relationships/rdRichValue" not in wbr:
        rel_miss.append("rdRichValue")
    if "xl/richData/rdrichvaluestructure.xml" in need and "relationships/rdRichValueStructure" not in wbr:
        rel_miss.append("rdRichValueStructure")
    vs, vd = sheet_vmcm(src), sheet_vmcm(dst)
    vm_diff = {nm: (vs.get(nm), vd.get(nm)) for nm in vs if vs.get(nm) != vd.get(nm)}
    bad_zip = zipfile.ZipFile(dst).testzip()
    for msg, cond in [("缺部件", miss), ("缺CT", ct_miss), ("缺rels", rel_miss),
                      ("vm/cm不一致", vm_diff), ("zip损坏", bad_zip)]:
        if cond:
            ok = False
            print(f"    [!] {tag} {msg}: {cond}")
    ns = len(parts_of(src))
    print(f"    {'OK ' if ok else 'FAIL'} {tag}  部件 {len(pd)}（源 {ns}）  富数据需要 {len(need)}  缺失 {len(miss)}  vm差异 {len(vm_diff)}")
    return ok


def main():
    apply = "--apply" in sys.argv
    books = sorted(f for f in os.listdir(SRC)
                   if f.endswith(".xlsx") and "xl/metadata.xml" in parts_of(os.path.join(SRC, f)))
    print(f"带富数据的簿 {len(books)} 本；模式 = {'APPLY(写回生产)' if apply else 'DRY(临时副本)'}")
    if apply:
        os.makedirs(BK, exist_ok=True)
    elif os.path.isdir(TMPDIR):
        shutil.rmtree(TMPDIR)
    if not apply:
        os.makedirs(TMPDIR, exist_ok=True)
    allok = True
    for f in books:
        src = os.path.join(SRC, f)
        dst = os.path.join(DST, f)
        if not os.path.exists(dst):
            print(f"    [skip] 生产缺 {f}")
            continue
        if apply:
            shutil.copy2(dst, os.path.join(BK, f))
            target = dst
        else:
            target = os.path.join(TMPDIR, f)
            shutil.copy2(dst, target)
        print(f"  == {f}")
        R.restore_parts(src, target)
        allok &= verify(src, target, "after")
        if not apply:
            # 干跑顺带对比：确认只多了富数据部件、没删别的
            a, b = parts_of(dst), parts_of(target)
            print(f"    新增部件: {sorted(b - a)}")
            print(f"    丢失部件: {sorted(a - b)}")
    if not apply:
        print(f"\n[干跑副本] {TMPDIR}")
    print("\n结论:", "全部 OK" if allok else "有 FAIL，见上")
    return 0 if allok else 1


if __name__ == "__main__":
    sys.exit(main())
