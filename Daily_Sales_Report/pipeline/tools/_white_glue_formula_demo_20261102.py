# -*- coding: utf-8 -*-
"""白边液核查 · 公式联动教学示例 —— 2026-11-02

把脚本「读出来塞进去」的做法改成纯 Excel 公式：
  组装表 SKU(带ZH) → 商品资料 O列反查 → 主SKU → A列查 → 状态/评级/ASIN/型号/品名

生成 5 个页签：
  商品资料52    真实列位复刻（A/B/F/H/L/M/O），学完公式能直接套到真实文件
  组装表52      A=SKU(ZH) B=商品名称 L=白边液用量
  公式版清单    A列静态SKU，B-H 全是 XLOOKUP 公式（改商品资料状态会自动联动）
  脚本对照      同样的列但静态值，预览/无公式环境能看到结果
  公式说明      每个公式的拆解 + 进阶 LET/FILTER 一键生成
"""
import json, os, re
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

DSR = r"C:\Users\dell\Desktop\work-main\sale_file_gen\Daily_Sales_Report"
OUT_DIR = os.path.join(DSR, "白边液核查", "20261102")
OUT = os.path.join(OUT_DIR, "白边液_公式联动示例.xlsx")

H_FILL = PatternFill("solid", fgColor="1F4E79")
H_FONT = Font(bold=True, color="FFFFFF", size=10)
GREY = PatternFill("solid", fgColor="F2F2F2")
NOTE_FONT = Font(italic=True, size=9, color="808080")
LINK_FONT = Font(color="0563C1", underline="single", size=10)
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def hdr(ws, row, n):
    for c in range(1, n + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = H_FILL
        cell.font = H_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BOX


def main():
    data = json.load(open(r"C:\Users\dell\Desktop\work-main\sale_file_gen\.workbuddy\wg_demo.json", encoding="utf-8"))
    asm = data["asm"]            # [(组装SKU, 商品名称, 用量), ...]
    info = data["info"]          # {主SKU: {asin,score,cn,model,state,o}}
    o2main = data["o2main"]
    os.makedirs(OUT_DIR, exist_ok=True)

    wb = openpyxl.Workbook()

    # ============ 1. 商品资料52（复刻真实列位 A/B/F/H/L/M/O）
    ws = wb.active
    ws.title = "商品资料52"
    ws["A1"] = "商品资料（示例 · 52 个白边液 SKU，列位与真实文件一致）"
    ws["A1"].font = Font(bold=True, size=12)
    ws.merge_cells("A1:O1")
    # 表头第2行
    labels = {1: "SKU（主）", 2: "ASIN", 6: "评级\n(9月得分)", 7: "类型", 8: "中文品名",
              12: "型号", 13: "状态", 14: "镜头膜", 15: "SKU（第二·带ZH）"}
    for col, t in labels.items():
        ws.cell(row=2, column=col, value=t)
    # 省略列标注
    skip = {3: "参考成本价", 4: "手机价格", 5: "—", 9: "品牌", 10: "系列", 11: "供应商"}
    for col, t in skip.items():
        c = ws.cell(row=2, column=col, value=f"（{t}·省略）")
        c.font = NOTE_FONT
        c.fill = GREY
    hdr_row = 2
    # 只给有 label 的列加表头样式
    for col in [1, 2, 6, 7, 8, 12, 13, 14, 15]:
        c = ws.cell(row=2, column=col)
        c.fill = H_FILL; c.font = H_FONT
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BOX
    # 数据：按组装表顺序，取映射后的主SKU行
    r = 3
    for a_sku, a_name, qty in asm:
        main_sku = o2main.get(a_sku, a_sku)
        d = info.get(main_sku, {})
        ws.cell(row=r, column=1, value=main_sku)
        ws.cell(row=r, column=2, value=d.get("asin", "") if re.fullmatch(r"B0[A-Z0-9]{8}", d.get("asin", "")) else "#N/A")
        ws.cell(row=r, column=6, value=d.get("score", ""))
        ws.cell(row=r, column=8, value=d.get("cn", ""))
        ws.cell(row=r, column=12, value=d.get("model", ""))
        ws.cell(row=r, column=13, value=d.get("state", ""))
        ws.cell(row=r, column=15, value=d.get("o", ""))
        r += 1
    widths = {1: 10, 2: 14, 3: 12, 4: 10, 5: 6, 6: 10, 7: 10, 8: 30, 9: 10, 10: 8, 11: 10, 12: 24, 13: 8, 14: 8, 15: 12}
    for c, w in widths.items():
        ws.column_dimensions[get_column_letter(c)].width = w
    ws.freeze_panes = "A3"

    # ============ 2. 组装表52（A=SKU B=商品名称 L=白边液）
    ws2 = wb.create_sheet("组装表52")
    ws2["A1"] = "钢化膜组装表（示例 · 白边液>0 的 52 行，列位与真实文件一致）"
    ws2["A1"].font = Font(bold=True, size=12)
    ws2.merge_cells("A1:L1")
    h2 = {1: "SKU（带ZH）", 2: "商品名称", 12: "白边液\n用量"}
    for col, t in h2.items():
        c = ws2.cell(row=2, column=col, value=t)
        c.fill = H_FILL; c.font = H_FONT
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BOX
    for col in [3, 4, 5, 6, 7, 8, 9, 10, 11]:
        c = ws2.cell(row=2, column=col, value="（省略）")
        c.font = NOTE_FONT; c.fill = GREY
    r = 3
    for a_sku, a_name, qty in asm:
        ws2.cell(row=r, column=1, value=a_sku)
        ws2.cell(row=r, column=2, value=a_name)
        ws2.cell(row=r, column=12, value=qty)
        r += 1
    for c, w in {1: 12, 2: 36, 3: 8, 4: 8, 5: 8, 6: 8, 7: 8, 8: 8, 9: 8, 10: 8, 11: 8, 12: 8}.items():
        ws2.column_dimensions[get_column_letter(c)].width = w
    ws2.freeze_panes = "A3"

    # ============ 3. 公式版清单（A列静态SKU，B-H 全公式）
    ws3 = wb.create_sheet("公式版清单")
    ws3["A1"] = "白边液清单 · 公式版（B-H 列全是公式，改「商品资料52」的状态会自动联动）"
    ws3["A1"].font = Font(bold=True, size=12)
    ws3.merge_cells("A1:H1")
    ws3["A2"] = "A 列 = 组装表 SKU（静态，手动粘进来）；B-H 列 = XLOOKUP 公式。选中任意单元格看编辑栏即公式。"
    ws3["A2"].font = NOTE_FONT
    ws3.merge_cells("A2:H2")
    h3 = ["组装表SKU\n(带ZH)", "主SKU\n(公式)", "状态\n(公式)", "评级\n(公式)",
          "ASIN\n(公式)", "型号\n(公式)", "中文品名\n(公式)", "白边液用量\n(公式)"]
    ws3.append([])
    ws3.append(h3)
    hdr(ws3, 4, len(h3))
    r = 5
    for a_sku, a_name, qty in asm:
        ws3.cell(row=r, column=1, value=a_sku)
        # B 主SKU：拿组装表SKU去商品资料O列(第二SKU)找，返回A列主SKU；找不到退回自身
        ws3.cell(row=r, column=2, value=f"=XLOOKUP(A{r},商品资料52!$O:$O,商品资料52!$A:$A,A{r})")
        # C 状态：用主SKU去A列查M列
        ws3.cell(row=r, column=3, value=f'=XLOOKUP($B{r},商品资料52!$A:$A,商品资料52!$M:$M,"查无")')
        # D 评级：F列
        ws3.cell(row=r, column=4, value=f'=XLOOKUP($B{r},商品资料52!$A:$A,商品资料52!$F:$F,"查无")')
        # E ASIN：B列
        ws3.cell(row=r, column=5, value=f'=XLOOKUP($B{r},商品资料52!$A:$A,商品资料52!$B:$B,"查无")')
        # F 型号：L列
        ws3.cell(row=r, column=6, value=f'=XLOOKUP($B{r},商品资料52!$A:$A,商品资料52!$L:$L,"查无")')
        # G 中文品名：H列
        ws3.cell(row=r, column=7, value=f'=XLOOKUP($B{r},商品资料52!$A:$A,商品资料52!$H:$H,"查无")')
        # H 白边液用量：从组装表L列查
        ws3.cell(row=r, column=8, value=f'=XLOOKUP($A{r},组装表52!$A:$A,组装表52!$L:$L,0)')
        r += 1
    for c, w in {1: 12, 2: 10, 3: 8, 4: 8, 5: 14, 6: 24, 7: 30, 8: 10}.items():
        ws3.column_dimensions[get_column_letter(c)].width = w
    ws3.freeze_panes = "A5"

    # ============ 4. 脚本对照（静态值，预览能看到结果）
    ws4 = wb.create_sheet("脚本对照")
    ws4["A1"] = "脚本对照（同样的列，但是脚本算好的静态值 —— 预览/无公式环境看这里）"
    ws4["A1"].font = Font(bold=True, size=12)
    ws4.merge_cells("A1:H1")
    ws4.append([])
    ws4.append(h3)
    hdr(ws4, 3, len(h3))
    r = 4
    for a_sku, a_name, qty in asm:
        main_sku = o2main.get(a_sku, a_sku)
        d = info.get(main_sku, {})
        asin = d.get("asin", "")
        asin = asin if re.fullmatch(r"B0[A-Z0-9]{8}", asin) else "#N/A"
        ws4.cell(row=r, column=1, value=a_sku)
        ws4.cell(row=r, column=2, value=main_sku)
        ws4.cell(row=r, column=3, value=d.get("state", ""))
        ws4.cell(row=r, column=4, value=d.get("score", "") or "—")
        ws4.cell(row=r, column=5, value=asin)
        ws4.cell(row=r, column=6, value=d.get("model", ""))
        ws4.cell(row=r, column=7, value=d.get("cn", ""))
        ws4.cell(row=r, column=8, value=qty)
        r += 1
    for c, w in {1: 12, 2: 10, 3: 8, 4: 8, 5: 14, 6: 24, 7: 30, 8: 10}.items():
        ws4.column_dimensions[get_column_letter(c)].width = w
    ws4.freeze_panes = "A4"

    # ============ 5. 公式说明
    ws5 = wb.create_sheet("公式说明")
    L = [
        ("白边液清单 · 公式联动拆解", Font(bold=True, size=14)),
        ("", None),
        ("一、为什么要「嵌套两步」", Font(bold=True, size=11, color="1F4E79")),
        ("钢化膜组装表里的 SKU 带 ZH 后缀（如 B172ZH），商品资料里也存了一份在 O 列（第二 SKU）。", None),
        ("而商品资料的状态/评级/ASIN 都挂在 A 列（主 SKU，如 G172）上。", None),
        ("所以不能一步查到，得先「ZH → 主SKU」再「主SKU → 状态」，这就是两步 XLOOKUP。", None),
        ("", None),
        ("二、每列公式（行号以第 5 行为例，往下拖）", Font(bold=True, size=11, color="1F4E79")),
        ("B 列·主SKU：=XLOOKUP(A5, 商品资料52!$O:$O, 商品资料52!$A:$A, A5)", None),
        ("    拿组装表 SKU 去 O 列找，找到返回对应 A 列主 SKU；找不到（说明本来就是主 SKU）退回 A5 自身。", Font(size=9, color="808080")),
        ("C 列·状态：=XLOOKUP($B5, 商品资料52!$A:$A, 商品资料52!$M:$M, \"查无\")", None),
        ("    用主 SKU 在 A 列定位，返回 M 列（状态）。$B5 锁列不锁行，方便横拖。", Font(size=9, color="808080")),
        ("D 列·评级：=XLOOKUP($B5, 商品资料52!$A:$A, 商品资料52!$F:$F, \"查无\")    ← F 列是评级", None),
        ("E 列·ASIN：=XLOOKUP($B5, 商品资料52!$A:$A, 商品资料52!$B:$B, \"查无\")     ← B 列是 ASIN", None),
        ("F 列·型号：=XLOOKUP($B5, 商品资料52!$A:$A, 商品资料52!$L:$L, \"查无\")     ← L 列是型号", None),
        ("G 列·品名：=XLOOKUP($B5, 商品资料52!$A:$A, 商品资料52!$H:$H, \"查无\")     ← H 列是中文品名", None),
        ("H 列·用量：=XLOOKUP($A5, 组装表52!$A:$A, 组装表52!$L:$L, 0)               ← 从组装表查白边液用量", None),
        ("", None),
        ("三、XLOOKUP 四个参数速记", Font(bold=True, size=11, color="1F4E79")),
        ("=XLOOKUP(找什么, 在哪一列找, 返回哪一列, 找不到时给什么)", None),
        ("对比 VLOOKUP：不用数第几列、不用 $ 锁起始列、默认精确匹配、找不到不会报 #N/A。", None),
        ("WPS 和 Excel 365 / 2021 都支持。老版 Excel 没有 XLOOKUP 就用 INDEX+MATCH：", None),
        ("    =INDEX(商品资料52!$M:$M, MATCH($B5, 商品资料52!$A:$A, 0))", Font(size=9, color="808080")),
        ("", None),
        ("四、进阶：一个公式生成整张表（需 Excel 365 / WPS 最新版，支持 FILTER+LET+LAMBDA）", Font(bold=True, size=11, color="1F4E79")),
        ("=LET(", None),
        ("  asm, FILTER(组装表52!A2:A500, 组装表52!L2:L500>0),", None),
        ("  main, MAP(asm, LAMBDA(x, XLOOKUP(x, 商品资料52!O:O, 商品资料52!A:A, x))),", None),
        ("  HSTACK(asm, main,", None),
        ("    MAP(main, LAMBDA(x, XLOOKUP(x, 商品资料52!A:A, 商品资料52!M:M))),", None),
        ("    MAP(main, LAMBDA(x, XLOOKUP(x, 商品资料52!A:A, 商品资料52!F:F))))", None),
        (")", None),
        ("把这段贴到一个单元格里，会自动 spill 成一张四列的表（组装SKU / 主SKU / 状态 / 评级）。", Font(size=9, color="808080")),
        ("FILTER 负责筛白边液>0，MAP 对每个 SKU 做映射，HSTACK 把列拼起来。", Font(size=9, color="808080")),
        ("", None),
        ("五、和脚本做法的关系", Font(bold=True, size=11, color="1F4E79")),
        ("脚本（_white_glue_check_20261030.py 的 load_white）做的事和上面公式一模一样：", None),
        ("  1. 读组装表筛 L>0 → asm 列表", None),
        ("  2. 建 o2main 字典（O列→A列）= 公式里的第一步 XLOOKUP", None),
        ("  3. 用主 SKU 取 info[sku]['state'] 等 = 公式里的第二步 XLOOKUP", None),
        ("区别：脚本把值「算死」写进单元格；公式保留「引用关系」，源表一改就联动。", None),
        ("做交付物（要发出去、要锁定口径）用脚本；做排查清单（要随源表更新）用公式。", None),
    ]
    for i, (t, f) in enumerate(L, 1):
        c = ws5.cell(row=i, column=1, value=t)
        if f:
            c.font = f
    ws5.column_dimensions["A"].width = 105

    wb.save(OUT)
    print("已生成：", OUT)


if __name__ == "__main__":
    main()
