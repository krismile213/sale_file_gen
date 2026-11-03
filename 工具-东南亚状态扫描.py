# -*- coding: utf-8 -*-
"""
扫描东南亚所有订单文件的「订单状态」实际取值分布，
并与脚本 month_sale_file.py 里的 status_map 对撞，列出：
  A. status_map 有、文件里没有（空配置，无害）
  B. 文件里有、status_map 没有（会变 NaN → 被静默丢弃，需警惕）
  C. 文件里有、status_map 里映射为 Cancelled（会被剔除，正常）
"""
import os
import sys
import pandas as pd
import chardet

sys.stdout.reconfigure(encoding='utf-8')
pd.set_option('display.unicode.east_asian_width', True)

BASE = r'D:\Desktop\work-main\sale_file_gen\trunk\data\数据归属周期\202608\快照日期\20260907\东南亚'

# ---- 脚本里现有的 status_map（逐站抄录，保持与代码一字一致）----
SHOPEE_MAP = {
    ('泰国', '本土'): {"สำเร็จแล้ว": "Completed", "ยกเลิกแล้ว": "Cancelled"},
    ('印尼', '本土'): {"Selesai": "Completed", "Batal": "Cancelled",
                       "Sedang Dikirim": "Shipping", "Telah Dikirim": "Delivered"},
    ('菲律宾', '本土'): {"Completed": "Completed", "Cancelled": "Cancelled",
                        "Shipping": "Shipping", "Delivered": "Delivered"},
    ('新加坡', '跨境'): {"Completed": "Completed", "Cancelled": "Cancelled",
                        "Shipping": "Shipping", "Delivered": "Delivered"},
    ('马来西亚', '跨境'): {"Completed": "Completed", "Cancelled": "Cancelled",
                          "Shipping": "Shipping", "Delivered": "Delivered"},
    ('越南', '跨境'): {"Hoàn thành": "Completed", "Đã hủy": "Cancelled", "Đã giao": "Delivered"},
}
LAZADA_MAP = {
    "confirmed": "confirmed", "shipped": "shipped", "delivered": "delivered",
    "canceled": "canceled", "returned": "returned",
    "Package Returned": "Package Returned",
    "In Transit: Returning to seller": "In Transit: Returning to seller",
}
TIKTOK_MAP = {"Completed": "Completed", "Shipped": "Shipped", "Canceled": "Canceled"}

# 退货窗口前缀（脚本 343-349 行）
RW_PREFIX = {
    "泰国": ("ผู้ซื้อได้รับสินค้าแล้ว", "สำเร็จแล้ว"),
    "印尼": ("Pesanan diterima", "Selesai"),
    "菲律宾": ("Order Received, but buyer can raise return/refund", "Completed"),
    "马来西亚": ("Order Received, but buyer can raise return/refund", "Completed"),
    "新加坡": ("Order Received, but buyer can raise return/refund", "Completed"),
}


def read_any(path):
    if path.lower().endswith('.csv'):
        with open(path, 'rb') as f:
            enc = chardet.detect(f.read())['encoding']
        return pd.read_csv(path, encoding=enc, dtype=str), enc
    return pd.read_excel(path, dtype=str), 'xlsx'


def scan(platform, filename, subdir, status_col, smap, site, entity):
    path = os.path.join(BASE, subdir, filename) if subdir else os.path.join(BASE, filename)
    if not os.path.exists(path):
        print(f'  ✘ 文件不存在：{path}')
        return
    df, enc = read_any(path)
    if status_col not in df.columns:
        print(f'  ✘ 找不到状态列「{status_col}」，实际列：{list(df.columns)[:8]}')
        return
    raw = df[status_col].astype(str)
    print(f'\n--- {platform} · {site} · {entity}   [{filename}]')
    print(f'    编码={enc}  行数={len(df)}  状态列=「{status_col}」')

    # A/B/C 分类
    cnt = raw.value_counts()
    rw_prefix = RW_PREFIX.get(site, (None, None))[0]
    for st, n in cnt.items():
        in_map = st in smap
        is_rw = rw_prefix and st.startswith(rw_prefix)
        if in_map:
            mapped = smap[st]
            flag = 'C 剔除(Cancelled)' if mapped == 'Cancelled' else f'✔保留→{mapped}'
        elif is_rw:
            mapped = RW_PREFIX[site][1]
            flag = f'✔保留→{mapped}  (退货窗口前缀归一)'
        else:
            flag = '⚠ B 未覆盖！→ NaN 被静默丢弃'
        print(f'      {n:6d} 行  {flag:38s}  「{st[:70]}」')

    # 反向：map 里有、文件里没有
    missing = [k for k in smap if k not in cnt.index]
    if missing:
        print(f'      · map 中配置了但文件里未出现（无害）：{missing}')


def main():
    print('=' * 110)
    print('【一】Shopee 6 站（本土 3 + 跨境 3）')
    print('=' * 110)
    for (site, entity), smap in SHOPEE_MAP.items():
        if entity == '本土':
            fn, sub = f'Shopee{site}站点202608订单报告.xlsx', ''
        else:
            fn, sub = f'Shopee跨境{site}站点202608订单报告.xlsx', '跨境'
        # 状态列名逐站不同
        col = {'泰国': 'สถานะการสั่งซื้อ', '印尼': 'Status Pesanan', '越南': '订单状态'}.get(site, 'Order Status')
        scan('Shopee', fn, sub, col, smap, site, entity)

    print('\n' + '=' * 110)
    print('【二】Lazada 7 站（本土 2 + 跨境 5）')
    print('=' * 110)
    for site, entity in [('菲律宾', '本土'), ('泰国', '本土'), ('新加坡', '跨境'),
                         ('泰国', '跨境'), ('菲律宾', '跨境'), ('越南', '跨境'), ('马来西亚', '跨境')]:
        fn = f'Lazada{site}站点202608订单报告.xlsx' if entity == '本土' else f'Lazada跨境{site}站点202608订单报告.xlsx'
        sub = '' if entity == '本土' else '跨境'
        scan('Lazada', fn, sub, 'status', LAZADA_MAP, site, entity)

    print('\n' + '=' * 110)
    print('【三】Tiktok 2 站（本土）')
    print('=' * 110)
    for site in ['印尼', '菲律宾']:
        scan('Tiktok', f'Tiktok{site}站点202608订单报告.csv', '', 'Order Status', TIKTOK_MAP, site, '本土')

    print('\n' + '=' * 110)
    print('【四】其它东南亚数据源（非 Shopee/Lazada/Tiktok 主链路，仅登记不展开）')
    print('=' * 110)
    for fn in ['印尼海外仓202608订单报告.xlsx', '印尼海外仓202608商品出库报告.xlsx',
               '菲律宾KSON仓202608线下订单报告.xlsx']:
        p = os.path.join(BASE, fn)
        if os.path.exists(p):
            try:
                df = pd.read_excel(p, dtype=str)
                print(f'  {fn}  行数={len(df)}  列={list(df.columns)[:10]}')
            except Exception as e:
                print(f'  {fn}  读取失败 {e}')


if __name__ == '__main__':
    main()
