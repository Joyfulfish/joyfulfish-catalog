#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
魚樂匯庫存更新腳本
使用方法：python3 update_inventory.py

在庫標記說明：
- V = 正常在庫販售
- S = 售完展示（顯示商品和圖片，但隱藏價格和訂購按鈕）
- 空白或其他 = 完全不顯示
"""

import pandas as pd
import re
import sys
import json
from pathlib import Path

# 設定檔案路徑（Excel檔案現在和腳本在同一目錄）
EXCEL_FILE = Path(__file__).parent / "庫存明細（魚樂匯）.xlsx"
HTML_FILE = Path(__file__).parent / "index.html"
SPECIAL_PRODUCTS_JSON = Path(__file__).parent / "../../特價選單用/網站檔案/products.json"

def load_special_products():
    """讀取特價商品資料"""
    try:
        if SPECIAL_PRODUCTS_JSON.exists():
            with open(SPECIAL_PRODUCTS_JSON, 'r', encoding='utf-8') as f:
                products = json.load(f)
            # 建立商品名稱到tiers的映射
            return {p['name']: p.get('tiers', []) for p in products}
        else:
            print(f"ℹ️  特價商品資料不存在：{SPECIAL_PRODUCTS_JSON}")
            return {}
    except Exception as e:
        print(f"⚠️  讀取特價商品資料失敗：{e}")
        return {}

def excel_to_js_array(excel_path, special_tiers_map):
    """讀取 Excel 並轉換為 JavaScript 陣列格式"""
    try:
        df = pd.read_excel(excel_path)
        print(f"✓ 成功讀取 Excel 檔案：{len(df)} 筆資料")
    except Exception as e:
        print(f"✗ 讀取 Excel 失敗：{e}")
        sys.exit(1)

    # 生成 JavaScript 資料
    js_lines = ['const RAW = [']

    # 檢查是否有「在庫」欄位
    has_in_stock_column = '在庫' in df.columns

    data_count = 0
    sold_out_count = 0
    for i, row in df.iterrows():
        # 如果有「在庫」欄位，處理標註 V 或 S 的商品
        if has_in_stock_column:
            in_stock = str(row['在庫']).strip().upper()
            if in_stock not in ['V', 'S']:
                continue  # 跳過沒有標註 V 或 S 的商品

        cat = str(row['分類']).replace('"', '\\"')
        name = str(row['品名']).replace('"', '\\"')
        size = str(row['尺寸']).replace('"', '\\"')
        stock = int(row['庫存數']) if pd.notna(row['庫存數']) else 0
        price = int(row['零售價']) if pd.notna(row['零售價']) else 0

        # 處理備註欄位
        note = ''
        if '備註' in df.columns and pd.notna(row['備註']):
            note = str(row['備註']).replace('"', '\\"').replace('\n', ' ')

        # S 標記代表售完展示模式：顯示商品但隱藏價格
        sold_out_display = 1 if in_stock == 'S' else 0

        # 檢查是否有特價階梯定價
        tiers_json = 'null'
        if name in special_tiers_map and len(special_tiers_map[name]) > 1:
            # 只在有多個階梯時才附加tiers數據
            tiers = special_tiers_map[name]
            tiers_json = json.dumps(tiers, ensure_ascii=False)

        if data_count > 0:
            js_lines.append(',')
        line = f'  {{cat:"{cat}",name:"{name}",size:"{size}",stock:{stock},price:{price},soldOut:{sold_out_display},note:"{note}",tiers:{tiers_json}}}'
        js_lines.append(line)
        data_count += 1
        if sold_out_display:
            sold_out_count += 1

    js_lines.append('];')

    return '\n'.join(js_lines), data_count, sold_out_count

def update_html(html_path, new_data):
    """更新 HTML 檔案中的資料"""
    try:
        with open(html_path, 'r', encoding='utf-8') as f:
            content = f.read()
        print(f"✓ 成功讀取 HTML 檔案")
    except Exception as e:
        print(f"✗ 讀取 HTML 失敗：{e}")
        sys.exit(1)

    # 使用正則表達式找到並替換 RAW 資料
    pattern = r'const RAW = \[.*?\];'

    if not re.search(pattern, content, re.DOTALL):
        print("✗ 找不到 RAW 資料區塊")
        sys.exit(1)

    # 替換資料
    new_content = re.sub(pattern, new_data, content, flags=re.DOTALL)

    # 寫回檔案
    try:
        with open(html_path, 'w', encoding='utf-8') as f:
            f.write(new_content)
        print(f"✓ 成功更新 HTML 檔案")
    except Exception as e:
        print(f"✗ 寫入 HTML 失敗：{e}")
        sys.exit(1)

def main():
    print("=" * 50)
    print("魚樂匯庫存資料更新工具")
    print("=" * 50)
    print()

    # 檢查檔案是否存在
    if not EXCEL_FILE.exists():
        print(f"✗ Excel 檔案不存在：{EXCEL_FILE}")
        print(f"  請確認路徑是否正確")
        sys.exit(1)

    if not HTML_FILE.exists():
        print(f"✗ HTML 檔案不存在：{HTML_FILE}")
        sys.exit(1)

    print(f"Excel 檔案：{EXCEL_FILE}")
    print(f"HTML 檔案：{HTML_FILE}")
    print()

    # 載入特價商品階梯定價
    print("載入特價商品資料...")
    special_tiers_map = load_special_products()
    if special_tiers_map:
        print(f"✓ 已載入 {len(special_tiers_map)} 個特價商品的階梯定價")

    # 轉換資料
    print("開始處理...")
    js_data, count, sold_out_count = excel_to_js_array(EXCEL_FILE, special_tiers_map)

    # 更新 HTML
    update_html(HTML_FILE, js_data)

    print()
    print("=" * 50)
    print(f"✓ 更新完成！共 {count} 筆商品資料")
    if sold_out_count > 0:
        print(f"  - 正常販售: {count - sold_out_count} 筆")
        print(f"  - 售完展示: {sold_out_count} 筆（顯示商品但隱藏價格）")
    print("=" * 50)
    print()
    print("請重新整理網頁查看更新結果。")

if __name__ == "__main__":
    main()
