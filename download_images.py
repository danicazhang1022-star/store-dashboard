#!/usr/bin/env python3
"""批量下载聚水潭产品图片到本地"""
import pandas as pd
import requests
import os
import time
import warnings
warnings.filterwarnings('ignore')

SOURCE_FILE = '/Users/danica/Downloads/店铺商品资料_20260407004004_152895639_1.xlsx'
OUTPUT_DIR = '/Users/danica/WorkBuddy/20260407210337/产品图片'

os.makedirs(OUTPUT_DIR, exist_ok=True)

df = pd.read_excel(SOURCE_FILE)
print(f'共 {len(df)} 条SKU，开始下载图片...\n')

success = 0
fail = 0
no_url = 0
results = []

for idx, row in df.iterrows():
    sku = str(row['原始商品编码']) if pd.notna(row['原始商品编码']) else f'sku_{idx}'
    img_url = str(row['图片']) if pd.notna(row['图片']) else ''
    
    if not img_url.startswith('http'):
        no_url += 1
        results.append({'sku': sku, 'filename': '', 'status': '无图片链接'})
        continue
    
    # 清理文件名中的非法字符
    safe_sku = sku.replace('/', '_').replace('\\', '_').replace(' ', '_')
    ext = '.jpg'
    filename = f'{safe_sku}{ext}'
    filepath = os.path.join(OUTPUT_DIR, filename)
    
    # 如果已存在就跳过
    if os.path.exists(filepath) and os.path.getsize(filepath) > 500:
        success += 1
        results.append({'sku': sku, 'filename': filename, 'status': '已存在'})
        continue
    
    try:
        headers = {
            'User-Agent': 'Mozilla/5.0 (iPhone; CPU iPhone OS 16_0 like Mac OS X)',
            'Referer': 'https://juishitan.com/'
        }
        resp = requests.get(img_url, headers=headers, timeout=10)
        if resp.status_code == 200 and len(resp.content) > 500:
            with open(filepath, 'wb') as f:
                f.write(resp.content)
            success += 1
            results.append({'sku': sku, 'filename': filename, 'status': '✅成功'})
        else:
            fail += 1
            results.append({'sku': sku, 'filename': '', 'status': f'下载失败(HTTP {resp.status_code})'})
    except Exception as e:
        fail += 1
        results.append({'sku': sku, 'filename': '', 'status': f'错误: {str(e)[:30]}'})
    
    # 避免请求太快
    time.sleep(0.05)
    
    if (idx + 1) % 50 == 0:
        print(f'  进度: {idx+1}/{len(df)} (成功{success} 失败{fail})')

print(f'\n═══ 下载完成 ═══')
print(f'成功: {success}  失败: {fail}  无链接: {no_url}')

# 保存结果映射表
result_df = pd.DataFrame(results)
result_df.to_csv('/Users/danica/WorkBuddy/20260407210337/图片下载结果.csv', index=False, encoding='utf-8-sig')
print(f'映射表已保存: 图片下载结果.csv')

# 列出下载的文件
files = [f for f in os.listdir(OUTPUT_DIR) if f.endswith('.jpg')]
print(f'图片文件夹中共 {len(files)} 张图片: {OUTPUT_DIR}')
