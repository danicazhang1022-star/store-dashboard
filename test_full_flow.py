#!/usr/bin/env python3
"""完整流程测试：上传 + 立即写入 + 验证"""
import base64
import json
import subprocess
import time

def mcporter_call(tool, args_dict):
    args_json = json.dumps(args_dict, ensure_ascii=False)
    cmd = f'''export PATH="/Users/danica/.workbuddy/binaries/node/versions/22.12.0/bin:$PATH" && mcporter call "tencent-docs" "{tool}" --args '{args_json}' 2>&1'''
    result = subprocess.run(cmd, shell=True, capture_output=True, text=True, timeout=120)
    try:
        return json.loads(result.stdout.strip())
    except:
        print(f"Parse error: {result.stdout[:200]}")
        return None

print('='*60)
print('完整流程测试：上传 + 立即写入')
print('='*60)

# 1. 上传图片
print('[1/3] 上传图片...')
with open('产品图片/NS2613-S.jpg', 'rb') as f:
    b64 = base64.b64encode(f.read()).decode('utf-8')

data = mcporter_call("upload_image", {"file_name": "NS2613-S.jpg", "image_base64": b64})
if not data or not data.get('image_id'):
    print('上传失败!')
    exit(1)

image_id = data['image_id']
print(f'获取 image_id: {image_id[:50]}...')

# 2. 立即写入（不等待）
print('[2/3] 立即写入记录...')
record_id = "rSn5SD"  # NS2613-S

update_data = mcporter_call("smartsheet.update_records", {
    "file_id": "DriTvwXLsoAA",
    "sheet_id": "t00i2h",
    "records": [{
        "record_id": record_id,
        "field_values": {"产品图片": image_id}  # 纯字符串格式
    }]
})

if update_data and not update_data.get('error'):
    print('写入成功!')
else:
    err = update_data.get('error') if update_data else 'API失败'
    print(f'写入失败: {err}')

# 3. 验证
print('[3/3] 验证写入结果...')
time.sleep(2)
verify_data = mcporter_call("smartsheet.list_records", {
    "file_id": "DriTvwXLsoAA",
    "sheet_id": "t00i2h",
    "record_ids": [record_id]
})

if verify_data and verify_data.get('records'):
    fv = verify_data['records'][0]['field_values']
    if '产品图片' in fv:
        img_val = fv['产品图片']
        print(f'验证成功! 产品图片字段存在: {img_val}')
    else:
        print('验证失败: 产品图片字段不存在')
        print(f'可用字段: {list(fv.keys())}')
else:
    print('验证查询失败')
