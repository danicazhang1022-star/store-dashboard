#!/usr/bin/env python3
"""测试图片上传 - 使用正确的数组格式"""
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
        return None

print('='*60)
print('测试第三张图片: NS2614-S.jpg (使用数组格式)')
print('='*60)

# 1. 上传图片
with open('产品图片/NS2614-S.jpg', 'rb') as f:
    b64 = base64.b64encode(f.read()).decode('utf-8')

print(f'[1/3] 图片base64大小: {len(b64)} bytes')
print('[2/3] 上传图片...')

data = mcporter_call("upload_image", {"file_name": "NS2614-S.jpg", "image_base64": b64})
if not data or not data.get('image_id'):
    print('上传失败!')
    exit(1)

image_id = data['image_id']
print(f'获取 image_id: {image_id[:40]}...')

# 2. 更新记录（使用数组格式 - 文档规定的正确格式）
print('[3/3] 更新记录（数组格式）...')
time.sleep(1)

# NS2614-S 的 record_id 是 rq8qse
data = mcporter_call("smartsheet.update_records", {
    "file_id": "DriTvwXLsoAA",
    "sheet_id": "t00i2h",
    "records": [{
        "record_id": "rq8qse",
        "field_values": {"产品图片": [{"image_id": image_id}]}
    }]
})

if data:
    err = data.get('error')
    if err:
        print(f'失败: {err}')
    else:
        tid = data.get('trace_id', 'N/A')
        print(f'成功! trace_id: {tid}')
else:
    print('API调用失败')
