#!/usr/bin/env python3
"""测试第二张图片上传"""
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

# 测试第二张图片
print('='*60)
print('测试第二张图片: NS2614-M.jpg')
print('='*60)

# 1. 上传图片
with open('产品图片/NS2614-M.jpg', 'rb') as f:
    b64 = base64.b64encode(f.read()).decode('utf-8')

print(f'[1/3] 图片base64大小: {len(b64)} bytes')
print('[2/3] 上传图片...')

data = mcporter_call("upload_image", {"file_name": "NS2614-M.jpg", "image_base64": b64})
if not data or not data.get('image_id'):
    print('上传失败!')
    exit(1)

image_id = data['image_id']
print(f'获取 image_id: {image_id[:30]}...')

# 2. 更新记录（使用纯字符串格式）
print('[3/3] 更新记录...')
time.sleep(1)

# NS2614-M 的 record_id 是 r8pgY9
data = mcporter_call("smartsheet.update_records", {
    "file_id": "DriTvwXLsoAA",
    "sheet_id": "t00i2h",
    "records": [{
        "record_id": "r8pgY9",
        "field_values": {"产品图片": image_id}
    }]
})

if data and not data.get('error'):
    print('✅ 测试成功！图片已写入')
else:
    err = data.get('error', '未知错误') if data else 'API调用失败'
    print(f'❌ 失败: {err}')
