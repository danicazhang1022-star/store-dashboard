#!/usr/bin/env python3
import json
import subprocess

image_id = "KlCYcLj1CTUoMfMX9KVQUtc8MaSC1BwDRoSUpXBoCtjNxwjzQl7xgwgdWOl3qM4rlWGJZeTSHiMDQakYM6mpG4byEao4fMHAdJVmAcWf7Km68ZqgHYXL6FE7byVxJzRpAA8Cw2QTQetXl0HPG/z0i8J6eeD3n/zX3tWq4j+47JXFBYmr3PvjXnhpIk1TBH9/7qV68GA0ubumuy9A7vrr5QdunXcF7cYkbIaeJ/OGdkQB+vwk2auwWu7hqtsMXg=="
record_id = "ryQRrU"  # NS2614-L 的 record_id

print('[3/3] 更新记录，写入图片...')

# 测试数组格式
args = {
    "file_id": "DriTvwXLsoAA",
    "sheet_id": "t00i2h",
    "records": [{
        "record_id": record_id,
        "field_values": {
            "产品图片": image_id
        }
    }]
}

args_json = json.dumps(args, ensure_ascii=False)
cmd = f'''export PATH="/Users/danica/.workbuddy/binaries/node/versions/22.12.0/bin:$PATH" && mcporter call "tencent-docs" "smartsheet.update_records" --args '{args_json}' 2>&1'''

result = subprocess.run(cmd, shell=True, capture_output=True, text=True, timeout=120)
print(f'更新结果: {result.stdout[:800]}')

try:
    data = json.loads(result.stdout.strip())
    if data.get('error'):
        err = data['error']
        print(f'更新失败: {err}')
    else:
        print('更新成功！')
        print(f'返回记录数: {len(data.get("records", []))}')
except Exception as e:
    print(f'解析错误: {e}')
