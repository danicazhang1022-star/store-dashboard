#!/usr/bin/env python3
"""
Upload product images to Tencent Docs Smartsheet - Product Library.
Final version: plain string image_id format, fixed pagination, 3s delay.
"""

import subprocess
import json
import sys
import os
import time
import base64
import csv
import glob

FILE_ID = "DriTvwXLsoAA"
SHEET_ID = "t00i2h"
IMAGE_FIELD_TITLE = "产品图片"
SKU_FIELD_TITLE = "货号SKU"
IMAGE_DIR = "/Users/danica/WorkBuddy/20260407210337/产品图片"
CSV_PATH = "/Users/danica/WorkBuddy/20260407210337/图片下载结果.csv"
DELAY = 3  # seconds between each upload+update
MAX_CONSECUTIVE_ERRORS = 3
COOLDOWN = 120


def mcporter_call(tool, args_dict):
    """Call mcporter tool and return parsed result dict."""
    args_json = json.dumps(args_dict, ensure_ascii=False)
    cmd = f'''export PATH="/Users/danica/.workbuddy/binaries/node/versions/22.12.0/bin:$PATH" && mcporter call "tencent-docs" "{tool}" --args '{args_json}' 2>/dev/null'''
    result = subprocess.run(cmd, shell=True, capture_output=True, text=True, timeout=120)
    try:
        data = json.loads(result.stdout.strip())
        if data.get("error"):
            print(f"  API ERROR: {data['error']}", file=sys.stderr)
            return None
        return data
    except Exception as e:
        print(f"  Parse error: {e}", file=sys.stderr)
        return None


def upload_image(file_path):
    """Upload a single image and return image_id."""
    file_name = os.path.basename(file_path)
    with open(file_path, 'rb') as f:
        b64 = base64.b64encode(f.read()).decode('utf-8')
    data = mcporter_call("upload_image", {
        "file_name": file_name,
        "image_base64": b64
    })
    if data and data.get("image_id"):
        return data["image_id"]
    return None


def get_all_records():
    """Get all records using fixed-offset pagination."""
    records = []
    sku_set = set()
    for offset in range(0, 500, 100):
        data = mcporter_call("smartsheet.list_records", {
            "file_id": FILE_ID,
            "sheet_id": SHEET_ID,
            "field_titles": [SKU_FIELD_TITLE],
            "offset": offset,
            "limit": 100
        })
        if not data:
            break
        for r in data.get("records", []):
            fv = r.get("field_values", {})
            sku_field = fv.get(SKU_FIELD_TITLE, [{}])
            sku = sku_field[0].get("text", "") if sku_field else ""
            if sku and sku not in sku_set and sku != "nan":
                sku_set.add(sku)
                # Check if has image (look at any field that starts with image content)
                has_image = False
                for k, v in fv.items():
                    if k == IMAGE_FIELD_TITLE and v:
                        has_image = True
                records.append({"record_id": r["record_id"], "sku": sku, "has_image": has_image})
        if len(data.get("records", [])) < 100:
            break
        time.sleep(1)
    return records


def update_record_image(record_id, image_id):
    """Update record with image_id in array format: [{"image_id": "xxx"}]."""
    data = mcporter_call("smartsheet.update_records", {
        "file_id": FILE_ID,
        "sheet_id": SHEET_ID,
        "records": [{
            "record_id": record_id,
            "field_values": {
                IMAGE_FIELD_TITLE: [{"image_id": image_id}]  # MUST be array format!
            }
        }]
    })
    return data is not None


def build_csv_mapping():
    """Build filename -> SKU mapping from CSV."""
    mapping = {}
    with open(CSV_PATH, "r", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        for row in reader:
            mapping[row["filename"]] = row["sku"]
    return mapping


def main():
    print("=" * 60)
    print("产品图片上传工具 v3 - 腾讯文档智能表格")
    print("=" * 60)
    
    # Step 1: Get all records
    print("\n[1/3] 获取产品库全部记录...")
    records = get_all_records()
    if not records:
        print("ERROR: 无法获取记录！")
        sys.exit(1)
    
    sku_to_record = {}
    for r in records:
        sku_to_record[r["sku"]] = {"record_id": r["record_id"], "has_image": r["has_image"]}
    
    print(f"  共 {len(sku_to_record)} 条唯一SKU记录")
    
    # Step 2: Load CSV and image files
    print("\n[2/3] 扫描图片文件...")
    csv_map = build_csv_mapping()
    image_files = glob.glob(os.path.join(IMAGE_DIR, "*.jpg"))
    print(f"  {len(image_files)} 张图片, {len(csv_map)} 条CSV映射")
    
    # Step 3: Upload
    print("\n[3/3] 开始上传...")
    matched = 0
    skipped_no_record = 0
    skipped_bad_sku = 0
    uploaded_ok = 0
    uploaded_fail = 0
    consecutive_errors = 0
    
    for i, img_path in enumerate(image_files):
        fname = os.path.basename(img_path)
        sku = csv_map.get(fname, "")
        
        if not sku or sku.startswith("sku_"):
            skipped_bad_sku += 1
            continue
        
        if sku not in sku_to_record:
            skipped_no_record += 1
            continue
        
        matched += 1
        
        # Upload
        print(f"  [{i+1}/{len(image_files)}] {sku}...", end=" ", flush=True)
        image_id = upload_image(img_path)
        
        if not image_id:
            uploaded_fail += 1
            consecutive_errors += 1
            print("上传失败")
            if consecutive_errors >= MAX_CONSECUTIVE_ERRORS:
                print(f"\n  ⚠ 连续失败 {MAX_CONSECUTIVE_ERRORS} 次，等 {COOLDOWN}s...")
                time.sleep(COOLDOWN)
                consecutive_errors = 0
            time.sleep(DELAY)
            continue
        
        consecutive_errors = 0
        time.sleep(1)  # small gap between upload and update
        
        # Update
        success = update_record_image(sku_to_record[sku]["record_id"], image_id)
        if success:
            uploaded_ok += 1
            print(f"✅ (成功{uploaded_ok})")
        else:
            uploaded_fail += 1
            print("❌ 写入失败")
        
        time.sleep(DELAY)
    
    print("\n" + "=" * 60)
    print(f"完成! 成功: {uploaded_ok}, 失败: {uploaded_fail}")
    print(f"跳过: 无匹配{skipped_no_record}, 无效SKU{skipped_bad_sku}")
    print("=" * 60)


if __name__ == "__main__":
    main()
