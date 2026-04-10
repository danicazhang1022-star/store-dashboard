#!/usr/bin/env python3
"""
Upload product images to GitHub and update Tencent Docs Smartsheet "图片链接" field.
"""

import subprocess
import json
import sys
import os
import time
import csv
import glob

FILE_ID = "DriTvwXLsoAA"
SHEET_ID = "t00i2h"
IMAGE_FIELD_TITLE = "图片链接"
SKU_FIELD_TITLE = "货号SKU"
IMAGE_DIR = "/Users/danica/WorkBuddy/20260407210337/产品图片"
CSV_PATH = "/Users/danica/WorkBuddy/20260407210337/图片下载结果.csv"
GITHUB_REPO = "danicazhang1022-star/store-dashboard"
GITHUB_BRANCH = "main"
DELAY = 2  # seconds between API calls


def mcporter_call(tool, args_dict):
    """Call mcporter tool and return parsed result dict."""
    args_json = json.dumps(args_dict, ensure_ascii=False)
    cmd = f'''export PATH="/Users/danica/.workbuddy/binaries/node/versions/22.12.0/bin:$PATH" && mcporter call "tencent-docs" "{tool}" --args '{args_json}' 2>/dev/null'''
    result = subprocess.run(cmd, shell=True, capture_output=True, text=True, timeout=120)
    try:
        data = json.loads(result.stdout.strip())
        if data and data.get("error"):
            print(f"  API ERROR: {data['error']}", file=sys.stderr)
            return None
        return data
    except Exception as e:
        print(f"  Parse error: {e}", file=sys.stderr)
        return None


def run_command(cmd, cwd=None):
    """Run shell command and return success status."""
    result = subprocess.run(cmd, shell=True, capture_output=True, text=True, cwd=cwd)
    if result.returncode != 0:
        print(f"  Command failed: {result.stderr}", file=sys.stderr)
        return False
    return True


def upload_image_to_github(image_path):
    """Upload image to GitHub repo and return raw URL."""
    file_name = os.path.basename(image_path)
    
    # Copy image to repo
    repo_dir = "/Users/danica/WorkBuddy/20260407210337"
    images_dir = os.path.join(repo_dir, "product-images")
    os.makedirs(images_dir, exist_ok=True)
    
    dest_path = os.path.join(images_dir, file_name)
    
    # Copy file
    with open(image_path, 'rb') as src, open(dest_path, 'wb') as dst:
        dst.write(src.read())
    
    # Git add, commit, push
    if not run_command(f'cd {repo_dir} && git add product-images/{file_name}'):
        return None
    
    # Check if there are changes to commit
    status_result = subprocess.run(f'cd {repo_dir} && git status --porcelain product-images/{file_name}', 
                                   shell=True, capture_output=True, text=True)
    if not status_result.stdout.strip():
        # File already exists and no changes
        print(f"  {file_name} already in repo")
    else:
        if not run_command(f'cd {repo_dir} && git commit -m "Add product image: {file_name}"'):
            return None
        if not run_command(f'cd {repo_dir} && git push origin {GITHUB_BRANCH}'):
            return None
        print(f"  Uploaded {file_name} to GitHub")
    
    # Return raw GitHub URL
    return f"https://raw.githubusercontent.com/{GITHUB_REPO}/{GITHUB_BRANCH}/product-images/{file_name}"


def get_all_records():
    """Get all records using pagination."""
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
                records.append({"record_id": r["record_id"], "sku": sku})
        if len(data.get("records", [])) < 100:
            break
        time.sleep(1)
    return records


def update_image_link(record_id, image_url):
    """Update record with image URL in 图片链接 field."""
    data = mcporter_call("smartsheet.update_records", {
        "file_id": FILE_ID,
        "sheet_id": SHEET_ID,
        "records": [{
            "record_id": record_id,
            "field_values": {
                IMAGE_FIELD_TITLE: {
                    "type": "url",
                    "link": image_url,
                    "text": image_url
                }
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
    print("产品图片上传工具 - GitHub + 图片链接字段更新")
    print("=" * 60)
    
    # Step 1: Get all records
    print("\n[1/4] 获取产品库全部记录...")
    records = get_all_records()
    if not records:
        print("ERROR: 无法获取记录！")
        sys.exit(1)
    
    sku_to_record = {r["sku"]: r["record_id"] for r in records}
    print(f"  共 {len(sku_to_record)} 条唯一SKU记录")
    
    # Step 2: Load CSV and image files
    print("\n[2/4] 扫描图片文件...")
    csv_map = build_csv_mapping()
    image_files = glob.glob(os.path.join(IMAGE_DIR, "*.jpg"))
    print(f"  {len(image_files)} 张图片, {len(csv_map)} 条CSV映射")
    
    # Step 3: Upload to GitHub
    print("\n[3/4] 上传图片到 GitHub...")
    matched = 0
    skipped_no_record = 0
    skipped_bad_sku = 0
    upload_ok = 0
    upload_fail = 0
    sku_to_url = {}
    
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
        print(f"  [{i+1}/{len(image_files)}] {sku}...", end=" ", flush=True)
        
        url = upload_image_to_github(img_path)
        if url:
            sku_to_url[sku] = url
            upload_ok += 1
            print(f"✅")
        else:
            upload_fail += 1
            print(f"❌")
        
        time.sleep(0.5)
    
    print(f"\n  上传完成: 成功 {upload_ok}, 失败 {upload_fail}")
    
    # Step 4: Update Smartsheet
    print("\n[4/4] 更新腾讯文档图片链接字段...")
    updated_ok = 0
    updated_fail = 0
    
    for sku, url in sku_to_url.items():
        record_id = sku_to_record[sku]
        print(f"  更新 {sku}...", end=" ", flush=True)
        
        if update_image_link(record_id, url):
            updated_ok += 1
            print(f"✅")
        else:
            updated_fail += 1
            print(f"❌")
        
        time.sleep(DELAY)
    
    print("\n" + "=" * 60)
    print(f"完成!")
    print(f"  GitHub上传: 成功 {upload_ok}, 失败 {upload_fail}")
    print(f"  字段更新: 成功 {updated_ok}, 失败 {updated_fail}")
    print(f"  跳过: 无匹配{skipped_no_record}, 无效SKU{skipped_bad_sku}")
    print("=" * 60)


if __name__ == "__main__":
    main()
