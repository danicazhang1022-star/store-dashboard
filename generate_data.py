#!/usr/bin/env python3
"""
生成本地数据文件
"""

import json
import requests
from datetime import datetime

FILE_ID = "DriTvwXLsoAA"
TOKEN = "099d0e8b0cad4948806f3e2ce92facf4"

SHEETS = {
    "daily_sales": "t00i2i",
    "salary": "t00i2j",
    "inventory": "t00i2k",
}

API_BASE = "https://docs.qq.com/openapi/mcp"

def call_api(tool, params):
    """调用 API"""
    headers = {
        "Authorization": TOKEN,
        "Content-Type": "application/json"
    }
    payload = {
        "name": tool,
        "arguments": params
    }
    try:
        resp = requests.post(f"{API_BASE}/call", headers=headers, json=payload, timeout=60)
        resp.raise_for_status()
        return resp.json()
    except Exception as e:
        print(f"API 错误: {e}")
        return None

def main():
    print("开始获取数据...")
    
    # 获取销售数据
    print("\n1. 获取销售数据...")
    sales_result = call_api("smartsheet.list_records", {
        "file_id": FILE_ID,
        "sheet_id": SHEETS["daily_sales"],
        "limit": 5000
    })
    
    sales_records = []
    if sales_result and "data" in sales_result:
        for r in sales_result["data"].get("records", []):
            fields = r.get("fields", {})
            sales_records.append({
                "date": fields.get("日期", ""),
                "salesperson": fields.get("销售员", ""),
                "product": fields.get("产品", ""),
                "amount": float(fields.get("实收金额", 0) or 0),
                "quantity": int(fields.get("数量", 0) or 0),
            })
    print(f"   获取到 {len(sales_records)} 条销售记录")
    
    # 获取库存数据
    print("\n2. 获取库存数据...")
    inv_result = call_api("smartsheet.list_records", {
        "file_id": FILE_ID,
        "sheet_id": SHEETS["inventory"],
        "limit": 5000
    })
    
    inventory = []
    if inv_result and "data" in inv_result:
        for r in inv_result["data"].get("records", []):
            fields = r.get("fields", {})
            inventory.append({
                "sku": fields.get("SKU", ""),
                "name": fields.get("产品名称", ""),
                "stock": int(fields.get("当前库存", 0) or 0),
            })
    print(f"   获取到 {len(inventory)} 条库存记录")
    
    # 计算统计
    today = datetime.now().strftime("%Y-%m-%d")
    current_month = datetime.now().strftime("%Y-%m")
    
    today_sales = [s for s in sales_records if s.get("date") == today]
    today_total = sum(s["amount"] for s in today_sales)
    
    month_sales = [s for s in sales_records if str(s.get("date", "")).startswith(current_month)]
    month_total = sum(s["amount"] for s in month_sales)
    
    total_stock = sum(i["stock"] for i in inventory)
    
    # 产品热销排行
    product_stats = {}
    for s in month_sales:
        p = s.get("product", "未知")
        if p not in product_stats:
            product_stats[p] = 0
        product_stats[p] += s["amount"]
    top_products = sorted([{"name": k, "amount": v} for k, v in product_stats.items()], key=lambda x: x["amount"], reverse=True)[:10]
    
    # 保存数据
    data = {
        "last_updated": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "summary": {
            "today": {"total_sales": today_total, "count": len(today_sales)},
            "month": {"total_sales": month_total, "count": len(month_sales)},
            "inventory": {"total_stock": total_stock, "count": len(inventory)},
            "top_products": top_products
        },
        "daily_sales": sales_records[-50:],
        "inventory": inventory
    }
    
    with open("data.json", "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    
    print(f"\n✓ data.json 已保存")
    print(f"  今日销售: ¥{today_total:.2f}")
    print(f"  本月销售: ¥{month_total:.2f}")
    print(f"  库存总量: {total_stock} 件")

if __name__ == "__main__":
    main()
