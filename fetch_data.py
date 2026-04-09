#!/usr/bin/env python3
"""
获取腾讯文档数据并生成仪表盘数据文件
使用与 auto_salary.py 相同的调用方式
"""

import json
import subprocess
import time
from datetime import datetime

FILE_ID = "DriTvwXLsoAA"

SHEETS = {
    "daily_sales": "VtXdmr",    # 每日销售表
    "salary": "VDPOkr",         # 薪资计算表
}

def mcporter_call(tool_name, args):
    """调用腾讯文档 MCP 工具"""
    cmd = [
        "/Users/danica/.workbuddy/binaries/node/versions/22.12.0/bin/node",
        "/Users/danica/.workbuddy/binaries/node/versions/22.12.0/lib/node_modules/mcporter/dist/cli.js",
        "call", "tencent-docs", tool_name,
        "--args", json.dumps(args, ensure_ascii=False)
    ]
    result = subprocess.run(cmd, capture_output=True, text=True, timeout=30)
    if result.returncode != 0:
        print(f"  ❌ MCP调用失败: {result.stderr[:200]}")
        return None
    try:
        output = result.stdout.strip()
        if output:
            return json.loads(output)
    except json.JSONDecodeError:
        print(f"  ⚠️ MCP返回非JSON: {result.stdout[:200]}")
    return None

def get_all_records(sheet_id):
    """获取工作表所有记录（自动分页）"""
    all_records = []
    offset = 0
    while True:
        data = mcporter_call("smartsheet.list_records", {
            "file_id": FILE_ID,
            "sheet_id": sheet_id,
            "offset": offset,
            "limit": 100
        })
        if not data or "records" not in data:
            break
        records = data["records"]
        if not records:
            break
        all_records.extend(records)
        if len(records) < 100:
            break
        offset += 100
        time.sleep(0.5)  # 避免限流
    return all_records

def extract_value(field_val):
    """提取字段值"""
    if isinstance(field_val, list) and len(field_val) > 0:
        if isinstance(field_val[0], dict):
            return field_val[0].get("text", "")
        return field_val[0]
    return field_val if field_val else ""

def main():
    print("=" * 50)
    print("获取腾讯文档数据")
    print("=" * 50)
    
    # 获取销售数据
    print("\n1. 获取每日销售数据...")
    sales_records = get_all_records(SHEETS["daily_sales"])
    print(f"   ✓ 共 {len(sales_records)} 条记录")
    
    sales_data = []
    for r in sales_records:
        fields = r.get("fields", {})
        try:
            amount = float(extract_value(fields.get("实收金额")) or 0)
        except:
            amount = 0
        try:
            qty = int(extract_value(fields.get("数量")) or 0)
        except:
            qty = 0
        sales_data.append({
            "date": extract_value(fields.get("日期")),
            "salesperson": extract_value(fields.get("销售员")),
            "product": extract_value(fields.get("产品")),
            "amount": amount,
            "quantity": qty,
            "order_id": extract_value(fields.get("订单号")),
        })
    
    # 计算统计
    today = datetime.now().strftime("%Y-%m-%d")
    current_month = datetime.now().strftime("%Y-%m")
    
    today_sales = [s for s in sales_data if s.get("date") == today]
    today_total = sum(s["amount"] for s in today_sales)
    
    month_sales = [s for s in sales_data if str(s.get("date", "")).startswith(current_month)]
    month_total = sum(s["amount"] for s in month_sales)
    
    # 销售排行
    sp_stats = {}
    for s in month_sales:
        sp = s.get("salesperson", "未知")
        if sp not in sp_stats:
            sp_stats[sp] = 0
        sp_stats[sp] += s["amount"]
    top_sp = sorted([{"name": k, "sales": v} for k, v in sp_stats.items()], key=lambda x: x["sales"], reverse=True)[:5]
    
    # 产品热销排行
    prod_stats = {}
    for s in month_sales:
        p = s.get("product", "未知")
        if p not in prod_stats:
            prod_stats[p] = {"qty": 0, "amount": 0}
        prod_stats[p]["qty"] += s["quantity"]
        prod_stats[p]["amount"] += s["amount"]
    top_prod = sorted([{"name": k, **v} for k, v in prod_stats.items()], key=lambda x: x["amount"], reverse=True)[:10]
    
    # 保存数据
    data = {
        "last_updated": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "summary": {
            "today": {
                "date": today,
                "total_sales": today_total,
                "order_count": len(set(s.get("order_id", "") for s in today_sales if s.get("order_id"))),
                "transaction_count": len(today_sales)
            },
            "month": {
                "month": current_month,
                "total_sales": month_total,
                "transaction_count": len(month_sales)
            },
            "top_salespeople": top_sp,
            "top_products": top_prod,
        },
        "daily_sales": sales_data[-50:],
    }
    
    with open("data.json", "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    
    print(f"\n{'=' * 50}")
    print("✓ data.json 已生成")
    print(f"  今日销售: ¥{today_total:.2f}")
    print(f"  本月销售: ¥{month_total:.2f}")
    print(f"  销售员: {len(sp_stats)} 人")
    print(f"  热销商品: {len(top_prod)} 款")
    print(f"{'=' * 50}")

if __name__ == "__main__":
    main()
