#!/usr/bin/env python3
"""
生成本地数据文件，用于仪表盘展示
在本地运行此脚本，然后将生成的 data.json 提交到 GitHub
"""

import json
import subprocess
from datetime import datetime

FILE_ID = "DriTvwXLsoAA"

SHEETS = {
    "daily_sales": "t00i2i",
    "salary": "t00i2j",
    "inventory": "t00i2k",
}


def mcp_call(tool_name, **params):
    """调用 MCP 工具"""
    cmd = ["/Users/danica/.workbuddy/binaries/node/versions/22.12.0/bin/node", 
           "/Users/danica/.workbuddy/binaries/node/versions/22.12.0/lib/node_modules/mcporter/dist/cli.js",
           "call", f"tencent-sheetengine.{tool_name}"]
    
    for key, value in params.items():
        if isinstance(value, str):
            cmd.append(f'{key}="{value}"')
        else:
            cmd.append(f'{key}={value}')
    
    result = subprocess.run(cmd, capture_output=True, text=True, timeout=60)
    
    if result.returncode != 0:
        print(f"错误: {result.stderr}")
        return None
    
    # 解析输出中的 JSON
    lines = result.stdout.strip().split('\n')
    for line in lines:
        line = line.strip()
        if line and line.startswith('{'):
            try:
                return json.loads(line)
            except:
                continue
    return None


def process_records(data, field_mapping):
    """处理记录数据"""
    if not data or "records" not in data:
        return []
    
    records = []
    for record in data["records"]:
        fields = record.get("fields", {})
        record_data = {"id": record.get("id", "")}
        
        for key, field_name in field_mapping.items():
            value = fields.get(field_name)
            
            # 处理文本数组
            if isinstance(value, list) and len(value) > 0:
                if isinstance(value[0], dict) and "text" in value[0]:
                    record_data[key] = value[0]["text"]
                else:
                    record_data[key] = value[0]
            # 处理数字
            elif isinstance(value, (int, float)):
                record_data[key] = value
            elif isinstance(value, str):
                try:
                    record_data[key] = float(value)
                except:
                    record_data[key] = value
            else:
                record_data[key] = value or ""
        
        records.append(record_data)
    
    return records


def main():
    print("=" * 50)
    print("生成仪表盘数据文件")
    print("=" * 50)
    
    # 获取每日销售数据
    print("\n1. 获取每日销售数据...")
    sales_data = mcp_call("smartsheet.list_records", 
                          file_id=FILE_ID, 
                          sheet_id=SHEETS["daily_sales"], 
                          limit=5000)
    
    sales_mapping = {
        "date": "日期",
        "salesperson": "销售员",
        "product": "产品",
        "quantity": "数量",
        "price": "单价",
        "amount": "实收金额",
        "order_id": "订单号",
        "payment": "收款方式",
        "note": "备注",
    }
    daily_sales = process_records(sales_data, sales_mapping)
    print(f"   获取到 {len(daily_sales)} 条销售记录")
    
    # 获取薪资数据
    print("\n2. 获取薪资数据...")
    salary_data = mcp_call("smartsheet.list_records",
                           file_id=FILE_ID,
                           sheet_id=SHEETS["salary"],
                           limit=5000)
    
    salary_mapping = {
        "date": "日期",
        "salesperson": "销售员",
        "daily_sales": "当日销售额",
        "commission_rate": "提成比例",
        "big_order_bonus": "大单奖励",
        "daily_salary": "固定日薪",
        "commission_subtotal": "提成小计",
        "daily_total": "当日薪资合计",
    }
    salary = process_records(salary_data, salary_mapping)
    print(f"   获取到 {len(salary)} 条薪资记录")
    
    # 获取库存数据
    print("\n3. 获取库存数据...")
    inventory_data = mcp_call("smartsheet.list_records",
                              file_id=FILE_ID,
                              sheet_id=SHEETS["inventory"],
                              limit=5000)
    
    inventory_mapping = {
        "sku": "SKU",
        "name": "产品名称",
        "category": "分类",
        "size": "尺码",
        "color": "颜色",
        "stock": "当前库存",
    }
    inventory = process_records(inventory_data, inventory_mapping)
    print(f"   获取到 {len(inventory)} 条库存记录")
    
    # 计算统计数据
    print("\n4. 计算统计数据...")
    today = datetime.now().strftime("%Y-%m-%d")
    current_month = datetime.now().strftime("%Y-%m")
    
    # 今日销售
    today_sales = [s for s in daily_sales if s.get("date") == today]
    today_total = sum(float(s.get("amount", 0) or 0) for s in today_sales)
    today_orders = len(set(s.get("order_id", "") for s in today_sales if s.get("order_id")))
    
    # 本月销售
    month_sales = [s for s in daily_sales if str(s.get("date", "")).startswith(current_month)]
    month_total = sum(float(s.get("amount", 0) or 0) for s in month_sales)
    
    # 库存统计
    total_stock = sum(int(i.get("stock", 0) or 0) for i in inventory)
    low_stock = [i for i in inventory if int(i.get("stock", 0) or 0) < 5]
    
    # 销售排行
    salesperson_stats = {}
    for sale in month_sales:
        sp = sale.get("salesperson", "未知")
        if sp not in salesperson_stats:
            salesperson_stats[sp] = {"sales": 0, "orders": set()}
        salesperson_stats[sp]["sales"] += float(sale.get("amount", 0) or 0)
        if sale.get("order_id"):
            salesperson_stats[sp]["orders"].add(sale.get("order_id"))
    
    top_salespeople = sorted(
        [{"name": k, "sales": v["sales"], "orders": len(v["orders"])} 
         for k, v in salesperson_stats.items()],
        key=lambda x: x["sales"],
        reverse=True
    )[:10]
    
    # 产品热销排行
    product_stats = {}
    for sale in month_sales:
        product = sale.get("product", "未知")
        if product not in product_stats:
            product_stats[product] = {"quantity": 0, "amount": 0}
        product_stats[product]["quantity"] += int(sale.get("quantity", 0) or 0)
        product_stats[product]["amount"] += float(sale.get("amount", 0) or 0)
    
    top_products = sorted(
        [{"name": k, "quantity": v["quantity"], "amount": v["amount"]} 
         for k, v in product_stats.items()],
        key=lambda x: x["amount"],
        reverse=True
    )[:10]
    
    # 滞销商品（本月销量为0的库存商品）
    sold_products = set(sale.get("product", "") for sale in month_sales)
    unsold = [i for i in inventory if i.get("name") not in sold_products and int(i.get("stock", 0) or 0) > 0]
    unsold_sorted = sorted(unsold, key=lambda x: int(x.get("stock", 0) or 0), reverse=True)[:10]
    
    # 构建数据包
    data_package = {
        "last_updated": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "summary": {
            "today": {
                "date": today,
                "total_sales": today_total,
                "order_count": today_orders,
                "transaction_count": len(today_sales)
            },
            "month": {
                "month": current_month,
                "total_sales": month_total,
                "transaction_count": len(month_sales)
            },
            "inventory": {
                "total_items": len(inventory),
                "total_stock": total_stock,
                "low_stock_count": len(low_stock),
                "low_stock_items": low_stock[:10]
            },
            "top_salespeople": top_salespeople,
            "top_products": top_products,
            "unsold_products": unsold_sorted
        },
        "daily_sales": daily_sales[-50:],
        "salary": salary[-50:],
        "inventory": inventory
    }
    
    # 保存文件
    print("\n5. 保存数据文件...")
    with open("data.json", "w", encoding="utf-8") as f:
        json.dump(data_package, f, ensure_ascii=False, indent=2)
    print("   ✓ data.json 已保存")
    
    print("\n" + "=" * 50)
    print("数据生成完成!")
    print(f"今日销售: ¥{today_total:.2f}")
    print(f"本月销售: ¥{month_total:.2f}")
    print(f"库存总量: {total_stock} 件")
    print(f"热销商品: {len(top_products)} 款")
    print(f"滞销商品: {len(unsold_sorted)} 款")
    print("=" * 50)


if __name__ == "__main__":
    main()
