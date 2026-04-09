#!/usr/bin/env python3
"""
GitHub Actions 数据同步脚本
从腾讯文档智能表格读取数据，生成仪表盘所需的 JSON 数据文件
"""

import os
import json
import requests
from datetime import datetime, timedelta

# ============ 配置 ============
FILE_ID = os.environ.get("FILE_ID", "DriTvwXLsoAA")
TOKEN = os.environ.get("TENCENT_DOCS_TOKEN", "")

# 工作表 ID
SHEETS = {
    "products": "t00i2h",      # 产品库
    "daily_sales": "t00i2i",   # 每日销售
    "salary": "t00i2j",        # 薪资计算
    "inventory": "t00i2k",     # 库存管理
    "inbound": "t00i2l",       # 入库记录
    "restock": "t00i2m",       # 补货导出
}

# 腾讯文档 API 端点
API_BASE_URL = "https://docs.qq.com/api/v6/sheet"


def api_call(endpoint, params):
    """调用腾讯文档 API"""
    headers = {
        "Authorization": f"Bearer {TOKEN}",
        "Content-Type": "application/json"
    }
    
    response = requests.post(
        f"{API_BASE_URL}/{endpoint}",
        headers=headers,
        json=params,
        timeout=60
    )
    response.raise_for_status()
    return response.json()


def fetch_sheet_data(sheet_id):
    """获取工作表数据"""
    try:
        result = api_call("getSheetData", {
            "fileID": FILE_ID,
            "sheetID": sheet_id,
            "limit": 5000
        })
        
        if result.get("code") == 0:
            return result.get("data", {})
        else:
            print(f"API 错误: {result.get('msg', '未知错误')}")
            return None
    except Exception as e:
        print(f"获取工作表 {sheet_id} 失败: {e}")
        return None


def process_daily_sales(data):
    """处理每日销售数据"""
    if not data or "records" not in data:
        return []
    
    records = []
    for record in data["records"]:
        fields = record.get("fields", {})
        
        # 提取字段值（处理文本数组格式）
        def get_text_value(field_val):
            if isinstance(field_val, list) and len(field_val) > 0:
                return field_val[0].get("text", "")
            return field_val if field_val else ""
        
        def get_number_value(field_val):
            if isinstance(field_val, (int, float)):
                return field_val
            if isinstance(field_val, str):
                try:
                    return float(field_val)
                except:
                    return 0
            return 0
        
        record_data = {
            "id": record.get("id", ""),
            "date": get_text_value(fields.get("日期")),
            "salesperson": get_text_value(fields.get("销售员")),
            "product": get_text_value(fields.get("产品")),
            "quantity": get_number_value(fields.get("数量")),
            "price": get_number_value(fields.get("单价")),
            "amount": get_number_value(fields.get("实收金额")),
            "order_id": get_text_value(fields.get("订单号")),
            "payment": get_text_value(fields.get("收款方式")),
            "note": get_text_value(fields.get("备注")),
        }
        records.append(record_data)
    
    return records


def process_salary_data(data):
    """处理薪资数据"""
    if not data or "records" not in data:
        return []
    
    records = []
    for record in data["records"]:
        fields = record.get("fields", {})
        
        def get_text_value(field_val):
            if isinstance(field_val, list) and len(field_val) > 0:
                return field_val[0].get("text", "")
            return field_val if field_val else ""
        
        def get_number_value(field_val):
            if isinstance(field_val, (int, float)):
                return field_val
            if isinstance(field_val, str):
                try:
                    return float(field_val)
                except:
                    return 0
            return 0
        
        record_data = {
            "id": record.get("id", ""),
            "date": get_text_value(fields.get("日期")),
            "salesperson": get_text_value(fields.get("销售员")),
            "daily_sales": get_number_value(fields.get("当日销售额")),
            "commission_rate": get_text_value(fields.get("提成比例")),
            "big_order_bonus": get_number_value(fields.get("大单奖励")),
            "daily_salary": get_number_value(fields.get("固定日薪")),
            "commission_subtotal": get_number_value(fields.get("提成小计")),
            "daily_total": get_number_value(fields.get("当日薪资合计")),
        }
        records.append(record_data)
    
    return records


def process_inventory_data(data):
    """处理库存数据"""
    if not data or "records" not in data:
        return []
    
    records = []
    for record in data["records"]:
        fields = record.get("fields", {})
        
        def get_text_value(field_val):
            if isinstance(field_val, list) and len(field_val) > 0:
                return field_val[0].get("text", "")
            return field_val if field_val else ""
        
        def get_number_value(field_val):
            if isinstance(field_val, (int, float)):
                return field_val
            if isinstance(field_val, str):
                try:
                    return float(field_val)
                except:
                    return 0
            return 0
        
        # 处理图片字段
        images = fields.get("图片", [])
        image_urls = []
        if isinstance(images, list):
            for img in images:
                if isinstance(img, dict) and "url" in img:
                    image_urls.append(img["url"])
        
        record_data = {
            "id": record.get("id", ""),
            "sku": get_text_value(fields.get("SKU")),
            "name": get_text_value(fields.get("产品名称")),
            "category": get_text_value(fields.get("分类")),
            "size": get_text_value(fields.get("尺码")),
            "color": get_text_value(fields.get("颜色")),
            "stock": get_number_value(fields.get("当前库存")),
            "images": image_urls,
        }
        records.append(record_data)
    
    return records


def generate_summary(daily_sales, salary_data, inventory):
    """生成汇总统计"""
    today = datetime.now().strftime("%Y-%m-%d")
    
    # 今日销售统计
    today_sales = [s for s in daily_sales if s.get("date") == today]
    today_total = sum(s.get("amount", 0) for s in today_sales)
    today_orders = len(set(s.get("order_id", "") for s in today_sales if s.get("order_id")))
    
    # 本月销售统计
    current_month = datetime.now().strftime("%Y-%m")
    month_sales = [s for s in daily_sales if s.get("date", "").startswith(current_month)]
    month_total = sum(s.get("amount", 0) for s in month_sales)
    
    # 库存统计
    low_stock = [i for i in inventory if i.get("stock", 0) < 5]
    total_stock = sum(i.get("stock", 0) for i in inventory)
    
    # 销售员业绩（本月）
    salesperson_stats = {}
    for sale in month_sales:
        sp = sale.get("salesperson", "未知")
        if sp not in salesperson_stats:
            salesperson_stats[sp] = {"sales": 0, "orders": 0}
        salesperson_stats[sp]["sales"] += sale.get("amount", 0)
    
    # 按销售额排序
    top_salespeople = sorted(
        [{"name": k, **v} for k, v in salesperson_stats.items()],
        key=lambda x: x["sales"],
        reverse=True
    )[:5]
    
    return {
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
            "low_stock_items": low_stock[:10]  # 只显示前10个
        },
        "top_salespeople": top_salespeople
    }


def main():
    """主函数"""
    print("=" * 50)
    print("开始同步数据...")
    print(f"时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"文档 ID: {FILE_ID}")
    print("=" * 50)
    
    # 获取各工作表数据
    print("\n1. 获取每日销售数据...")
    daily_sales_raw = fetch_sheet_data(SHEETS["daily_sales"])
    daily_sales = process_daily_sales(daily_sales_raw)
    print(f"   获取到 {len(daily_sales)} 条销售记录")
    
    print("\n2. 获取薪资数据...")
    salary_raw = fetch_sheet_data(SHEETS["salary"])
    salary_data = process_salary_data(salary_raw)
    print(f"   获取到 {len(salary_data)} 条薪资记录")
    
    print("\n3. 获取库存数据...")
    inventory_raw = fetch_sheet_data(SHEETS["inventory"])
    inventory = process_inventory_data(inventory_raw)
    print(f"   获取到 {len(inventory)} 条库存记录")
    
    # 生成汇总统计
    print("\n4. 生成汇总统计...")
    summary = generate_summary(daily_sales, salary_data, inventory)
    
    # 保存数据文件
    print("\n5. 保存数据文件...")
    
    data_package = {
        "last_updated": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "summary": summary,
        "daily_sales": daily_sales[-100:],  # 最近100条
        "salary": salary_data[-100:],       # 最近100条
        "inventory": inventory
    }
    
    with open("data.json", "w", encoding="utf-8") as f:
        json.dump(data_package, f, ensure_ascii=False, indent=2)
    print("   ✓ data.json 已保存")
    
    # 同时保存单独的文件供详细查看
    with open("daily_sales.json", "w", encoding="utf-8") as f:
        json.dump(daily_sales, f, ensure_ascii=False, indent=2)
    print("   ✓ daily_sales.json 已保存")
    
    with open("salary.json", "w", encoding="utf-8") as f:
        json.dump(salary_data, f, ensure_ascii=False, indent=2)
    print("   ✓ salary.json 已保存")
    
    with open("inventory.json", "w", encoding="utf-8") as f:
        json.dump(inventory, f, ensure_ascii=False, indent=2)
    print("   ✓ inventory.json 已保存")
    
    print("\n" + "=" * 50)
    print("数据同步完成!")
    print(f"今日销售: ¥{summary['today']['total_sales']:.2f}")
    print(f"本月销售: ¥{summary['month']['total_sales']:.2f}")
    print(f"库存总量: {summary['inventory']['total_stock']} 件")
    print("=" * 50)


if __name__ == "__main__":
    main()
