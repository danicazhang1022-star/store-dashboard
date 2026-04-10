#!/usr/bin/env python3
"""
生成本地数据文件，用于仪表盘展示
在本地运行此脚本，然后将生成的 data.json 提交到 GitHub
"""

import json
import subprocess
from datetime import datetime, timedelta
from collections import defaultdict

FILE_ID = "DriTvwXLsoAA"

# 正确的工作表 ID（与 auto_salary.py 一致）
SHEETS = {
    "daily_sales": "VtXdmr",    # 每日销售表
    "salary": "VDPOkr",         # 薪资计算表
    "inventory": "gsnTK8",      # 库存管理表（店铺库存管理）
}


def mcporter_call(tool_name, args):
    """调用腾讯文档 MCP 工具（使用与 auto_salary.py 相同的格式）"""
    cmd = [
        "/Users/danica/.workbuddy/binaries/node/versions/22.12.0/bin/node",
        "/Users/danica/.workbuddy/binaries/node/versions/22.12.0/lib/node_modules/mcporter/dist/cli.js",
        "call", "tencent-docs", tool_name,
        "--args", json.dumps(args, ensure_ascii=False)
    ]
    result = subprocess.run(cmd, capture_output=True, text=True, timeout=30)
    if result.returncode != 0:
        print(f"  ❌ MCP调用失败: {tool_name}")
        print(f"     stderr: {result.stderr[:200]}")
        return None
    try:
        output = result.stdout.strip()
        if output:
            return json.loads(output)
    except json.JSONDecodeError:
        print(f"  ⚠️  MCP返回非JSON: {result.stdout[:200]}")
    return None


def get_all_records(sheet_id, field_titles=None):
    """获取工作表所有记录（自动分页）"""
    all_records = []
    offset = 0
    while True:
        args = {"file_id": FILE_ID, "sheet_id": sheet_id, "offset": offset, "limit": 100}
        if field_titles:
            args["field_titles"] = field_titles
        
        data = mcporter_call("smartsheet.list_records", args)
        if not data or "records" not in data:
            break
        
        records = data["records"]
        if not records:
            break
        
        all_records.extend(records)
        if len(records) < 100:
            break
        
        offset += 100
        import time
        time.sleep(1)  # 避免限流
    
    return all_records


def parse_date(val):
    """解析日期值，支持毫秒时间戳和日期字符串"""
    if not val:
        return None
    if isinstance(val, list):
        val = val[0] if val else None
    if not val:
        return None
    
    val_str = str(val).strip()
    
    # 毫秒时间戳（13位数字）
    if val_str.isdigit() and len(val_str) >= 10:
        try:
            ts = int(val_str)
            if ts > 1e10:
                ts = ts / 1000
            import datetime
            tz_cst = datetime.timezone(datetime.timedelta(hours=8))
            return datetime.datetime.fromtimestamp(ts, tz=tz_cst).strftime("%Y-%m-%d")
        except Exception:
            return None
    
    # 日期字符串
    if len(val_str) >= 10 and val_str[4] == "-":
        return val_str[:10]
    
    return None


def parse_number(val):
    """安全解析数字"""
    if val is None:
        return 0
    if isinstance(val, (int, float)):
        return val
    if isinstance(val, list):
        val = val[0] if val else 0
    try:
        return float(str(val).replace(",", "").replace("¥", "").strip())
    except (ValueError, TypeError):
        return 0


def parse_text(val):
    """解析文本字段"""
    if not val:
        return ""
    if isinstance(val, list) and val:
        if isinstance(val[0], dict):
            return val[0].get("text", "")
        return str(val[0])
    return str(val)


def main():
    print("=" * 50)
    print("生成仪表盘数据文件")
    print("=" * 50)
    
    # 每日销售表字段
    SALES_FIELDS = {
        "date": "日期",
        "order_id": "订单号（第几单）",
        "salesperson": "销售员",
        "amount": "实收金额",
        "product": "产品",
        "quantity": "数量",
    }
    
    # 库存表字段
    INVENTORY_FIELDS = {
        "sku": "SKU",
        "name": "产品名称",
        "stock": "当前库存",
        "category": "分类",
    }
    
    # 获取每日销售数据
    print("\n1. 获取每日销售数据...")
    sales_field_titles = list(SALES_FIELDS.values())
    sales_records = get_all_records(SHEETS["daily_sales"], field_titles=sales_field_titles)
    print(f"   获取到 {len(sales_records)} 条销售记录")
    
    # 处理销售记录
    daily_sales = []
    for rec in sales_records:
        field_values = rec.get("field_values", {})
        rec_date = parse_date(field_values.get(SALES_FIELDS["date"], ""))
        if not rec_date:
            continue
        
        daily_sales.append({
            "date": rec_date,
            "order_id": parse_text(field_values.get(SALES_FIELDS["order_id"], "")),
            "salesperson": parse_text(field_values.get(SALES_FIELDS["salesperson"], "")),
            "amount": parse_number(field_values.get(SALES_FIELDS["amount"], 0)),
            "product": parse_text(field_values.get(SALES_FIELDS["product"], "")),
            "quantity": parse_number(field_values.get(SALES_FIELDS["quantity"], 1)),
        })
    
    # 获取库存数据
    print("\n2. 获取库存数据...")
    inventory_records = get_all_records(SHEETS["inventory"])
    print(f"   获取到 {len(inventory_records)} 条库存记录")
    
    # 处理库存记录
    inventory = []
    for rec in inventory_records:
        field_values = rec.get("field_values", {})
        inventory.append({
            "sku": parse_text(field_values.get(INVENTORY_FIELDS["sku"], "")),
            "name": parse_text(field_values.get(INVENTORY_FIELDS["name"], "")),
            "stock": int(parse_number(field_values.get(INVENTORY_FIELDS["stock"], 0))),
            "category": parse_text(field_values.get(INVENTORY_FIELDS["category"], "")),
        })
    
    # 计算统计数据
    print("\n3. 计算统计数据...")
    today = datetime.now().strftime("%Y-%m-%d")
    current_month = datetime.now().strftime("%Y-%m")
    
    # 30天前日期（用于滞销计算）
    date_30d_ago = (datetime.now() - timedelta(days=30)).strftime("%Y-%m-%d")
    
    # 今日销售
    today_sales = [s for s in daily_sales if s.get("date") == today]
    today_total = sum(s.get("amount", 0) for s in today_sales)
    today_orders = len(set(s.get("order_id", "") for s in today_sales if s.get("order_id")))
    
    # 本月销售
    month_sales = [s for s in daily_sales if str(s.get("date", "")).startswith(current_month)]
    month_total = sum(s.get("amount", 0) for s in month_sales)
    
    # 最近30天销售（用于滞销计算）
    sales_30d = [s for s in daily_sales if s.get("date", "") >= date_30d_ago]
    
    # 库存统计
    total_stock = sum(i.get("stock", 0) for i in inventory)
    low_stock = [i for i in inventory if i.get("stock", 0) < 5 and i.get("stock", 0) > 0]
    
    # 销售员排行
    salesperson_stats = defaultdict(lambda: {"sales": 0, "orders": set()})
    for sale in month_sales:
        sp = sale.get("salesperson", "未知")
        if sp:
            salesperson_stats[sp]["sales"] += sale.get("amount", 0)
            if sale.get("order_id"):
                salesperson_stats[sp]["orders"].add(sale.get("order_id"))
    
    top_salespeople = sorted(
        [{"name": k, "sales": v["sales"], "orders": len(v["orders"])} 
         for k, v in salesperson_stats.items()],
        key=lambda x: x["sales"],
        reverse=True
    )[:10]
    
    # 款式热销 Top10（按销量排序）
    product_stats = defaultdict(lambda: {"quantity": 0, "sales": 0})
    for sale in month_sales:
        product_name = sale.get("product", "")
        if product_name:
            # 尝试匹配 SKU
            sku = ""
            for inv in inventory:
                if inv.get("name") == product_name:
                    sku = inv.get("sku", "")
                    break
            
            product_stats[product_name]["quantity"] += sale.get("quantity", 1)
            product_stats[product_name]["sales"] += sale.get("amount", 0)
            if "sku" not in product_stats[product_name]:
                product_stats[product_name]["sku"] = sku
    
    top_products = sorted(
        [{"name": k, "sku": v.get("sku", ""), "quantity": v["quantity"], "sales": v["sales"]} 
         for k, v in product_stats.items()],
        key=lambda x: x["quantity"],
        reverse=True
    )[:10]
    
    # 滞销排行（库存高但30天销量低）
    # 计算每个商品的30天销量
    product_sales_30d = defaultdict(int)
    for sale in sales_30d:
        product_name = sale.get("product", "")
        if product_name:
            product_sales_30d[product_name] += sale.get("quantity", 1)
    
    # 滞销商品：库存 > 10 且 30天销量 < 5
    slow_products = []
    for inv in inventory:
        stock = inv.get("stock", 0)
        name = inv.get("name", "")
        sales_30d_count = product_sales_30d.get(name, 0)
        
        if stock > 0:
            ratio = round(stock / max(sales_30d_count, 1), 1)  # 库销比
            if stock >= 10 and sales_30d_count < 5:
                slow_products.append({
                    "sku": inv.get("sku", ""),
                    "name": name,
                    "stock": stock,
                    "sales_30d": sales_30d_count,
                    "ratio": ratio,
                })
    
    # 按库销比排序（越高越滞销）
    slow_products = sorted(slow_products, key=lambda x: x["ratio"], reverse=True)[:10]
    
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
            "slow_products": slow_products,
        },
        "daily_sales": daily_sales[-50:],
        "inventory": inventory
    }
    
    # 保存文件
    print("\n4. 保存数据文件...")
    with open("data.json", "w", encoding="utf-8") as f:
        json.dump(data_package, f, ensure_ascii=False, indent=2)
    print("   ✓ data.json 已保存")
    
    print("\n" + "=" * 50)
    print("数据生成完成!")
    print(f"今日销售: ¥{today_total:.2f}")
    print(f"本月销售: ¥{month_total:.2f}")
    print(f"库存总量: {total_stock} 件")
    print(f"热销商品: {len(top_products)} 款")
    print(f"滞销商品: {len(slow_products)} 款")
    print("=" * 50)


if __name__ == "__main__":
    main()
