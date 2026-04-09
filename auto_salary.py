#!/usr/bin/env python3
"""
店铺薪资自动计算脚本
=====================
功能：
1. 读取「每日销售」表 → 按日期+销售员汇总
2. 自动计算：日销售额、阶梯提成、大单奖励、日薪、合计当日薪资
3. 写入「薪资计算」表（追加新记录，不覆盖已有数据）

使用方式：
  python3 auto_salary.py              # 计算今天的数据
  python3 auto_salary.py 2026-04-07   # 计算指定日期

依赖：mcporter CLI 已安装并完成授权
"""

import json
import subprocess
import sys
import os
from datetime import date, timedelta
from collections import defaultdict

# ============ 配置区（根据你的实际字段名修改）============

FILE_ID = "DriTvwXLsoAA"

# 每日销售表
SALES_SHEET_ID = "VtXdmr"

# 薪资计算表
SALARY_SHEET_ID = "VDPOkr"

# 每日销售表字段名（已确认）
SALES_FIELDS = {
    "date": "日期",
    "order_id": "订单号（第几单）",   # 同一单多件填同一个号，用于大单汇总
    "salesperson": "销售员",
    "amount": "实收金额",
    "big_order_bonus": "大单奖励",
}

# 薪资计算表字段名（已确认）
SALARY_FIELDS = {
    "date": "日期",
    "salesperson": "销售员",
    "daily_sales": "当日销售额",
    "staff_count": "当日在岗人数",
    "commission_rate": "日提成比例",   # 文本字段，值为 "2%" / "2.5%" / "3%" / "4%"
    "commission_subtotal": "提成小计",
    "big_order_bonus": "大单奖励",
    "daily_salary": "固定日薪",
    "total_salary": "合计当日薪资",
    "remark": "备注",
}

# 薪资参数
SALARY_PARAMS = {
    # 阶梯提成（按当日店铺总销售额）
    "tiers": [
        {"max": 3000,  "rate": 0.02},
        {"max": 6000,  "rate": 0.025},
        {"max": 10000, "rate": 0.03},
        {"max": None,  "rate": 0.04},
    ],
    # 大单奖励（归成交人）
    "big_order_tiers": [
        {"min": 2000, "bonus": 20},
        {"min": 1000, "bonus": 10},
    ],
    "big_order_daily_limit": 2,  # 每日最多计2笔大单

    # 日薪（按销售员，后续可通过查询薪资表参数区自动获取）
    "daily_salary": {
        "default": 260,  # 新手期默认
        # "里里": 280,   # 稳定期
        # "小何": 260,   # 新手期
    },
}

# 参与薪资计算的销售员白名单（不在此名单的自动排除）
SALARY_WHITELIST = {"里里", "嘉悦"}

# 不参与薪资计算的人员（即使录了销售也不计入）
SALARY_EXCLUDE = {"余淮", "小何"}


# ============ MCP 调用工具 ============

def mcporter_call(tool_name, args):
    """调用腾讯文档 MCP 工具"""
    # 直接用 node 运行 mcporter 的 cli.js，避免 env: node: No such file or directory 错误
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


def list_tables():
    """获取所有工作表"""
    return mcporter_call("smartsheet.list_tables", {"file_id": FILE_ID})


def list_fields(sheet_id, field_titles=None):
    """获取工作表字段列表"""
    args = {"file_id": FILE_ID, "sheet_id": sheet_id}
    if field_titles:
        args["field_titles"] = field_titles
    return mcporter_call("smartsheet.list_fields", args)


def list_records(sheet_id, field_titles=None, offset=0, limit=100):
    """分页获取记录"""
    args = {"file_id": FILE_ID, "sheet_id": sheet_id, "offset": offset, "limit": limit}
    if field_titles:
        args["field_titles"] = field_titles
    return mcporter_call("smartsheet.list_records", args)


def add_records(sheet_id, records):
    """批量添加记录"""
    args = {
        "file_id": FILE_ID,
        "sheet_id": sheet_id,
        "records": records,
    }
    return mcporter_call("smartsheet.add_records", args)


def update_records(sheet_id, records):
    """批量更新记录"""
    args = {
        "file_id": FILE_ID,
        "sheet_id": sheet_id,
        "records": records,
    }
    return mcporter_call("smartsheet.update_records", args)


# ============ 核心计算逻辑 ============

def get_all_records(sheet_id, field_titles=None):
    """获取工作表所有记录（自动分页）"""
    all_records = []
    offset = 0
    while True:
        data = list_records(sheet_id, field_titles=field_titles, offset=offset, limit=100)
        if not data or "records" not in data:
            break
        records = data["records"]
        if not records:
            break
        all_records.extend(records)
        if len(records) < 100:
            break
        offset += 100
        # 避免限流
        import time
        time.sleep(1)
    return all_records


def calc_commission_rate(total_sales):
    """根据当日店铺总销售额计算提成比例"""
    for tier in SALARY_PARAMS["tiers"]:
        if tier["max"] is None or total_sales <= tier["max"]:
            return tier["rate"]
    return SALARY_PARAMS["tiers"][-1]["rate"]


def calc_commission(total_sales, staff_count):
    """计算提成分配：总提成 / 在岗人数，小数向上取整"""
    import math
    rate = calc_commission_rate(total_sales)
    total_commission = total_sales * rate
    per_person = total_commission / staff_count if staff_count > 0 else total_commission
    # 向上取整
    per_person = math.ceil(per_person)
    return total_commission, per_person, rate


def calc_big_order_bonus(sales_records):
    """
    计算大单奖励（按订单号汇总）
    - 同一订单号的多件衣服金额合计后判断
    - 单笔 ≥1000: +10元（取最高档）
    - 单笔 ≥2000: +20元
    - 每日最多计2笔（按人）
    - 归成交人，不平分
    
    Args:
        sales_records: 当天销售记录列表，每条包含:
            - salesperson: 销售员
            - amount: 单件实收金额
            - order_id: 订单号（同一单多件填同一个号）
    """
    # 按订单号汇总金额
    order_summary = defaultdict(lambda: {"amount": 0, "salesperson": ""})
    for rec in sales_records:
        order_id = rec.get("order_id", "") or ""
        # 如果没有订单号，则每条记录视为独立单
        if not order_id:
            order_id = f"_auto_{id(rec)}"
        order_summary[order_id]["amount"] += rec.get("amount", 0)
        # 记录销售员（同一单应该只有一个人）
        sp = rec.get("salesperson", "")
        if sp:
            order_summary[order_id]["salesperson"] = sp

    # 判断哪些订单达到大单门槛
    big_orders = []
    for order_id, info in order_summary.items():
        total_amount = info["amount"]
        sp = info["salesperson"]
        if not sp:
            continue
        
        # 找最高档奖励
        bonus = 0
        for tier in SALARY_PARAMS["big_order_tiers"]:
            if total_amount >= tier["min"] and tier["bonus"] > bonus:
                bonus = tier["bonus"]
        
        if bonus > 0:
            big_orders.append({
                "order_id": order_id,
                "salesperson": sp,
                "amount": total_amount,
                "bonus": bonus,
            })

    # 按金额从高到低排序，每人每日最多计2笔（优先算金额大的）
    big_orders.sort(key=lambda x: x["amount"], reverse=True)

    person_count = defaultdict(int)
    person_bonus = defaultdict(int)
    for order in big_orders:
        sp = order["salesperson"]
        if person_count[sp] < SALARY_PARAMS["big_order_daily_limit"]:
            person_bonus[sp] += order["bonus"]
            person_count[sp] += 1
            print(f"    🎯 大单: {sp} | 订单{order['order_id']} | ¥{int(order['amount'])} → +¥{order['bonus']}")

    return dict(person_bonus)


def calc_daily_salary(salesperson, total_sales, avg_daily_sales=2500):
    """
    根据销售员状态计算日薪
    - 新手期（入职~2周）：¥260/天
    - 稳定期（满2周后）：¥280/天
    - 熟练期（稳定达标满2周）：¥300/天
    """
    if salesperson in SALARY_PARAMS["daily_salary"]:
        return SALARY_PARAMS["daily_salary"][salesperson]
    return SALARY_PARAMS["daily_salary"]["default"]


def summarize_day(sales_records, target_date):
    """
    汇总一天的销售额数据，返回按销售员分组的结果
    
    Args:
        sales_records: 当天的所有销售记录列表
        target_date: 目标日期字符串 "2026-04-08"
    
    Returns:
        dict: {销售员: {日期, 销售员, 当日销售额, ...}}
    """
    # 按销售员分组
    by_person = defaultdict(list)
    for rec in sales_records:
        sp = rec.get("salesperson", "")
        if not sp:
            continue
        by_person[sp].append(rec)

    # 当日店铺总销售额（所有人加在一起）
    total_daily_sales = sum(
        (rec.get("amount") or 0) for rec in sales_records
    )
    staff_count = len(by_person)  # 当日在岗人数

    # 计算提成分配
    total_commission, per_person_commission, commission_rate = calc_commission(
        total_daily_sales, staff_count
    )

    # 计算大单奖励（归成交人）
    big_order_bonuses = calc_big_order_bonus(sales_records)

    # 薪资表日期字段是文本类型，直接传字符串
    _date_str = target_date  # "2026-04-08"

    # 过滤掉不参与薪资计算的人员
    filtered_by_person = {sp: recs for sp, recs in by_person.items() if sp not in SALARY_EXCLUDE}

    # 重新计算在岗人数（只算参与薪资的人）
    staff_count = len(filtered_by_person)

    # 重新计算提成分配（基于过滤后的人数）
    total_commission, per_person_commission, commission_rate = calc_commission(
        total_daily_sales, staff_count
    )

    # 计算大单奖励（基于过滤后的销售记录）
    filtered_records = [r for r in sales_records if r.get("salesperson") not in SALARY_EXCLUDE]
    big_order_bonuses = calc_big_order_bonus(filtered_records)

    # 构建每人结果
    results = {}
    for sp, records in filtered_by_person.items():
        person_sales = sum((r.get("amount") or 0) for r in records)
        person_big_bonus = big_order_bonuses.get(sp, 0)
        person_daily_salary = calc_daily_salary(sp, person_sales)
        person_total = int(per_person_commission + person_big_bonus + person_daily_salary)
        # 提成比例格式化：2% -> "2%", 2.5% -> "2.5%", 去掉末尾多余的0
        rate_percent = commission_rate * 100
        if rate_percent == int(rate_percent):
            commission_rate_str = f"{int(rate_percent)}%"
        else:
            commission_rate_str = f"{rate_percent}%".replace(".0%", "%")

        # 文本字段需要用数组格式 [{"text": "值", "type": "text"}]
        results[sp] = {
            SALARY_FIELDS["date"]: [{"text": _date_str, "type": "text"}],
            SALARY_FIELDS["salesperson"]: [{"text": sp, "type": "text"}],
            SALARY_FIELDS["daily_sales"]: int(person_sales),
            SALARY_FIELDS["staff_count"]: staff_count,
            SALARY_FIELDS["commission_rate"]: [{"text": commission_rate_str, "type": "text"}],
            SALARY_FIELDS["commission_subtotal"]: int(per_person_commission),
            SALARY_FIELDS["big_order_bonus"]: person_big_bonus,
            SALARY_FIELDS["daily_salary"]: person_daily_salary,
            SALARY_FIELDS["total_salary"]: person_total,
            SALARY_FIELDS["remark"]: [{"text": f"总提成¥{int(total_commission)}÷{staff_count}人，日薪档¥{person_daily_salary}", "type": "text"}],
        }

    return results


# ============ 主流程 ============

def main():
    print("=" * 50)
    print("  扭捏niunie 店铺薪资自动计算")
    print("=" * 50)

    # 1. 确定目标日期
    target_date = sys.argv[1] if len(sys.argv) > 1 else date.today().strftime("%Y-%m-%d")
    print(f"\n📅 目标日期: {target_date}")
    print(f"   每日销售表: {SALES_SHEET_ID}")
    print(f"   薪资计算表: {SALARY_SHEET_ID}")

    # 2. 读取每日销售表字段（确认字段存在）
    print(f"\n📋 读取每日销售表字段 (sheet: {SALES_SHEET_ID})...")
    sales_fields_data = list_fields(SALES_SHEET_ID)
    if not sales_fields_data or "fields" not in sales_fields_data:
        print("  ❌ 无法获取字段列表")
        return

    # 打印字段映射
    print("  字段列表:")
    for field in sales_fields_data["fields"]:
        print(f"    - {field.get('title', '?')} (id: {field.get('field_id', '?')}, type: {field.get('field_type', '?')})")

    # 4. 读取每日销售记录
    print(f"\n📋 读取每日销售记录...")
    sales_field_titles = list(SALES_FIELDS.values())
    all_sales = get_all_records(SALES_SHEET_ID, field_titles=sales_field_titles)
    if not all_sales:
        print("  ⚠️  每日销售表无记录")
        return

    print(f"  共 {len(all_sales)} 条记录")

    # 5. 筛选目标日期的记录
    date_field = SALES_FIELDS["date"]

    day_records = []
    for rec in all_sales:
        field_values = rec.get("field_values", {})
        rec_date_raw = field_values.get(date_field, "")

        # API 返回的日期是毫秒时间戳字符串，需要转换
        rec_date = _parse_date(rec_date_raw)
        if not rec_date:
            continue

        # 匹配日期
        if rec_date == target_date:
            # 提取订单号
            order_id = field_values.get(SALES_FIELDS["order_id"], "")
            if isinstance(order_id, list) and order_id:
                order_id = order_id[0].get("text", "") if isinstance(order_id[0], dict) else str(order_id[0])
            elif isinstance(order_id, list):
                order_id = ""
            order_id = str(order_id).strip()

            # 提取销售员
            salesperson = field_values.get(SALES_FIELDS["salesperson"], "")
            if isinstance(salesperson, list) and salesperson:
                salesperson = salesperson[0].get("text", "") if isinstance(salesperson[0], dict) else str(salesperson[0])

            day_records.append({
                "order_id": order_id,
                "salesperson": str(salesperson).strip(),
                "amount": _parse_number(field_values.get(SALES_FIELDS["amount"], 0)),
            })

    if not day_records:
        print(f"  ⚠️  {target_date} 无销售记录")
        return

    print(f"  {target_date} 共 {len(day_records)} 笔销售")

    # 6. 汇总计算
    print(f"\n📊 开始计算...")
    results = summarize_day(day_records, target_date)

    for sp, data in results.items():
        print(f"\n  ── {sp} ──")
        print(f"    当日销售额: ¥{data[SALARY_FIELDS['daily_sales']]}")
        print(f"    在岗人数:   {data[SALARY_FIELDS['staff_count']]}人")
        print(f"    提成比例:   {data[SALARY_FIELDS['commission_rate']]}")
        print(f"    提成小计:   ¥{data[SALARY_FIELDS['commission_subtotal']]}")
        print(f"    大单奖励:   ¥{data[SALARY_FIELDS['big_order_bonus']]}")
        print(f"    固定日薪:   ¥{data[SALARY_FIELDS['daily_salary']]}")
        print(f"    ────────────")
        print(f"    当日合计:   ¥{data[SALARY_FIELDS['total_salary']]}")

    # 7. 检查薪资表是否已有该日期+销售员记录（避免重复）
    print(f"\n📋 检查薪资表是否已有 {target_date} 的记录...")
    salary_records = get_all_records(SALARY_SHEET_ID)
    existing_keys = set()  # (日期, 销售员) 组合
    if salary_records:
        for rec in salary_records:
            vals = rec.get("field_values", {})
            d = _parse_date(vals.get(SALARY_FIELDS["date"], ""))
            sp_raw = vals.get(SALARY_FIELDS["salesperson"], "")
            if isinstance(sp_raw, list) and sp_raw:
                sp_val = sp_raw[0].get("text", "") if isinstance(sp_raw[0], dict) else str(sp_raw[0])
            else:
                sp_val = str(sp_raw)
            if d and sp_val:
                existing_keys.add((d, sp_val.strip()))

    # 筛选需要写入的销售员
    new_records = []
    for sp, data in results.items():
        if (target_date, sp) in existing_keys:
            print(f"  ⏭️  {sp} 在 {target_date} 已有记录，跳过")
            continue
        new_records.append(data)

    if not new_records:
        print(f"\n  ✅ 所有记录已存在，无需新增")
        return

    # 8. 写入薪资计算表
    print(f"\n📋 写入薪资计算表 ({len(new_records)} 条新记录)...")
    for rec in new_records:
        print(f"  → {rec[SALARY_FIELDS['salesperson']]}: ¥{rec[SALARY_FIELDS['total_salary']]}")

    # 转换为 API 需要的格式: [{"field_values": {...}}, ...]
    api_records = [{"field_values": rec} for rec in new_records]
    write_data = add_records(SALARY_SHEET_ID, api_records)
    
    # 调试：打印API返回结果
    print(f"\n  [调试] API返回: {json.dumps(write_data, ensure_ascii=False, indent=2)[:500]}")
    
    # 判断成功：没有error字段，或者有error但值为空/None/False
    if write_data:
        error_val = write_data.get("error")
        if error_val is None or error_val == "" or error_val == False:
            print(f"\n  ✅ 成功写入 {len(new_records)} 条薪资记录！")
        else:
            print(f"\n  ❌ 写入失败: {error_val}")
    else:
        print(f"\n  ❌ 写入失败: 无返回数据")


def _parse_number(val):
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


def _parse_date(val):
    """
    解析日期值，支持：
    - 毫秒时间戳（整数或字符串，如 1775642258133）
    - 日期字符串（如 2026-04-08）
    返回 "YYYY-MM-DD" 格式，无法解析则返回 None
    """
    import datetime
    if not val:
        return None
    # 处理列表
    if isinstance(val, list):
        val = val[0] if val else None
    if not val:
        return None
    val_str = str(val).strip()
    # 毫秒时间戳（13位数字）
    if val_str.isdigit() and len(val_str) >= 10:
        try:
            ts = int(val_str)
            # 毫秒转秒
            if ts > 1e10:
                ts = ts / 1000
            # 使用北京时间（UTC+8）
            tz_cst = datetime.timezone(datetime.timedelta(hours=8))
            return datetime.datetime.fromtimestamp(ts, tz=tz_cst).strftime("%Y-%m-%d")
        except Exception:
            return None
    # 已经是日期字符串
    if len(val_str) >= 10 and val_str[4] == "-":
        return val_str[:10]
    return None


if __name__ == "__main__":
    main()
