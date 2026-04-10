#!/bin/bash
# 店铺薪资自动同步脚本 - 每晚10点运行

WORK_DIR="/Users/danica/WorkBuddy/20260407210337"
LOG_FILE="$WORK_DIR/salary_sync.log"

echo "===== $(date) =====" >> "$LOG_FILE"

cd "$WORK_DIR" || exit 1

# 运行薪资计算脚本（计算今天）
/Users/danica/.workbuddy/binaries/python/versions/3.13.12/bin/python3 auto_salary.py >> "$LOG_FILE" 2>&1

# 也可以计算昨天（防止漏算）
# /Users/danica/.workbuddy/binaries/python/versions/3.13.12/bin/python3 auto_salary.py $(date -v-1d +%Y-%m-%d) >> "$LOG_FILE" 2>&1

echo "薪资计算完成" >> "$LOG_FILE"
echo "" >> "$LOG_FILE"
