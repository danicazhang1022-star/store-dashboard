#!/bin/bash
# 店铺数据自动同步脚本

WORK_DIR="/Users/danica/WorkBuddy/20260407210337"
LOG_FILE="$WORK_DIR/sync.log"

echo "===== $(date) =====" >> "$LOG_FILE"

cd "$WORK_DIR" || exit 1

# 运行数据生成脚本
/Users/danica/.workbuddy/binaries/python/versions/3.13.12/bin/python3 generate_dashboard_data.py >> "$LOG_FILE" 2>&1

# 检查是否有数据更新
if git diff --quiet data.json; then
    echo "数据无变化" >> "$LOG_FILE"
else
    # 提交并推送
    git add data.json
    git commit -m "自动更新数据 $(date '+%Y-%m-%d %H:%M')" >> "$LOG_FILE" 2>&1
    git push origin main >> "$LOG_FILE" 2>&1
    echo "数据已推送" >> "$LOG_FILE"
fi

echo "" >> "$LOG_FILE"
