#!/bin/bash
# 使用 Personal Access Token 推送到 GitHub

echo "请先去 https://github.com/settings/tokens/new 创建 Token"
echo "需要勾选 'repo' 权限"
echo ""
read -sp "然后粘贴你的 Token 到这里（不会显示）: " TOKEN
echo ""

git remote remove origin 2>/dev/null
git remote add origin "https://danicazhang1022-star:${TOKEN}@github.com/danicazhang1022-star/store-dashboard.git"

echo "正在推送..."
git push -u origin main

# 清除 token
git remote remove origin
git remote add origin "https://github.com/danicazhang1022-star/store-dashboard.git"

echo "✅ 推送完成！"
