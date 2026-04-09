#!/bin/bash
# 店铺智能管理系统 - GitHub 仓库初始化脚本

set -e

echo "=========================================="
echo "  店铺智能管理系统 - GitHub 部署初始化"
echo "=========================================="
echo ""

# 检查 git 是否安装
if ! command -v git &> /dev/null; then
    echo "❌ 错误：未安装 Git，请先安装 Git"
    exit 1
fi

# GitHub 用户名
GITHUB_USER="danicazhang1022-star"
REPO_NAME="store-dashboard"

echo "📋 配置信息："
echo "   GitHub 用户名: $GITHUB_USER"
echo "   仓库名称: $REPO_NAME"
echo ""

# 检查当前目录
if [ ! -f "index.html" ] || [ ! -f "sync_data.py" ]; then
    echo "❌ 错误：请在项目根目录运行此脚本"
    exit 1
fi

echo "🚀 开始初始化..."
echo ""

# 初始化 git 仓库
if [ ! -d ".git" ]; then
    echo "1️⃣  初始化 Git 仓库..."
    git init
    git branch -m main
else
    echo "1️⃣  Git 仓库已存在，跳过初始化"
fi

# 添加文件
echo "2️⃣  添加文件到 Git..."
git add .

# 提交
echo "3️⃣  提交更改..."
git commit -m "Initial commit: 店铺智能管理系统仪表盘" || echo "   无新更改需要提交"

# 检查远程仓库
if git remote get-url origin &> /dev/null; then
    echo "4️⃣  远程仓库已存在，更新 URL..."
    git remote set-url origin "https://github.com/$GITHUB_USER/$REPO_NAME.git"
else
    echo "4️⃣  添加远程仓库..."
    git remote add origin "https://github.com/$GITHUB_USER/$REPO_NAME.git"
fi

echo ""
echo "=========================================="
echo "  ✅ 本地初始化完成！"
echo "=========================================="
echo ""
echo "📌 接下来请在 GitHub 上完成以下步骤："
echo ""
echo "1. 创建仓库"
echo "   访问: https://github.com/new"
echo "   仓库名: $REPO_NAME"
echo "   选择: Public (公开)"
echo "   不要勾选: Add a README"
echo ""
echo "2. 配置 Secrets"
echo "   进入: https://github.com/$GITHUB_USER/$REPO_NAME/settings/secrets/actions"
echo "   添加以下 Secrets："
echo "   • TENCENT_DOCS_TOKEN = 099d0e8b0cad4948806f3e2ce92facf4"
echo "   • FILE_ID = DriTvwXLsoAA"
echo ""
echo "3. 启用 GitHub Pages"
echo "   进入: https://github.com/$GITHUB_USER/$REPO_NAME/settings/pages"
echo "   Source: 选择 GitHub Actions"
echo ""
echo "4. 推送代码到 GitHub"
echo "   运行: git push -u origin main"
echo ""
echo "5. 手动触发部署"
echo "   进入: https://github.com/$GITHUB_USER/$REPO_NAME/actions"
echo "   点击: Deploy Dashboard → Run workflow"
echo ""
echo "📊 部署完成后访问："
echo "   https://$GITHUB_USER.github.io/$REPO_NAME/"
echo ""
