# 店铺智能管理系统 - 数据仪表盘

自动同步腾讯文档智能表格数据，生成可视化仪表盘，部署到 GitHub Pages。

## 功能特性

- 📊 **实时数据展示**：今日/本月销售额、库存状态、销售排行
- ⏰ **自动同步**：每天北京时间晚上10点自动更新数据
- 🖱️ **手动刷新**：支持随时手动触发数据同步
- 📱 **移动端适配**：手机随时查看最新数据
- 💰 **零成本部署**：使用 GitHub Actions + GitHub Pages，完全免费

## 部署步骤

### 1. 创建 GitHub 仓库

在你的电脑上执行以下命令：

```bash
# 进入项目目录
cd /Users/danica/WorkBuddy/20260407210337

# 初始化 Git 仓库
git init

# 添加所有文件
git add .

# 提交
git commit -m "Initial commit: 店铺智能管理系统仪表盘"

# 添加远程仓库（用你的用户名替换 danicazhang1022-star）
git remote add origin https://github.com/danicazhang1022-star/store-dashboard.git

# 推送到 GitHub
git push -u origin main
```

### 2. 配置 GitHub Secrets

在 GitHub 仓库页面，进入 **Settings → Secrets and variables → Actions**，添加以下 Secrets：

| Secret 名称 | 值 |
|------------|-----|
| `TENCENT_DOCS_TOKEN` | `099d0e8b0cad4948806f3e2ce92facf4` |
| `FILE_ID` | `DriTvwXLsoAA` |

### 3. 启用 GitHub Pages

1. 进入仓库 **Settings → Pages**
2. Source 选择 **GitHub Actions**

### 4. 运行工作流

1. 进入仓库的 **Actions** 页面
2. 点击 **Deploy Dashboard** 工作流
3. 点击 **Run workflow** 手动触发首次部署

## 访问仪表盘

部署完成后，访问地址：

```
https://danicazhang1022-star.github.io/store-dashboard/
```

## 数据更新

- **自动更新**：每天北京时间晚上10点（UTC 14:00）
- **手动更新**：在 Actions 页面点击 "Run workflow"

## 项目结构

```
.
├── .github/
│   └── workflows/
│       └── deploy.yml      # GitHub Actions 工作流
├── index.html              # 仪表盘页面
├── sync_data.py            # 数据同步脚本
└── README.md               # 本文件
```

## 数据来源

- 腾讯文档智能表格：店铺智能管理系统
- 文档 ID：`DriTvwXLsoAA`

## 技术栈

- 前端：HTML + Tailwind CSS + Vanilla JS
- 后端：Python + GitHub Actions
- 部署：GitHub Pages
