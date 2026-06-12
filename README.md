# 产品目录自动生成工具

内部使用的产品目录系统 — 上传 Excel 自动生成网页预览、PDF 目录和 PNG 总览图。

适用于向客户快速推送产品信息：PNG 快速预览 → PDF 详细资料。

## 快速开始

### 本地运行

```bash
# 1. 安装依赖
pip install -r requirements.txt

# 2. 安装 Playwright 浏览器 (首次需要)
playwright install chromium

# 3. 启动服务
cd app
uvicorn main:app --host 0.0.0.0 --port 8000 --reload

# 4. 打开浏览器访问
# http://localhost:8000
```

**注意**: 如果从项目根目录启动，需使用:
```bash
uvicorn app.main:app --host 0.0.0.0 --port 8000 --reload
```

### Docker 本地测试

```bash
# 构建镜像
docker build -t product-catalog .

# 运行容器
docker run -p 8000:8000 product-catalog

# 访问 http://localhost:8000
```

## Railway 部署

### 一键部署

1. 将项目推送到 GitHub
2. 在 [Railway](https://railway.com) 新建项目 → **Deploy from GitHub repo**
3. 选择本仓库 → Railway 自动读取 `Dockerfile` 构建
4. 构建完成（约 3-5 分钟）后即可访问

### 手动配置

Railway 会自动注入 `PORT` 环境变量，无需手动设置。
如需自定义端口或其他变量，在 Railway Dashboard → Variables 中添加。

### Dockerfile 说明

基于 `python:3.13-slim`，自动安装：
- Playwright 所需的系统库（libnss3, libgbm1, libcairo2 等）
- Playwright Chromium 浏览器
- CJK 字体（fonts-noto-cjk，支持中文）

## 使用流程

```
上传 Excel → 预览产品 → 生成 PDF → 下载 PNG
```

### 1. 上传 Excel

打开网页 → 选择 `.xlsx` 文件 → 点击 **Generate Preview**

系统自动解析 Excel 并在页面展示产品卡片预览。

### 2. 生成 PDF

在预览区点击 **Download PDF** 下载，或 **View PDF** 浏览器内预览。

PDF 包含：封面 → 产品总览 → 逐品详情（含卖点+规格+二维码） → 对比表 → 备注页

### 3. 生成 PNG

点击 **Download Overview PNG** 下载产品总览宣传图。

适合微信、邮件、WhatsApp 发送，客户快速浏览后点击 PDF 查看详情。

## Excel 字段要求

Excel 第一行必须是以下字段名（**区分大小写**）：

| 字段 | 说明 | 必填 |
|------|------|------|
| Product Name | 产品名称 | 否 |
| Brand | 品牌 | 否 |
| Price | 价格 | 否 |
| Category | 类目 | 否 |
| Weight | 重量 | 否 |
| Dimensions | 尺寸 | 否 |
| Selling Point 1 | 卖点 1 | 否 |
| Selling Point 2 | 卖点 2 | 否 |
| Selling Point 3 | 卖点 3 | 否 |
| Promotion Angle | 推广方向 | 否 |
| Product URL | 产品链接 | 否 |
| Image URL | 产品图片链接 | 否 |

- 所有字段均为可选，缺失的在页面上留空
- 空行自动跳过
- Image URL 为空时自动显示占位图

## 项目结构

```
product-catalog/
├── app/
│   ├── main.py                  # FastAPI 后端
│   ├── templates/
│   │   ├── index.html           # Web 前端页面
│   │   ├── catalog.html         # PDF 模板
│   │   └── overview.html        # PNG 总览模板
│   └── static/
│       ├── uploads/             # 上传文件暂存
│       └── outputs/             # 生成文件输出
├── sample_products.xlsx         # 测试数据
├── requirements.txt             # Python 依赖
├── Dockerfile                   # Docker 构建
├── railway.json                 # Railway 配置
└── README.md
```

## 常见问题

### 本地 PDF/PNG 生成失败

```bash
# 确认 Playwright 浏览器已安装
playwright install chromium

# 确认系统依赖齐全 (Ubuntu/Debian)
apt-get install -y libnss3 libgbm1 libcairo2 libpango-1.0-0
```

### Railway 部署后页面无法访问

1. 检查 Railway 构建日志，确认 `playwright install chromium --with-deps` 执行成功
2. 检查 `PORT` 环境变量是否正确
3. 确认 `railway.json` 中 `healthcheckPath` 为 `/`

### 中文显示为方框

Dockerfile 已安装 `fonts-noto-cjk` 中文支持字体。如果仍然出现方框：
- 确认 Dockerfile 中字体安装步骤未被删除
- 重新构建镜像

### Excel 上传后无产品数据

- 检查表头字段名是否完全匹配（区分大小写）
- 确认文件名以 `.xlsx` 结尾
- 检查 Excel 第 1 行为表头、第 2 行起为数据

### PDF/PNG 生成超时

- 产品图片 URL 较多时，Playwright 需要等待加载。默认超时 30 秒
- 如果图片 URL 不可访问，系统会自动跳过，不影响生成

## 技术栈

| 组件 | 技术 |
|------|------|
| 后端框架 | FastAPI |
| 模板引擎 | Jinja2 |
| Excel 解析 | openpyxl |
| PDF/PNG 渲染 | Playwright (Chromium) |
| QR 码生成 | qrcode[pil] |
| 部署 | Docker + Railway |
