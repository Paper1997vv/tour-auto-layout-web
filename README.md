# TourAutoLayout Web

面向个人办公场景的旅游行程自动排版 Web 工具。当前版本支持上传模板底图和 `.doc/.docx` 文档，在服务器端生成可继续编辑的 `.docx` 成品，并通过浏览器下载结果。

## 功能

- 浏览器上传模板图和多份旅游文档
- 支持 `.doc` 自动转换为 `.docx`
- 整篇迁移源文档正文到模板页内，尽量保留原顺序和常见格式
- 批量任务异步处理，支持单文件下载和整批 ZIP 下载
- 固定输出版式：
  - A4 竖版
  - 页边距：上 `4.1cm`，下/左/右 `1cm`
  - 模板底图写入页眉背景

## 本地运行

```bash
cd /Users/X/Documents/自动化转化
swift run TourAutoLayoutWeb
```

默认访问地址：

- `http://127.0.0.1:8080`

可选环境变量：

- `PORT`
- `STORAGE_ROOT`
- `MAX_CONCURRENT_JOBS`
- `MAX_UPLOAD_SIZE_MB`
- `PUBLIC_BASE_URL`
- `ACCESS_PASSWORD`
- `LIBREOFFICE_PATH`

## Docker 部署

```bash
cd /Users/X/Documents/自动化转化
docker compose up --build
```

启动后默认映射：

- Web 服务：`http://127.0.0.1:8080`
- 本地持久化目录：`./storage`

## 测试

```bash
cd /Users/X/Documents/自动化转化
swift test
```

## 项目结构

- [Package.swift](/Users/X/Documents/自动化转化/Package.swift)：Swift Package 定义
- [Sources/TourAutoLayoutCore](/Users/X/Documents/自动化转化/Sources/TourAutoLayoutCore)：共享文档处理核心
- [Sources/TourAutoLayoutWeb](/Users/X/Documents/自动化转化/Sources/TourAutoLayoutWeb)：Vapor Web 服务
- [Public](/Users/X/Documents/自动化转化/Public)：浏览器前端静态页面
- [docker-compose.yml](/Users/X/Documents/自动化转化/docker-compose.yml)：本地容器编排

## 部署说明

- GitHub 负责托管代码仓库，不直接承载运行服务。
- 若需要公网访问，请把仓库部署到支持 Docker 的服务器或平台，例如 VPS、Railway、Render、Fly.io。
- 若启用 `ACCESS_PASSWORD`，页面会先要求输入访问密码，然后才能创建和下载任务。
