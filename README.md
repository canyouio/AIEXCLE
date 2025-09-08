<div align="center">
  <img src="./docs/img/logo.png" alt="ExcelGenius Logo" width="200" />
  <h1>ExcelGenius</h1>
  <p>
    <strong>一款基于AI的智能Excel生成与编辑工具，能够通过自然语言描述快速创建和修改Excel表格，大大提高办公效率。</strong>
  </p>
  <p>
    <a href="https://gitee.com/yafengke/excel-genius/stargazers"><img src="https://gitee.com/yafengke/excel-genius/badge/star.svg?theme=dark" alt="star"></a>
    <a href="https://gitee.com/yafengke/excel-genius/blob/master/LICENSE"><img src="https://img.shields.io/github/license/dream-num/excelgenius?style=flat-square" alt="license"></a>
    <a href="https://github.com/dream-num/excelgenius/actions/workflows/build.yml"><img src="https://img.shields.io/github/actions/workflow/status/dream-num/excelgenius/build.yml?style=flat-square" alt="build"></a>
  </p>
</div>

## 📖 项目简介

ExcelGenius 是一款由AI驱动的智能Excel生成与编辑工具，致力于通过自然语言交互，为用户带来前所未有的智能化办公体验。告别繁琐的手动操作，让Excel处理变得轻松高效。

---

## 🌈 项目亮点

- 🧠 **AI驱动**：通过自然语言描述快速生成和编辑Excel表格，无需手动操作。
- 🚀 **高效便捷**：极大提升办公效率，降低Excel操作的学习成本。
- 📊 **可视化编辑**：提供直观的表格编辑界面，支持基本的编辑和格式设置。
- 🔌 **灵活扩展**：模块化设计，易于扩展新功能和集成其他AI服务。
- 🔧 **监控系统**：集成Prometheus和Grafana监控，实时跟踪应用性能。
- 🖥️ **前后端分离**：基于FastAPI和Vue3的现代化架构，提供流畅的用户体验。
- 📱 **多平台兼容**：支持在不同操作系统和设备上使用。

---

## ✨ 功能特性

### 📊 Excel生成与编辑

- **自然语言生成Excel**：通过简单的文字描述，快速生成符合需求的Excel表格。
- **智能编辑现有Excel**：上传Excel文件后，通过自然语言指令进行编辑操作。
- **在线表格编辑器**：提供直观的表格编辑界面，支持基本的编辑和格式设置功能。
- **模拟数据模式**：在没有OpenAI API密钥的情况下，使用内置的模拟数据展示功能。

### 🔧 系统功能

- **文件管理**：支持Excel文件的上传分析、下载和临时存储。
- **配置管理**：灵活的配置选项，支持模型选择和端口设置。
- **异常处理**：完善的错误处理和提示机制。
- **监控集成**：与Prometheus和Grafana无缝集成，提供实时性能监控。

---

## 🏗️ 项目架构

ExcelGenius采用前后端分离的现代化架构设计，确保系统的可扩展性和可维护性。

```
├── backend/                  # 后端服务 (Python)
│   ├── config.py             # 配置 (AI模型选择、端口)
│   ├── excel_utils.py        # Excel生成/编辑工具类
│   ├── logger_config.py      # 日志配置
│   ├── main.py               # 核心API (自然语言解析、Excel生成)
│   └── requirements.txt      # Python依赖
├── frontend/                 # 前端页面 (Vue3)
│   ├── public/               # 公共资源
│   ├── src/                  # 前端源码
│   │   ├── api/              # API请求工具
│   │   ├── assets/           # 静态资源
│   │   ├── components/       # Vue组件
│   │   ├── App.vue           # 根组件
│   │   └── main.js           # 前端入口逻辑
│   ├── index.html            # 入口页面
│   ├── package.json          # 前端依赖配置
│   └── vite.config.js        # Vite构建配置
├── logstash/                 # 日志处理配置
├── prometheus.yml            # Prometheus监控配置
└── excelgenius_monitoring_dashboard.json # Grafana看板配置
```
---

## 🔧 技术栈

| 分类       | 技术                                        |
| ---------- | ------------------------------------------- |
| **后端**   | Python, FastAPI, OpenAI API, openpyxl       |
| **前端**   | Vue 3, Element Plus, Tailwind CSS, Axios, XLSX.js |
| **构建工具** | Vite                                        |
| **监控系统** | Prometheus, Grafana                         |
| **日志管理** | Logstash                                    |

---

## 🚀 快速开始

### 1. Vercel一键部署 (推荐)

[![Deploy to Vercel](https://vercel.com/button)](https://vercel.com/new/clone?repository-url=https://gitee.com/yafengke/excel-genius)

点击上方按钮，即可轻松将ExcelGenius部署到Vercel平台。

### 2. 本地部署

**环境准备：**

- Python 3.12+
- Node.js 20.0+
- OpenAI API 密钥 (可选，若不提供将使用模拟数据)
- Prometheus & Grafana (用于监控功能)

#### 后端安装与运行

1.  **进入后端目录**
    ```bash
    cd backend
    ```
2.  **安装依赖**
    ```bash
    pip install -r requirements.txt
    ```
3.  **配置环境变量 (可选)**
    - 创建 `.env` 文件。
    - 在文件中添加 `DEEPSEEK_API_KEY="您的API密钥"`。
4.  **启动后端服务**
    ```bash
    uvicorn backend.main:app --host 127.0.0.1 --port 8000 --reload
    ```
    服务将运行在 `http://127.0.0.1:8000`。

#### 前端安装与运行

1.  **进入前端目录**
    ```bash
    cd frontend
    ```
2.  **安装依赖**
    ```bash
    npm install
    ```
3.  **启动前端开发服务器**
    ```bash
    npm run dev
    ```
    前端页面将运行在 `http://127.0.0.1:3001`。

---

## 📊 API 端点

| 方法   | 路径                  | 描述             |
| ------ | --------------------- | ---------------- |
| `POST` | `/generate_excel`     | 根据文本描述生成Excel |
| `POST` | `/edit_excel`         | 上传并根据指令编辑Excel |
| `GET`  | `/download/{file_name}` | 下载指定的Excel文件   |

---

## 📈 监控系统

ExcelGenius集成了完善的监控系统，助您实时掌握应用的运行状态。

### Prometheus 配置

项目根目录下的 `prometheus.yml` 文件包含了详细的监控配置，主要监控目标包括：
- Prometheus自身
- ExcelGenius后端服务 (端口 `8000`)
- Windows系统资源 (通过 `windows-exporter`，端口 `9182`)

### Grafana 看板

我们提供了预配置的Grafana看板文件 `excelgenius_monitoring_dashboard.json`，导入后即可获得丰富的监控视图：
- 系统资源使用情况 (CPU, 内存, 磁盘)
- API请求统计 (QPS, 延迟)
- Excel生成/编辑性能指标
- 错误日志监控

---

## 🐱 项目演示

<div align="center">
  <img src="docs/img/examples-demo.png" alt="项目首页" width="700" />
  <p><em>项目首页</em></p>
</div>

<br>

<div align="center">
  <img src="docs/img/grafana-dashboard.png" alt="Grafana看板" width="700" />
  <p><em>Grafana 监控看板</em></p>
</div>

---

## 💻 使用说明

### 生成Excel
1.  在首页的"**生成Excel**"标签页中，输入您想要的Excel内容描述。
2.  (可选) 输入自定义文件名。
3.  点击"**生成Excel**"按钮。
4.  生成完成后，可以**下载文件**或在**表格编辑器**中查看。

### 编辑Excel
1.  在首页的"**编辑Excel**"标签页中，选择要编辑的Excel文件。
2.  输入编辑指令，例如："在表格末尾添加一行汇总数据"。
3.  点击"**编辑Excel**"按钮。
4.  编辑完成后，可以**下载文件**或在**表格编辑器**中查看。

### 表格编辑器
1.  在首页的"**表格编辑器**"标签页中，可以查看和编辑Excel内容。
2.  支持**添加行**、**添加列**、**删除选中内容**、**清空表格**等操作。
3.  支持基本的格式设置，如字体大小、粗体、斜体、下划线和文字颜色。
4.  可以**导出**编辑后的内容为Excel文件。

---

## ⚠️ 注意事项

1.  **OpenAI API 密钥**：如果没有配置有效的OpenAI API密钥，系统将自动切换到**模拟数据模式**。
2.  **文件存储**：生成和编辑的Excel文件默认存储在项目的 `temp` 目录中。
3.  **模拟数据模式**：在此模式下，系统会根据描述生成预定义的示例数据，仅供演示使用。
4.  **安全性**：请勿在公共环境中暴露包含敏感信息的API密钥。

---

### 部署建议

- **后端**：推荐使用 Uvicorn 或 Gunicorn 作为ASGI服务器。
- **前端**：构建后的静态文件可以使用 Nginx 等Web服务器进行托管。
- **监控系统**：可以与主应用一同部署，或独立部署。
- **容器化**：建议使用 Docker 将整个应用容器化，以便于管理和迁移。

---

## 🤝 贡献

我们欢迎各种形式的贡献，您可以：
- 提交Bug报告和功能请求
- 改进文档和示例代码
- 提交代码贡献 (Pull Request)
- 分享您的使用经验和案例

---

## ❤️ 赞助

ExcelGenius 的持续发展离不开社区的支持。如果您觉得这个项目对您有帮助，请考虑成为我们的赞助者，您的支持是我们前进的最大动力！

---

## 📄 许可

本项目采用 [MIT License](./LICENSE) 开源许可。