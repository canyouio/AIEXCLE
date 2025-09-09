<div align="center">
  <a href="https://gitee.com/yafengke/excel-genius">
    <img src="./docs/img/logo.png" alt="ExcelGenius Logo" width="200" />
  </a>
  <h1 align="center">ExcelGenius</h1>
  <p align="center">
    <strong>:zap: 一款由AI驱动的智能Excel生成、编辑与分析工具，助您通过自然语言驰骋于数据世界。</strong>
  </p>
  <p align="center">
    <a href="https://gitee.com/yafengke/excel-genius/stargazers"><img src="https://gitee.com/yafengke/excel-genius/badge/star.svg?theme=dark" alt="Gitee star"></a>
    <a href="https://gitee.com/yafengke/excel-genius/blob/master/LICENSE"><img src="https://img.shields.io/badge/license-MIT-blue.svg" alt="License"></a>
    <a href="#"><img src="https://img.shields.io/badge/Python-3.10+-blue?logo=python" alt="Python Version"></a>
    <a href="#"><img src="https://img.shields.io/badge/Vue.js-3.x-green?logo=vue.js" alt="Vue.js Version"></a>
    <a href="#"><img src="https://img.shields.io/badge/FastAPI-latest-teal?logo=fastapi" alt="FastAPI"></a>
    <a href="#"><img src="https://img.shields.io/badge/Univer.js-latest-orange" alt="Univer.js"></a>
  </p>
</div>

---

### :sparkles: 目录 (Table of Contents)

- [📖 项目简介](#-项目简介)
- [✨ 为什么选择 ExcelGenius?](#-为什么选择-excelgenius)
- [💡 详细文档分析](#-详细文档分析)
- [🏗️ 项目架构](#️-项目架构)
- [🔧 技术栈](#-技术栈)
- [🛠️ 快速开始](#️-快速开始)
- [📊 API 端点](#-api-端点)
- [📈 监控系统](#-监控系统)
- [💻 项目演示](#-项目演示)
- [🗺️ 未来路线图 (Roadmap)](#️-未来路线图-roadmap)
- [🤝 贡献指南](#-贡献指南)
- [📄 开源许可](#-开源许可)

---

## 📖 项目简介

**ExcelGenius** 是一款将大型语言模型（LLM）与现代Web技术深度融合的智能数据处理工作台。它彻底颠覆了传统的Excel操作模式，旨在：

1.  **赋能所有人：** 无论你是否精通Excel，现在都可以通过简单的自然语言对话，完成从数据生成到深度分析的全流程操作。
2.  **效率革命：** 将过去需要数小时手动操作和编写公式的复杂任务，压缩到几次点击和几句话之内。
3.  **提供洞察：** 不仅仅是处理数据，更是利用AI从数据中挖掘潜在的商业价值、趋势和异常，为您提供决策支持。

项目核心集成了顶级的开源在线表格引擎 **[Univer.js](https://univer.ai/)**，确保了流畅、专业、可扩展的在线编辑体验。

---

## ✨ 为什么选择 ExcelGenius?

| 特性 | 传统方式 | **ExcelGenius 方式** |
| :--- | :--- | :--- |
| **创建表格** | 手动输入标题、调整格式、填充数据 | :speech_balloon: **一句话描述**: "创建一个包含员工姓名、部门、职位、入职日期和薪资的表格" |
| **数据分析** | 筛选、排序、编写复杂的`VLOOKUP`, `SUMIF`公式, 创建数据透视表 | :mag_right: **一键上传**: AI自动分析数据，生成**摘要、洞察、趋势、异常**和**可视化图表** |
| **在线编辑** | 需要昂贵的Office 365订阅或功能受限的Web版 | :computer: **高性能在线编辑**: 由Univer.js驱动，提供流畅的编辑、保存和实时反馈 |
| **系统监控** | 需要专业的运维知识来配置和搭建 | :chart_with_upwards_trend: **开箱即用**: 集成Prometheus & Grafana，提供覆盖全链路的专业级监控仪表盘 |

---

## 💡 详细文档分析

您可以通过点击下方链接，查看由AI驱动生成的、对本项目代码的全面深度分析文档，以获取更完整的技术解读、设计思路与使用指南。

> **:link: 查看详细文档分析: <a href="https://deepwiki.com/canyouio/ExcelGenius" target="_blank">https://deepwiki.com/canyouio/ExcelGenius</a>**

---

## 🏗️ 项目架构

ExcelGenius采用前后端分离的现代化架构，确保了高度的可扩展性和可维护性。

```
├── backend/                  # 后端服务 (Python, FastAPI)
│   ├── config.py             # 配置模块 (AI模型, API密钥, 端口)
│   ├── excel_utils.py        # Excel核心工具 (读/写, 模拟数据)
│   ├── logger_config.py      # 日志配置
│   ├── main.py               # 主API服务 (生成, 编辑, 分析, 保存)
│   └── requirements.txt      # Python依赖
├── frontend/                 # 前端应用 (Vue 3, Vite)
│   ├── public/               # 公共静态资源
│   ├── src/                  # 前端源码
│   │   ├── api/              # API请求封装 & 日志上报
│   │   ├── assets/           # 静态资源 (图片, 全局CSS)
│   │   ├── components/       # Vue核心组件 (主面板, 弹窗, Univer编辑器)
│   │   ├── App.vue           # 根组件
│   │   └── main.js           # 前端入口
│   ├── index.html            # HTML入口
│   ├── package.json          # 前端依赖
│   └── vite.config.js        # Vite构建配置
├── logstash/                 # 日志聚合配置 (可选)
├── prometheus.yml            # Prometheus监控配置
└── excelgenius_monitoring_dashboard.json # Grafana预设仪表盘
```

---

## 🔧 技术栈

| 分类 | 技术 |
| :--- | :--- |
| **后端** | `Python`, `FastAPI`, `DeepSeek API`, `openpyxl`, `uvicorn` |
| **前端** | `Vue 3`, `Vite`, `Univer.js`, `Chart.js`, `Element Plus`, `Tailwind CSS`, `Axios` |
| **可观测性** | `Prometheus`, `Grafana`, `Logstash` (可选) |
| **开发工具** | `VS Code`, `Git`, `Pip`, `NPM` |

---

## 🛠️ 快速开始

### 环境准备

- Python 3.10+ & Pip
- Node.js 18.0+ & NPM
- [Git](https://git-scm.com/)

### 安装与运行

#### 1. 克隆仓库

```bash
git clone https://gitee.com/yafengke/excel-genius.git
cd excel-genius
```

#### 2. 启动后端服务

```bash
cd backend

# (推荐) 创建并激活Python虚拟环境
python -m venv venv
source venv/bin/activate  # on Windows use `venv\Scripts\activate`

# 安装依赖
pip install -r requirements.txt

# 配置环境变量 (如果需要使用AI分析)
# 编辑 .env 文件并填入你的 DEEPSEEK_API_KEY
# DEEPSEEK_API_KEY="YOUR_API_KEY_HERE"

# 启动服务
uvicorn backend.main:app --host 127.0.0.1 --port 8000 --reload
```
:white_check_mark: 后端服务现在应该运行在 `http://127.0.0.1:8000`。

#### 3. 启动前端服务

*打开一个新的终端窗口*
```bash
cd frontend

# 安装依赖
npm install

# 启动开发服务器
npm run dev
```
:white_check_mark: 前端应用现在应该运行在 `http://localhost:3001` (或终端提示的其他端口)。在浏览器中打开它即可开始使用！

---

## 📊 API 端点

| 方法   | 路径                  | 描述                       |
| ------ | --------------------- | -------------------------- |
| `POST` | `/generate_excel`     | 根据文本描述生成Excel并返回数据 |
| `POST` | `/analyze_excel`      | 上传Excel并返回AI分析报告   |
| `POST` | `/api/excel/save_data`| 保存修改后的在线表格数据     |
| `GET`  | `/download/{file_name}` | 下载服务器端的指定Excel文件 |

---

## 📈 监控系统

通过预置的 `prometheus.yml` 和 `excelgenius_monitoring_dashboard.json` 文件，您可以轻松搭建覆盖应用全链路的监控仪表盘，实时洞察系统性能。

---

## 💻 项目演示

<div align="center">
  <img src="./docs/img/examples-demo.png" alt="项目首页" width="700" />
  <p><em>项目主界面与智能分析</em></p>
</div>

<br>

<div align="center">
  <img src="./docs/img/grafana-dashboard.png" alt="Grafana看板" width="700" />
  <p><em>Grafana 监控看板</em></p>
</div>

---

## 🗺️ 未来路线图 (Roadmap)

我们对ExcelGenius的未来充满期待！以下是我们计划中的一些功能：

- [ ] **增强的AI能力**
  - [ ] 支持通过自然语言生成**图表**
  - [ ] 支持通过自然语言**执行公式计算**
  - [ ] 支持更复杂的Excel文件结构（多Sheet页）
- [ ] **在线协作**
  - [ ] 基于WebSocket实现**多人实时协同编辑**
- [ ] **用户系统**
  - [ ] 用户注册与登录
  - [ ] 文件的云端持久化存储
- [ ] **容器化部署**
  - [ ] 提供完整的 `docker-compose.yml` 文件，实现一键部署整个技术栈（包括Prometheus, Grafana）

---

## 🤝 贡献指南

我们热烈欢迎所有形式的贡献！无论是**提交一个Issue**来报告Bug或提出功能建议，还是**提交一个Pull Request**来直接贡献代码，都将帮助ExcelGenius变得更好。

---

## 📄 开源许可

本项目基于 [MIT License](./LICENSE) 开源。