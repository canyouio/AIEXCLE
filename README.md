# ExcelGenius - 智能Excel生成与编辑工具

ExcelGenius是一款基于AI的智能Excel生成与编辑工具，能够通过自然语言描述快速创建和修改Excel表格，大大提高办公效率。

## 项目结构

```
├─ backend/           # 后端服务（Python） 
│  ├─ main.py         # 核心API（自然语言解析、Excel生成） 
│  ├─ config.py       # 配置（AI模型选择、端口） 
│  ├─ excel_utils.py  # Excel生成/编辑工具类 
│  └─ .env            # 环境变量（如OpenAI APIKey，可选） 
├─ frontend/          # 前端页面（Vue3） 
│  ├─ index.html      # 入口页面 
│  ├─ package.json    # 前端依赖配置 
│  ├─ vite.config.js  # Vite构建配置
│  ├─ src/ 
│  │  ├─ main.js      # 前端入口逻辑 
│  │  ├─ App.vue      # 根组件
│  │  ├─ components/ 
│  │  │  └─ ExcelEditor.vue  # 表格可视化编辑组件 
│  │  └─ api/ 
│  │     └─ request.js       # 后端API请求工具 
└─ temp/              # 临时文件目录（存储生成的Excel，本地自动创建）
```

## 功能特点

1. **自然语言生成Excel**：通过简单的文字描述，快速生成符合需求的Excel表格
2. **智能编辑现有Excel**：上传Excel文件后，通过自然语言指令进行编辑操作
3. **在线表格编辑器**：提供直观的表格编辑界面，支持基本的编辑和格式设置功能
4. **模拟数据模式**：在没有OpenAI API密钥的情况下，使用内置的模拟数据展示功能

## 技术栈

- **后端**：Python, FastAPI, OpenAI API, openpyxl
- **前端**：Vue 3, Element Plus, Tailwind CSS, Axios, XLSX.js
- **构建工具**：Vite

## 快速开始

### 前置要求

- Python 3.8+ 和 Node.js 16+ 环境
- OpenAI API 密钥（可选，不提供时将使用模拟数据）

### 后端安装与运行

1. 进入backend目录
```bash
cd backend
```

2. 安装依赖
```bash
pip install fastapi uvicorn openai python-dotenv openpyxl
```

3. 配置环境变量（可选）
   - 复制 `.env` 文件中的示例配置
   - 将 `OPENAI_API_KEY` 设置为您的实际API密钥

4. 启动后端服务
```bash
python main.py
```

后端服务默认运行在 `http://127.0.0.1:8000`

### 前端安装与运行

1. 进入frontend目录
```bash
cd frontend
```

2. 安装依赖
```bash
npm install
```

3. 启动前端开发服务器
```bash
npm run dev
```

前端服务默认运行在 `http://127.0.0.1:3000`

## API 端点

### 生成Excel
- **POST** `/generate_excel`
- **参数**：
  - `description`：Excel内容描述（必填）
  - `file_name`：文件名（可选）
- **返回**：生成的Excel文件信息和下载链接

### 编辑Excel
- **POST** `/edit_excel`
- **参数**：
  - `file`：上传的Excel文件（必填）
  - `instructions`：编辑指令（必填）
- **返回**：编辑后的Excel文件信息和下载链接

### 下载文件
- **GET** `/download/{file_name}`
- **返回**：Excel文件下载

## 使用说明

### 生成Excel
1. 在首页的"生成Excel"标签页中，输入您想要的Excel内容描述
2. 可选：输入文件名
3. 点击"生成Excel"按钮
4. 生成完成后，可以下载文件或在表格编辑器中查看

### 编辑Excel
1. 在首页的"编辑Excel"标签页中，选择要编辑的Excel文件
2. 输入编辑指令，例如"在表格末尾添加一行汇总数据"
3. 点击"编辑Excel"按钮
4. 编辑完成后，可以下载文件或在表格编辑器中查看

### 表格编辑器
1. 在首页的"表格编辑器"标签页中，可以查看和编辑Excel内容
2. 支持添加行、添加列、删除选中内容、清空表格等操作
3. 支持基本的格式设置，如字体大小、粗体、斜体、下划线和文字颜色
4. 可以导出编辑后的内容为Excel文件

## 注意事项

1. **OpenAI API 密钥**：如果没有配置有效的OpenAI API密钥，系统将自动切换到模拟数据模式
2. **文件存储**：生成和编辑的Excel文件默认存储在项目的 `temp` 目录中
3. **模拟数据模式**：在模拟数据模式下，系统会根据描述生成预定义的示例数据，仅供演示使用
4. **安全性**：不要在公共环境中暴露包含敏感信息的API密钥

## 开发说明

### 后端开发
- 配置文件位于 `config.py`，可以修改端口、模型等配置
- Excel处理逻辑位于 `excel_utils.py`，可以扩展更多Excel操作功能
- API端点位于 `main.py`，可以添加更多API功能

### 前端开发
- 使用Vite作为构建工具，可以通过 `vite.config.js` 配置构建选项
- 组件位于 `src/components` 目录
- API请求工具位于 `src/api/request.js`

## 构建与部署

### 前端构建
```bash
cd frontend
npm run build
```
构建后的文件将位于项目根目录的 `dist` 目录中

### 部署建议
- 后端可以使用Uvicorn或Gunicorn作为ASGI服务器
- 前端构建后的静态文件可以使用Nginx等Web服务器提供服务
- 可以使用Docker容器化部署整个应用

## License

MIT License