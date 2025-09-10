import os
import re
import json
import time
from fastapi import FastAPI, UploadFile, File, Form, HTTPException, Body, WebSocket, WebSocketDisconnect
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
import uvicorn
import requests
import logging
from redis import Redis
from typing import List
from .excel_utils import generate_analysis_report
# 导入Prometheus监控库
from prometheus_client import Counter, Histogram, Gauge, generate_latest, CONTENT_TYPE_LATEST
from .config import Config
from .excel_utils import ExcelUtils
from .logger_config import api_logger, excel_logger
from typing import Dict, Any


# 初始化FastAPI应用
app = FastAPI(title="ExcelGenius API", description="自然语言生成和编辑Excel文件")

origins = [
    "http://localhost:3001",
    "http://127.0.0.1:3001"
]

app.add_middleware(
      CORSMiddleware,
    allow_origins=["*"], # 允许所有来源，在开发中最方便
    allow_credentials=True,
    allow_methods=["GET", "POST", "PUT", "DELETE", "OPTIONS"],
    allow_headers=["*"],
)

# 初始化Prometheus监控指标

# Token消耗相关指标
token_consumed = Counter(
    'excelgenius_token_consumed_total',
    'Total tokens consumed by the application',
    ['api_endpoint']
)

token_consumed_current = Gauge(
    'excelgenius_token_consumed_current',
    'Current tokens consumed by the application',
    ['api_endpoint']
)

# API请求相关指标
api_requests_total = Counter(
    'excelgenius_api_requests_total',
    'Total number of API requests',
    ['api_endpoint', 'status_code']
)

api_request_duration_seconds = Histogram(
    'excelgenius_api_request_duration_seconds',
    'API request processing time in seconds',
    ['api_endpoint']
)

# Excel处理相关指标
excel_files_processed = Counter(
    'excelgenius_excel_files_processed_total',
    'Total number of Excel files processed',
    ['operation_type']  # generate, edit, analyze
)

excel_rows_processed = Counter(
    'excelgenius_excel_rows_processed_total',
    'Total number of Excel rows processed',
    ['operation_type']
)

excel_processing_time_seconds = Histogram(
    'excelgenius_excel_processing_time_seconds',
    'Excel processing time in seconds',
    ['operation_type']
)

# 系统资源相关指标
active_requests = Gauge(
    'excelgenius_active_requests',
    'Number of active requests being processed'
)

# 成功率指标
success_rate = Gauge(
    'excelgenius_success_rate',
    'Success rate of API requests',
    ['api_endpoint']
)

# 配置CORS
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# 初始化配置和工具类
config = Config()
excel_utils = ExcelUtils()

# 初始化 Redis（使用配置或默认）
redis_client = Redis(host=os.getenv("REDIS_HOST", "localhost"), port=int(os.getenv("REDIS_PORT", "6379")), db=0, decode_responses=True)

# 活跃连接管理： instance_id -> websocket
active_connections: dict[str, WebSocket] = {}

class ExcelWebSocketManager:
    async def connect(self, websocket: WebSocket, instance_id: str):
        await websocket.accept()
        active_connections[instance_id] = websocket
        # 如果Redis有会话数据则推送回前端
        cached = redis_client.get(f"excel:{instance_id}")
        if cached:
            await websocket.send_text(json.dumps({"type": "load_history", "data": json.loads(cached)}))

    def disconnect(self, instance_id: str):
        if instance_id in active_connections:
            del active_connections[instance_id]

    async def send_to(self, instance_id: str, message: dict):
        ws = active_connections.get(instance_id)
        if ws:
            await ws.send_text(json.dumps(message))

ws_manager = ExcelWebSocketManager()

def save_history(instance_id: str, excel_data: list):
    key = f"excel:history:{instance_id}"
    history = json.loads(redis_client.get(key) or "[]")
    history.append({"timestamp": int(time.time()), "data": excel_data})
    # 保留最近 10 条
    if len(history) > 10:
        history = history[-10:]
    redis_client.set(key, json.dumps(history))

@app.websocket("/ws/excel/{instance_id}")
async def websocket_excel(websocket: WebSocket, instance_id: str):
    await ws_manager.connect(websocket, instance_id)
    try:
        while True:
            text = await websocket.receive_text()
            op = json.loads(text)
            op_type = op.get("type")
            # 从 Redis 读取当前会话数据（Luckysheet format）
            raw = redis_client.get(f"excel:{instance_id}") or "[]"
            excel_data = json.loads(raw)

            # 简化示例：仅处理 edit_cell 和 insert_row（真实项目需适配 Luckysheet 的完整事件）
            if op_type == "edit_cell":
                r = op["data"]["row"]
                c = op["data"]["col"]
                v = op["data"]["value"]
                # 确保行存在
                while r >= len(excel_data):
                    excel_data.append({"cell": {}})
                row_cell = excel_data[r].setdefault("cell", {})
                row_cell[str(c)] = v

            elif op_type == "insert_row":
                r = op["data"]["row"]
                count = op["data"].get("count", 1)
                for _ in range(count):
                    excel_data.insert(r, {"cell": {}})

            # TODO: handle insert_col, merge_cell, formula updates, styles...

            # 写回 Redis（过期2小时）
            redis_client.setex(f"excel:{instance_id}", 7200, json.dumps(excel_data))

            # 保存历史记录
            save_history(instance_id, excel_data)

            # 自动保存提示：每30秒发送一次（示例逻辑）
            last_save = int(redis_client.get(f"excel:last_save:{instance_id}") or "0")
            if int(time.time()) - last_save > 30:
                await ws_manager.send_to(instance_id, {"type": "auto_save_success", "message": "已自动保存"})
                redis_client.set(f"excel:last_save:{instance_id}", int(time.time()))

    except WebSocketDisconnect:
        ws_manager.disconnect(instance_id)
    except Exception as e:
        api_logger.error(f"WebSocket error: {e}")
        ws_manager.disconnect(instance_id)

@app.get("/api/excel/load/{file_id}")
async def load_excel(file_id: str, instance_id: str):
    upload_file_path = os.path.join(config.TEMP_DIR, f"{file_id}.xlsx")
    if not os.path.exists(upload_file_path):
        raise HTTPException(status_code=404, detail="文件不存在")
    luckysheet_data = excel_utils.parse_excel_to_luckysheet(upload_file_path)
    redis_client.setex(f"excel:{instance_id}", 7200, json.dumps(luckysheet_data))
    return {"status": "success", "data": luckysheet_data}

@app.post("/api/excel/save/{instance_id}")
async def save_excel(instance_id: str, fileName: str = Body(..., embed=True)):
    cached = redis_client.get(f"excel:{instance_id}")
    if not cached:
        raise HTTPException(status_code=400, detail="无编辑数据，请先加载Excel")
    luckysheet_data = json.loads(cached)
    # 将 Luckysheet JSON 转回 Excel 文件（使用 ExcelUtils 或 pyexcelerate）
    output_file = excel_utils.luckysheet_to_xlsx(luckysheet_data, file_name=fileName)
    return {"status": "success", "fileName": output_file, "downloadUrl": f"/download/{output_file}"}

@app.post("/api/excel/upload")
async def upload_excel(file: UploadFile = File(...)):
    # 生成 file_id 并保存到 TEMP_DIR（兼容现有结构）
    file_id = f"upload_{int(time.time())}_{file.filename.replace('.xlsx','')}"
    upload_path = os.path.join(config.TEMP_DIR, f"{file_id}.xlsx")
    with open(upload_path, "wb") as f:
        f.write(await file.read())
    return {"status": "success", "file_id": file_id}



# 设置DeepSeek API密钥
api_logger.info(f"config.DEEPSEEK_API_KEY存在: {bool(config.DEEPSEEK_API_KEY)}")
api_logger.info(f"config.DEEPSEEK_API_KEY != 'placeholder-api-key': {config.DEEPSEEK_API_KEY != 'placeholder-api-key'}")

if config.DEEPSEEK_API_KEY and config.DEEPSEEK_API_KEY != "placeholder-api-key":
    use_mock = False
    api_logger.info("已成功启用DeepSeek API模式")
else:
    use_mock = True
    api_logger.warning("使用模拟数据模式：未配置有效的DeepSeek API密钥")

# 从excel_utils导入模拟响应生成函数
from .excel_utils import generate_mock_response, generate_analysis_report

class GenerateExcelRequest(BaseModel):
    description: str
    file_name: str = None

@app.post("/api/generate_excel")
async def generate_excel(
    request: GenerateExcelRequest
):
    """
    根据自然语言描述生成Excel文件
    """
    endpoint = "/generate_excel"
    active_requests.inc()
    start_time = time.time()
    
    try:
        # 记录API请求
        api_requests_total.labels(api_endpoint=endpoint, status_code="200").inc()
        
        if use_mock:
            # 使用模拟数据
            api_logger.info("使用模拟数据生成Excel")
            result = generate_mock_response(request.description)
            sheet_name = result.get('sheet_name', 'Sheet1')
            data = result.get('data', [])
            
            # 模拟token消耗 - 实际项目中应根据API响应获取真实token消耗
            mock_tokens_used = 50 + len(request.description) // 10
            token_consumed.labels(api_endpoint=endpoint).inc(mock_tokens_used)
            token_consumed_current.labels(api_endpoint=endpoint).set(mock_tokens_used)
        else:
            # 使用DeepSeek API解析自然语言描述
            url = "https://api.deepseek.com/chat/completions"
            headers = {
                "Content-Type": "application/json",
                "Authorization": f"Bearer {config.DEEPSEEK_API_KEY}"
            }
            data = {
                "model": config.DEEPSEEK_MODEL,
                "messages": [
                    {"role": "system", "content": "你是一个Excel生成专家，需要根据用户描述创建Excel表格数据。"},
                    {"role": "user", "content": f"请根据以下描述生成Excel数据：{request.description}\n返回格式应为JSON，包含'sheet_name'和'data'字段，其中data是二维数组。"}
                ],
                "max_tokens": config.DEEPSEEK_MAX_TOKENS,
                "temperature": config.DEEPSEEK_TEMPERATURE
            }
            
            # 发送请求到DeepSeek API
            response = requests.post(url, headers=headers, json=data)
            response_json = response.json()
            
            # 解析DeepSeek响应
            result = json.loads(response_json['choices'][0]['message']['content'])
            sheet_name = result.get('sheet_name', 'Sheet1')
            data = result.get('data', [])
            
            # 记录token消耗
            if 'usage' in response_json:
                tokens_used = response_json['usage']['total_tokens']
                token_consumed.labels(api_endpoint=endpoint).inc(tokens_used)
                token_consumed_current.labels(api_endpoint=endpoint).set(tokens_used)
        
        print(f"收到生成Excel请求: description={request.description}, file_name={request.file_name}")
        
        # 生成Excel文件
        if not request.file_name:
            file_name = f"generated_{int(time.time())}.xlsx"
        else:
            file_name = request.file_name
            # 确保文件名包含.xlsx扩展名
            if not file_name.lower().endswith('.xlsx'):
                file_name += '.xlsx'
        
        file_path = os.path.join(config.TEMP_DIR, file_name)
        print(f"准备创建Excel文件: {file_path}")
        
        # 记录Excel处理开始时间
        excel_process_start = time.time()
        success = excel_utils.create_excel(file_path, sheet_name, data)
        # 记录Excel处理时间和行数
        excel_processing_time_seconds.labels(operation_type="generate").observe(time.time() - excel_process_start)
        
        if success:
            print(f"Excel文件创建成功: {file_path}")
            
            # 记录处理的Excel文件和行数
            excel_files_processed.labels(operation_type="generate").inc()
            excel_rows_processed.labels(operation_type="generate").inc(len(data))

            # 记录响应时间
            api_request_duration_seconds.labels(api_endpoint=endpoint).observe(time.time() - start_time)
            active_requests.dec()

            excel_content = excel_utils.read_excel(file_path)
            
            return {
                "status": "success",
                "file_name": file_name,
                "download_url": f"/download/{file_name}",  # <--- 确认这里有逗号
                "excel_data": excel_content.get('data', [])
            }
        else:
            print(f"Excel文件创建失败")
            api_requests_total.labels(api_endpoint=endpoint, status_code="500").inc()
            api_request_duration_seconds.labels(api_endpoint=endpoint).observe(time.time() - start_time)
            active_requests.dec()
            raise HTTPException(status_code=500, detail="创建Excel文件失败")
        
    except Exception as e:
        print(f"生成Excel时出错：{str(e)}")
        api_requests_total.labels(api_endpoint=endpoint, status_code="500").inc()
        api_request_duration_seconds.labels(api_endpoint=endpoint).observe(time.time() - start_time)
        active_requests.dec()
        
        # 如果API调用失败，回退到模拟数据
        if not use_mock:
            print("API调用失败，回退到模拟数据")
            try:
                result = generate_mock_response(request.description)
                sheet_name = result.get('sheet_name', 'Sheet1')
                data = result.get('data', [])
                
                if not file_name:
                    file_name = f"generated_{int(time.time())}.xlsx"
                
                file_path = os.path.join(config.TEMP_DIR, file_name)
                excel_utils.create_excel(file_path, sheet_name, data)
                
                # 记录处理的Excel文件和行数
                excel_files_processed.labels(operation_type="generate").inc()
                excel_rows_processed.labels(operation_type="generate").inc(len(data))
                
                return {
                    "status": "success",
                    "file_name": file_name,
                    "download_url": f"/download/{file_name}",
                    "message": "使用模拟数据生成的Excel文件"
                }
            except:
                pass
        
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/edit_excel")
async def edit_excel(
    file: UploadFile = File(...),
    instructions: str = Form(...)
):
    """
    根据自然语言指令编辑已有的Excel文件
    """
    endpoint = "/edit_excel"
    active_requests.inc()
    start_time = time.time()
    
    try:
        # 记录API请求
        api_requests_total.labels(api_endpoint=endpoint, status_code="200").inc()
        
        # 保存上传的文件
        temp_input_path = os.path.join(config.TEMP_DIR, f"temp_{int(time.time())}_{file.filename}")
        with open(temp_input_path, "wb") as buffer:
            buffer.write(await file.read())
        
        # 读取Excel内容
        excel_content = excel_utils.read_excel(temp_input_path)
        
        if use_mock:
            # 使用模拟数据进行编辑
            print("使用模拟数据编辑Excel")
            # 创建一个简单的编辑操作列表作为模拟响应
            edit_operations = [
                {
                    "type": "add_row",
                    "index": 1,
                    "data": ["模拟数据已添加"]
                }
            ]
            
            # 模拟token消耗
            mock_tokens_used = 80 + len(instructions) // 10
            token_consumed.labels(api_endpoint=endpoint).inc(mock_tokens_used)
            token_consumed_current.labels(api_endpoint=endpoint).set(mock_tokens_used)
        else:
            # 使用DeepSeek API解析编辑指令
            url = "https://api.deepseek.com/chat/completions"
            headers = {
                "Content-Type": "application/json",
                "Authorization": f"Bearer {config.DEEPSEEK_API_KEY}"
            }
            data = {
                "model": config.DEEPSEEK_MODEL,
                "messages": [
                    {"role": "system", "content": "你是一个Excel编辑专家，需要根据用户指令修改现有的Excel表格。"},
                    {"role": "user", "content": f"现有Excel内容：{excel_content}\n编辑指令：{instructions}\n返回格式应为JSON，包含需要执行的操作列表，如添加行、修改单元格等。"}
                ],
                "max_tokens": config.DEEPSEEK_MAX_TOKENS,
                "temperature": config.DEEPSEEK_TEMPERATURE
            }
            
            # 发送请求到DeepSeek API
            response = requests.post(url, headers=headers, json=data)
            response_json = response.json()
            
            # 解析编辑操作
            edit_operations = json.loads(response_json['choices'][0]['message']['content'])
            
            # 记录token消耗
            if 'usage' in response_json:
                tokens_used = response_json['usage']['total_tokens']
                token_consumed.labels(api_endpoint=endpoint).inc(tokens_used)
                token_consumed_current.labels(api_endpoint=endpoint).set(tokens_used)
        
        # 记录Excel处理开始时间
        excel_process_start = time.time()
        # 执行编辑操作
        output_file_name = f"edited_{int(time.time())}_{file.filename}"
        output_path = os.path.join(config.TEMP_DIR, output_file_name)
        success = excel_utils.edit_excel(temp_input_path, output_path, edit_operations)
        # 记录Excel处理时间
        excel_processing_time_seconds.labels(operation_type="edit").observe(time.time() - excel_process_start)
        
        # 清理临时文件
        os.remove(temp_input_path)
        
        # 记录处理的Excel文件
        if success:
            excel_files_processed.labels(operation_type="edit").inc()
        
        # 记录响应时间
        api_request_duration_seconds.labels(api_endpoint=endpoint).observe(time.time() - start_time)
        active_requests.dec()
        
        return {
            "status": "success",
            "file_name": output_file_name,
            "download_url": f"/download/{output_file_name}"
        }
        
    except Exception as e:
        print(f"编辑Excel时出错：{str(e)}")
        api_requests_total.labels(api_endpoint=endpoint, status_code="500").inc()
        api_request_duration_seconds.labels(api_endpoint=endpoint).observe(time.time() - start_time)
        active_requests.dec()
        
        # 即使出错也尝试返回原始文件的副本
        try:
            output_file_name = f"backup_{int(time.time())}_{file.filename}"
            output_path = os.path.join(config.TEMP_DIR, output_file_name)
            
            # 复制原始文件作为备份
            import shutil
            if os.path.exists(temp_input_path):
                shutil.copy2(temp_input_path, output_path)
                os.remove(temp_input_path)
            else:
                # 如果临时文件不存在，创建一个空文件
                with open(output_path, 'wb'):
                    pass
            
            # 记录处理的Excel文件（作为备份）
            excel_files_processed.labels(operation_type="edit").inc()
            
            return {
                "status": "partial_success",
                "file_name": output_file_name,
                "download_url": f"/download/{output_file_name}",
                "message": "无法完成编辑操作，返回原始文件的副本"
            }
        except:
            pass
        
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/download/{file_name}")
async def download_file(file_name: str):
    """
    下载生成或编辑后的Excel文件
    """
    file_path = os.path.join(config.TEMP_DIR, file_name)
    if not os.path.exists(file_path):
        raise HTTPException(status_code=404, detail="文件不存在")
    
    return FileResponse(
        path=file_path,
        filename=file_name,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

class LogEntry(BaseModel):
    level: str
    message: str
    source: str = "frontend"
    timestamp: str
    additional_data: dict = None

class BatchLogRequest(BaseModel):
    logs: list[LogEntry]

@app.post("/log")
async def log_entry(log_entry: LogEntry):
    """
    接收前端发送的单条日志（兼容旧版）
    """
    try:
        # 根据日志级别记录到相应的logger
        if log_entry.level.upper() == "ERROR":
            api_logger.error(f"{log_entry.source}: {log_entry.message}" + (f" - {log_entry.additional_data}" if log_entry.additional_data else ""))
        elif log_entry.level.upper() == "WARNING":
            api_logger.warning(f"{log_entry.source}: {log_entry.message}" + (f" - {log_entry.additional_data}" if log_entry.additional_data else ""))
        elif log_entry.level.upper() == "INFO":
            api_logger.info(f"{log_entry.source}: {log_entry.message}" + (f" - {log_entry.additional_data}" if log_entry.additional_data else ""))
        else:
            api_logger.debug(f"{log_entry.source}: {log_entry.message}" + (f" - {log_entry.additional_data}" if log_entry.additional_data else ""))
        return {"status": "success"}
    except Exception as e:
        print(f"记录日志时出错：{str(e)}")
        return {"status": "error"}

@app.post("/log/batch")
async def log_batch(request: BatchLogRequest):
    """
    批量接收前端发送的日志
    """
    try:
        for log_entry in request.logs:
            # 根据日志级别记录到相应的logger
            if log_entry.level.upper() == "ERROR":
                api_logger.error(f"{log_entry.source}: {log_entry.message}" + (f" - {log_entry.additional_data}" if log_entry.additional_data else ""))
            elif log_entry.level.upper() == "WARNING":
                api_logger.warning(f"{log_entry.source}: {log_entry.message}" + (f" - {log_entry.additional_data}" if log_entry.additional_data else ""))
            elif log_entry.level.upper() == "INFO":
                api_logger.info(f"{log_entry.source}: {log_entry.message}" + (f" - {log_entry.additional_data}" if log_entry.additional_data else ""))
            else:
                api_logger.debug(f"{log_entry.source}: {log_entry.message}" + (f" - {log_entry.additional_data}" if log_entry.additional_data else ""))
        return {"status": "success", "count": len(request.logs)}
    except Exception as e:
        print(f"批量记录日志时出错：{str(e)}")
        return {"status": "error"}

# 数据分析

@app.post("/analyze_excel")
async def analyze_excel(
    file: UploadFile = File(...)
):
    """
    分析Excel文件数据，生成关键洞察、趋势分析和数据异常检测结果
    """
    # ----------------------------------------------------------------
    # 你的数据清洗函数 (我们把它放在 analyze_excel 函数内部，作为辅助函数)
    # ----------------------------------------------------------------
    def _sanitize_visualization_data(report: Dict[str, Any]) -> Dict[str, Any]:
        """
        清洗 visualization_data，确保数值字段是 float 或 int。
        这是嵌套在 analyze_excel 内部的辅助函数。
        """
        if "visualization_data" not in report or not isinstance(report.get("visualization_data"), list):
            return report

        for chart in report["visualization_data"]:
            if not isinstance(chart, dict) or "data" not in chart or not isinstance(chart.get("data"), list):
                continue
                
            clean_data = []
            
            numeric_keys = set()
            if chart.get("y_axis"):
                numeric_keys.add(chart.get("y_axis"))
            
            if chart.get("type") == "multi_line_chart":
                x_axis_key = chart.get("x_axis")
                if chart["data"] and isinstance(chart["data"][0], dict):
                    numeric_keys.update([k for k in chart["data"][0].keys() if k != x_axis_key])

            for row in chart.get("data", []):
                if not isinstance(row, dict):
                    clean_data.append(row)
                    continue

                clean_row = {}
                for k, v in row.items():
                    if k in numeric_keys:
                        if isinstance(v, str):
                            cleaned_v = v.replace("%", "").replace(",", "").replace("元", "").strip()
                            try:
                                clean_row[k] = float(cleaned_v)
                            except (ValueError, TypeError):
                                clean_row[k] = 0
                        elif isinstance(v, (int, float)):
                             clean_row[k] = v
                        else:
                            clean_row[k] = 0
                    else:
                        clean_row[k] = v
                clean_data.append(clean_row)
            chart["data"] = clean_data
        return report
    
    def _ensure_chart_metadata(report: Dict[str, Any]) -> Dict[str, Any]:
        """
        检查并补充 visualization_data 中缺失的 x_axis 和 y_axis 元数据。
        """
        if "visualization_data" not in report or not isinstance(report.get("visualization_data"), list):
            return report
            
        for chart in report["visualization_data"]:
            if not isinstance(chart, dict) or "data" not in chart or not chart.get("data"):
                continue

            # 检查 bar_chart 或 line_chart
            if chart.get("type") in ["bar_chart", "line_chart"]:
                first_row = chart["data"][0]
                if isinstance(first_row, dict):
                    keys = list(first_row.keys())
                    if "x_axis" not in chart and len(keys) > 0:
                        chart["x_axis"] = keys[0] # 假设第一个key是x轴
                    if "y_axis" not in chart and len(keys) > 1:
                        chart["y_axis"] = keys[1] # 假设第二个key是y轴
            
            # 检查 pie_chart (它使用'name'和'value'作为隐式轴)
            elif chart.get("type") == "pie_chart":
                first_row = chart["data"][0]
                if isinstance(first_row, dict):
                    if "name" not in first_row and "category" in first_row:
                         # 有时AI会用'category'，我们把它规范化为'name'
                         for row in chart["data"]:
                             row["name"] = row.pop("category")
        return report

    # ----------------------------------------------------------------
    # 主函数逻辑开始
    # ----------------------------------------------------------------
    endpoint = "/analyze_excel"
    active_requests.inc()
    start_time = time.time()
    
    temp_input_path = None
    try:
        api_requests_total.labels(api_endpoint=endpoint, status_code="200").inc()
        
        temp_input_path = os.path.join(config.TEMP_DIR, f"temp_analyze_{int(time.time())}_{file.filename}")
        with open(temp_input_path, "wb") as buffer:
            buffer.write(await file.read())
        
        excel_content = excel_utils.read_excel(temp_input_path)
        
        if "error" in excel_content:
            api_requests_total.labels(api_endpoint=endpoint, status_code="400").inc()
            raise HTTPException(status_code=400, detail=excel_content["error"])
        
        excel_process_start = time.time()
        
        if use_mock:
            api_logger.info("使用增强版模拟函数生成Excel分析报告")
            report = generate_analysis_report(excel_content)
        else:
            api_logger.info("调用DeepSeek API进行深度数据分析")
            
            system_prompt = """
            你是一位顶尖的数据分析师。你的回答必须严格遵循JSON格式。
            分析任务：
            1. `summary`: 一句话总结数据核心。
            2. `insights`: 提供3个最深刻的洞察，必须是完整的句子。
            3. `trends`: 提供1-2个数据趋势，必须是完整的句子。
            4. `anomalies`: 提供1-2个数据异常点，必须是完整的句子。
            5. `visualization_data`: **这是强制任务，你必须返回一个非空的JSON数组。**
                - **首选：** 如果数据包含时间和至少两个数值列（如“销售额”和“目标”），生成一个 `multi_line_chart` (多线图) 进行对比。
                - **次选：** 如果只有时间和单个数值列，生成一个 `line_chart` (单线图)。
                - **备选：** 如果有类别和数值列，生成 `bar_chart` (柱状图) 或 `pie_chart` (饼图)。
                - **最后手段：** 如果以上都无法生成，你必须生成一个关于数据行数和列数的 `bar_chart` 作为概览。
                - 格式: `[{"type": "...", "title": "...", "data": [...]}]`
            你的回答必须是纯粹的JSON，不含任何解释。
            """
            
            headers = {
                "Content-Type": "application/json",
                "Authorization": f"Bearer {config.DEEPSEEK_API_KEY}"
            }
            
            data_payload = {
                "model": config.DEEPSEEK_MODEL,
                "messages": [
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": f"请分析以下Excel JSON数据: {json.dumps(excel_content, ensure_ascii=False)}"}
                ],
                "max_tokens": config.DEEPSEEK_MAX_TOKENS,
                "temperature": config.DEEPSEEK_TEMPERATURE
            }
            
            response = requests.post(
                "https://api.deepseek.com/chat/completions",
                headers=headers,
                json=data_payload
            )
            response.raise_for_status()
            response_json = response.json()
            
            content = response_json.get('choices', [{}])[0].get('message', {}).get('content', '{}')
            clean_content = content.strip().lstrip('```json').rstrip('```')
            report = json.loads(clean_content)

        excel_processing_time_seconds.labels(operation_type="analyze").observe(time.time() - excel_process_start)
        
        api_logger.info(f"AI分析完成，清洗前的数据: {report}")
        report = _sanitize_visualization_data(report)
        report = _ensure_chart_metadata(report)
        api_logger.info(f"清洗后的数据: {report}")

        excel_files_processed.labels(operation_type="analyze").inc()
        data_rows = len(excel_content.get('data', [])) - 1
        if data_rows > 0:
            excel_rows_processed.labels(operation_type="analyze").inc(data_rows)
        
        if not report or not isinstance(report, dict):
            api_requests_total.labels(api_endpoint=endpoint, status_code="500").inc()
            raise HTTPException(status_code=500, detail="AI未能生成有效的分析报告")
        
        if "status" not in report:
            report["status"] = "success"
            
        return report
        
    except Exception as e:
        api_logger.error(f"分析Excel时发生错误: {str(e)}", exc_info=True)
        api_requests_total.labels(api_endpoint=endpoint, status_code="500").inc()
        
        if not use_mock and 'excel_content' in locals() and "error" not in excel_content:
            try:
                api_logger.warning("AI API调用失败，回退到模拟分析")
                report = generate_analysis_report(excel_content)
                report = _sanitize_visualization_data(report)
                excel_files_processed.labels(operation_type="analyze").inc()
                data_rows = len(excel_content.get('data', [])) - 1
                if data_rows > 0:
                    excel_rows_processed.labels(operation_type="analyze").inc(data_rows)
                return report
            except Exception as fallback_e:
                api_logger.error(f"回退到模拟分析也失败了: {fallback_e}", exc_info=True)
        
        raise HTTPException(status_code=500, detail=str(e))
    
    finally:
        if temp_input_path and os.path.exists(temp_input_path):
            os.remove(temp_input_path)
        
        api_request_duration_seconds.labels(api_endpoint=endpoint).observe(time.time() - start_time)
        active_requests.dec()


@app.get("/metrics")
async def metrics():
    """
    Prometheus指标端点，提供监控数据
    """
    return Response(content=generate_latest(), media_type=CONTENT_TYPE_LATEST)

@app.get("/")
async def root():
    """
    API根路径，返回基本信息
    """
    return {
        "app": "ExcelGenius",
        "version": "1.0.0",
        "endpoints": [
            {"path": "/generate_excel", "method": "POST", "description": "根据自然语言描述生成Excel文件"},
            {"path": "/edit_excel", "method": "POST", "description": "根据自然语言指令编辑Excel文件"},
            {"path": "/analyze_excel", "method": "POST", "description": "分析Excel文件数据并生成洞察"},
            {"path": "/download/{file_name}", "method": "GET", "description": "下载Excel文件"},
            {"path": "/log", "method": "POST", "description": "接收前端日志"},
            {"path": "/metrics", "method": "GET", "description": "Prometheus监控指标"}
        ]
    }

class SaveDataRequest(BaseModel):
    file_name: str
    sheet_data: List[List[str | None]] # 允许单元格为 null (None)

@app.post("/api/excel/save_data")
async def save_excel_data(request: SaveDataRequest):
    """
    接收前端发送的、修改后的二维数组数据，
    并将其覆盖保存到指定的Excel文件中。
    """
    # 确保文件名是安全的，防止路径遍历攻击 (一个好的安全实践)
    safe_file_name = os.path.basename(request.file_name)
    file_path = os.path.join(config.TEMP_DIR, safe_file_name)
    
    try:
        # 我们复用早已写好的 create_excel 函数，它会直接覆盖同名文件
        success = excel_utils.create_excel(file_path, "Sheet1", request.sheet_data)
        
        if success:
            excel_logger.info(f"文件已成功更新并保存: {file_path}")
            return {"status": "success", "message": "文件已成功保存"}
        else:
            raise HTTPException(status_code=500, detail="使用 excel_utils 保存文件时发生错误")
            
    except Exception as e:
        excel_logger.error(f"保存文件 '{safe_file_name}' 失败: {str(e)}")
        raise HTTPException(status_code=500, detail=str(e))
        
if __name__ == "__main__":
    # 启动服务器
    uvicorn.run(
        "main:app",
        host=config.HOST,
        port=config.PORT,
        reload=config.DEBUG
    )