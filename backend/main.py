import os
import re
import json
import time
import os
from fastapi import FastAPI, UploadFile, File, Form, HTTPException, Body
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
import uvicorn
import requests
import logging

# 导入Prometheus监控库
from prometheus_client import Counter, Histogram, Gauge, generate_latest, CONTENT_TYPE_LATEST
from fastapi import Response

from .config import Config
from .excel_utils import ExcelUtils
from .logger_config import api_logger, excel_logger

# 初始化FastAPI应用
app = FastAPI(title="ExcelGenius API", description="自然语言生成和编辑Excel文件")

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

@app.post("/generate_excel")
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
            
            return {
                "status": "success",
                "file_name": file_name,
                "download_url": f"/download/{file_name}"
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

@app.post("/analyze_excel")
async def analyze_excel(
    file: UploadFile = File(...)
):
    """
    分析Excel文件数据，生成关键洞察、趋势分析和数据异常检测结果
    """
    endpoint = "/analyze_excel"
    active_requests.inc()
    start_time = time.time()
    
    try:
        # 记录API请求
        api_requests_total.labels(api_endpoint=endpoint, status_code="200").inc()
        
        # 保存上传的文件
        temp_input_path = os.path.join(config.TEMP_DIR, f"temp_analyze_{int(time.time())}_{file.filename}")
        with open(temp_input_path, "wb") as buffer:
            buffer.write(await file.read())
        
        # 读取Excel内容
        excel_content = excel_utils.read_excel(temp_input_path)
        
        # 检查是否读取成功
        if "error" in excel_content:
            api_requests_total.labels(api_endpoint=endpoint, status_code="400").inc()
            api_request_duration_seconds.labels(api_endpoint=endpoint).observe(time.time() - start_time)
            active_requests.dec()
            raise HTTPException(status_code=400, detail=excel_content["error"])
        
        # 记录Excel处理开始时间
        excel_process_start = time.time()
        
        if use_mock:
            # 使用模拟函数生成分析报告
            api_logger.info("使用模拟函数生成Excel分析报告")
            report = generate_analysis_report(excel_content)
            
            # 模拟token消耗
            data_rows = len(excel_content.get('data', [])) - 1  # 减去表头
            mock_tokens_used = 100 + data_rows * 2
            token_consumed.labels(api_endpoint=endpoint).inc(mock_tokens_used)
            token_consumed_current.labels(api_endpoint=endpoint).set(mock_tokens_used)
        else:
            # 使用DeepSeek API进行数据分析
            url = "https://api.deepseek.com/chat/completions"
            headers = {
                "Content-Type": "application/json",
                "Authorization": f"Bearer {config.DEEPSEEK_API_KEY}"
            }
            
            # 准备请求数据
            system_prompt = """
            你是一个数据分析师，需要分析Excel表格数据并生成详细的分析报告。
            请返回JSON格式的分析结果，包含以下字段：
            - status: "success" 或 "error"
            - sheet_name: 工作表名称
            - summary: 数据摘要
            - insights: 关键洞察（字符串数组）
            - trends: 趋势分析（字符串数组）
            - anomalies: 数据异常（字符串数组）
            - visualization_data: 可视化数据（数组，每项包含type、title和data等字段）
            """
            
            data = {
                "model": config.DEEPSEEK_MODEL,
                "messages": [
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": f"请分析以下Excel数据：{excel_content}"}
                ],
                "max_tokens": config.DEEPSEEK_MAX_TOKENS,
                "temperature": config.DEEPSEEK_TEMPERATURE
            }
            
            # 发送请求到DeepSeek API
            response = requests.post(url, headers=headers, json=data)
            response_json = response.json()
            
            # 解析API响应
            report = json.loads(response_json['choices'][0]['message']['content'])
            
            # 记录token消耗
            if 'usage' in response_json:
                tokens_used = response_json['usage']['total_tokens']
                token_consumed.labels(api_endpoint=endpoint).inc(tokens_used)
                token_consumed_current.labels(api_endpoint=endpoint).set(tokens_used)
        
        # 记录Excel处理时间
        excel_processing_time_seconds.labels(operation_type="analyze").observe(time.time() - excel_process_start)
        
        # 添加详细日志，查看生成的报告内容
        print(f"生成的分析报告: {report}")
        print(f"可视化数据长度: {len(report.get('visualization_data', []))}")
        print(f"可视化数据内容: {report.get('visualization_data', [])}")
        
        # 清理临时文件
        os.remove(temp_input_path)
        
        # 记录处理的Excel文件和行数
        excel_files_processed.labels(operation_type="analyze").inc()
        data_rows = len(excel_content.get('data', [])) - 1  # 减去表头
        if data_rows > 0:
            excel_rows_processed.labels(operation_type="analyze").inc(data_rows)
        
        # 确保返回的是有效的分析报告
        if not report or report.get("status") != "success":
            api_requests_total.labels(api_endpoint=endpoint, status_code="500").inc()
            api_request_duration_seconds.labels(api_endpoint=endpoint).observe(time.time() - start_time)
            active_requests.dec()
            raise HTTPException(status_code=500, detail="生成分析报告失败")
        
        # 记录响应时间
        api_request_duration_seconds.labels(api_endpoint=endpoint).observe(time.time() - start_time)
        active_requests.dec()
        
        return report
        
    except Exception as e:
        print(f"分析Excel时出错：{str(e)}")
        api_requests_total.labels(api_endpoint=endpoint, status_code="500").inc()
        api_request_duration_seconds.labels(api_endpoint=endpoint).observe(time.time() - start_time)
        active_requests.dec()
        
        # 即使出错也尝试清理临时文件
        try:
            if os.path.exists(temp_input_path):
                os.remove(temp_input_path)
        except:
            pass
        
        # 如果API调用失败，尝试使用模拟函数作为回退
        if not use_mock:
            try:
                if "excel_content" in locals() and not isinstance(excel_content, dict) or "error" not in excel_content:
                    report = generate_analysis_report(excel_content)
                    
                    # 记录处理的Excel文件和行数（回退模式）
                    excel_files_processed.labels(operation_type="analyze").inc()
                    data_rows = len(excel_content.get('data', [])) - 1  # 减去表头
                    if data_rows > 0:
                        excel_rows_processed.labels(operation_type="analyze").inc(data_rows)
                    
                    return report
            except:
                pass
        
        raise HTTPException(status_code=500, detail=str(e))


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

if __name__ == "__main__":
    # 启动服务器
    uvicorn.run(
        "main:app",
        host=config.HOST,
        port=config.PORT,
        reload=config.DEBUG
    )