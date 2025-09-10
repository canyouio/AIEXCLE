import os
from dotenv import load_dotenv
import logging

# 配置基本日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger('config')

class Config:
    def __init__(self):
        # 显式指定.env文件路径
        env_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), '.env')
        logger.info(f"尝试加载.env文件: {env_path}")
        
        # 检查.env文件是否存在
        if os.path.exists(env_path):
            logger.info(f".env文件存在，路径: {env_path}")
            # 加载环境变量
            load_dotenv(dotenv_path=env_path)
        else:
            logger.warning(f".env文件不存在: {env_path}")
            load_dotenv()  # 尝试默认加载
        
        # 服务器配置
        self.HOST = os.getenv("HOST", "127.0.0.1")
        self.PORT = int(os.getenv("PORT", "8000"))
        self.DEBUG = os.getenv("DEBUG", "True").lower() == "true"
        
        # DeepSeek配置
        self.DEEPSEEK_API_KEY = os.getenv("DEEPSEEK_API_KEY", "placeholder-api-key")  # DeepSeek API Key
        logger.info(f"加载的DEEPSEEK_API_KEY: {'已设置(部分隐藏)' if self.DEEPSEEK_API_KEY and self.DEEPSEEK_API_KEY != 'placeholder-api-key' else '未设置或为默认值'}")
        logger.info(f"DEEPSEEK_API_KEY == 'placeholder-api-key': {self.DEEPSEEK_API_KEY == 'placeholder-api-key'}")
        self.DEEPSEEK_MODEL = os.getenv("DEEPSEEK_MODEL", "deepseek-chat")
        self.DEEPSEEK_MAX_TOKENS = int(os.getenv("DEEPSEEK_MAX_TOKENS", "2000"))
        self.DEEPSEEK_TEMPERATURE = float(os.getenv("DEEPSEEK_TEMPERATURE", "0.7"))
        
        # 应用配置
        self.APP_NAME = "ExcelGenius"
        self.APP_VERSION = "1.0.0"
        
        # 目录配置
        self.BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        # 使用相对路径
        self.TEMP_DIR = os.path.join(self.BASE_DIR, "temp")
        # 确保临时目录存在
        if not os.path.exists(self.TEMP_DIR):
            os.makedirs(self.TEMP_DIR)
        
        # Excel配置
        self.EXCEL_DEFAULT_SHEET = "Sheet1"
        self.EXCEL_MAX_ROWS = int(os.getenv("EXCEL_MAX_ROWS", "1000"))
        self.EXCEL_MAX_COLS = int(os.getenv("EXCEL_MAX_COLS", "100"))
        
        # 日志配置
        self.LOG_LEVEL = os.getenv("LOG_LEVEL", "INFO")
        self.LOG_FORMAT = "%(asctime)s - %(name)s - %(levelname)s - %(message)s"
        self.LOG_FILE = os.path.join(self.BASE_DIR, "logs", "app.log")
        
        # 确保日志目录存在
        if not os.path.exists(os.path.dirname(self.LOG_FILE)):
            os.makedirs(os.path.dirname(self.LOG_FILE))
            
        # 安全配置
        self.API_KEY_REQUIRED = os.getenv("API_KEY_REQUIRED", "False").lower() == "true"
        self.UPLOAD_MAX_SIZE_MB = int(os.getenv("UPLOAD_MAX_SIZE_MB", "10"))
        self.ALLOWED_FILE_TYPES = ["xlsx", "xls"]
        
        # 缓存配置
        self.CACHE_ENABLED = os.getenv("CACHE_ENABLED", "True").lower() == "true"
        self.CACHE_TTL = int(os.getenv("CACHE_TTL", "3600"))  # 秒