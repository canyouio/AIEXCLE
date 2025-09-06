import logging
import os
from logging.handlers import RotatingFileHandler
from .config import Config

class LoggerConfig:
    @staticmethod
    def setup_logger():
        """
        设置全局日志配置，替换print语句
        """
        # 加载配置
        config = Config()
        
        # 创建logger对象
        logger = logging.getLogger()
        logger.setLevel(getattr(logging, config.LOG_LEVEL))
        
        # 避免重复添加handler
        if not logger.handlers:
            # 创建轮转文件handler，限制单个文件大小为10MB，最多保留5个备份
            file_handler = RotatingFileHandler(
                config.LOG_FILE, 
                maxBytes=10*1024*1024,  # 10MB
                backupCount=5,          # 保留5个备份文件
                encoding='utf-8'
            )
            file_handler.setLevel(getattr(logging, config.LOG_LEVEL))
            
            # 创建控制台handler
            console_handler = logging.StreamHandler()
            console_handler.setLevel(getattr(logging, config.LOG_LEVEL))
            
            # 设置日志格式
            formatter = logging.Formatter(config.LOG_FORMAT)
            file_handler.setFormatter(formatter)
            console_handler.setFormatter(formatter)
            
            # 添加handler到logger
            logger.addHandler(file_handler)
            logger.addHandler(console_handler)
        
        return logger

# 创建全局logger实例，供其他模块导入使用
global_logger = LoggerConfig.setup_logger()

# 为不同模块创建专用logger
excel_logger = logging.getLogger('app.excel')
api_logger = logging.getLogger('app.api')
utils_logger = logging.getLogger('app.utils')