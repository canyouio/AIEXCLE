import axios from './request';

class Logger {
  constructor() {
    this.isEnabled = true;
    this.sendToBackend = true;
    this.environment = import.meta.env.MODE || 'development';
    // 日志缓冲队列
    this.logBuffer = [];
    // 缓冲时间（毫秒）
    this.bufferTime = 1000;
    // 上次发送时间
    this.lastSendTime = Date.now();
    // 批量发送定时器
    this.sendTimer = null;
    // 节流控制，避免短时间内重复发送相同日志
    this.throttleMap = new Map();
    // 节流时间（毫秒）
    this.throttleTime = 5000;
    // 仅在生产环境收集的日志级别
    this.productionLevels = ['ERROR', 'WARNING'];
  }

  /**
   * 格式化日期时间
   * @returns {string} 格式化后的日期时间字符串
   */
  _formatDateTime() {
    const now = new Date();
    return now.toISOString().replace('T', ' ').substring(0, 23);
  }

  /**
   * 检查是否需要节流
   * @param {string} key - 节流的键
   * @returns {boolean} 是否需要节流
   */
  _shouldThrottle(key) {
    const now = Date.now();
    const lastTime = this.throttleMap.get(key);
    
    if (lastTime && now - lastTime < this.throttleTime) {
      return true;
    }
    
    this.throttleMap.set(key, now);
    // 清理过期的节流记录
    setTimeout(() => {
      this.throttleMap.delete(key);
    }, this.throttleTime);
    
    return false;
  }

  /**
   * 批量发送日志到后端
   */
  async _sendBatchLogs() {
    if (!this.sendToBackend || this.logBuffer.length === 0) return;

    // 复制并清空缓冲队列
    const logsToSend = [...this.logBuffer];
    this.logBuffer = [];
    
    try {
      // 批量发送日志
      await axios.post('/log/batch', {
        logs: logsToSend
      });
    } catch (error) {
      // 如果发送失败，保留最后10条日志到缓冲队列
      console.error('Failed to send batch logs to backend:', error);
      this.logBuffer = [...logsToSend.slice(-10), ...this.logBuffer];
    }
  }

  /**
   * 添加日志到缓冲队列，并在合适的时机发送
   * @param {string} level - 日志级别
   * @param {string} message - 日志消息
   * @param {Object} additionalData - 附加数据
   */
  _addLogToBuffer(level, message, additionalData = null) {
    // 生成节流键
    const throttleKey = `${level}:${message.substring(0, 100)}`; // 截取前100个字符作为键
    
    // 节流检查
    if (this._shouldThrottle(throttleKey)) {
      return;
    }
    
    const logEntry = {
      level,
      message,
      timestamp: this._formatDateTime(),
      source: window.location.pathname || 'frontend',
      additional_data: additionalData
    };
    
    // 添加到缓冲队列
    this.logBuffer.push(logEntry);
    
    const now = Date.now();
    // 如果缓冲队列超过20条或者距离上次发送超过1秒，则发送
    if (this.logBuffer.length >= 20 || now - this.lastSendTime >= this.bufferTime) {
      // 清除之前的定时器
      if (this.sendTimer) {
        clearTimeout(this.sendTimer);
      }
      
      // 立即发送
      this._sendBatchLogs();
      this.lastSendTime = now;
    } else {
      // 否则设置定时器延迟发送
      if (this.sendTimer) {
        clearTimeout(this.sendTimer);
      }
      
      this.sendTimer = setTimeout(() => {
        this._sendBatchLogs();
        this.lastSendTime = Date.now();
      }, this.bufferTime);
    }
  }

  /**
   * 记录日志
   * @param {string} level - 日志级别
   * @param {string} message - 日志消息
   * @param {Object} additionalData - 附加数据
   */
  _log(level, message, additionalData = null) {
    if (!this.isEnabled) return;
    
    // 根据级别调用对应的console方法
    const consoleMethod = console[level.toLowerCase()] || console.log;
    consoleMethod(`[${level.toUpperCase()}] ${message}`, additionalData);
    
    // 检查是否需要发送到后端
    if (this.sendToBackend) {
      // 在开发环境下，只发送ERROR和WARNING级别的日志
      if (this.environment === 'development') {
        if (['ERROR', 'WARNING'].includes(level.toUpperCase())) {
          this._addLogToBuffer(level, message, additionalData);
        }
      } else {
        // 在生产环境下，只发送配置中指定的级别
        if (this.productionLevels.includes(level.toUpperCase())) {
          this._addLogToBuffer(level, message, additionalData);
        }
      }
    }
  }

  /**
   * 记录调试日志
   * @param {string} message - 日志消息
   * @param {Object} additionalData - 附加数据
   */
  debug(message, additionalData = null) {
    this._log('debug', message, additionalData);
  }

  /**
   * 记录信息日志
   * @param {string} message - 日志消息
   * @param {Object} additionalData - 附加数据
   */
  info(message, additionalData = null) {
    this._log('info', message, additionalData);
  }

  /**
   * 记录警告日志
   * @param {string} message - 日志消息
   * @param {Object} additionalData - 附加数据
   */
  warning(message, additionalData = null) {
    this._log('warning', message, additionalData);
  }

  /**
   * 记录错误日志
   * @param {string} message - 日志消息
   * @param {Object|Error} error - 错误对象或附加数据
   */
  error(message, error = null) {
    let additionalData = null;
    if (error instanceof Error) {
      additionalData = {
        error: error.message,
        stack: error.stack
      };
    } else {
      additionalData = error;
    }
    
    this._log('error', message, additionalData);
  }

  /**
   * 设置日志是否启用
   * @param {boolean} enabled - 是否启用
   */
  setEnabled(enabled) {
    this.isEnabled = enabled;
  }

  /**
   * 设置是否发送日志到后端
   * @param {boolean} send - 是否发送
   */
  setSendToBackend(send) {
    this.sendToBackend = send;
  }
  
  /**
   * 设置生产环境下收集的日志级别
   * @param {Array} levels - 日志级别数组
   */
  setProductionLevels(levels) {
    this.productionLevels = levels;
  }
  
  /**
   * 立即发送所有缓冲的日志
   */
  async flushLogs() {
    if (this.sendTimer) {
      clearTimeout(this.sendTimer);
      this.sendTimer = null;
    }
    await this._sendBatchLogs();
  }
}

// 创建单例实例
export default new Logger();