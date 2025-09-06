import axios from 'axios';
import { ElMessage, ElMessageBox } from 'element-plus';
import logger from './logger';

// 创建axios实例
const service = axios.create({
  baseURL: 'http://127.0.0.1:8000', // 后端API基础URL
  timeout: 30000, // 请求超时时间（毫秒）
  headers: {
    'Content-Type': 'application/json'
  }
});

// 请求拦截器
service.interceptors.request.use(
  config => {
    // 在发送请求之前做些什么
    
    // 记录请求信息
    const requestStart = Date.now();
    config.meta = {
      startTime: requestStart,
      requestId: Math.random().toString(36).substring(2, 15)
    };
    
    logger.info(`API请求: ${config.method.toUpperCase()} ${config.baseURL}${config.url}`, {
      requestId: config.meta.requestId,
      params: config.params || {},
      data: config.data || {}
    });
    
    // 如果有token，可以在这里添加token
    const token = localStorage.getItem('token');
    if (token) {
      config.headers['Authorization'] = `Bearer ${token}`;
    }
    
    // 显示加载状态
    // 可以在这里添加全局加载指示器
    
    return config;
  },
  error => {
    // 对请求错误做些什么
    logger.error('请求配置错误', error);
    return Promise.reject(error);
  }
);

// 响应拦截器
service.interceptors.response.use(
  response => {
    // 对响应数据做点什么
    const res = response.data;
    
    // 根据后端API的统一响应格式处理
    // 这里假设后端返回的数据格式为 { code: 0, message: '', data: {} }
    
    // 如果后端没有统一的响应格式，可以直接返回response.data
    return res;
  },
  response => {
    // 对响应数据做点什么
    const res = response.data;
    
    // 计算请求耗时
    const requestTime = Date.now() - response.config.meta.startTime;
    
    logger.info(`API响应: ${response.config.method.toUpperCase()} ${response.config.baseURL}${response.config.url} - 状态码: ${response.status} - 耗时: ${requestTime}ms`, {
      requestId: response.config.meta.requestId,
      status: response.status,
      responseTime: requestTime,
      data: res
    });
    
    // 根据后端API的统一响应格式处理
    // 这里假设后端返回的数据格式为 { code: 0, message: '', data: {} }
    
    // 如果后端没有统一的响应格式，可以直接返回response.data
    return res;
  },
  error => {
    // 对响应错误做点什么
    const requestConfig = error.config || {};
    const requestId = requestConfig.meta?.requestId || 'unknown';
    
    let errorInfo = {
      requestId,
      method: requestConfig.method || 'unknown',
      url: requestConfig.baseURL + requestConfig.url || 'unknown'
    };
    
    if (error.response) {
      // 服务器返回了错误状态码
      errorInfo.status = error.response.status;
      errorInfo.data = error.response.data;
      errorInfo.responseTime = Date.now() - (requestConfig.meta?.startTime || 0);
      
      logger.error(`API错误响应: ${errorInfo.method.toUpperCase()} ${errorInfo.url} - 状态码: ${errorInfo.status}`, errorInfo);
      
      switch (error.response.status) {
        case 401:
          // 未授权，跳转到登录页或显示登录对话框
          ElMessageBox.confirm('您的登录已过期，请重新登录', '登录过期', {
            confirmButtonText: '重新登录',
            cancelButtonText: '取消',
            type: 'warning'
          }).then(() => {
            // 这里可以清除登录状态并跳转到登录页
            localStorage.removeItem('token');
            // 可以根据实际需求进行页面跳转
          });
          break;
          
        case 403:
          // 禁止访问
          ElMessage.error('没有权限访问该资源');
          break;
          
        case 404:
          // 请求的资源不存在
          ElMessage.error('请求的资源不存在');
          break;
          
        case 500:
          // 服务器内部错误
          ElMessage.error('服务器内部错误，请稍后重试');
          break;
          
        default:
          ElMessage.error(`请求失败：${error.response.status}`);
      }
    } else if (error.request) {
      // 请求已发送但没有收到响应
      logger.error('API请求无响应', errorInfo);
      ElMessage.error('网络错误，请检查您的网络连接');
    } else {
      // 请求配置出错
      logger.error(`API请求配置错误: ${error.message}`, errorInfo);
      ElMessage.error(`请求配置错误：${error.message}`);
    }
    
    return Promise.reject(error);
  }
);

// 封装常用的请求方法
export default {
  // GET请求
  get(url, params = {}) {
    return service.get(url, {
      params
    });
  },
  
  // POST请求
  post(url, data = {}, config = {}) {
    return service.post(url, data, config);
  },
  
  // PUT请求
  put(url, data = {}) {
    return service.put(url, data);
  },
  
  // DELETE请求
  delete(url, params = {}) {
    return service.delete(url, {
      params
    });
  },
  
  // 文件上传
  upload(url, file, onUploadProgress = null) {
    const formData = new FormData();
    formData.append('file', file);
    
    return service.post(url, formData, {
      headers: {
        'Content-Type': 'multipart/form-data'
      },
      onUploadProgress
    });
  },
  
  // 下载文件
  download(url, params = {}) {
    return service.get(url, {
      params,
      responseType: 'blob' // 设置响应类型为blob
    }).then(response => {
      // 处理文件下载
      const blob = new Blob([response.data]);
      const url = window.URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      
      // 从响应头获取文件名
      const contentDisposition = response.headers['content-disposition'];
      let fileName = 'download';
      if (contentDisposition) {
        const match = contentDisposition.match(/filename="(.+)"/);
        if (match && match[1]) {
          fileName = match[1];
        }
      }
      
      link.download = fileName;
      document.body.appendChild(link);
      link.click();
      
      // 清理
      document.body.removeChild(link);
      window.URL.revokeObjectURL(url);
      
      return { success: true, fileName };
    });
  },
  
  // 设置基础URL
  setBaseURL(baseURL) {
    service.defaults.baseURL = baseURL;
  },
  
  // 设置超时时间
  setTimeout(timeout) {
    service.defaults.timeout = timeout;
  },
  
  // 设置请求头
  setHeaders(headers) {
    Object.assign(service.defaults.headers, headers);
  },
  
  // 添加请求拦截器
  addRequestInterceptor(onFulfilled, onRejected) {
    return service.interceptors.request.use(onFulfilled, onRejected);
  },
  
  // 添加响应拦截器
  addResponseInterceptor(onFulfilled, onRejected) {
    return service.interceptors.response.use(onFulfilled, onRejected);
  },
  
  // 移除请求拦截器
  ejectRequestInterceptor(interceptorId) {
    service.interceptors.request.eject(interceptorId);
  },
  
  // 移除响应拦截器
  ejectResponseInterceptor(interceptorId) {
    service.interceptors.response.eject(interceptorId);
  }
};

// 导出原始的axios实例，以便在需要时可以直接使用
export const axiosInstance = service;

// 导出工具函数，用于处理API错误
export const handleApiError = (error, customMessage = '', context = {}) => {
  logger.error('API请求错误', {
    message: customMessage || 'API请求失败',
    error: error,
    context: context
  });
  
  // 优先使用自定义错误信息
  if (customMessage) {
    ElMessage.error(customMessage);
    return;
  }
  
  // 根据错误类型显示不同的错误信息
  if (error.response) {
    // 服务器返回了错误状态码
    const status = error.response.status;
    const errorData = error.response.data;
    
    if (errorData && errorData.message) {
      ElMessage.error(errorData.message);
    } else {
      switch (status) {
        case 400:
          ElMessage.error('请求参数错误');
          break;
        case 401:
          ElMessage.error('未授权，请重新登录');
          break;
        case 403:
          ElMessage.error('拒绝访问');
          break;
        case 404:
          ElMessage.error('请求的资源不存在');
          break;
        case 500:
          ElMessage.error('服务器内部错误');
          break;
        default:
          ElMessage.error(`请求失败：${status}`);
      }
    }
  } else if (error.request) {
    // 请求已发送但没有收到响应
    ElMessage.error('网络错误，请检查您的网络连接');
  } else {
    // 请求配置出错
    ElMessage.error(`请求错误：${error.message}`);
  }
};

// 导出工具函数，用于显示请求成功信息
export const showSuccessMessage = (message = '操作成功') => {
  ElMessage.success(message);
};

// 导出工具函数，用于显示警告信息
export const showWarningMessage = (message = '警告') => {
  ElMessage.warning(message);
};

// 导出工具函数，用于显示信息提示
export const showInfoMessage = (message = '提示') => {
  ElMessage.info(message);
};