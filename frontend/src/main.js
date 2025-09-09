import { createApp } from 'vue';
import ElementPlus from 'element-plus';
import 'element-plus/dist/index.css';
import axios from 'axios';
import App from './App.vue';
import logger from './api/logger';
// 【新增】引入 Univer 的样式文件
import '@univerjs/design/lib/index.css';

// 创建Vue应用实例
const app = createApp(App);

// 使用Element Plus
app.use(ElementPlus);

// 全局配置
app.config.globalProperties.$apiBaseUrl = 'http://127.0.0.1:8000';
app.config.globalProperties.$axios = axios;
app.config.globalProperties.$logger = logger;

// 配置全局错误处理
app.config.errorHandler = (err, instance, info) => {
  logger.error('Vue组件错误', {
    error: err,
    component: instance?.$options?.name || 'unknown',
    info: info
  });
};

// 捕获全局未处理的Promise错误
window.addEventListener('unhandledrejection', (event) => {
  logger.error('未处理的Promise拒绝', {
    reason: event.reason,
    promise: event.promise
  });
  // 不阻止默认行为，让浏览器也记录这个错误
});

// 捕获全局未捕获的JavaScript错误
window.addEventListener('error', (event) => {
  logger.error('全局未捕获错误', {
    message: event.message,
    filename: event.filename,
    lineno: event.lineno,
    colno: event.colno,
    error: event.error
  });
  // 不阻止默认行为，让浏览器也记录这个错误
});

// 挂载应用
app.mount('#app');

// 记录应用启动信息
logger.info(`应用启动成功 - 环境: ${import.meta.env.MODE}`);