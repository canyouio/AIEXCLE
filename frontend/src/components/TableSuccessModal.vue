<template>
  <!-- 恢复visible条件，但确保样式在Windows环境下正确显示 -->
  <div v-show="visible" class="fixed inset-0 z-9999 flex items-center justify-center p-4 bg-black/50 backdrop-blur-sm" style="position: fixed !important; z-index: 9999 !important; display: flex !important;">
    <div class="relative w-full max-w-md rounded-2xl overflow-hidden glass-effect bg-white/90 backdrop-blur-xl border border-white/70 shadow-[0_10px_40px_rgba(0,0,0,0.15)] animate-[scaleIn_0.3s_ease-out]">
      <!-- 弹窗头部 -->
      <div class="p-6 border-b border-white/70">
        <h3 class="text-xl font-semibold text-gray-800 text-center">表格生成成功</h3>
      </div>
      
      <!-- 弹窗内容 -->
      <div class="p-6">
        <div class="flex flex-col items-center space-y-6">
          <!-- 文件信息展示 -->
          <div class="flex flex-col items-center justify-center w-full p-4 rounded-xl bg-white/70 backdrop-blur-sm border border-white/50 shadow-[0_4px_20px_rgba(0,0,0,0.08)]">
            <div class="w-16 h-16 rounded-full bg-green-100 flex items-center justify-center mb-3">
              <i class="fa fa-file-excel-o text-green-500 text-3xl"></i>
            </div>
            <div class="text-center">
              <div class="font-medium text-gray-800">{{ fileName }}</div>
              <div class="text-sm text-gray-500">生成时间: {{ formatDate(new Date()) }}</div>
            </div>
          </div>
          
          <!-- 按钮组 -->
          <div class="flex w-full space-x-4">
            <!-- 在线编辑按钮 -->
            <button 
              @click="handleOnlineEdit" 
              class="group flex-1 relative glass-btn px-6 py-3 bg-blue-500/90 hover:bg-blue-500 text-white rounded-xl text-sm font-medium shadow-[0_4px_16px_rgba(59,130,246,0.3),inset_0_1px_1px_rgba(255,255,255,0.7)] focus:outline-none transition-all duration-300 hover:scale-102 active:scale-98 flex items-center justify-center"
            >
              <div class="absolute top-2 left-2 h-2 w-2 rounded-full bg-white/50 opacity-0 blur-sm transition-opacity duration-300 group-hover:opacity-100"></div>
              <i class="fa fa-pencil mr-2"></i>
              <span>在线编辑</span>
            </button>
            
            <!-- 下载到本地按钮 -->
            <button 
              @click="handleDownload"
              class="group flex-1 relative glass-btn px-6 py-3 border border-white/70 rounded-xl text-sm text-gray-700 bg-white/60 backdrop-blur-sm hover:bg-white/80 shadow-[0_4px_16px_rgba(0,0,0,0.08),inset_0_1px_1px_rgba(255,255,255,0.6)] focus:outline-none transition-all duration-300 hover:scale-102 active:scale-98 flex items-center justify-center"
            >
              <div class="absolute top-2 left-2 h-2 w-2 rounded-full bg-white/50 opacity-0 blur-sm transition-opacity duration-300 group-hover:opacity-100"></div>
              <i class="fa fa-download mr-2"></i>
              <span>下载到本地</span>
            </button>
          </div>
        </div>
      </div>
      
      <!-- 关闭按钮 -->
      <button 
        @click="handleClose" 
        class="absolute top-4 right-4 w-8 h-8 rounded-full glass-btn bg-white/60 backdrop-blur-sm border border-white/70 flex items-center justify-center text-gray-500 hover:text-gray-700 shadow-[0_2px_8px_rgba(0,0,0,0.06)] transition-all duration-300"
      >
        <i class="fa fa-times"></i>
      </button>
    </div>
  </div>
  
  <!-- 在线编辑弹窗 -->
  <div v-if="showEditor" class="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/50 backdrop-blur-sm">
    <div class="relative w-full max-w-5xl max-h-[90vh] rounded-2xl overflow-hidden glass-effect bg-white/95 backdrop-blur-xl border border-white/70 shadow-[0_10px_40px_rgba(0,0,0,0.15)] animate-[scaleIn_0.3s_ease-out]">
      <!-- 编辑器头部 -->
      <div class="p-4 border-b border-white/70 flex items-center justify-between">
        <h3 class="text-xl font-semibold text-gray-800">在线编辑表格</h3>
        <button 
          @click="showEditor = false" 
          class="w-8 h-8 rounded-full glass-btn bg-white/60 backdrop-blur-sm border border-white/70 flex items-center justify-center text-gray-500 hover:text-gray-700 shadow-[0_2px_8px_rgba(0,0,0,0.06)] transition-all duration-300"
        >
          <i class="fa fa-times"></i>
        </button>
      </div>
      
      <!-- 编辑器内容 -->
      <div class="p-4 overflow-auto max-h-[calc(90vh-120px)]">
        <ExcelEditor 
          ref="excelEditorRef" 
          :initialData="initialData" 
          :fileName="fileName" 
          @saveAndDownload="handleSaveAndDownload"
        />
      </div>
    </div>
  </div>
</template>

<script>
import { ref, defineComponent } from 'vue';
import ExcelEditor from './ExcelEditor.vue';
import axios from 'axios';

export default defineComponent({
  name: 'TableSuccessModal',
  components: {
    ExcelEditor
  },
  props: {
    visible: {
      type: Boolean,
      default: false
    },
    fileName: {
      type: String,
      default: ''
    },
    fileUrl: {
      type: String,
      default: ''
    },
    fileData: {
      type: Array,
      default: () => []
    }
  },
  emits: ['close', 'download', 'edit'],
  setup(props, { emit }) {
    const showEditor = ref(false);
    const initialData = ref([]);
    const excelEditorRef = ref(null);
    
    // 格式化日期
    const formatDate = (date) => {
      return new Intl.DateTimeFormat('zh-CN', {
        year: 'numeric',
        month: '2-digit',
        day: '2-digit',
        hour: '2-digit',
        minute: '2-digit'
      }).format(date);
    };
    
    // 处理在线编辑
    const handleOnlineEdit = () => {
      // 直接使用传入的文件数据
      initialData.value = props.fileData || [];
      showEditor.value = true;
      // 触发编辑事件
      emit('edit');
    };
    
    // 处理下载
    const handleDownload = () => {
      if (props.fileUrl) {
        window.open(props.fileUrl, '_blank');
      }
      // 触发下载事件
      emit('download');
    };
    
    // 处理保存并下载
    const handleSaveAndDownload = () => {
      if (excelEditorRef.value && excelEditorRef.value.exportToExcel) {
        excelEditorRef.value.exportToExcel();
      }
      showEditor.value = false;
      emit('close');
    };
    
    // 处理关闭
    const handleClose = () => {
      emit('close');
    };
    
    return {
      showEditor,
      initialData,
      excelEditorRef,
      formatDate,
      handleOnlineEdit,
      handleDownload,
      handleSaveAndDownload,
      handleClose
    };
  }
});
</script>

<style scoped>
/* 动画样式 */
@keyframes scaleIn {
  from {
    opacity: 0;
    transform: scale(0.9);
  }
  to {
    opacity: 1;
    transform: scale(1);
  }
}

/* 自定义滚动条样式 */
.overflow-auto {
  scrollbar-width: thin;
  scrollbar-color: #c0c4cc #f0f0f0;
}

.overflow-auto::-webkit-scrollbar {
  width: 8px;
  height: 8px;
}

.overflow-auto::-webkit-scrollbar-track {
  background: #f0f0f0;
}

.overflow-auto::-webkit-scrollbar-thumb {
  background-color: #c0c4cc;
  border-radius: 4px;
}

/* 玻璃态效果 */
.glass-effect {
  background: rgba(255, 255, 255, 0.9);
  backdrop-filter: blur(12px);
  -webkit-backdrop-filter: blur(12px);
  border: 1px solid rgba(255, 255, 255, 0.7);
}

/* 提升按钮交互体验 */
button {
  user-select: none;
  -webkit-user-select: none;
}

/* 确保弹窗在Windows上显示正常 */
.fixed {
  position: fixed !important;
}

.z-50 {
  z-index: 50 !important;
}
</style>