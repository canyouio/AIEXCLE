<!-- 文件路径: frontend/src/components/TableSuccessModal.vue -->
<template>
  <div v-if="visible" class="fixed inset-0 z-[100] flex items-center justify-center p-4 bg-black/50 backdrop-blur-sm">

    <!-- 成功信息视图 -->
    <div v-if="!showEditor" 
         class="relative w-full max-w-md rounded-2xl overflow-hidden glass-effect bg-white/90 border border-white/70 shadow-xl animate-[scaleIn_0.3s_ease-out]">
      <div class="p-6 border-b border-white/70">
        <h3 class="text-xl font-semibold text-gray-800 text-center">表格生成成功</h3>
      </div>
      <div class="p-6">
        <div class="flex flex-col items-center space-y-6">
          <div class="flex flex-col items-center justify-center w-full p-4 rounded-xl bg-white/70 border border-white/50">
            <div class="w-16 h-16 rounded-full bg-green-100 flex items-center justify-center mb-3">
              <i class="fa fa-file-excel-o text-green-500 text-3xl"></i>
            </div>
            <div class="text-center">
              <div class="font-medium text-gray-800">{{ fileName }}</div>
              <div class="text-sm text-gray-500">生成时间: {{ formatDate(new Date()) }}</div>
            </div>
          </div>
          <div class="flex w-full space-x-4">
            <button @click="handleOnlineEdit" class="group flex-1 relative glass-btn px-6 py-3 bg-blue-500/90 hover:bg-blue-500 text-white rounded-xl text-sm font-medium shadow-[0_4px_16px_rgba(59,130,246,0.3),inset_0_1px_1px_rgba(255,255,255,0.7)] focus:outline-none transition-all duration-300 hover:scale-102 active:scale-98 flex items-center justify-center">
              <i class="fa fa-pencil mr-2"></i>
              <span>在线编辑</span>
            </button>
            <button @click="handleDownload" class="group flex-1 relative glass-btn px-6 py-3 border border-white/70 rounded-xl text-sm text-gray-700 bg-white/60 backdrop-blur-sm hover:bg-white/80 shadow-[0_4px_16px_rgba(0,0,0,0.08),inset_0_1px_1px_rgba(255,255,255,0.6)] focus:outline-none transition-all duration-300 hover:scale-102 active:scale-98 flex items-center justify-center">
              <i class="fa fa-download mr-2"></i>
              <span>下载到本地</span>
            </button>
          </div>
        </div>
      </div>
      <button @click="handleClose" class="absolute top-4 right-4 w-8 h-8 rounded-full glass-btn bg-white/60 backdrop-blur-sm border border-white/70 flex items-center justify-center text-gray-500 hover:text-gray-700 shadow-[0_2px_8px_rgba(0,0,0,0.06)] transition-all duration-300">
        <i class="fa fa-times"></i>
      </button>
    </div>
    
    <!-- 在线编辑器视图 -->
    <div v-else 
         class="relative w-full max-w-6xl h-[90vh] rounded-2xl overflow-hidden bg-white shadow-xl animate-[scaleIn_0.3s_ease-out] flex flex-col">
      <div class="p-4 border-b border-gray-200 flex items-center justify-between flex-shrink-0">
        <h3 class="text-xl font-semibold text-gray-800">在线编辑: {{ fileName }}</h3>
        <div>
            <button 
                @click="handleSave"
                :disabled="isSaving"
                class="bg-blue-500 hover:bg-blue-600 text-white font-bold py-2 px-4 rounded-lg transition duration-300 ease-in-out disabled:bg-gray-400 flex items-center"
            >
                <i v-if="isSaving" class="fa fa-spinner fa-spin mr-2"></i>
                <i v-else class="fa fa-save mr-2"></i>
                保存并关闭
            </button>
        </div>
      </div>
      <!-- 【终极CSS修复】这是唯一正确的布局 -->
      <div class="flex-grow min-h-0">
        <UniverEditor v-if="fileData"
          ref="univerEditorRef" 
          :initialData="fileData"
        />
        <div v-else class="text-center p-10 text-gray-500">
            数据加载中...
        </div>
      </div>
    </div>
  </div>
</template>

<script>
import { ref, defineComponent, watch } from 'vue';
import UniverEditor from './UniverEditor.vue';
import axios from 'axios';
import { ElMessage } from 'element-plus';

export default defineComponent({
  name: 'TableSuccessModal',
  components: { UniverEditor },
  props: {
    visible: { type: Boolean, default: false },
    fileName: { type: String, default: '' },
    fileUrl: { type: String, default: '' },
    fileData: { type: Array, default: () => [] }
  },
  emits: ['close', 'download'],
  setup(props, { emit }) {
    const showEditor = ref(false);
    const univerEditorRef = ref(null);
    const isSaving = ref(false);

    watch(() => props.visible, (isVisible) => {
      if (isVisible) showEditor.value = false;
    });
    
    const formatDate = (date) => new Intl.DateTimeFormat('zh-CN', { year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit' }).format(date);
    
    const handleOnlineEdit = () => {
      showEditor.value = true;
    };
    
    const handleDownload = () => {
      if (props.fileUrl) window.open(props.fileUrl, '_blank');
      emit('download');
    };
    
    const handleSave = async () => {
        if (!univerEditorRef.value) {
            ElMessage.error('编辑器实例不存在！');
            return;
        }
        isSaving.value = true;
        try {
            const latestData = univerEditorRef.value.getCurrentData();
            if (!latestData || latestData.length === 0) {
              ElMessage.warning('没有可保存的数据。');
              isSaving.value = false;
              return;
            }
            await axios.post('/api/excel/save_data', {
                file_name: props.fileName,
                sheet_data: latestData
            });
            ElMessage.success('保存成功！');
            emit('close');
        } catch (error) {
            console.error('保存失败:', error);
            ElMessage.error('保存失败，请查看控制台');
        } finally {
            isSaving.value = false;
        }
    };
    
    const handleClose = () => {
      emit('close');
    };
    
    return {
      showEditor, univerEditorRef, isSaving,
      formatDate, handleOnlineEdit, handleDownload, handleSave, handleClose
    };
  }
});
</script>

<style scoped>
@keyframes scaleIn { from { opacity: 0; transform: scale(0.95); } to { opacity: 1; transform: scale(1); } }
.glass-effect { background: rgba(255, 255, 255, 0.85); backdrop-filter: blur(16px); -webkit-backdrop-filter: blur(16px); }
</style>