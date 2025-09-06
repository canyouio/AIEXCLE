<template>
  <!-- 一个非常简单的测试弹窗，始终显示 -->
  <div class="fixed inset-0 z-9999 flex items-center justify-center p-4 bg-black/50 backdrop-blur-sm" style="position: fixed !important; z-index: 9999 !important; display: flex !important;">
    <div class="relative w-full max-w-md rounded-2xl overflow-hidden bg-white border border-gray-200 shadow-lg">
      <!-- 弹窗头部 -->
      <div class="p-6 border-b border-gray-200">
        <h3 class="text-xl font-semibold text-gray-800 text-center">测试弹窗</h3>
      </div>
      
      <!-- 弹窗内容 -->
      <div class="p-6">
        <div class="flex flex-col items-center space-y-6">
          <div class="text-center">
            <p class="text-gray-600">这是一个测试弹窗，如果您能看到这个消息，说明弹窗功能正常工作。</p>
          </div>
          
          <!-- 关闭按钮 -->
          <button 
            @click="handleClose"
            class="px-6 py-2.5 bg-blue-500 text-white font-medium rounded-lg shadow hover:bg-blue-600 focus:outline-none focus:ring-4 focus:ring-blue-500/30 transition-all duration-300"
          >
            关闭
          </button>
        </div>
      </div>
    </div>
  </div>
</template>

<script>
import { defineComponent } from 'vue';

export default defineComponent({
  name: 'TestModal',
  emits: ['close'],
  setup(props, { emit }) {
    // 处理关闭
    const handleClose = () => {
      emit('close');
    };
    
    return {
      handleClose
    };
  }
});
</script>