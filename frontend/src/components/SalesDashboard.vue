<template>
  <!-- iOS 26液态背景 -->
  <div class="absolute inset-0 bg-gradient-to-br from-blue-100/80 via-indigo-100/70 to-purple-200/80"></div>
  <div class="absolute top-1/4 left-1/5 h-96 w-96 animate-pulse rounded-full bg-gradient-to-br from-blue-400/15 to-indigo-400/15 blur-3xl"></div>
  <div class="absolute right-1/5 bottom-1/4 h-80 w-80 animate-pulse rounded-full bg-gradient-to-br from-purple-400/15 to-pink-400/15 blur-3xl delay-1000"></div>
  

  
  <div class="flex h-full relative z-10">
    <!-- 侧边栏 - iOS 26风格 -->
    <aside class="w-64 flex flex-col glass-effect bg-white/30 backdrop-blur-xl border-r border-white/70 shadow-[0_4px_20px_rgba(0,0,0,0.08),inset_0_1px_1px_rgba(255,255,255,0.5)]">
      <div class="p-5 border-b border-gray-100">
        <div class="flex items-center space-x-3">
          <!-- 品牌LOGO -->
          <div class= "relative w-10 h-10 rounded-2xl overflow-hidden 
                      bg-white/40 backdrop-blur-2xl backdrop-saturate-150 
                      border border-white/50 shadow-[0_4px_16px_rgba(0,0,0,0.1)] flex items-center justify-center">
            <i class="fa fa-file-excel-o text-blue-500 text-lg"></i>
          </div>
          <h2 class="text-xl font-bold text-gray-800">ExcelGenius</h2>
        </div>
      </div>
      
      <!-- 导航菜单 -->
      <nav class="flex-1 overflow-y-auto py-4">
        <div class="px-5 mb-4 text-xs font-semibold text-gray-400 uppercase tracking-wider">
          功能菜单
        </div>
        
        <!-- 高仿苹果动态手风琴按钮 - 功能菜单 -->
        <div class="space-y-4 px-3">
          <!-- 首页按钮 -->
          <div class="w-14 group transition-all duration-300" :class="activeTab === 'overview' ? 'w-full' : ''">
            <div class="relative overflow-hidden rounded-full">
              <button 
                @click="activeTab = 'overview'" 
                class="flex items-center h-14 glass-btn bg-white/50 backdrop-blur-sm border border-white/60 shadow-[0_2px_8px_rgba(0,0,0,0.06),inset_0_1px_1px_rgba(255,255,255,0.6)] transition-all duration-300 w-full"
              >
                <div class="w-14 h-14 rounded-full flex items-center justify-center">
                  <i class="fa fa-home text-xl" :class="activeTab === 'overview' ? 'text-blue-500' : 'text-gray-600 group-hover:text-blue-500'" style="transition: all 0.3s ease-out;"></i>
                </div>
                <span class="ml-2 font-medium opacity-0 group-hover:opacity-100 transition-all duration-300" :style="activeTab === 'overview' ? { opacity: '1' } : {}">首页</span>
              </button>
            </div>
          </div>
          
          <!-- 创建Excel按钮 -->
          <div class="w-14 group transition-all duration-300" :class="activeTab === 'create' ? 'w-full' : ''">
            <div class="relative overflow-hidden rounded-full">
              <button 
                @click="activeTab = 'create'" 
                class="flex items-center h-14 glass-btn bg-white/50 backdrop-blur-sm border border-white/60 shadow-[0_2px_8px_rgba(0,0,0,0.06),inset_0_1px_1px_rgba(255,255,255,0.6)] transition-all duration-300 w-full"
              >
                <div class="w-14 h-14 rounded-full flex items-center justify-center">
                  <i class="fa fa-plus-circle text-xl" :class="activeTab === 'create' ? 'text-blue-500' : 'text-gray-600 group-hover:text-blue-500'" style="transition: all 0.3s ease-out;"></i>
                </div>
                <span class="ml-2 font-medium opacity-0 group-hover:opacity-100 transition-all duration-300" :style="activeTab === 'create' ? { opacity: '1' } : {}">创建Excel</span>
              </button>
            </div>
          </div>
          
          <!-- 编辑Excel按钮 -->
          <div class="w-14 group transition-all duration-300" :class="activeTab === 'edit' ? 'w-full' : ''">
            <div class="relative overflow-hidden rounded-full">
              <button 
                @click="activeTab = 'edit'" 
                class="flex items-center h-14 glass-btn bg-white/50 backdrop-blur-sm border border-white/60 shadow-[0_2px_8px_rgba(0,0,0,0.06),inset_0_1px_1px_rgba(255,255,255,0.6)] transition-all duration-300 w-full"
              >
                <div class="w-14 h-14 rounded-full flex items-center justify-center">
                  <i class="fa fa-pencil text-xl" :class="activeTab === 'edit' ? 'text-blue-500' : 'text-gray-600 group-hover:text-blue-500'" style="transition: all 0.3s ease-out;"></i>
                </div>
                <span class="ml-2 font-medium opacity-0 group-hover:opacity-100 transition-all duration-300" :style="activeTab === 'edit' ? { opacity: '1' } : {}">编辑Excel</span>
              </button>
            </div>
          </div>
          
          <!-- 已保存文件按钮 -->
          <div class="w-14 group transition-all duration-300" :class="activeTab === 'saved' ? 'w-full' : ''">
            <div class="relative overflow-hidden rounded-full">
              <button 
                @click="activeTab = 'saved'" 
                class="flex items-center h-14 glass-btn bg-white/50 backdrop-blur-sm border border-white/60 shadow-[0_2px_8px_rgba(0,0,0,0.06),inset_0_1px_1px_rgba(255,255,255,0.6)] transition-all duration-300 w-full"
              >
                <div class="w-14 h-14 rounded-full flex items-center justify-center">
                  <i class="fa fa-floppy-o text-xl" :class="activeTab === 'saved' ? 'text-blue-500' : 'text-gray-600 group-hover:text-blue-500'" style="transition: all 0.3s ease-out;"></i>
                </div>
                <span class="ml-2 font-medium opacity-0 group-hover:opacity-100 transition-all duration-300" :style="activeTab === 'saved' ? { opacity: '1' } : {} ">已保存文件</span>
              </button>
            </div>
          </div>
          
          <!-- 数据分析按钮 -->
          <div class="w-14 group transition-all duration-300" :class="activeTab === 'analyze' ? 'w-full' : ''">
            <div class="relative overflow-hidden rounded-full">
              <button 
                @click="activeTab = 'analyze'" 
                class="flex items-center h-14 glass-btn bg-white/50 backdrop-blur-sm border border-white/60 shadow-[0_2px_8px_rgba(0,0,0,0.06),inset_0_1px_1px_rgba(255,255,255,0.6)] transition-all duration-300 w-full"
              >
                <div class="w-14 h-14 rounded-full flex items-center justify-center">
                  <i class="fa fa-bar-chart text-xl" :class="activeTab === 'analyze' ? 'text-blue-500' : 'text-gray-600 group-hover:text-blue-500'" style="transition: all 0.3s ease-out;"></i>
                </div>
                <span class="ml-2 font-medium opacity-0 group-hover:opacity-100 transition-all duration-300" :style="activeTab === 'analyze' ? { opacity: '1' } : {} ">数据分析</span>
              </button>
            </div>
          </div>
        </div>
        
        <div class="px-5 mt-8 mb-4 text-xs font-semibold text-gray-400 uppercase tracking-wider">
          模板库
        </div>
        
        <!-- 高仿苹果动态手风琴按钮 - 模板库 -->
        <div class="space-y-4 px-3">
          <!-- 销售报表模板 -->
          <div class="w-14 group transition-all duration-300">
            <div class="relative overflow-hidden rounded-full">
              <button 
                @click="useTemplate('sales')" 
                class="flex items-center h-14 glass-btn bg-white/50 backdrop-blur-sm border border-white/60 shadow-[0_2px_8px_rgba(0,0,0,0.06),inset_0_1px_1px_rgba(255,255,255,0.6)] transition-all duration-300 w-full"
              >
                <div class="w-14 h-14 rounded-full flex items-center justify-center">
                  <i class="fa fa-line-chart text-xl text-gray-600 group-hover:text-blue-500" style="transition: all 0.3s ease-out;"></i>
                </div>
                <span class="ml-2 font-medium opacity-0 group-hover:opacity-100 transition-all duration-300">销售报表</span>
              </button>
            </div>
          </div>
          
          <!-- 库存管理模板 -->
          <div class="w-14 group transition-all duration-300">
            <div class="relative overflow-hidden rounded-full">
              <button 
                @click="useTemplate('inventory')" 
                class="flex items-center h-14 glass-btn bg-white/50 backdrop-blur-sm border border-white/60 shadow-[0_2px_8px_rgba(0,0,0,0.06),inset_0_1px_1px_rgba(255,255,255,0.6)] transition-all duration-300 w-full"
              >
                <div class="w-14 h-14 rounded-full flex items-center justify-center">
                  <i class="fa fa-cubes text-xl text-gray-600 group-hover:text-blue-500" style="transition: all 0.3s ease-out;"></i>
                </div>
                <span class="ml-2 font-medium opacity-0 group-hover:opacity-100 transition-all duration-300">库存管理</span>
              </button>
            </div>
          </div>
          
          <!-- 财务报表模板 -->
          <div class="w-14 group transition-all duration-300">
            <div class="relative overflow-hidden rounded-full">
              <button 
                @click="useTemplate('finance')" 
                class="flex items-center h-14 glass-btn bg-white/50 backdrop-blur-sm border border-white/60 shadow-[0_2px_8px_rgba(0,0,0,0.06),inset_0_1px_1px_rgba(255,255,255,0.6)] transition-all duration-300 w-full"
              >
                <div class="w-14 h-14 rounded-full flex items-center justify-center">
                  <i class="fa fa-money text-xl text-gray-600 group-hover:text-blue-500" style="transition: all 0.3s ease-out;"></i>
                </div>
                <span class="ml-2 font-medium opacity-0 group-hover:opacity-100 transition-all duration-300">财务报表</span>
              </button>
            </div>
          </div>
        </div>
      </nav>
    </aside>
    
    <!-- 主内容区域 -->
    <main class="flex-1 flex flex-col">
      <!-- 顶部导航 - iOS 26风格 -->
      <header class="h-16 border-b border-white/70 flex items-center px-6 glass-effect bg-white/30 backdrop-blur-xl shadow-[0_4px_20px_rgba(0,0,0,0.08),inset_0_1px_1px_rgba(255,255,255,0.5)]">
        <div class="flex items-center space-x-6">
          <h1 class="text-xl font-bold text-gray-800" v-if="activeTab === 'overview'">欢迎使用 ExcelGenius</h1>
          <h1 class="text-xl font-bold text-gray-800" v-else-if="activeTab === 'create'">创建 Excel</h1>
          <h1 class="text-xl font-bold text-gray-800" v-else-if="activeTab === 'edit'">编辑 Excel</h1>
          <h1 class="text-xl font-bold text-gray-800" v-else-if="activeTab === 'saved'">已保存文件</h1>
          <h1 class="text-xl font-bold text-gray-800" v-else-if="activeTab === 'analyze'">Excel 数据分析</h1>
        </div>
        
        <div class="ml-auto flex items-center space-x-4">
          <!-- iOS 26风格按钮 -->
          <button class="w-10 h-10 rounded-full glass-btn bg-white/40 backdrop-blur-sm border border-white/60 shadow-[0_2px_8px_rgba(0,0,0,0.06),inset_0_1px_1px_rgba(255,255,255,0.6)] flex items-center justify-center text-gray-600 hover:bg-white/60 transition-all duration-300 group">
            <div class="absolute top-1 left-1 h-2 w-2 rounded-full bg-white/50 opacity-0 blur-sm transition-opacity duration-300 group-hover:opacity-100"></div>
            <i class="fa fa-bell"></i>
          </button>
          
          <!-- 用户信息 -->
          <div class="flex items-center space-x-3 cursor-pointer p-2 rounded-full glass-btn hover:bg-white/40 backdrop-blur-sm transition-all duration-300">
            <div class="w-10 h-10 rounded-full overflow-hidden border-2 border-white/80 shadow-[0_2px_12px_rgba(0,0,0,0.1)]">
              <img src="/src/assets/karina.jpg" alt="用户头像" class="w-full h-full object-cover">
            </div>
            <span class="text-gray-700 font-medium">用户</span>
          </div>
        </div>
      </header>
      
      <!-- 内容区域 -->
      <div class="flex-1 overflow-auto p-6">
        <!-- 首页内容 -->
        <div v-if="activeTab === 'overview'" class="space-y-6">
          <!-- 欢迎卡片 - iOS风格 -->
            <div class="rounded-2xl p-6 glass-effect bg-white/60 backdrop-blur-xl border border-white/70 shadow-[0_8px_32px_rgba(0,0,0,0.12)]">
            <h1 class="text-2xl font-bold text-gray-800 mb-4">欢迎使用ExcelGenius</h1>
            <p class="text-gray-600 mb-6">使用AI创建和编辑Excel表格，提高您的工作效率。</p>
            
            <div class="grid grid-cols-1 md:grid-cols-3 gap-6">
              <!-- iOS 26风格功能卡片 -->
            <div class="group relative rounded-2xl p-6 cursor-pointer hover:bg-white/60 transition-all duration-300 transform hover:-translate-y-1 glass-effect bg-white/30 backdrop-blur-sm border border-white/50 shadow-[0_4px_20px_rgba(0,0,0,0.08),inset_0_1px_1px_rgba(255,255,255,0.5)]" @click="activeTab = 'create'">
              <div class="absolute top-2 left-2 h-3 w-3 rounded-full bg-white/50 opacity-0 blur-sm transition-opacity duration-300 group-hover:opacity-100"></div>
                <i class="fa fa-plus-circle text-3xl text-blue-500 mb-3"></i>
                <h3 class="text-lg font-semibold text-gray-800 mb-1">创建新表格</h3>
                <p class="text-gray-500 text-sm">通过自然语言描述快速创建Excel表格</p>
              </div>
              
              <div class="group relative rounded-2xl p-6 cursor-pointer hover:bg-white/60 transition-all duration-300 transform hover:-translate-y-1 glass-effect bg-white/30 backdrop-blur-sm border border-white/50 shadow-[0_4px_20px_rgba(0,0,0,0.08),inset_0_1px_1px_rgba(255,255,255,0.5)]" @click="activeTab = 'edit'">
                <div class="absolute top-2 left-2 h-3 w-3 rounded-full bg-white/50 opacity-0 blur-sm transition-opacity duration-300 group-hover:opacity-100"></div>
                <i class="fa fa-pencil text-3xl text-purple-500 mb-3"></i>
                <h3 class="text-lg font-semibold text-gray-800 mb-1">编辑现有表格</h3>
                <p class="text-gray-500 text-sm">上传Excel文件并通过自然语言指令进行编辑</p>
              </div>
              
              <div class="group relative rounded-2xl p-6 cursor-pointer hover:bg-white/60 transition-all duration-300 transform hover:-translate-y-1 glass-effect bg-white/30 backdrop-blur-sm border border-white/50 shadow-[0_4px_20px_rgba(0,0,0,0.08),inset_0_1px_1px_rgba(255,255,255,0.5)]" @click="activeTab = 'saved'">
                <div class="absolute top-2 left-2 h-3 w-3 rounded-full bg-white/50 opacity-0 blur-sm transition-opacity duration-300 group-hover:opacity-100"></div>
                <i class="fa fa-folder-open text-3xl text-green-500 mb-3"></i>
                <h3 class="text-lg font-semibold text-gray-800 mb-1">查看历史文件</h3>
                <p class="text-gray-500 text-sm">访问和管理您之前创建或编辑的文件</p>
              </div>
            </div>
          </div>
          
          <div class="grid grid-cols-1 md:grid-cols-2 gap-6">
            <!-- 模板卡片 - iOS风格 -->
            <div class="rounded-2xl p-6 glass-effect bg-white/60 backdrop-blur-xl border border-white/70 shadow-[0_8px_32px_rgba(0,0,0,0.12)]">
              <h2 class="text-xl font-semibold text-gray-800 mb-4">最近使用的模板</h2>
              <div class="space-y-3">
                <div class="flex items-center justify-between p-3 rounded-xl hover:bg-gray-50/70 cursor-pointer transition-colors duration-200 bg-white/40 backdrop-blur-sm border border-white/50">
                  <div class="flex items-center space-x-3">
                    <i class="fa fa-line-chart text-blue-500"></i>
                    <span class="text-gray-700">销售报表模板</span>
                  </div>
                  <button class="group relative glass-btn text-blue-500 hover:text-blue-600 px-3 py-1.5 rounded-xl hover:bg-white/40 transition-all duration-300 hover:shadow-[0_2px_8px_rgba(0,0,0,0.06)]">
                <span>使用</span>
                <div class="absolute top-1 left-1 h-2 w-2 rounded-full bg-white/50 opacity-0 blur-sm transition-opacity duration-300 group-hover:opacity-100"></div>
              </button>
                </div>
              </div>
            </div>
            
            <!-- 提示卡片 - iOS风格 -->
            <div class="rounded-2xl p-6 glass-effect bg-white/60 backdrop-blur-xl border border-white/70 shadow-[0_8px_32px_rgba(0,0,0,0.12)]">
              <h2 class="text-xl font-semibold text-gray-800 mb-4">使用提示</h2>
              <ul class="space-y-3 text-gray-600">
                <li class="flex items-start space-x-3 p-3 rounded-xl hover:bg-gray-50/70 transition-colors duration-200 bg-white/40 backdrop-blur-sm border border-white/50">
                  <i class="fa fa-lightbulb-o text-yellow-500 mt-1"></i>
                  <span>使用清晰简洁的语言描述您需要的Excel表格结构</span>
                </li>
                <li class="flex items-start space-x-3 p-3 rounded-xl hover:bg-gray-50/70 transition-colors duration-200 bg-white/40 backdrop-blur-sm border border-white/50">
                  <i class="fa fa-lightbulb-o text-yellow-500 mt-1"></i>
                  <span>编辑时，请具体说明需要修改的内容和位置</span>
                </li>
                <li class="flex items-start space-x-3 p-3 rounded-xl hover:bg-gray-50/70 transition-colors duration-200 bg-white/40 backdrop-blur-sm border border-white/50">
                  <i class="fa fa-lightbulb-o text-yellow-500 mt-1"></i>
                  <span>生成的文件将自动保存，您可以在"已保存文件"中查看</span>
                </li>
              </ul>
            </div>
          </div>
        </div>
        
        <!-- 创建Excel内容 -->
        <div v-if="activeTab === 'create'" class="space-y-6">
          <!-- iOS风格创建表格卡片 -->
              <div class="rounded-2xl p-6 glass-effect bg-white/60 backdrop-blur-xl border border-white/70 shadow-[0_8px_32px_rgba(0,0,0,0.12)]">
            <h1 class="text-2xl font-bold text-gray-800 mb-4">创建Excel表格</h1>
            <p class="text-gray-600 mb-6">请用自然语言描述您需要的Excel表格结构，AI将为您自动创建。</p>
            
            <div class="space-y-5">
              <div>
                <label class="block text-sm font-medium text-gray-700 mb-1">表格描述</label>
                <!-- iOS 26风格文本域 -->
              <textarea 
                v-model="excelDescription" 
                placeholder="例如：创建一个销售业绩表，包含产品名称、销售额、目标、完成率等字段..."
                class="w-full px-4 py-3 border border-white/70 rounded-xl bg-white/40 backdrop-blur-sm shadow-[0_2px_8px_rgba(0,0,0,0.06)] focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent min-h-[120px] transition-all duration-300"
              ></textarea>
              </div>
              
              <div class="grid grid-cols-1 md:grid-cols-2 gap-4">
                <div>
                  <label class="block text-sm font-medium text-gray-700 mb-1">工作表名称</label>
                  <!-- iOS 26风格输入框 -->
                  <input 
                    v-model="sheetName" 
                    type="text" 
                    placeholder="例如：销售数据"
                    class="w-full px-4 py-3 border border-white/70 rounded-xl bg-white/40 backdrop-blur-sm shadow-[0_2px_8px_rgba(0,0,0,0.06)] focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent transition-all duration-300"
                  >
                </div>
                
                <div>
                  <label class="block text-sm font-medium text-gray-700 mb-1">文件名称</label>
                  <!-- iOS 26风格输入框 -->
                  <input 
                    v-model="fileName" 
                    type="text" 
                    placeholder="例如：2023年销售报表"
                    class="w-full px-4 py-3 border border-white/70 rounded-xl bg-white/40 backdrop-blur-sm shadow-[0_2px_8px_rgba(0,0,0,0.06)] focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent transition-all duration-300"
                  >
                </div>
              </div>
              
              <div class="flex justify-end space-x-3 pt-4">
              <!-- iOS 26风格按钮 -->
                <button class="group relative glass-btn px-6 py-2.5 border border-white/70 rounded-xl text-sm text-gray-700 bg-white/40 backdrop-blur-sm hover:bg-white/60 shadow-[0_2px_8px_rgba(0,0,0,0.06),inset_0_1px_1px_rgba(255,255,255,0.6)] focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-blue-500 transition-all duration-300 hover:scale-102 active:scale-95" @click="resetCreateForm">
                  重置
                  <div class="absolute top-1 left-2 h-2 w-2 rounded-full bg-white/40 opacity-0 blur-sm transition-opacity duration-300 group-hover:opacity-100"></div>
                </button>

                <button class="group relative glass-btn px-6 py-2.5 bg-blue-500/80 hover:bg-blue-500 text-white rounded-xl text-sm font-medium shadow-[0_2px_8px_rgba(0,0,0,0.1),inset_0_1px_1px_rgba(255,255,255,0.7)] focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-blue-500 transition-all duration-300 flex items-center space-x-1 hover:scale-102 active:scale-95" @click="generateExcel" :disabled="isGenerating">
                  <div class="absolute top-1 left-2 h-2 w-2 rounded-full bg-white/50 opacity-0 blur-sm transition-opacity duration-300 group-hover:opacity-100"></div>
                  <i v-if="isGenerating" class="fa fa-spinner fa-spin"></i>
                  <span>生成Excel</span>
                </button>
              </div>
            </div>
          </div>
          
          <!-- 生成结果区域 -->
            <div v-if="generatedFile" class="rounded-2xl p-6 glass-effect bg-white/60 backdrop-blur-xl border border-white/70 shadow-[0_8px_32px_rgba(0,0,0,0.12)]">
            <h2 class="text-xl font-semibold text-gray-800 mb-4">生成结果</h2>
            <div class="flex items-center justify-between rounded-xl bg-white/70 backdrop-blur-sm p-4 mb-4 border border-white/50 shadow-[0_2px_12px_rgba(0,0,0,0.05)]">
              <div class="flex items-center space-x-3">
                <i class="fa fa-file-excel-o text-green-500 text-2xl"></i>
                <div>
                  <div class="font-medium text-gray-800">{{ generatedFile.fileName }}</div>
                  <div class="text-sm text-gray-500">刚刚生成</div>
                </div>
              </div>
              <button class="group relative glass-btn text-blue-500 hover:text-blue-600 px-3 py-1.5 rounded-xl hover:bg-white/40 transition-all duration-300 flex items-center space-x-1 text-sm hover:shadow-[0_2px_8px_rgba(0,0,0,0.06)]">
                  <div class="absolute top-1 left-1 h-2 w-2 rounded-full bg-white/50 opacity-0 blur-sm transition-opacity duration-300 group-hover:opacity-100"></div>
                <i class="fa fa-download"></i>
                <span>下载</span>
              </button>
            </div>
          </div>
        </div>
        
        <!-- 编辑Excel内容 -->
        <div v-if="activeTab === 'edit'" class="space-y-6">
          <!-- iOS风格编辑表格卡片 -->
              <div class="rounded-2xl p-6 glass-effect bg-white/60 backdrop-blur-xl border border-white/70 shadow-[0_8px_32px_rgba(0,0,0,0.12)]">
            <h1 class="text-2xl font-bold text-gray-800 mb-4">编辑Excel表格</h1>
            <p class="text-gray-600 mb-6">上传Excel文件并通过自然语言指令进行编辑。</p>
            
            <div class="space-y-5">
              <div>
                <label class="block text-sm font-medium text-gray-700 mb-1">上传Excel文件</label>
                <!-- iOS 26风格文件上传区域 -->
                  <div class="border-2 border-dashed border-white/70 rounded-xl p-6 text-center hover:border-blue-300 transition-colors cursor-pointer glass-effect bg-white/30 backdrop-blur-sm shadow-[0_2px_8px_rgba(0,0,0,0.06)]" @click="fileInput.click()">
                  <input ref="fileInput" type="file" accept=".xlsx, .xls" class="hidden" @change="handleFileUpload">
                  <i class="fa fa-cloud-upload text-gray-400 text-3xl mb-2"></i>
                  <p class="text-gray-500" v-if="!uploadedFile">拖放文件到此处，或点击选择文件</p>
                  <div v-else class="flex items-center justify-center space-x-2">
                    <i class="fa fa-file-excel-o text-green-500"></i>
                    <span class="text-gray-700">{{ uploadedFile.name }}</span>
                    <button class="group relative glass-btn text-gray-400 hover:text-red-500 p-1.5 rounded-full hover:bg-white/40 transition-colors duration-300">
                    <div class="absolute top-1 left-1 h-1 w-1 rounded-full bg-white/50 opacity-0 blur-sm transition-opacity duration-300 group-hover:opacity-100"></div>
                      <i class="fa fa-times"></i>
                    </button>
                  </div>
                </div>
              </div>
              
              <div v-if="uploadedFile">
                <label class="block text-sm font-medium text-gray-700 mb-1">编辑指令</label>
                <!-- iOS 26风格文本域 -->
                  <textarea 
                    v-model="editInstructions" 
                    placeholder="例如：在第一行添加表头，将第二列数据格式化为货币..."
                    class="w-full px-4 py-3 border border-white/70 rounded-xl bg-white/40 backdrop-blur-sm shadow-[0_2px_8px_rgba(0,0,0,0.06)] focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent min-h-[120px] transition-all duration-300"
                  ></textarea>
              </div>
              
              <div class="flex justify-end space-x-3 pt-4">
                <!-- iOS 26风格按钮 -->
                <button class="group relative glass-btn px-6 py-2.5 border border-white/70 rounded-xl text-sm text-gray-700 bg-white/40 backdrop-blur-sm hover:bg-white/60 shadow-[0_2px_8px_rgba(0,0,0,0.06),inset_0_1px_1px_rgba(255,255,255,0.6)] focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-blue-500 transition-all duration-300 hover:scale-102 active:scale-95" @click="resetEditForm" :disabled="!uploadedFile">
                  重置
                  <div class="absolute top-1 left-2 h-2 w-2 rounded-full bg-white/40 opacity-0 blur-sm transition-opacity duration-300 group-hover:opacity-100"></div>
                </button>
                <button class="group relative glass-btn px-6 py-2.5 bg-blue-500/80 hover:bg-blue-500 text-white rounded-xl text-sm font-medium shadow-[0_2px_8px_rgba(0,0,0,0.1),inset_0_1px_1px_rgba(255,255,255,0.7)] focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-blue-500 transition-all duration-300 flex items-center space-x-1 hover:scale-102 active:scale-95" @click="editExcel" :disabled="isEditing || !uploadedFile">
                  <div class="absolute top-1 left-2 h-2 w-2 rounded-full bg-white/50 opacity-0 blur-sm transition-opacity duration-300 group-hover:opacity-100"></div>
                  <i v-if="isEditing" class="fa fa-spinner fa-spin"></i>
                  <span>编辑Excel</span>
                </button>
              </div>
            </div>
          </div>
          
          <!-- 编辑结果区域 -->
            <div v-if="editedFile" class="rounded-2xl p-6 glass-effect bg-white/60 backdrop-blur-xl border border-white/70 shadow-[0_8px_32px_rgba(0,0,0,0.12)]">
            <h2 class="text-xl font-semibold text-gray-800 mb-4">编辑结果</h2>
            <div class="flex items-center justify-between rounded-xl bg-white/70 backdrop-blur-sm p-4 mb-4 border border-white/50 shadow-[0_2px_12px_rgba(0,0,0,0.05)]">
              <div class="flex items-center space-x-3">
                <i class="fa fa-file-excel-o text-green-500 text-2xl"></i>
                <div>
                  <div class="font-medium text-gray-800">{{ editedFile.fileName }}</div>
                  <div class="text-sm text-gray-500">刚刚编辑</div>
                </div>
              </div>
              <button class="glass-btn text-blue-500 hover:text-blue-600 px-3 py-1.5 rounded-lg hover:bg-blue-50 transition-colors duration-200 flex items-center space-x-1 text-sm">
                <i class="fa fa-download"></i>
                <span>下载</span>
              </button>
            </div>
          </div>
        </div>
        
        <!-- 数据分析内容 -->
        <div v-if="activeTab === 'analyze'" class="space-y-6">
          <!-- iOS风格数据分析卡片 -->
          <div class="rounded-2xl p-6 glass-effect bg-white/60 backdrop-blur-xl border border-white/70 shadow-[0_8px_32px_rgba(0,0,0,0.12)]">
            <h1 class="text-2xl font-bold text-gray-800 mb-4">Excel数据分析</h1>
            <p class="text-gray-600 mb-6">上传Excel文件，AI将自动分析数据并生成关键洞察、趋势分析和数据异常检测结果。</p>
            
            <div class="space-y-5">
              <div>
                <label class="block text-sm font-medium text-gray-700 mb-1">上传Excel文件</label>
                <!-- iOS 26风格文件上传区域 -->
                <div class="border-2 border-dashed border-white/70 rounded-xl p-6 text-center hover:border-blue-300 transition-colors cursor-pointer glass-effect bg-white/30 backdrop-blur-sm shadow-[0_2px_8px_rgba(0,0,0,0.06)]" @click="analysisFileInput.click()">
                  <input ref="analysisFileInput" type="file" accept=".xlsx, .xls" class="hidden" @change="handleAnalysisFileUpload">
                  <i class="fa fa-cloud-upload text-gray-400 text-3xl mb-2"></i>
                  <p class="text-gray-500" v-if="!analysisFile">拖放文件到此处，或点击选择文件</p>
                  <div v-else class="flex items-center justify-center space-x-2">
                    <i class="fa fa-file-excel-o text-green-500"></i>
                    <span class="text-gray-700">{{ analysisFile.name }}</span>
                    <button class="group relative glass-btn text-gray-400 hover:text-red-500 p-1.5 rounded-full hover:bg-white/40 transition-colors duration-300" @click.stop="removeAnalysisFile">
                      <div class="absolute top-1 left-1 h-1 w-1 rounded-full bg-white/50 opacity-0 blur-sm transition-opacity duration-300 group-hover:opacity-100"></div>
                      <i class="fa fa-times"></i>
                    </button>
                  </div>
                </div>
              </div>
              
              <div class="flex justify-end pt-4">
                <!-- iOS 26风格按钮 -->
                <button class="group relative glass-btn px-6 py-2.5 bg-blue-500/80 hover:bg-blue-500 text-white rounded-xl text-sm font-medium shadow-[0_2px_8px_rgba(0,0,0,0.1),inset_0_1px_1px_rgba(255,255,255,0.7)] focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-blue-500 transition-all duration-300 flex items-center space-x-1 hover:scale-102 active:scale-95" @click="analyzeExcel" :disabled="isAnalyzing || !analysisFile">
                  <div class="absolute top-1 left-2 h-2 w-2 rounded-full bg-white/50 opacity-0 blur-sm transition-opacity duration-300 group-hover:opacity-100"></div>
                  <i v-if="isAnalyzing" class="fa fa-spinner fa-spin"></i>
                  <span>分析数据</span>
                </button>
              </div>
            </div>
          </div>
          
          <!-- 分析结果区域 -->
          <div v-if="analysisResult" class="space-y-6">
            <!-- 分析摘要卡片 -->
            <div class="rounded-2xl p-6 glass-effect bg-white/60 backdrop-blur-xl border border-white/70 shadow-[0_8px_32px_rgba(0,0,0,0.12)]">
              <h2 class="text-xl font-semibold text-gray-800 mb-3">{{ analysisResult.sheet_name }} - 分析摘要</h2>
              <p class="text-gray-600 mb-4">{{ analysisResult.summary }}</p>
            </div>
            
            <!-- 关键洞察卡片 -->
            <div class="rounded-2xl p-6 glass-effect bg-white/60 backdrop-blur-xl border border-white/70 shadow-[0_8px_32px_rgba(0,0,0,0.12)]">
              <h2 class="text-xl font-semibold text-gray-800 mb-3">关键洞察</h2>
              <ul class="space-y-2">
                <li v-for="(insight, index) in analysisResult.insights" :key="index" class="flex items-start space-x-3 p-3 rounded-xl bg-white/40 backdrop-blur-sm border border-white/50">
                  <i class="fa fa-lightbulb-o text-yellow-500 mt-1"></i>
                  <span class="text-gray-700">{{ insight }}</span>
                </li>
              </ul>
            </div>
            
            <!-- 趋势分析卡片 -->
            <div class="rounded-2xl p-6 glass-effect bg-white/60 backdrop-blur-xl border border-white/70 shadow-[0_8px_32px_rgba(0,0,0,0.12)]">
              <h2 class="text-xl font-semibold text-gray-800 mb-3">趋势分析</h2>
              <ul v-if="analysisResult.trends && analysisResult.trends.length > 0" class="space-y-2">
                <li v-for="(trend, index) in analysisResult.trends" :key="index" class="flex items-start space-x-3 p-3 rounded-xl bg-white/40 backdrop-blur-sm border border-white/50">
                  <i class="fa fa-line-chart text-blue-500 mt-1"></i>
                  <span class="text-gray-700">{{ trend }}</span>
                </li>
              </ul>
              <p v-else class="text-gray-500">暂无明显趋势数据</p>
            </div>
            
            <!-- 数据异常卡片 -->
            <div class="rounded-2xl p-6 glass-effect bg-white/60 backdrop-blur-xl border border-white/70 shadow-[0_8px_32px_rgba(0,0,0,0.12)]">
              <h2 class="text-xl font-semibold text-gray-800 mb-3">数据异常</h2>
              <ul v-if="analysisResult.anomalies && analysisResult.anomalies.length > 0" class="space-y-2">
                <li v-for="(anomaly, index) in analysisResult.anomalies" :key="index" class="flex items-start space-x-3 p-3 rounded-xl bg-white/40 backdrop-blur-sm border border-white/50">
                  <i class="fa fa-exclamation-triangle text-orange-500 mt-1"></i>
                  <span class="text-gray-700">{{ anomaly }}</span>
                </li>
              </ul>
              <p v-else class="text-gray-500">未发现数据异常</p>
            </div>
            
            <!-- 可视化图表卡片 -->
            <div class="rounded-2xl p-6 glass-effect bg-white/60 backdrop-blur-xl border border-white/70 shadow-[0_8px_32px_rgba(0,0,0,0.12)]">
              <h2 class="text-xl font-semibold text-gray-800 mb-3">数据可视化</h2>
              <div id="visualization-area" class="space-y-6">
                <!-- 图表容器会在这里动态创建 -->
                <div v-if="!analysisResult || (!analysisResult.visualization_data || analysisResult.visualization_data.length === 0)" class="p-8 text-center text-gray-500">
                  <i class="fa fa-bar-chart text-4xl mb-3"></i>
                  <p>上传Excel文件进行分析，查看数据可视化结果</p>
                </div>
                <div v-else>
                  <!-- 图表将通过JavaScript动态创建 -->
                </div>
              </div>
            </div>
          </div>
        </div>
        
        <!-- 已保存文件内容 -->
        <div v-if="activeTab === 'saved'" class="space-y-6">
          <!-- iOS风格已保存文件卡片 -->
            <div class="rounded-2xl p-6 glass-effect bg-white/60 backdrop-blur-xl border border-white/70 shadow-[0_8px_32px_rgba(0,0,0,0.12)]">
            <h1 class="text-2xl font-bold text-gray-800 mb-4">已保存文件</h1>
            
            <div class="overflow-x-auto">
              <table class="min-w-full divide-y divide-gray-100">
                <thead>
                  <tr class="bg-white/40 backdrop-blur-sm">
                    <th scope="col" class="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">文件名</th>
                    <th scope="col" class="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">类型</th>
                    <th scope="col" class="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">创建时间</th>
                    <th scope="col" class="px-6 py-3 text-right text-xs font-medium text-gray-500 uppercase tracking-wider">操作</th>
                  </tr>
                </thead>
                <tbody class="bg-white/20 backdrop-blur-sm divide-y divide-gray-50">
                  <tr v-for="file in savedFiles" :key="file.id" class="hover:bg-white/50 transition-colors duration-150">
                    <td class="px-6 py-4 whitespace-nowrap">
                      <div class="flex items-center">
                        <i class="fa fa-file-excel-o text-green-500 mr-3"></i>
                        <div class="text-sm font-medium text-gray-900">{{ file.name }}</div>
                      </div>
                    </td>
                    <td class="px-6 py-4 whitespace-nowrap">
                      <span class="px-2.5 inline-flex text-xs leading-5 font-semibold rounded-full bg-blue-50 text-blue-700">
                        Excel
                      </span>
                    </td>
                    <td class="px-6 py-4 whitespace-nowrap text-sm text-gray-500">
                      {{ formatDate(file.createdAt) }}
                    </td>
                    <td class="px-6 py-4 whitespace-nowrap text-right text-sm font-medium">
                      <button class="group relative glass-btn text-blue-500 hover:text-blue-600 mr-4 px-2 py-1 rounded-xl hover:bg-white/40 transition-all duration-300">
                  下载
                  <div class="absolute top-1 left-1 h-1 w-1 rounded-full bg-white/50 opacity-0 blur-sm transition-opacity duration-300 group-hover:opacity-100"></div>
                </button>
                <button class="group relative glass-btn text-red-500 hover:text-red-600 px-2 py-1 rounded-xl hover:bg-white/40 transition-all duration-300">
                  删除
                  <div class="absolute top-1 left-1 h-1 w-1 rounded-full bg-white/50 opacity-0 blur-sm transition-opacity duration-300 group-hover:opacity-100"></div>
                </button>
                    </td>
                  </tr>
                </tbody>
              </table>
            </div>
            
            <div v-if="savedFiles.length === 0" class="text-center py-12">
              <i class="fa fa-folder-open text-gray-300 text-5xl mb-3"></i>
              <p class="text-gray-500">暂无保存的文件</p>
            </div>
          </div>
        </div>
      </div>
    </main>
  </div>
</template>

<script>
import { ref, onMounted } from 'vue';
import axios from 'axios';

export default {
  name: 'SalesDashboard',
  setup() {
    // 状态管理
    const activeTab = ref('overview');
    const excelDescription = ref('');
    const sheetName = ref('数据表格');
    const fileName = ref('');
    const editInstructions = ref('');
    const uploadedFile = ref(null);
    const generatedFile = ref(null);
    const editedFile = ref(null);
    const isGenerating = ref(false);
    const isEditing = ref(false);
    const fileInput = ref(null);
    const savedFiles = ref([]);
    // 数据分析相关状态
    const analysisFile = ref(null);
    const analysisFileInput = ref(null);
    const analysisResult = ref(null);
    const isAnalyzing = ref(false);
    
    // 数据状态
    const generatedFileData = ref([]);
    const isLoadingFileData = ref(false);
    
    // 初始化已保存文件列表
    const initSavedFiles = () => {
      // 在实际应用中，这里应该从服务器获取已保存的文件列表
      // 这里使用模拟数据
      savedFiles.value = [
        { id: 1, name: '销售报表.xlsx', type: 'excel', createdAt: new Date(Date.now() - 86400000) },
        { id: 2, name: '库存管理.xlsx', type: 'excel', createdAt: new Date(Date.now() - 172800000) }
      ];
    };
    
    // 格式化日期
    const formatDate = (date) => {
      if (typeof date === 'string') {
        date = new Date(date);
      }
      return new Intl.DateTimeFormat('zh-CN', {
        year: 'numeric',
        month: '2-digit',
        day: '2-digit',
        hour: '2-digit',
        minute: '2-digit'
      }).format(date);
    };
    
    // 生成Excel
    const generateExcel = async () => {
      if (!excelDescription.value.trim()) {
        alert('请输入表格描述');
        return;
      }
      
      isGenerating.value = true;
      
      try {
        // 真正调用后端API生成Excel文件
        console.log('开始调用后端API生成Excel文件');
        
        const response = await axios.post('http://localhost:8000/generate_excel', {
          description: excelDescription.value,
          file_name: fileName.value || '生成的表格.xlsx'
        });
        
        generatedFile.value = {
          fileName: response.data.file_name,
          downloadUrl: response.data.download_url
        };
        
        // 将生成的文件添加到已保存文件列表
        savedFiles.value.unshift({
          id: Date.now(),
          name: response.data.file_name,
          type: 'excel',
          createdAt: new Date()
        });
        
        // 显示保存成功提示
        setTimeout(() => {
          alert('Excel文件已成功生成并保存到E:/ExcelGenius/temp目录');
        }, 100);
        
      } catch (error) {
        console.error('生成Excel失败:', error);
        
        // 使用setTimeout确保DOM更新后再设置弹窗状态
        setTimeout(() => {
          alert('生成Excel失败，请稍后重试');
        }, 100);
      } finally {
        isGenerating.value = false;
      }
    };
    

    

    
    // 处理文件上传
    const handleFileUpload = (event) => {
      const file = event.target.files[0];
      if (file) {
        uploadedFile.value = file;
      }
    };
    
    // 移除上传的文件
    const removeFile = () => {
      uploadedFile.value = null;
      if (fileInput.value) {
        fileInput.value.value = '';
      }
    };
    
    // 移除分析的文件
    const removeAnalysisFile = () => {
      analysisFile.value = null;
      if (analysisFileInput.value) {
        analysisFileInput.value.value = '';
      }
      analysisResult.value = null;
    };
    
    // 编辑Excel
    const editExcel = async () => {
      if (!editInstructions.value.trim()) {
        alert('请输入编辑指令');
        return;
      }
      
      isEditing.value = true;
      
      try {
        const formData = new FormData();
        formData.append('file', uploadedFile.value);
        formData.append('instructions', editInstructions.value);
        
        const response = await axios.post('http://localhost:8000/edit_excel', formData, {
          headers: {
            'Content-Type': 'multipart/form-data'
          }
        });
        
        editedFile.value = {
          fileName: response.data.file_name,
          downloadUrl: response.data.download_url
        };
        
        // 将编辑的文件添加到已保存文件列表
        savedFiles.value.unshift({
          id: Date.now(),
          name: response.data.file_name,
          type: 'excel',
          createdAt: new Date()
        });
        
      } catch (error) {
        console.error('编辑Excel失败:', error);
        alert('编辑Excel失败，请稍后重试');
      } finally {
        isEditing.value = false;
      }
    };
    
    // 处理分析文件上传
    const handleAnalysisFileUpload = (event) => {
      const file = event.target.files[0];
      if (file) {
        analysisFile.value = file;
        analysisResult.value = null;
      }
    };
    
    // 分析Excel数据
    const analyzeExcel = async () => {
      if (!analysisFile.value) {
        alert('请先上传Excel文件');
        return;
      }
      
      isAnalyzing.value = true;
      
      try {
        const formData = new FormData();
        formData.append('file', analysisFile.value);
        
        const response = await axios.post('http://localhost:8000/analyze_excel', formData, {
          headers: {
            'Content-Type': 'multipart/form-data'
          }
        });
        
        analysisResult.value = response.data;
        
        // 添加日志输出，查看返回的分析结果
        console.log('分析结果:', analysisResult.value);
        console.log('可视化数据:', analysisResult.value?.visualization_data || '无');
        
        // 如果没有可视化数据，尝试使用模拟数据进行调试
        if (!analysisResult.value.visualization_data || analysisResult.value.visualization_data.length === 0) {
          console.log('没有可视化数据，使用模拟数据进行调试');
          analysisResult.value.visualization_data = [
            {
              "type": "bar_chart",
              "title": "测试柱状图",
              "data": [
                {"category": "A", "value": 10},
                {"category": "B", "value": 25},
                {"category": "C", "value": 15},
                {"category": "D", "value": 30}
              ],
              "x_axis": "category",
              "y_axis": "value"
            },
            {
              "type": "pie_chart",
              "title": "测试饼图",
              "data": [
                {"name": "A", "value": 30},
                {"name": "B", "value": 20},
                {"name": "C", "value": 50}
              ]
            },
            {
              "type": "line_chart",
              "title": "测试折线图",
              "data": [
                {"month": "1月", "value": 10},
                {"month": "2月", "value": 15},
                {"month": "3月", "value": 12},
                {"month": "4月", "value": 18},
                {"month": "5月", "value": 25}
              ],
              "x_axis": "month",
              "y_axis": "value"
            }
          ];
        }
        
        // 延迟调用渲染图表函数，确保DOM已经更新
        setTimeout(() => {
          renderCharts();
        }, 100);
        
      } catch (error) {
        console.error('分析Excel失败:', error);
        // 添加详细的错误信息日志
        console.error('错误详情:', error.response?.data || error.message);
        alert('分析Excel失败，请稍后重试');
      } finally {
        isAnalyzing.value = false;
      }
    };
    
    // 渲染图表
    const renderCharts = () => {
      if (!analysisResult.value || !analysisResult.value.visualization_data) {
        return;
      }
      
      // 检查是否已经在HTML中加载了Chart.js库
      if (typeof Chart === 'undefined') {
        // 如果没有加载，动态创建script标签加载Chart.js
        const script = document.createElement('script');
        script.src = 'https://cdn.jsdelivr.net/npm/chart.js';
        script.onload = () => {
          // Chart.js加载完成后再渲染图表
          createCharts();
        };
        document.head.appendChild(script);
      } else {
        // 如果已经加载了Chart.js，直接渲染图表
        createCharts();
      }
    };
    
    // 创建图表
    const createCharts = () => {
      // 确保analysisResult和visualization_data存在
      if (!analysisResult.value) {
        console.log('没有分析结果数据，使用模拟数据');
        analysisResult.value = {
          visualization_data: [
            {
              "type": "bar_chart",
              "title": "测试柱状图",
              "data": [
                {"category": "A", "value": 10},
                {"category": "B", "value": 25},
                {"category": "C", "value": 15},
                {"category": "D", "value": 30}
              ],
              "x_axis": "category",
              "y_axis": "value"
            },
            {
              "type": "pie_chart",
              "title": "测试饼图",
              "data": [
                {"name": "A", "value": 30},
                {"name": "B", "value": 20},
                {"name": "C", "value": 50}
              ]
            },
            {
              "type": "line_chart",
              "title": "测试折线图",
              "data": [
                {"month": "1月", "value": 10},
                {"month": "2月", "value": 15},
                {"month": "3月", "value": 12},
                {"month": "4月", "value": 18},
                {"month": "5月", "value": 25}
              ],
              "x_axis": "month",
              "y_axis": "value"
            }
          ]
        };
      }
      
      if (!analysisResult.value.visualization_data || analysisResult.value.visualization_data.length === 0) {
        console.log('没有可视化数据，使用默认模拟数据');
        analysisResult.value.visualization_data = [
          {
            "type": "bar_chart",
            "title": "数据分布概览",
            "data": [
              {"field": "表头数量", "value": 10},
              {"field": "数据行数", "value": 100}
            ],
            "x_axis": "field",
            "y_axis": "value"
          }
        ];
      }
      
      // 清空现有的图表容器
      const visualizationArea = document.getElementById('visualization-area');
      if (visualizationArea) {
        visualizationArea.innerHTML = '';
      }
      
      analysisResult.value.visualization_data.forEach((chartData, index) => {
        console.log(`处理图表 ${index}:`, chartData);
        
        const chartId = `chart-${index}`;
        const canvasId = `canvas-${index}`;
        
        // 创建新的图表容器
        const newContainer = document.createElement('div');
        newContainer.id = chartId;
        newContainer.className = 'rounded-xl bg-white/40 backdrop-blur-sm p-4 border border-white/50';
        newContainer.style.width = '100%';
        newContainer.style.height = '400px';
        
        // 创建标题
        const title = document.createElement('h3');
        title.textContent = chartData.title || `图表 ${index + 1}`;
        title.className = 'text-lg font-medium text-gray-800 mb-4';
        newContainer.appendChild(title);
        
        // 创建canvas元素
        const canvas = document.createElement('canvas');
        canvas.id = canvasId;
        canvas.style.width = '100%';
        canvas.style.height = '300px';
        newContainer.appendChild(canvas);
        
        // 将容器添加到可视化区域
        if (visualizationArea) {
          visualizationArea.appendChild(newContainer);
        }
        
        // 获取canvas上下文
        const ctx = document.getElementById(canvasId);
        if (!ctx) {
          console.error(`无法找到canvas元素: ${canvasId}`);
          return;
        }
        
        // 根据图表类型创建不同的图表
        if (chartData.type === 'line_chart') {
          new Chart(ctx, {
            type: 'line',
            data: {
              labels: chartData.data.map(item => item[chartData.x_axis]),
              datasets: [{
                label: chartData.title,
                data: chartData.data.map(item => item[chartData.y_axis]),
                borderColor: 'rgba(59, 130, 246, 1)',
                backgroundColor: 'rgba(59, 130, 246, 0.2)',
                borderWidth: 2,
                tension: 0.3,
                fill: true
              }]
            },
            options: {
              responsive: true,
              maintainAspectRatio: false,
              plugins: {
                legend: {
                  display: false
                }
              },
              scales: {
                y: {
                  beginAtZero: true,
                  grid: {
                    color: 'rgba(0, 0, 0, 0.05)'
                  }
                },
                x: {
                  grid: {
                    display: false
                  }
                }
              }
            }
          });
        } else if (chartData.type === 'pie_chart') {
          new Chart(ctx, {
            type: 'pie',
            data: {
              labels: chartData.data.map(item => item.name),
              datasets: [{
                data: chartData.data.map(item => item.value),
                backgroundColor: [
                  'rgba(59, 130, 246, 0.8)',
                  'rgba(16, 185, 129, 0.8)',
                  'rgba(236, 72, 153, 0.8)',
                  'rgba(234, 179, 8, 0.8)',
                  'rgba(124, 58, 237, 0.8)',
                  'rgba(239, 68, 68, 0.8)'
                ],
                borderWidth: 1,
                borderColor: 'rgba(255, 255, 255, 0.8)'
              }]
            },
            options: {
              responsive: true,
              maintainAspectRatio: false,
              plugins: {
                legend: {
                  position: 'right'
                }
              }
            }
          });
        } else if (chartData.type === 'bar_chart') {
          new Chart(ctx, {
            type: 'bar',
            data: {
              labels: chartData.data.map(item => item[chartData.x_axis]),
              datasets: [{
                label: chartData.title,
                data: chartData.data.map(item => item[chartData.y_axis]),
                backgroundColor: [
                  'rgba(59, 130, 246, 0.8)',
                  'rgba(16, 185, 129, 0.8)',
                  'rgba(236, 72, 153, 0.8)',
                  'rgba(234, 179, 8, 0.8)',
                  'rgba(124, 58, 237, 0.8)',
                  'rgba(239, 68, 68, 0.8)'
                ],
                borderColor: [
                  'rgba(59, 130, 246, 1)',
                  'rgba(16, 185, 129, 1)',
                  'rgba(236, 72, 153, 1)',
                  'rgba(234, 179, 8, 1)',
                  'rgba(124, 58, 237, 1)',
                  'rgba(239, 68, 68, 1)'
                ],
                borderWidth: 1
              }]
            },
            options: {
              responsive: true,
              maintainAspectRatio: false,
              plugins: {
                legend: {
                  display: false
                }
              },
              scales: {
                y: {
                  beginAtZero: true,
                  grid: {
                    color: 'rgba(0, 0, 0, 0.05)'
                  }
                },
                x: {
                  grid: {
                    display: false
                  }
                }
              }
            }
          });
        }
      });
    };
    
    // 下载文件
    const downloadFile = (fileName) => {
      // 使用request.js中的download方法来确保可靠下载
      axios.get(`http://localhost:8000/download/${fileName}`, {
        responseType: 'blob'
      }).then(response => {
        const blob = new Blob([response.data]);
        const url = window.URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.download = fileName;
        document.body.appendChild(link);
        link.click();
        
        // 清理
        document.body.removeChild(link);
        window.URL.revokeObjectURL(url);
        
        // 显示保存成功的提示
        alert('文件已成功保存到临时文件夹');
      }).catch(error => {
        console.error('下载文件失败:', error);
        alert('下载文件失败，请稍后重试');
      });
    };
    
    // 删除文件
    const deleteFile = (fileId) => {
      if (confirm('确定要删除这个文件吗？')) {
        savedFiles.value = savedFiles.value.filter(file => file.id !== fileId);
      }
    };
    
    // 使用模板
    const useTemplate = (templateType) => {
      activeTab.value = 'create';
      
      // 根据模板类型填充描述
      switch (templateType) {
        case 'sales':
          excelDescription.value = '创建一个销售业绩表，包含以下字段：月份、销售额、目标、达成率。数据包含1-12月的模拟数据，达成率为销售额除以目标的百分比。';
          sheetName.value = '销售业绩';
          fileName.value = '销售报表';
          break;
        case 'inventory':
          excelDescription.value = '创建一个库存管理表，包含以下字段：产品ID、产品名称、类别、单价、库存数量、供应商。包含10种不同产品的模拟数据。';
          sheetName.value = '库存数据';
          fileName.value = '库存管理';
          break;
        case 'finance':
          excelDescription.value = '创建一个财务报表，包含以下字段：日期、摘要、收入、支出、余额。日期从本月1号开始，包含日常收支的模拟数据。';
          sheetName.value = '财务数据';
          fileName.value = '财务报表';
          break;
      }
    };
    
    // 重置创建表单
    const resetCreateForm = () => {
      excelDescription.value = '';
      sheetName.value = '数据表格';
      fileName.value = '';
      generatedFile.value = null;
    };
    
    // 重置编辑表单
    const resetEditForm = () => {
      editInstructions.value = '';
      removeFile();
      editedFile.value = null;
    };
    
    // 生命周期钩子
    onMounted(() => {
      initSavedFiles();
    });
    
    return {
      activeTab,
      excelDescription,
      sheetName,
      fileName,
      editInstructions,
      uploadedFile,
      generatedFile,
      editedFile,
      isGenerating,
      isEditing,
      fileInput,
      savedFiles,
      generateExcel,
      handleFileUpload,
      removeFile,
      editExcel,
      downloadFile,
      deleteFile,
      useTemplate,
      resetCreateForm,
      resetEditForm,
      formatDate,
      // 数据分析相关方法
      analysisFile,
      analysisFileInput,
      analysisResult,
      isAnalyzing,
      handleAnalysisFileUpload,
      removeAnalysisFile,
      analyzeExcel

    };
  }
};
</script>

<style scoped>
/* 自定义滚动条样式 - iOS风格 */
::-webkit-scrollbar {
  width: 6px;
  height: 6px;
}

::-webkit-scrollbar-track {
  background: transparent;
}

::-webkit-scrollbar-thumb {
  background: rgba(156, 163, 175, 0.5);
  border-radius: 3px;
}

::-webkit-scrollbar-thumb:hover {
  background: rgba(156, 163, 175, 0.8);
}

/* 过渡动画 */
.fade-enter-active, .fade-leave-active {
  transition: opacity 0.3s ease;
}

.fade-enter-from, .fade-leave-to {
  opacity: 0;
}

/* 平滑滚动 */
html {
  scroll-behavior: smooth;
}

/* 标题文字效果 */
h1, h2, h3 {
  font-weight: bold;
}

/* 移除点击时的焦点轮廓 */
button:focus {
  outline: none;
}

/* 提升按钮交互体验 */
button {
  user-select: none;
  -webkit-user-select: none;
}

/* 输入框聚焦效果增强 */
input:focus, textarea:focus {
  box-shadow: 0 0 0 3px rgba(59, 130, 246, 0.15);
}

/* 优化移动端触摸效果 */
@media (hover: none) {
  a:hover, button:hover {
    background-color: initial !important;
  }
}

/* 玻璃态效果增强 */
.glass-effect {
  background: rgba(255, 255, 255, 0.6);
  backdrop-filter: blur(12px);
  -webkit-backdrop-filter: blur(12px);
  border: 1px solid rgba(255, 255, 255, 0.7);
}

/* 深度阴影效果 */
.ios-shadow {
  box-shadow: 0 8px 32px rgba(0, 0, 0, 0.12);
}
</style>