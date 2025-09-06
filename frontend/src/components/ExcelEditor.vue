<template>
  <div class="bg-white rounded-lg shadow-soft p-6">
    <h2 class="text-xl font-semibold mb-4">Excel表格编辑器</h2>
    
    <!-- 工具栏 -->
    <div class="flex flex-wrap gap-2 mb-4">
      <el-button type="primary" @click="addRow">
        <i class="fa fa-plus mr-1"></i>添加行
      </el-button>
      <el-button type="primary" @click="addColumn">
        <i class="fa fa-plus-square-o mr-1"></i>添加列
      </el-button>
      <el-button type="danger" @click="deleteSelected" :disabled="!selectedCells.length">
        <i class="fa fa-trash mr-1"></i>删除选中
      </el-button>
      <el-button @click="clearTable">
        <i class="fa fa-refresh mr-1"></i>清空表格
      </el-button>
      <el-button @click="exportToExcel">
        <i class="fa fa-file-excel-o mr-1"></i>导出Excel
      </el-button>
      
      <!-- 格式设置 -->
      <div class="ml-auto flex items-center gap-2">
        <el-select v-model="fontSize" placeholder="字体大小" @change="formatSelectedCells">
          <el-option v-for="size in [10, 12, 14, 16, 18, 20]" :key="size" :label="size" :value="size"></el-option>
        </el-select>
        <el-button icon="el-icon-bold" @click="toggleBold"></el-button>
        <el-button icon="el-icon-italic" @click="toggleItalic"></el-button>
        <el-button icon="el-icon-underline" @click="toggleUnderline"></el-button>
        <el-color-picker v-model="textColor" @change="changeTextColor"></el-color-picker>
      </div>
    </div>
    
    <!-- 表格区域 -->
    <div class="overflow-auto max-h-[600px] border rounded">
      <table class="min-w-full border-collapse">
        <thead>
          <tr class="bg-gray-100 sticky top-0 z-10">
            <th class="border p-2 w-10 text-center bg-gray-200">
              <el-checkbox v-model="selectAll" @change="handleSelectAll"></el-checkbox>
            </th>
            <th 
              v-for="(col, index) in columns" 
              :key="index" 
              class="border p-2 text-center min-w-[100px]"
              :style="getColumnStyle(index)"
            >
              <div class="flex items-center justify-center gap-1">
                <span>{{ col }}</span>
                <i 
                  class="fa fa-ellipsis-v cursor-pointer text-gray-500 hover:text-primary" 
                  @click="showColumnMenu(index, $event)"
                ></i>
              </div>
            </th>
          </tr>
        </thead>
        <tbody>
          <tr 
            v-for="(row, rowIndex) in tableData" 
            :key="rowIndex"
            @click="handleRowClick(rowIndex, $event)"
          >
            <td class="border p-2 text-center bg-gray-100 sticky left-0">
              <el-checkbox 
                v-model="row.selected" 
                @change="handleRowSelect(rowIndex)"
              ></el-checkbox>
            </td>
            <td 
              v-for="(cell, colIndex) in row.cells" 
              :key="colIndex"
              class="border p-2 min-w-[100px]"
              :style="getCellStyle(rowIndex, colIndex)"
              @dblclick="startEditing(rowIndex, colIndex)"
            >
              <div v-if="editingCell.row === rowIndex && editingCell.col === colIndex">
                <el-input 
                  v-model="editingValue" 
                  ref="editInput"
                  @blur="finishEditing"
                  @keyup.enter="finishEditing"
                  @keyup.esc="cancelEditing"
                  size="small"
                ></el-input>
              </div>
              <div v-else class="whitespace-nowrap overflow-hidden text-ellipsis max-w-[200px]">
                {{ cell.value || '' }}
              </div>
            </td>
          </tr>
        </tbody>
      </table>
    </div>
    
    <!-- 统计信息 -->
    <div class="mt-4 text-gray-600">
      <span>行数: {{ tableData.length }}</span>
      <span class="mx-2">|</span>
      <span>列数: {{ columns.length }}</span>
      <span class="mx-2">|</span>
      <span>选中单元格: {{ selectedCells.length }}</span>
    </div>
    
    <!-- 右键菜单 -->
    <div 
      v-if="showMenu" 
      class="absolute bg-white shadow-lg rounded border z-20" 
      :style="{ left: menuPosition.x + 'px', top: menuPosition.y + 'px' }"
    >
      <div class="py-1">
        <div class="px-4 py-2 hover:bg-gray-100 cursor-pointer" @click="copySelected">复制</div>
        <div class="px-4 py-2 hover:bg-gray-100 cursor-pointer" @click="pasteSelected">粘贴</div>
        <div class="px-4 py-2 hover:bg-gray-100 cursor-pointer" @click="cutSelected">剪切</div>
        <div class="border-t my-1"></div>
        <div class="px-4 py-2 hover:bg-gray-100 cursor-pointer" @click="insertRowAbove">在上方插入行</div>
        <div class="px-4 py-2 hover:bg-gray-100 cursor-pointer" @click="insertRowBelow">在下方插入行</div>
        <div class="border-t my-1"></div>
        <div class="px-4 py-2 hover:bg-gray-100 cursor-pointer" @click="deleteSelectedRows">删除选中行</div>
      </div>
    </div>
  </div>
</template>

<script>
import { ref, reactive, onMounted, nextTick, watch } from 'vue';
import { ElMessage, ElMessageBox } from 'element-plus';
import * as XLSX from 'xlsx';

// 生成列名（A, B, C, ..., Z, AA, AB, ...）
const generateColumnName = (index) => {
  let name = '';
  let num = index;
  
  while (num >= 0) {
    name = String.fromCharCode(65 + (num % 26)) + name;
    num = Math.floor(num / 26) - 1;
  }
  
  return name;
};

export default {
  name: 'ExcelEditor',
  props: {
    initialData: {
      type: Array,
      default: () => []
    },
    fileName: {
      type: String,
      default: 'ExcelGenius_Export'
    }
  },
  emits: ['saveAndDownload'],
  setup(props, { emit }) {
    // 表格数据
    const tableData = ref([]);
    const columns = ref([]);
    const selectedCells = ref([]);
    const selectAll = ref(false);
    
    // 编辑状态
    const editingCell = reactive({ row: -1, col: -1 });
    const editingValue = ref('');
    
    // 格式设置
    const fontSize = ref(12);
    const isBold = ref(false);
    const isItalic = ref(false);
    const isUnderline = ref(false);
    const textColor = ref('#000000');
    const cellStyles = reactive({});
    const columnStyles = reactive({});
    
    // 右键菜单
    const showMenu = ref(false);
    const menuPosition = reactive({ x: 0, y: 0 });
    const contextMenuTarget = reactive({ row: -1, col: -1 });
    
    // 初始化表格数据
    const initTable = () => {
      if (props.initialData && props.initialData.length > 0) {
        // 使用传入的初始数据
        const data = props.initialData;
        
        // 生成列名
        const colCount = data[0]?.length || 10;
        columns.value = Array.from({ length: colCount }, (_, i) => generateColumnName(i));
        
        // 转换数据格式
        tableData.value = data.map((row, rowIndex) => ({
          selected: false,
          cells: row.map((cellValue, colIndex) => ({
            value: cellValue,
            style: rowIndex === 0 ? {
              fontWeight: 'bold',
              textAlign: 'center',
              backgroundColor: '#f0f0f0'
            } : {}
          }))
        }));
      } else {
        // 默认生成10列
        columns.value = Array.from({ length: 10 }, (_, i) => generateColumnName(i));
        
        // 默认生成10行mock数据
        tableData.value = Array.from({ length: 10 }, (_, rowIndex) => ({
          selected: false,
          cells: Array.from({ length: 10 }, (_, colIndex) => ({
            value: rowIndex === 0 ? `标题${colIndex + 1}` : `数据${rowIndex + 1}-${colIndex + 1}`,
            style: {}
          }))
        }));
        
        // 设置第一行为表头样式
        if (tableData.value.length > 0) {
          tableData.value[0].cells.forEach((cell, index) => {
            cell.style = {
              fontWeight: 'bold',
              textAlign: 'center',
              backgroundColor: '#f0f0f0'
            };
          });
        }
      }
    };
    
    // 获取单元格样式
    const getCellStyle = (rowIndex, colIndex) => {
      const cell = tableData.value[rowIndex]?.cells[colIndex];
      if (cell && cell.style) {
        return cell.style;
      }
      return {};
    };
    
    // 获取列样式
    const getColumnStyle = (colIndex) => {
      return columnStyles[colIndex] || {};
    };
    
    // 开始编辑单元格
    const startEditing = (rowIndex, colIndex) => {
      editingCell.row = rowIndex;
      editingCell.col = colIndex;
      editingValue.value = tableData.value[rowIndex].cells[colIndex].value || '';
      
      nextTick(() => {
        const input = document.querySelector('.el-input__inner:focus');
        if (input) {
          input.select();
        }
      });
    };
    
    // 完成编辑
    const finishEditing = () => {
      if (editingCell.row >= 0 && editingCell.col >= 0) {
        tableData.value[editingCell.row].cells[editingCell.col].value = editingValue.value;
      }
      
      // 重置编辑状态
      editingCell.row = -1;
      editingCell.col = -1;
      editingValue.value = '';
    };
    
    // 取消编辑
    const cancelEditing = () => {
      editingCell.row = -1;
      editingCell.col = -1;
      editingValue.value = '';
    };
    
    // 处理行点击
    const handleRowClick = (rowIndex, event) => {
      // 如果点击的是复选框或编辑框，不进行选择
      if (event.target.type === 'checkbox' || event.target.closest('.el-input')) {
        return;
      }
      
      // 多选逻辑
      if (event.ctrlKey) {
        // Ctrl+点击：切换选择状态
        tableData.value[rowIndex].selected = !tableData.value[rowIndex].selected;
      } else if (event.shiftKey) {
        // Shift+点击：选择连续行
        // 这里简化处理，实际应该记录上次点击的行
      } else {
        // 普通点击：单选
        tableData.value.forEach(row => {
          row.selected = false;
        });
        tableData.value[rowIndex].selected = true;
      }
      
      // 更新选中的单元格
      updateSelectedCells();
    };
    
    // 处理行选择
    const handleRowSelect = (rowIndex) => {
      tableData.value[rowIndex].selected = !tableData.value[rowIndex].selected;
      updateSelectedCells();
    };
    
    // 处理全选
    const handleSelectAll = () => {
      tableData.value.forEach(row => {
        row.selected = selectAll.value;
      });
      updateSelectedCells();
    };
    
    // 更新选中的单元格
    const updateSelectedCells = () => {
      const selected = [];
      tableData.value.forEach((row, rowIndex) => {
        if (row.selected) {
          row.cells.forEach((_, colIndex) => {
            selected.push({ row: rowIndex, col: colIndex });
          });
        }
      });
      selectedCells.value = selected;
      
      // 更新全选状态
      const allSelected = tableData.value.length > 0 && 
                         tableData.value.every(row => row.selected);
      selectAll.value = allSelected;
    };
    
    // 添加行
    const addRow = () => {
      const newRow = {
        selected: false,
        cells: Array.from({ length: columns.value.length }, (_, colIndex) => ({
          value: '',
          style: {}
        }))
      };
      
      tableData.value.push(newRow);
      ElMessage.success('已添加新行');
    };
    
    // 添加列
    const addColumn = () => {
      const newColName = generateColumnName(columns.value.length);
      columns.value.push(newColName);
      
      // 为每行添加新单元格
      tableData.value.forEach(row => {
        row.cells.push({
          value: '',
          style: {}
        });
      });
      
      ElMessage.success('已添加新列');
    };
    
    // 删除选中的单元格内容
    const deleteSelected = () => {
      if (selectedCells.value.length === 0) {
        ElMessage.warning('请先选择要删除的单元格');
        return;
      }
      
      selectedCells.value.forEach(cell => {
        tableData.value[cell.row].cells[cell.col].value = '';
      });
      
      ElMessage.success('已删除选中内容');
    };
    
    // 清空表格
    const clearTable = () => {
      ElMessageBox.confirm('确定要清空表格所有数据吗？', '确认', {
        confirmButtonText: '确定',
        cancelButtonText: '取消',
        type: 'warning'
      }).then(() => {
        tableData.value.forEach(row => {
          row.cells.forEach(cell => {
            cell.value = '';
          });
        });
        ElMessage.success('表格已清空');
      }).catch(() => {
        // 用户取消操作
      });
    };
    
    // 导出Excel
    const exportToExcel = () => {
      // 创建工作簿
      const wb = XLSX.utils.book_new();
      
      // 准备数据
      const data = [];
      
      // 添加表头
      data.push(columns.value);
      
      // 添加数据行
      tableData.value.forEach(row => {
        const rowData = row.cells.map(cell => cell.value || '');
        data.push(rowData);
      });
      
      // 创建工作表
      const ws = XLSX.utils.aoa_to_sheet(data);
      
      // 添加工作表到工作簿
      XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
      
      // 导出文件
      const exportFileName = props.fileName.endsWith('.xlsx') ? props.fileName : `${props.fileName}.xlsx`;
      XLSX.writeFile(wb, exportFileName);
      
      ElMessage.success('Excel文件已导出');
      
      // 触发保存并下载事件
      emit('saveAndDownload');
    };
    
    // 格式化选中的单元格
    const formatSelectedCells = () => {
      if (selectedCells.value.length === 0) return;
      
      selectedCells.value.forEach(cell => {
        if (!tableData.value[cell.row].cells[cell.col].style) {
          tableData.value[cell.row].cells[cell.col].style = {};
        }
        
        if (fontSize.value) {
          tableData.value[cell.row].cells[cell.col].style.fontSize = `${fontSize.value}px`;
        }
        
        if (isBold.value) {
          tableData.value[cell.row].cells[cell.col].style.fontWeight = 'bold';
        } else {
          delete tableData.value[cell.row].cells[cell.col].style.fontWeight;
        }
        
        if (isItalic.value) {
          tableData.value[cell.row].cells[cell.col].style.fontStyle = 'italic';
        } else {
          delete tableData.value[cell.row].cells[cell.col].style.fontStyle;
        }
        
        if (isUnderline.value) {
          tableData.value[cell.row].cells[cell.col].style.textDecoration = 'underline';
        } else {
          delete tableData.value[cell.row].cells[cell.col].style.textDecoration;
        }
        
        if (textColor.value) {
          tableData.value[cell.row].cells[cell.col].style.color = textColor.value;
        }
      });
    };
    
    // 切换粗体
    const toggleBold = () => {
      isBold.value = !isBold.value;
      formatSelectedCells();
    };
    
    // 切换斜体
    const toggleItalic = () => {
      isItalic.value = !isItalic.value;
      formatSelectedCells();
    };
    
    // 切换下划线
    const toggleUnderline = () => {
      isUnderline.value = !isUnderline.value;
      formatSelectedCells();
    };
    
    // 改变文字颜色
    const changeTextColor = () => {
      formatSelectedCells();
    };
    
    // 显示列菜单
    const showColumnMenu = (colIndex, event) => {
      event.stopPropagation();
      // 实际项目中应该实现列菜单功能
      ElMessage.info('列操作菜单将在这里显示');
    };
    
    // 复制选中内容
    const copySelected = () => {
      // 简化实现，实际应该复制单元格内容到剪贴板
      ElMessage.success('内容已复制到剪贴板');
    };
    
    // 粘贴内容
    const pasteSelected = () => {
      // 简化实现，实际应该从剪贴板粘贴内容
      ElMessage.success('已粘贴内容');
    };
    
    // 剪切内容
    const cutSelected = () => {
      // 简化实现，实际应该剪切单元格内容到剪贴板
      ElMessage.success('内容已剪切到剪贴板');
    };
    
    // 在上方插入行
    const insertRowAbove = () => {
      const rowIndex = contextMenuTarget.row;
      if (rowIndex >= 0) {
        const newRow = {
          selected: false,
          cells: Array.from({ length: columns.value.length }, (_, colIndex) => ({
            value: '',
            style: {}
          }))
        };
        
        tableData.value.splice(rowIndex, 0, newRow);
        ElMessage.success('已在上方插入新行');
      }
      showMenu.value = false;
    };
    
    // 在下方插入行
    const insertRowBelow = () => {
      const rowIndex = contextMenuTarget.row;
      if (rowIndex >= 0) {
        const newRow = {
          selected: false,
          cells: Array.from({ length: columns.value.length }, (_, colIndex) => ({
            value: '',
            style: {}
          }))
        };
        
        tableData.value.splice(rowIndex + 1, 0, newRow);
        ElMessage.success('已在下方插入新行');
      }
      showMenu.value = false;
    };
    
    // 删除选中行
    const deleteSelectedRows = () => {
      const selectedRows = tableData.value.filter(row => row.selected);
      if (selectedRows.length === 0) {
        ElMessage.warning('请先选择要删除的行');
        return;
      }
      
      ElMessageBox.confirm(`确定要删除选中的${selectedRows.length}行吗？`, '确认', {
        confirmButtonText: '确定',
        cancelButtonText: '取消',
        type: 'warning'
      }).then(() => {
        tableData.value = tableData.value.filter(row => !row.selected);
        selectAll.value = false;
        selectedCells.value = [];
        ElMessage.success(`已删除${selectedRows.length}行`);
      }).catch(() => {
        // 用户取消操作
      });
      
      showMenu.value = false;
    };
    
    // 处理右键菜单
    const handleContextMenu = (event) => {
      event.preventDefault();
      showMenu.value = true;
      menuPosition.x = event.clientX;
      menuPosition.y = event.clientY;
      
      // 查找点击的单元格
      const cell = event.target.closest('td');
      if (cell) {
        const row = cell.parentElement.rowIndex;
        const col = cell.cellIndex - 1; // 减去复选框列
        contextMenuTarget.row = row - 1; // 减去表头行
        contextMenuTarget.col = col;
      }
    };
    
    // 关闭右键菜单
    const closeContextMenu = () => {
      showMenu.value = false;
    };
    
    // 监听初始数据变化
    watch(() => props.initialData, (newData) => {
      if (newData && newData.length > 0) {
        initTable();
      }
    }, { deep: true });
    
    // 生命周期钩子
    onMounted(() => {
      // 初始化表格数据
      initTable();
      
      // 添加右键菜单事件监听
      document.addEventListener('contextmenu', handleContextMenu);
      document.addEventListener('click', closeContextMenu);
    });
    
    return {
      tableData,
      columns,
      selectedCells,
      selectAll,
      editingCell,
      editingValue,
      fontSize,
      isBold,
      isItalic,
      isUnderline,
      textColor,
      showMenu,
      menuPosition,
      
      getCellStyle,
      getColumnStyle,
      startEditing,
      finishEditing,
      cancelEditing,
      handleRowClick,
      handleRowSelect,
      handleSelectAll,
      addRow,
      addColumn,
      deleteSelected,
      clearTable,
      exportToExcel,
      formatSelectedCells,
      toggleBold,
      toggleItalic,
      toggleUnderline,
      changeTextColor,
      showColumnMenu,
      copySelected,
      pasteSelected,
      cutSelected,
      insertRowAbove,
      insertRowBelow,
      deleteSelectedRows
    };
  }
};
</script>

<style scoped>
/* 自定义样式 */
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

/* 单元格编辑时的样式 */
.el-input__inner:focus {
  outline: none;
  box-shadow: 0 0 0 2px rgba(64, 158, 255, 0.2);
}
</style>