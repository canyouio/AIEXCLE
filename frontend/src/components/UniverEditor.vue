<template>
  <div ref="univerContainerRef" class="univer-container"></div>
</template>

<script>
import { ref, onMounted, onUnmounted, watch } from 'vue';
import { createUniver, UniverInstanceType, LocaleType, mergeLocales } from '@univerjs/presets';
import { UniverSheetsCorePreset } from '@univerjs/preset-sheets-core';
import { FUniver } from '@univerjs/core/facade';
import univerZhCN from '@univerjs/preset-sheets-core/locales/zh-CN';

export default {
  name: 'UniverEditor',
  props: {
    initialData: { type: Array, default: () => [] }
  },
  setup(props, { expose }) {
    const univerContainerRef = ref(null);
    let univerAPI = null;
    let workbook = null;

    const convertToUniverFormat = (data) => {
      const cellData = {};
      let rowCount = 0;
      let columnCount = 0;
      if (data && data.length > 0) {
        rowCount = data.length;
        columnCount = data[0]?.length || 0;
        data.forEach((row, r) => {
          cellData[r] = {};
          row.forEach((cellValue, c) => {
            if (cellValue !== null && cellValue !== undefined) {
              cellData[r][c] = { v: cellValue };
            }
          });
        });
      }
      return {
        id: `workbook-${Date.now()}`, sheetOrder: ['sheet-01'],
        sheets: { 'sheet-01': { id: 'sheet-01', name: 'Sheet1', cellData, rowCount, columnCount: columnCount + 10 } },
      };
    };
    
    const initUniver = (data) => {
      if (univerAPI) univerAPI.dispose();
      
      const { univer } = createUniver({
        locale: LocaleType.ZH_CN,
        locales: { [LocaleType.ZH_CN]: mergeLocales(univerZhCN) },
        presets: [ UniverSheetsCorePreset({ container: univerContainerRef.value }) ],
      });
      
      univerAPI = FUniver.newAPI(univer);
      const workbookData = convertToUniverFormat(data);
      workbook = univerAPI.createWorkbook(workbookData);
    };

    onMounted(() => {
      setTimeout(() => {
        if (univerContainerRef.value) initUniver(props.initialData);
      }, 100);
    });

    onUnmounted(() => {
      if (univerAPI) {
        univerAPI.dispose();
        univerAPI = null;
      }
    });

    // 【最终修复】完全采纳你的正确方案，进行数据类型转换
    const getCurrentData = () => {
      if (!workbook) return [];
      const activeSheet = workbook.getActiveSheet();
      if (!activeSheet) return [];
      
      const lastRow = activeSheet.getLastRow();
      const lastCol = activeSheet.getLastColumn();

      if (lastRow < 0 || lastCol < 0) return [[]];

      const rawData = activeSheet.getRange(0, 0, lastRow + 1, lastCol + 1).getValues();
      
      // 将所有单元格数据转换为字符串或 null，以匹配后端 Pydantic 模型
      return rawData.map(row => 
        row.map(cell => {
          if (cell === null || cell === undefined) {
            return null;
          }
          // 将所有非空值（包括数字、布尔值等）转换为字符串
          return String(cell);
        })
      );
    };

    expose({
      getCurrentData
    });

    return { univerContainerRef };
  }
};
</script>

<style>
@import '@univerjs/preset-sheets-core/lib/index.css';
.univer-container {
  width: 100%;
  height: 100%;
  overflow: hidden;
}
</style>