<script setup lang="ts">
/** File > Import: an Excel or CSV file. */
import { ref } from 'vue';
import { useUI } from '../context';
import Icon from '../Icon.vue';

const emit = defineEmits<{ close: [] }>();
const ui = useUI();
const kind = ref<'excel' | 'csv'>('excel');
const run = () => {
  emit('close');
  ui.value.files.importFile(kind.value);
};
</script>

<template>
  <h1>Import</h1>
  <div class="sh-bs-columns">
    <div class="sh-bs-list" role="menu" aria-label="Import">
      <button type="button" role="menuitemradio" :aria-checked="kind === 'excel'" :class="kind === 'excel' ? 'active' : ''" data-testid="import-excel" @click="kind = 'excel'">Excel File</button>
      <button type="button" role="menuitemradio" :aria-checked="kind === 'csv'" :class="kind === 'csv' ? 'active' : ''" data-testid="import-csv" @click="kind = 'csv'">CSV File</button>
    </div>
    <div class="sh-bs-detail">
      <h2>{{ kind === 'excel' ? 'Excel File' : 'CSV File' }}</h2>
      <p class="sh-bs-lead">{{ kind === 'excel' ? 'Opens an Excel workbook (.xlsx, .xlsm, .xltx, .xltm, Excel 97-2003 .xls, binary .xlsb) with its formulas, styles, tables, charts, pictures, comments and conditional formats.' : 'Opens a comma, semicolon or tab separated text file. The separator and the encoding are detected.' }}</p>
      <button type="button" class="sh-bs-action" data-testid="import-file" @click="run">
        <Icon name="upload_file" />
        <span>{{ kind === 'excel' ? 'Import Excel File' : 'Import CSV File' }}</span>
      </button>
    </div>
  </div>
</template>
