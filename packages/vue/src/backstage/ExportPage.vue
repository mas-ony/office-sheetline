<script setup lang="ts">
/** File > Export: a copy as an Excel workbook or template, CSV, or PDF (through Print). */
import { type ExportFormat, fileBaseName } from '@sheetline/core/views';
import { ref } from 'vue';
import { useUI } from '../context';
import Icon from '../Icon.vue';

const emit = defineEmits<{ close: [] }>();

const EXCEL_TYPES: { value: ExportFormat; label: string }[] = [
  { value: 'xlsx', label: 'Excel Workbook (*.xlsx)' },
  { value: 'xltx', label: 'Excel Template (*.xltx)' },
  { value: 'xlsm', label: 'Excel Macro-Enabled Workbook (*.xlsm)' },
  { value: 'xltm', label: 'Excel Macro-Enabled Template (*.xltm)' },
];
const KINDS = ['excel', 'csv', 'pdf'] as const;
const LEADS = {
  excel: 'Saves a copy of the workbook. Parts of an opened file that the editor does not change are kept as they were.',
  csv: 'Saves the values of the active sheet as comma separated text (UTF-8).',
  pdf: 'Creates a PDF from the print settings and the page setup of the sheet (File > Print shows the pages first).',
};

const ui = useUI();
const kind = ref<(typeof KINDS)[number]>('excel');
const type = ref<ExportFormat>('xlsx');
const name = ref(fileBaseName(ui.value.doc.fileName));
const kindLabel = (k: (typeof KINDS)[number]) => (k === 'excel' ? 'Excel File' : k === 'csv' ? 'CSV File' : 'PDF File');
const run = () => {
  const format: ExportFormat = kind.value === 'excel' ? type.value : kind.value;
  emit('close');
  if (format === 'pdf') ui.value.openBackstage('print');
  else ui.value.files.exportAs(format, name.value.trim() || 'Book1');
};
</script>

<template>
  <h1>Export</h1>
  <div class="sh-bs-columns">
    <div class="sh-bs-list" role="menu" aria-label="Export">
      <button
        v-for="k in KINDS"
        :key="k"
        type="button"
        role="menuitemradio"
        :aria-checked="kind === k"
        :class="kind === k ? 'active' : ''"
        :data-testid="`export-${k}`"
        @click="kind = k"
      >{{ kindLabel(k) }}</button>
    </div>
    <div class="sh-bs-detail">
      <h2>{{ kindLabel(kind) }}</h2>
      <label v-if="kind !== 'pdf'" class="sh-field">
        <span>File name</span>
        <input class="sh-input" :value="name" data-testid="export-name" @input="name = ($event.target as HTMLInputElement).value" />
      </label>
      <label v-if="kind === 'excel'" class="sh-field">
        <span>File type</span>
        <select class="sh-select" :value="type" data-testid="export-type" @change="type = ($event.target as HTMLSelectElement).value as ExportFormat">
          <option v-for="t in EXCEL_TYPES" :key="t.value" :value="t.value">{{ t.label }}</option>
        </select>
      </label>
      <p class="sh-bs-lead">{{ LEADS[kind] }}</p>
      <button type="button" class="sh-bs-action" data-testid="export-file" @click="run">
        <Icon :name="kind === 'pdf' ? 'picture_as_pdf' : 'file_download'" />
        <span>{{ kind === 'excel' ? 'Export Excel File' : kind === 'csv' ? 'Export CSV File' : 'Export PDF File' }}</span>
      </button>
    </div>
  </div>
</template>
