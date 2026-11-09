<script setup lang="ts">
/** File > Open: browse for a workbook or a text file. */
import { useUI } from '../context';
import Icon from '../Icon.vue';

const emit = defineEmits<{ close: [] }>();
const ui = useUI();
const browse = () => {
  emit('close');
  ui.value.files.open();
};
</script>

<template>
  <h1>Open</h1>
  <p class="sh-bs-lead">Open an Excel workbook (.xlsx, .xlsm, .xltx, .xltm, .xls, .xlsb) or a text file (.csv, .tsv, .txt) from this device. You can also drop a file anywhere on the window.</p>
  <button type="button" class="sh-bs-action" data-testid="open-browse" @click="browse">
    <Icon name="folder_open" />
    <span>Browse</span>
  </button>
</template>
