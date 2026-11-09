<script setup lang="ts">
/** File > Save: the file name, then the workbook is saved. */
import { ref } from 'vue';
import { fileBaseName } from '@sheetline/core/views';
import { useUI } from '../context';
import Icon from '../Icon.vue';

const emit = defineEmits<{ close: [] }>();
const ui = useUI();
const name = ref(fileBaseName(ui.value.doc.fileName));
const ext = /\.xlsm$/i.test(ui.value.doc.fileName) ? '.xlsm' : '.xlsx';
const save = () => {
  const ctx = ui.value;
  if (name.value.trim() && name.value.trim() !== fileBaseName(ctx.doc.fileName)) ctx.files.rename(name.value.trim());
  emit('close');
  ctx.files.save();
};
</script>

<template>
  <h1>Save</h1>
  <p class="sh-bs-lead">Save the workbook as an Excel file. Where the browser allows it, the file you opened is updated; otherwise a copy is downloaded.</p>
  <div class="sh-bs-form">
    <label class="sh-field">
      <span>File name</span>
      <span class="sh-inline">
        <input class="sh-input" :value="name" data-testid="save-name" @input="name = ($event.target as HTMLInputElement).value" />
        <span>{{ ext }}</span>
      </span>
    </label>
    <button type="button" class="sh-bs-action" data-testid="save-button" @click="save">
      <Icon name="save" />
      <span>Save</span>
    </button>
  </div>
</template>
