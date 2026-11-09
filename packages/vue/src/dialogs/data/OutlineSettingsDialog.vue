<script setup lang="ts">
/** Outline settings: where summary rows and columns are, and Create (an automatic outline). */
import { ref } from 'vue';
import { useUI } from '../../context';
import Check from '../Check.vue';
import Dialog from '../Dialog.vue';

const emit = defineEmits<{ close: [] }>();

const ui = useUI();
const sheet = ui.value.doc.sheet;
const below = ref(sheet.outlineBelow);
const right = ref(sheet.outlineRight);
const ok = () => ui.value.doc.setOutlineSettings({ summaryBelow: below.value, summaryRight: right.value });
const create = () => {
  const ctx = ui.value;
  ctx.doc.setOutlineSettings({ summaryBelow: below.value, summaryRight: right.value });
  if (!ctx.doc.autoOutline()) ctx.notify('Cannot create an outline.');
  emit('close');
};
</script>

<template>
  <Dialog title="Settings" :width="320" test-id="outline-settings" :on-ok="ok" @close="emit('close')">
    <fieldset class="sh-group">
      <legend>Direction</legend>
      <Check label="Summary rows below detail" :checked="below" @change="below = $event" />
      <Check label="Summary columns to right of detail" :checked="right" @change="right = $event" />
    </fieldset>
    <template #footer-left>
      <button type="button" class="sh-btn" @click="create">Create</button>
    </template>
  </Dialog>
</template>
