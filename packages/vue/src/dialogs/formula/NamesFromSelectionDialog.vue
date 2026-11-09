<script setup lang="ts">
/** Create Names from Selection: names from the labels in the top row, left column, bottom row or right column. */
import { makeRange, rangeText } from '@sheetline/core/views';
import { computed, ref } from 'vue';
import { useUI } from '../../context';
import Check from '../Check.vue';
import Dialog from '../Dialog.vue';

const emit = defineEmits<{ close: [] }>();

const ui = useUI();
const range = computed(() => {
  const { doc } = ui.value;
  return doc.bounded(doc.range, doc.sheet) ?? doc.range;
});
const top = ref(range.value.r1 > range.value.r0);
const left = ref(range.value.c1 > range.value.c0);
const bottom = ref(false);
const right = ref(false);

const ok = () => {
  const { doc } = ui.value;
  const sheet = doc.sheet;
  const rg = range.value;
  const q = /[^\w]/.test(sheet.name) ? `'${sheet.name.replace(/'/g, "''")}'` : sheet.name;
  const errors: string[] = [];
  const add = (label: string, r0: number, c0: number, r1: number, c1: number) => {
    const name = label.trim().replace(/\s+/g, '_').replace(/^(\d)/, '_$1');
    if (!name) return;
    const error = doc.defineName({ name, formula: `${q}!${rangeText(makeRange(r0, c0, r1, c1), sheet)}` }, doc.wb.names.some((x) => x.name.toUpperCase() === name.toUpperCase() && x.localSheet === undefined) ? { name } : undefined);
    if (error) errors.push(name);
  };
  const r0 = rg.r0 + (top.value ? 1 : 0);
  const r1 = rg.r1 - (bottom.value ? 1 : 0);
  const c0 = rg.c0 + (left.value ? 1 : 0);
  const c1 = rg.c1 - (right.value ? 1 : 0);
  doc.transact('Create names', () => {
    if (top.value) for (let c = c0; c <= c1; c++) add(doc.displayText(rg.r0, c), r0, c, r1, c);
    if (bottom.value) for (let c = c0; c <= c1; c++) add(doc.displayText(rg.r1, c), r0, c, r1, c);
    if (left.value) for (let r = r0; r <= r1; r++) add(doc.displayText(r, rg.c0), r, c0, r, c1);
    if (right.value) for (let r = r0; r <= r1; r++) add(doc.displayText(r, rg.c1), r, c0, r, c1);
  });
  if (errors.length) ui.value.notify(`Some names could not be created: ${errors.slice(0, 5).join(', ')}`);
  return true;
};
</script>

<template>
  <Dialog title="Create Names from Selection" :width="320" test-id="names-from-selection" :on-ok="ok" @close="emit('close')">
    <p class="sh-lead">Create names from values in the:</p>
    <Check label="Top row" :checked="top" test-id="names-top" @change="top = $event" />
    <Check label="Left column" :checked="left" test-id="names-left" @change="left = $event" />
    <Check label="Bottom row" :checked="bottom" @change="bottom = $event" />
    <Check label="Right column" :checked="right" @change="right = $event" />
  </Dialog>
</template>
