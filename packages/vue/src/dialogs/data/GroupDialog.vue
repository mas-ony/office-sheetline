<script setup lang="ts">
/** Group / Ungroup: rows or columns of the selection. */
import { commit } from '@sheetline/core/views';
import { ref } from 'vue';
import { useUI } from '../../context';
import Dialog from '../Dialog.vue';
import Radio from '../Radio.vue';

const props = defineProps<{ ungroup?: boolean }>();
const emit = defineEmits<{ close: [] }>();

const ui = useUI();
const axis = ref<'row' | 'col'>('row');
const ok = () => {
  const ctx = ui.value;
  commit(ctx);
  const r = ctx.doc.range;
  const [from, to] = axis.value === 'row' ? [r.r0, r.r1] : [r.c0, r.c1];
  if (props.ungroup) {
    if (!ctx.doc.ungroupLines(axis.value, from, to)) ctx.notify('The selection is not grouped.');
    return true;
  }
  const error = ctx.doc.groupLines(axis.value, from, to);
  if (error) ctx.notify(error);
  return true;
};
</script>

<template>
  <Dialog :title="ungroup ? 'Ungroup' : 'Group'" :width="260" test-id="group-dialog" :on-ok="ok" @close="emit('close')">
    <Radio name="group-axis" label="Rows" :checked="axis === 'row'" test-id="group-rows" @change="axis = 'row'" />
    <Radio name="group-axis" label="Columns" :checked="axis === 'col'" test-id="group-cols" @change="axis = 'col'" />
  </Dialog>
</template>
