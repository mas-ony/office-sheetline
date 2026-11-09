<script setup lang="ts">
/** Insert or Delete cells: shift the cells, or whole rows or columns. */
import { cmd } from '@sheetline/core/views';
import { ref } from 'vue';
import { useUI } from '../../context';
import Dialog from '../Dialog.vue';
import Radio from '../Radio.vue';

const props = defineProps<{ remove?: boolean }>();
const emit = defineEmits<{ close: [] }>();

const ui = useUI();
const mode = ref<'shift' | 'shift2' | 'rows' | 'cols'>('shift');
const ok = () => {
  const ctx = ui.value;
  const { remove } = props;
  cmd.commit(ctx);
  const range = ctx.doc.range;
  if (mode.value === 'rows') {
    if (remove) cmd.deleteRows(ctx);
    else cmd.insertRows(ctx);
    return true;
  }
  if (mode.value === 'cols') {
    if (remove) cmd.deleteCols(ctx);
    else cmd.insertCols(ctx);
    return true;
  }
  const error = remove ? ctx.doc.deleteCells(range, mode.value === 'shift' ? 'left' : 'up') : ctx.doc.insertCells(range, mode.value === 'shift' ? 'right' : 'down');
  if (error) {
    ctx.notify(error);
    return false;
  }
  return true;
};
</script>

<template>
  <Dialog :title="remove ? 'Delete' : 'Insert'" :width="260" :test-id="remove ? 'delete-cells' : 'insert-cells'" :on-ok="ok" @close="emit('close')">
    <fieldset class="sh-group">
      <legend>{{ remove ? 'Delete' : 'Insert' }}</legend>
      <Radio name="cells" :label="remove ? 'Shift cells left' : 'Shift cells right'" :checked="mode === 'shift'" test-id="cells-shift-h" @change="mode = 'shift'" />
      <Radio name="cells" :label="remove ? 'Shift cells up' : 'Shift cells down'" :checked="mode === 'shift2'" test-id="cells-shift-v" @change="mode = 'shift2'" />
      <Radio name="cells" label="Entire row" :checked="mode === 'rows'" test-id="cells-rows" @change="mode = 'rows'" />
      <Radio name="cells" label="Entire column" :checked="mode === 'cols'" test-id="cells-cols" @change="mode = 'cols'" />
    </fieldset>
  </Dialog>
</template>
