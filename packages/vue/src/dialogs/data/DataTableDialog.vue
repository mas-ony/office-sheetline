<script setup lang="ts">
/** Data Table: the row and column input cells of a what-if table over the selection. */
import { commit, parseRangeText } from '@sheetline/core/views';
import { ref } from 'vue';
import { useUI } from '../../context';
import Dialog from '../Dialog.vue';
import Field from '../Field.vue';
import RangeField from '../RangeField.vue';

const emit = defineEmits<{ close: [] }>();

const ui = useUI();
const rowInput = ref('');
const colInput = ref('');
const ok = () => {
  const ctx = ui.value;
  const { doc } = ctx;
  const cell = (t: string) => {
    if (!t.trim()) return null;
    const hit = parseRangeText(doc, t);
    return hit && hit.range.r0 === hit.range.r1 && hit.range.c0 === hit.range.c1 && hit.sheet === doc.sheet ? { r: hit.range.r0, c: hit.range.c0 } : undefined;
  };
  const ri = cell(rowInput.value);
  const ci = cell(colInput.value);
  if (ri === undefined || ci === undefined) {
    ctx.notify('Input cell reference is not valid.');
    return false;
  }
  commit(ctx);
  const error = doc.dataTable(doc.range, ri, ci);
  if (error) {
    ctx.notify(error);
    return false;
  }
  return true;
};
</script>

<template>
  <Dialog title="Data Table" :width="340" test-id="data-table" :on-ok="ok" @close="emit('close')">
    <Field label="Row input cell:" wide>
      <RangeField :value="rowInput" label="Row input cell" test-id="data-table-row" @change="rowInput = $event" />
    </Field>
    <Field label="Column input cell:" wide>
      <RangeField :value="colInput" label="Column input cell" test-id="data-table-col" @change="colInput = $event" />
    </Field>
    <p class="sh-hint">The results are calculated once and stored as values.</p>
  </Dialog>
</template>
