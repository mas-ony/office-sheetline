<script setup lang="ts">
/** Create Table: the range of the table, and whether it has headers. */
import { chartSource, commit, parseRangeText, rangeText } from '@sheetline/core/views';
import { ref } from 'vue';
import { useUI } from '../../context';
import Check from '../Check.vue';
import Dialog from '../Dialog.vue';
import Field from '../Field.vue';
import RangeField from '../RangeField.vue';

const props = defineProps<{ tableStyle?: string }>();
const emit = defineEmits<{ close: [] }>();

const ui = useUI();
const start = ui.value;
const rangeRef = ref(rangeText(chartSource(start) ?? start.doc.range, start.doc.sheet));
const headers = ref(true);
const ok = () => {
  const ctx = ui.value;
  const { doc } = ctx;
  const hit = parseRangeText(doc, rangeRef.value);
  if (!hit || hit.sheet !== doc.sheet) {
    ctx.notify('The reference isn\'t valid. Select a range on this sheet.');
    return false;
  }
  commit(ctx);
  const t = doc.createTable(hit.range, { hasHeaders: headers.value, style: props.tableStyle });
  if (typeof t === 'string') {
    ctx.notify(t);
    return false;
  }
  doc.selectCell(t.ref.r0 + (t.headerRow ? 1 : 0), t.ref.c0);
  ctx.showTab('table-design');
  return true;
};
</script>

<template>
  <Dialog title="Create Table" :width="360" test-id="create-table" :on-ok="ok" @close="emit('close')">
    <Field label="Where is the data for your table?" wide>
      <RangeField :value="rangeRef" label="Table range" test-id="create-table-range" auto-focus @change="rangeRef = $event" />
    </Field>
    <Check label="My table has headers" :checked="headers" test-id="create-table-headers" @change="headers = $event" />
  </Dialog>
</template>
