<script setup lang="ts">
/** Resize Table: the new range of the table. */
import { parseRangeText, rangeText } from '@sheetline/core/views';
import { computed, ref } from 'vue';
import { useUI } from '../../context';
import Dialog from '../Dialog.vue';
import Field from '../Field.vue';
import RangeField from '../RangeField.vue';

const props = defineProps<{ tableId: number }>();
const emit = defineEmits<{ close: [] }>();

const ui = useUI();
const hit = computed(() => ui.value.doc.tableById(props.tableId));
const rangeRef = ref(hit.value ? rangeText(hit.value.table.ref, hit.value.sheet) : '');
const ok = () => {
  const ctx = ui.value;
  const h = hit.value;
  if (!h) return;
  const r = parseRangeText(ctx.doc, rangeRef.value, h.sheet);
  if (!r || r.sheet !== h.sheet) {
    ctx.notify('The reference isn\'t valid.');
    return false;
  }
  const error = ctx.doc.resizeTable(props.tableId, r.range);
  if (error) {
    ctx.notify(error);
    return false;
  }
  return true;
};
</script>

<template>
  <Dialog v-if="hit" title="Resize Table" :width="360" test-id="resize-table" :on-ok="ok" @close="emit('close')">
    <Field label="Select the new data range for your table:" wide>
      <RangeField :value="rangeRef" label="Table range" test-id="resize-table-range" :home="hit.sheet" auto-focus @change="rangeRef = $event" />
    </Field>
    <p class="sh-hint">Note: The headers must remain in the same row, and the resulting table range must overlap the original table range.</p>
  </Dialog>
</template>
