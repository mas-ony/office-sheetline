<script setup lang="ts">
/** Subtotal: a function of chosen columns at each change in a column. */
import { colName, commit, dataRegion, SUBTOTAL_FUNCTIONS } from '@sheetline/core/views';
import { ref, shallowRef } from 'vue';
import { useUI } from '../../context';
import Check from '../Check.vue';
import Dialog from '../Dialog.vue';
import Field from '../Field.vue';

const emit = defineEmits<{ close: [] }>();

const ui = useUI();
const start = ui.value;
const region = dataRegion(start) ?? start.doc.range;
const cols = Array.from({ length: region.c1 - region.c0 + 1 }, (_, i) => region.c0 + i);
const name = (c: number) => ui.value.doc.displayText(region.r0, c) || `Column ${colName(c)}`;
const groupBy = ref(region.c0);
const code = ref(9);
const columns = shallowRef<Set<number>>(new Set([region.c1]));
const replace = ref(true);
const pageBreaks = ref(false);
const below = ref(true);
const choose = (c: number, v: boolean) => {
  const n = new Set(columns.value);
  if (v) n.add(c);
  else n.delete(c);
  columns.value = n;
};

const removeAll = () => {
  const ctx = ui.value;
  commit(ctx);
  ctx.doc.removeSubtotals(region);
  emit('close');
};
const ok = () => {
  const ctx = ui.value;
  commit(ctx);
  const error = ctx.doc.applySubtotals(region, { groupBy: groupBy.value, code: code.value, columns: cols.filter((c) => columns.value.has(c)), replace: replace.value, pageBreaks: pageBreaks.value, summaryBelow: below.value });
  if (error) {
    ctx.notify(error);
    return false;
  }
  return true;
};
</script>

<template>
  <Dialog title="Subtotal" :width="380" test-id="subtotal-dialog" :on-ok="ok" @close="emit('close')">
    <Field label="At each change in:" wide>
      <select :value="groupBy" data-testid="subtotal-groupby" @change="groupBy = Number(($event.target as HTMLSelectElement).value)">
        <option v-for="c in cols" :key="c" :value="c">{{ name(c) }}</option>
      </select>
    </Field>
    <Field label="Use function:" wide>
      <select :value="code" data-testid="subtotal-function" @change="code = Number(($event.target as HTMLSelectElement).value)">
        <option v-for="f in SUBTOTAL_FUNCTIONS" :key="f.code" :value="f.code">{{ f.label }}</option>
      </select>
    </Field>
    <div class="sh-label">Add subtotal to:</div>
    <div class="sh-checklist" data-testid="subtotal-columns">
      <Check v-for="c in cols" :key="c" :label="name(c)" :checked="columns.has(c)" @change="choose(c, $event)" />
    </div>
    <Check label="Replace current subtotals" :checked="replace" @change="replace = $event" />
    <Check label="Page break between groups" :checked="pageBreaks" @change="pageBreaks = $event" />
    <Check label="Summary below data" :checked="below" @change="below = $event" />
    <template #footer-left>
      <button type="button" class="sh-btn" data-testid="subtotal-remove" @click="removeAll">Remove All</button>
    </template>
  </Dialog>
</template>
