<script setup lang="ts">
/**
 * Goal Seek, then Goal Seek Status in the same dialog frame: OK keeps the value found, Cancel
 * restores the input cell.
 */
import { cellName, commit, makeRange, parseRangeText, rangeText } from '@sheetline/core/views';
import { ref, shallowRef } from 'vue';
import { useUI } from '../../context';
import Dialog from '../Dialog.vue';
import Field from '../Field.vue';
import RangeField from '../RangeField.vue';

const emit = defineEmits<{ close: [] }>();

const ui = useUI();
const doc0 = ui.value.doc;
const setCell = ref(rangeText(makeRange(doc0.selection.active.r, doc0.selection.active.c), doc0.sheet));
const to = ref('');
const by = ref('');
const status = shallowRef<{ text: string; target: number; result: number } | null>(null);

const seek = () => {
  const ctx = ui.value;
  const { doc } = ctx;
  const set = parseRangeText(doc, setCell.value);
  const changing = parseRangeText(doc, by.value);
  const target = Number(to.value);
  if (!set || set.range.r0 !== set.range.r1 || set.range.c0 !== set.range.c1) {
    ctx.notify('Set cell must contain a single cell reference.');
    return false;
  }
  if (!Number.isFinite(target) || to.value.trim() === '') {
    ctx.notify('To value must contain a number.');
    return false;
  }
  if (!changing || changing.range.r0 !== changing.range.r1 || changing.range.c0 !== changing.range.c1 || changing.sheet !== set.sheet) {
    ctx.notify('By changing cell must contain a single cell reference on the same sheet.');
    return false;
  }
  commit(ctx);
  const res = doc.goalSeek({ r: set.range.r0, c: set.range.c0 }, target, { r: changing.range.r0, c: changing.range.c0 }, set.sheet);
  if (typeof res === 'string') {
    ctx.notify(res);
    return false;
  }
  const ref = cellName(set.range.r0, set.range.c0);
  status.value = { text: res.found ? `Goal Seeking with Cell ${ref} found a solution.` : `Goal Seeking with Cell ${ref} may not have found a solution.`, target, result: res.result };
  return false;
};
const keep = () => {
  // OK keeps the value found (closing directly skips the Cancel path).
  emit('close');
  return false;
};
const close = () => {
  // Cancel restores the input cell.
  if (status.value) ui.value.doc.undo();
  emit('close');
};
</script>

<template>
  <Dialog :title="status ? 'Goal Seek Status' : 'Goal Seek'" :width="340" :test-id="status ? 'goal-seek-status' : 'goal-seek'" :on-ok="status ? keep : seek" @close="close">
    <template v-if="status">
      <p>{{ status.text }}</p>
      <div class="sh-kv"><span>Target value:</span><b>{{ status.target }}</b></div>
      <div class="sh-kv"><span>Current value:</span><b data-testid="goal-seek-result">{{ Number(status.result.toPrecision(12)) }}</b></div>
    </template>
    <template v-else>
      <Field label="Set cell:" wide>
        <RangeField :value="setCell" label="Set cell" test-id="goal-seek-set" @change="setCell = $event" />
      </Field>
      <Field label="To value:" wide>
        <input :value="to" inputmode="decimal" data-testid="goal-seek-to" @input="to = ($event.target as HTMLInputElement).value" />
      </Field>
      <Field label="By changing cell:" wide>
        <RangeField :value="by" label="By changing cell" test-id="goal-seek-by" @change="by = $event" />
      </Field>
    </template>
  </Dialog>
</template>
