<script setup lang="ts">
/** Remove Duplicates: the columns compared, with or without a header row. */
import { colName, commit, dataRegion, looksLikeHeader } from '@sheetline/core/views';
import { ref, shallowRef } from 'vue';
import { useUI } from '../../context';
import Check from '../Check.vue';
import Dialog from '../Dialog.vue';

const emit = defineEmits<{ close: [] }>();

const ui = useUI();
const start = ui.value;
const region = dataRegion(start) ?? start.doc.range;
const headers = ref(looksLikeHeader(start, region, region.c0));
const cols = Array.from({ length: region.c1 - region.c0 + 1 }, (_, i) => region.c0 + i);
const chosen = shallowRef<Set<number>>(new Set(cols));
const name = (c: number) => (headers.value ? ui.value.doc.displayText(region.r0, c) || `Column ${colName(c)}` : `Column ${colName(c)}`);
const choose = (c: number, v: boolean) => {
  const n = new Set(chosen.value);
  if (v) n.add(c);
  else n.delete(c);
  chosen.value = n;
};

const ok = () => {
  const ctx = ui.value;
  if (!chosen.value.size) {
    ctx.notify('Select at least one column.');
    return false;
  }
  commit(ctx);
  const res = ctx.doc.removeDuplicates(region, cols.filter((c) => chosen.value.has(c)), headers.value);
  ctx.notify(res.removed ? `${res.removed} duplicate value${res.removed > 1 ? 's' : ''} found and removed; ${res.remaining} unique value${res.remaining === 1 ? '' : 's'} remain.` : 'No duplicate values found.');
  return true;
};
</script>

<template>
  <Dialog title="Remove Duplicates" :width="420" test-id="remove-duplicates" :on-ok="ok" @close="emit('close')">
    <p class="sh-lead">To delete duplicate values, select one or more columns that contain duplicates.</p>
    <div class="sh-row wrap">
      <button type="button" class="sh-btn" @click="chosen = new Set(cols)">Select All</button>
      <button type="button" class="sh-btn" @click="chosen = new Set()">Unselect All</button>
      <span class="grow" />
      <Check label="My data has headers" :checked="headers" test-id="remove-duplicates-headers" @change="headers = $event" />
    </div>
    <div class="sh-checklist" data-testid="remove-duplicates-columns">
      <Check v-for="c in cols" :key="c" :label="name(c)" :checked="chosen.has(c)" @change="choose(c, $event)" />
    </div>
  </Dialog>
</template>
