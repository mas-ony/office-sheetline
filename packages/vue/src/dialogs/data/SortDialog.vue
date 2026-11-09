<script setup lang="ts">
/** Sort: levels (column, sort on, order) applied in turn, with the Sort Options dialog. */
import { commit, CUSTOM_LISTS, dataRegion, looksLikeHeader, type SortLevel, sortLineName, sortLineNumeric, sortLines } from '@sheetline/core/views';
import { computed, ref, shallowRef } from 'vue';
import { useUI } from '../../context';
import { vValue } from '../../directives';
import Check from '../Check.vue';
import Dialog from '../Dialog.vue';
import Radio from '../Radio.vue';

const emit = defineEmits<{ close: [] }>();

const ui = useUI();
const start = ui.value;
const region = dataRegion(start) ?? start.doc.range;
const firstLine = Math.max(region.c0, Math.min(region.c1, start.doc.selection.active.c));
let levelKey = 0;
const headers = ref(looksLikeHeader(start, region, firstLine));
const byColumns = ref(false);
const caseSensitive = ref(false);
const options = ref(false);
const levels = shallowRef<(SortLevel & { key: number })[]>([{ key: levelKey++, index: firstLine, order: 'asc', sortOn: 'values' }]);
const sel = ref(0);
const lines = computed(() => sortLines(region, byColumns.value));
const lineName = (i: number) => sortLineName(ui.value.doc, region, byColumns.value, headers.value, i);
const numeric = (i: number) => sortLineNumeric(ui.value.doc, region, byColumns.value, headers.value, i);
const set = (k: number, patch: Partial<SortLevel>) => (levels.value = levels.value.map((l) => (l.key === k ? { ...l, ...patch } : l)));

const ok = () => {
  const ctx = ui.value;
  commit(ctx);
  ctx.doc.sortLevels(region, levels.value, { hasHeader: headers.value && !byColumns.value, byColumns: byColumns.value, caseSensitive: caseSensitive.value });
  return true;
};
const addLevel = () => {
  const list = levels.value;
  const unused = lines.value.find((i) => !list.some((l) => l.index === i)) ?? lines.value[0];
  levels.value = [...list, { key: levelKey++, index: unused, order: 'asc', sortOn: 'values' }];
  sel.value = list.length;
};
const deleteLevel = () => {
  const s = sel.value;
  levels.value = levels.value.filter((_, i) => i !== s);
  sel.value = Math.max(0, s - 1);
};
const copyLevel = () => {
  const s = sel.value;
  const list = levels.value;
  const l = list[s];
  if (l) {
    levels.value = [...list.slice(0, s + 1), { ...l, key: levelKey++ }, ...list.slice(s + 1)];
    sel.value = s + 1;
  }
};
const moveLevel = (dir: -1 | 1) => {
  const s = sel.value;
  const n = levels.value.slice();
  [n[s + dir], n[s]] = [n[s], n[s + dir]];
  levels.value = n;
  sel.value = s + dir;
};
const pickSortOn = (k: number, v: string) => set(k, { sortOn: v as SortLevel['sortOn'], order: v === 'values' ? 'asc' : 'asc', color: v === 'values' ? undefined : '' });
const pickOrder = (k: number, v: string) => {
  if (v.startsWith('custom:')) set(k, { order: 'custom', list: CUSTOM_LISTS[Number(v.slice(7))] });
  else set(k, { order: v as 'asc' | 'desc', list: undefined });
};
const orderValue = (l: SortLevel) => (l.order === 'custom' ? `custom:${CUSTOM_LISTS.findIndex((x) => x === l.list)}` : l.order);
const leftToRight = () => {
  byColumns.value = true;
  levels.value = levels.value.map((l) => ({ ...l, index: region.r0 }));
};
</script>

<template>
  <Dialog title="Sort" :width="640" test-id="sort-dialog" :on-ok="ok" @close="emit('close')">
    <div class="sh-row wrap sh-toolbar-row">
      <button type="button" class="sh-btn" data-testid="sort-add-level" @click="addLevel">+ Add Level</button>
      <button type="button" class="sh-btn" :disabled="levels.length <= 1" @click="deleteLevel">✕ Delete Level</button>
      <button type="button" class="sh-btn" @click="copyLevel">Copy Level</button>
      <button type="button" class="sh-btn icon" aria-label="Move up" :disabled="sel <= 0" @click="moveLevel(-1)">▲</button>
      <button type="button" class="sh-btn icon" aria-label="Move down" :disabled="sel >= levels.length - 1" @click="moveLevel(1)">▼</button>
      <button type="button" class="sh-btn" @click="options = true">Options...</button>
      <span class="grow" />
      <Check label="My data has headers" :checked="headers" :disabled="byColumns" test-id="sort-headers" @change="headers = $event" />
    </div>
    <div class="sh-table sort" role="table">
      <div class="sh-table-head" role="row">
        <span>{{ byColumns ? 'Row' : 'Column' }}</span>
        <span>Sort On</span>
        <span>Order</span>
      </div>
      <div class="sh-table-body">
        <div v-for="(l, i) in levels" :key="l.key" role="row" :class="`sh-table-row${i === sel ? ' selected' : ''}`" @click="sel = i">
          <span>
            <span class="sh-sort-by">{{ i === 0 ? 'Sort by' : 'Then by' }}</span>
            <select :value="l.index" :data-testid="`sort-by-${i}`" aria-label="Sort by" @change="set(l.key, { index: Number(($event.target as HTMLSelectElement).value) })">
              <option v-for="x in lines" :key="x" :value="x">{{ lineName(x) }}</option>
            </select>
          </span>
          <span>
            <select :value="l.sortOn" aria-label="Sort on" @change="pickSortOn(l.key, ($event.target as HTMLSelectElement).value)">
              <option value="values">Cell Values</option>
              <option value="cellColor">Cell Color</option>
              <option value="fontColor">Font Color</option>
            </select>
          </span>
          <span>
            <select v-if="l.sortOn === 'values'" :value="orderValue(l)" aria-label="Order" :data-testid="`sort-order-${i}`" @change="pickOrder(l.key, ($event.target as HTMLSelectElement).value)">
              <option value="asc">{{ numeric(l.index) ? 'Smallest to Largest' : 'A to Z' }}</option>
              <option value="desc">{{ numeric(l.index) ? 'Largest to Smallest' : 'Z to A' }}</option>
              <option v-for="(list, k) in CUSTOM_LISTS" :key="k" :value="`custom:${k}`">{{ list.slice(0, 3).join(', ') }}…</option>
            </select>
            <span v-else class="sh-row">
              <input v-value="l.color ?? ''" placeholder="#RRGGBB" :style="{ width: '80px' }" aria-label="Color" @input="set(l.key, { color: ($event.target as HTMLInputElement).value.toUpperCase() })" />
              <select :value="l.order" aria-label="Position" @change="set(l.key, { order: ($event.target as HTMLSelectElement).value as 'asc' | 'desc' })">
                <option value="asc">On Top</option>
                <option value="desc">On Bottom</option>
              </select>
            </span>
          </span>
        </div>
      </div>
    </div>
  </Dialog>
  <Dialog v-if="options" title="Sort Options" :width="300" test-id="sort-options" @close="options = false">
    <Check label="Case sensitive" :checked="caseSensitive" @change="caseSensitive = $event" />
    <fieldset class="sh-group">
      <legend>Orientation</legend>
      <Radio name="sort-orientation" label="Sort top to bottom" :checked="!byColumns" @change="byColumns = false" />
      <Radio name="sort-orientation" label="Sort left to right" :checked="byColumns" @change="leftToRight" />
    </fieldset>
  </Dialog>
</template>
