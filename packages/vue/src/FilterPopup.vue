<script setup lang="ts">
/** The drop-down of a filter button: sorting, the condition filters, and the values to show. */
import { type ColumnFilter, type MenuItem, NUMBER_CONDITIONS, TEXT_CONDITIONS } from '@sheetline/core/views';
import { computed, ref, shallowRef } from 'vue';
import Anchored from './Anchored.vue';
import { useUI } from './context';
import CustomFilterDialog from './CustomFilterDialog.vue';
import { vChecked } from './directives';
import Icon from './Icon.vue';
import Menu from './Menu.vue';

const props = defineProps<{ col: number; at: { x: number; y: number }; tableId?: number }>();
const emit = defineEmits<{ close: [] }>();

const ui = useUI();
// A new object on every change: the filters are updated in place.
const view = computed(() => {
  const { doc } = ui.value;
  const table = props.tableId !== undefined ? doc.tableById(props.tableId)?.table : undefined;
  const range = table ? { ...table.ref, r1: table.ref.r1 - (table.totalsRow ? 1 : 0) } : doc.sheet.autoFilter?.ref;
  const current: ColumnFilter | undefined = table ? table.filters.get(props.col) : doc.sheet.autoFilter?.columns.get(props.col);
  return { doc, table, range, current };
});
type FilterValues = ReturnType<typeof ui.value.doc.filterValues>;
let memo: { doc: unknown; col: number; range: unknown; values: FilterValues } | null = null;
const values = computed(() => {
  const { doc, range } = view.value;
  if (memo && memo.doc === doc && memo.col === props.col && memo.range === range) return memo.values;
  const v = doc.filterValues(props.col, doc.sheet, range);
  memo = { doc, col: props.col, range, values: v };
  return v;
});
const numeric = computed(() => values.value.filter((v) => v.text !== '').every((v) => v.numeric) && values.value.some((v) => v.numeric));
const header = computed(() => {
  const { doc, range } = view.value;
  return range ? doc.displayText(range.r0, props.col) || `Column ${props.col + 1}` : '';
});
const search = ref('');
const custom = shallowRef<ColumnFilter | null>(null);
const sub = shallowRef<DOMRect | null>(null);
const initial = view.value.current;
const checked = shallowRef<Set<string>>(
  initial?.values ? new Set([...initial.values.map((v) => v.toLowerCase()), ...(initial.blanks ? [''] : [])]) : new Set(values.value.map((v) => v.text.toLowerCase())),
);
const shown = computed(() => values.value.filter((v) => !search.value || v.text.toLowerCase().includes(search.value.toLowerCase())));
const all = computed(() => shown.value.every((v) => checked.value.has(v.text.toLowerCase())));
const rect = computed(() => new DOMRect(props.at.x, props.at.y, 0, 0));

const setFilter = (f: ColumnFilter | null) => {
  const { doc, table } = view.value;
  if (table) doc.setTableFilter(table.id, props.col, f);
  else doc.setColumnFilter(props.col, f);
};
const apply = () => {
  const visible = search.value ? shown.value : values.value;
  const allChecked = values.value.every((v) => checked.value.has(v.text.toLowerCase())) && !search.value;
  if (allChecked) setFilter(null);
  else {
    const keep = visible.filter((v) => checked.value.has(v.text.toLowerCase()) && v.text !== '').map((v) => v.text);
    setFilter({ values: keep, blanks: visible.some((v) => v.text === '') && checked.value.has('') });
  }
  emit('close');
};
const sort = (asc: boolean) => {
  const { doc, table } = view.value;
  if (table) doc.sortTable(table.id, props.col, asc);
  else doc.sortFilterColumn(props.col, asc);
  emit('close');
};
const clear = () => {
  setFilter(null);
  emit('close');
};
const subItems = computed<MenuItem[]>(() => [
  ...(numeric.value ? NUMBER_CONDITIONS : TEXT_CONDITIONS).map((c) => ({ title: c.label, onClick: () => (custom.value = { conditions: [{ op: c.op, value: '' }], and: true }) })),
  { type: 'line' },
  { title: 'Custom Filter...', onClick: () => (custom.value = view.value.current ?? { conditions: [], and: true }) },
]);
const toggleAll = (e: Event) => {
  const on = (e.target as HTMLInputElement).checked;
  const next = new Set(checked.value);
  for (const v of shown.value) {
    if (on) next.add(v.text.toLowerCase());
    else next.delete(v.text.toLowerCase());
  }
  checked.value = next;
};
const toggle = (text: string, e: Event) => {
  const next = new Set(checked.value);
  if ((e.target as HTMLInputElement).checked) next.add(text.toLowerCase());
  else next.delete(text.toLowerCase());
  checked.value = next;
};
const close = () => {
  if (!sub.value) emit('close');
};
const applyCustom = (f: ColumnFilter) => {
  setFilter(f.conditions?.length ? f : null);
  emit('close');
};
</script>

<template>
  <CustomFilterDialog v-if="custom" :title="header" :initial="custom" @close="emit('close')" @apply="applyCustom" />
  <Anchored v-else :rect="rect" below box-class="sh-filter" test-id="filter-popup" label="Filter" @close="close">
    <button type="button" class="sh-filter-item" data-testid="filter-sort-asc" @click="sort(true)">
      <Icon name="sort_by_alpha" /> {{ numeric ? 'Sort Smallest to Largest' : 'Sort A to Z' }}
    </button>
    <button type="button" class="sh-filter-item" data-testid="filter-sort-desc" @click="sort(false)">
      <Icon name="sort" /> {{ numeric ? 'Sort Largest to Smallest' : 'Sort Z to A' }}
    </button>
    <div class="sh-filter-sep" />
    <button type="button" class="sh-filter-item" :disabled="!view.current" data-testid="filter-clear" @click="clear">
      <Icon name="filter_alt_off" /> Clear Filter From “{{ header }}”</button>
    <button type="button" class="sh-filter-item" data-testid="filter-conditions" @click="sub = ($event.currentTarget as HTMLElement).getBoundingClientRect()">
      <Icon name="filter_list" /> {{ numeric ? 'Number Filters' : 'Text Filters' }} <Icon name="arrow_right" class="sub" />
    </button>
    <input class="sh-input sh-filter-search" placeholder="Search" :value="search" aria-label="Search" data-testid="filter-search" @input="search = ($event.target as HTMLInputElement).value" @keydown.enter="apply" />
    <div class="sh-filter-values" data-testid="filter-values">
      <label><input v-checked="all" type="checkbox" @change="toggleAll" />{{ search ? '(Select All Search Results)' : '(Select All)' }}</label>
      <label v-for="v in shown" :key="v.text || '(blank)'"><input v-checked="checked.has(v.text.toLowerCase())" type="checkbox" @change="toggle(v.text, $event)" />{{ v.text === '' ? '(Blanks)' : v.text }}</label>
    </div>
    <div class="sh-filter-actions">
      <button type="button" class="sh-btn-primary" data-testid="filter-ok" @click="apply">OK</button>
      <button type="button" class="sh-btn-secondary" @click="emit('close')">Cancel</button>
    </div>
    <Menu v-if="sub" :at="sub" :items="subItems" @close="sub = null" />
  </Anchored>
</template>
