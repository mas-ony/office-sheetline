<script setup lang="ts">
/** Name Manager: the defined names and the table names, filtered, with New, Edit and Delete. */
import { CellError, type DefinedName, rangeText } from '@sheetline/core/views';
import { computed, ref, shallowRef } from 'vue';
import { useUI } from '../../context';
import Dialog from '../Dialog.vue';
import Field from '../Field.vue';
import NewNameDialog from './NewNameDialog.vue';

type Filter = 'all' | 'sheet' | 'workbook' | 'errors' | 'noErrors' | 'defined' | 'tables';

interface NameRow {
  key: string;
  name: string;
  formula: string;
  localSheet?: number;
  comment?: string;
  table?: boolean;
}

const emit = defineEmits<{ close: [] }>();

const ui = useUI();
const filter = ref<Filter>('all');
const sel = ref<string | null>(null);
const edit = shallowRef<{ row: NameRow | null } | null>(null);
const refersTo = ref<string | null>(null);
const version = ref(0);
const rows = computed<NameRow[]>(() => {
  void version.value;
  const { doc } = ui.value;
  const names: NameRow[] = doc.wb.names.filter((n) => !n.hidden && !n.name.startsWith('_xlnm.')).map((n: DefinedName) => ({ key: `${n.name}|${n.localSheet ?? ''}`, name: n.name, formula: n.formula, localSheet: n.localSheet, comment: n.comment }));
  const tables: NameRow[] = doc.wb.sheets.flatMap((s) => s.tables.map((t) => ({ key: `table|${t.name}`, name: t.name, formula: `${s.name}!${rangeText(t.ref, s)}`, table: true })));
  return [...names, ...tables].sort((a, b) => a.name.localeCompare(b.name));
});
const value = (r: NameRow) => {
  const { doc } = ui.value;
  try {
    const v = doc.engine.evaluateText(r.formula, doc.wb.sheets[r.localSheet ?? doc.activeIndex] ?? doc.sheet, 0, 0);
    if (v instanceof CellError) return v.code;
    if (v && typeof v === 'object' && 'areas' in (v as object)) return '{…}';
    return v === null ? '' : String(v);
  } catch {
    return '#NAME?';
  }
};
const shown = computed(() =>
  rows.value.filter((r) => {
    switch (filter.value) {
      case 'sheet': return r.localSheet !== undefined;
      case 'workbook': return r.localSheet === undefined;
      case 'errors': return value(r).startsWith('#');
      case 'noErrors': return !value(r).startsWith('#');
      case 'defined': return !r.table;
      case 'tables': return !!r.table;
      default: return true;
    }
  }),
);
const current = computed(() => shown.value.find((r) => r.key === sel.value) ?? null);
const scopeName = (r: NameRow) => (r.localSheet === undefined ? 'Workbook' : (ui.value.doc.wb.sheets[r.localSheet]?.name ?? ''));

const remove = async () => {
  const c = current.value;
  if (!c) return;
  const ctx = ui.value;
  if (!(await ctx.confirm(`Are you sure you want to delete the name ${c.name}?`, { title: 'Name Manager', ok: 'OK' }))) return;
  ctx.doc.removeName(c.name, c.localSheet);
  sel.value = null;
  version.value++;
};
const pick = (r: NameRow) => {
  sel.value = r.key;
  refersTo.value = null;
};
const commitRefersTo = () => {
  const c = current.value;
  if (!c || refersTo.value === null) return;
  const error = ui.value.doc.defineName({ name: c.name, formula: refersTo.value, localSheet: c.localSheet, comment: c.comment }, { name: c.name, localSheet: c.localSheet });
  if (error) ui.value.notify(error);
  refersTo.value = null;
  version.value++;
};
const editDone = () => {
  edit.value = null;
  version.value++;
};
</script>

<template>
  <Dialog title="Name Manager" close-only :width="720" test-id="name-manager" @close="emit('close')">
    <div class="sh-row wrap sh-toolbar-row">
      <button type="button" class="sh-btn" data-testid="name-new" @click="edit = { row: null }">New...</button>
      <button type="button" class="sh-btn" :disabled="!current || current.table" @click="current && (edit = { row: current })">Edit...</button>
      <button type="button" class="sh-btn" :disabled="!current || current.table" data-testid="name-delete" @click="remove">Delete</button>
      <span class="grow" />
      <Field label="Filter:">
        <select :value="filter" @change="filter = ($event.target as HTMLSelectElement).value as Filter">
          <option value="all">Clear Filter</option>
          <option value="sheet">Names Scoped to Worksheet</option>
          <option value="workbook">Names Scoped to Workbook</option>
          <option value="errors">Names with Errors</option>
          <option value="noErrors">Names without Errors</option>
          <option value="defined">Defined Names</option>
          <option value="tables">Table Names</option>
        </select>
      </Field>
    </div>
    <div class="sh-table names" role="table" data-testid="name-list">
      <div class="sh-table-head" role="row">
        <span>Name</span>
        <span>Value</span>
        <span>Refers To</span>
        <span>Scope</span>
        <span>Comment</span>
      </div>
      <div class="sh-table-body">
        <div v-for="r in shown" :key="r.key" role="row" :class="`sh-table-row${r.key === sel ? ' selected' : ''}`" @click="pick(r)" @dblclick="!r.table && (edit = { row: r })">
          <span>{{ r.table ? '▦ ' : '' }}{{ r.name }}</span>
          <span>{{ value(r) }}</span>
          <span>={{ r.formula }}</span>
          <span>{{ scopeName(r) }}</span>
          <span>{{ r.comment }}</span>
        </div>
        <div v-if="!shown.length" class="sh-table-empty">No names.</div>
      </div>
    </div>
    <div class="sh-row">
      <span class="sh-label">Refers to:</span>
      <input class="grow" :value="refersTo ?? (current ? `=${current.formula}` : '')" :disabled="!current || current.table" aria-label="Refers to" @input="refersTo = ($event.target as HTMLInputElement).value" />
      <button type="button" class="sh-btn icon" aria-label="Commit" :disabled="refersTo === null || !current" @click="commitRefersTo">✓</button>
      <button type="button" class="sh-btn icon" aria-label="Cancel" :disabled="refersTo === null" @click="refersTo = null">✕</button>
    </div>
  </Dialog>
  <NewNameDialog v-if="edit" :name="edit.row?.name" :scope="edit.row?.localSheet" @close="editDone" />
</template>
