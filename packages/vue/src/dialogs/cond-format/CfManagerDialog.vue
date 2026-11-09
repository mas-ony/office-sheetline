<script setup lang="ts">
/** Conditional Formatting Rules Manager: the rules of the selection or the sheet, edited, reordered and applied together. */
import { type CfRule, commit, describeRule, parseRangeText, type Range, rangeText, type XmlElement } from '@sheetline/core/views';
import { computed, ref, shallowRef } from 'vue';
import { useUI } from '../../context';
import { vChecked } from '../../directives';
import Dialog from '../Dialog.vue';
import RulePreview from './RulePreview.vue';
import RuleEditorDialog from './RuleEditorDialog.vue';

const emit = defineEmits<{ close: [] }>();

interface Row {
  key: string;
  id?: string;
  rule: CfRule | null;
  raw?: XmlElement;
  ranges: Range[];
  text: string;
  stopIfTrue: boolean;
}

type Scope = 'selection' | 'sheet';

const ui = useUI();
let rowKey = 0;
const load = (s: Scope): Row[] => {
  const { doc } = ui.value;
  return doc.cfRules(doc.sheet, s === 'selection' ? doc.selection.ranges : undefined).map((e) => ({
    key: `r${rowKey++}`,
    id: e.id,
    rule: e.rule,
    raw: e.raw,
    ranges: e.ranges,
    text: e.ranges.map((r) => rangeText(r, doc.sheet)).join(','),
    stopIfTrue: e.stopIfTrue,
  }));
};
const scope = ref<Scope>('selection');
const rows = shallowRef<Row[]>(load('selection'));
const sel = ref<string | null>(rows.value[0]?.key ?? null);
const editor = shallowRef<{ row: Row | null } | null>(null);
const dirty = ref(false);
const index = computed(() => rows.value.findIndex((r) => r.key === sel.value));

const apply = (): boolean => {
  const ctx = ui.value;
  const { doc } = ctx;
  const list = [];
  for (const r of rows.value) {
    const ranges = r.text
      .split(',')
      .map((t) => parseRangeText(doc, t)?.range)
      .filter((x): x is Range => !!x);
    if (!ranges.length) {
      ctx.notify(`"${r.text}" is not a valid reference.`);
      return false;
    }
    list.push({ id: r.id, rule: r.rule, raw: r.raw, ranges, stopIfTrue: r.stopIfTrue });
  }
  commit(ctx);
  doc.applyRuleList(list, scope.value === 'selection' ? doc.selection.ranges : null);
  rows.value = load(scope.value);
  dirty.value = false;
  return true;
};
const ok = () => (dirty.value ? apply() : true);

const move = (dir: -1 | 1) => {
  const i = index.value;
  if (i < 0) return;
  const j = i + dir;
  if (j < 0 || j >= rows.value.length) return;
  const next = rows.value.slice();
  [next[i], next[j]] = [next[j], next[i]];
  rows.value = next;
  dirty.value = true;
};
const pickScope = (e: Event) => {
  const s = (e.target as HTMLSelectElement).value as Scope;
  scope.value = s;
  const next = load(s);
  rows.value = next;
  sel.value = next[0]?.key ?? null;
};
const remove = () => {
  const list = rows.value;
  const i = index.value;
  rows.value = list.filter((r) => r.key !== sel.value);
  sel.value = list[i + 1]?.key ?? list[i - 1]?.key ?? null;
  dirty.value = true;
};
const duplicate = () => {
  const list = rows.value;
  const i = index.value;
  const copy = { ...list[i], key: `r${rowKey++}`, id: undefined };
  rows.value = [...list.slice(0, i + 1), copy, ...list.slice(i + 1)];
  dirty.value = true;
};
const update = (key: string, patch: Partial<Row>) => {
  rows.value = rows.value.map((x) => (x.key === key ? { ...x, ...patch } : x));
  dirty.value = true;
};
const describe = (r: Row) => describeRule(r.rule, r.raw ?? ({ type: 'element', name: 'cfRule', attrs: {}, children: [] } as XmlElement));
const save = (rule: CfRule) => {
  const row = editor.value?.row;
  if (row) rows.value = rows.value.map((x) => (x.key === row.key ? { ...x, rule, raw: undefined } : x));
  else {
    const { doc } = ui.value;
    const added: Row = { key: `r${rowKey++}`, rule, ranges: doc.selection.ranges, text: doc.selection.ranges.map((r) => rangeText(r, doc.sheet)).join(','), stopIfTrue: false };
    rows.value = [added, ...rows.value];
    sel.value = added.key;
  }
  dirty.value = true;
};
</script>

<template>
  <Dialog title="Conditional Formatting Rules Manager" :on-ok="ok" :width="720" test-id="cf-manager" @close="emit('close')">
    <div class="sh-row">
      <span>Show formatting rules for:</span>
      <select :value="scope" data-testid="cf-manager-scope" @change="pickScope">
        <option value="selection">Current Selection</option>
        <option value="sheet">This Worksheet</option>
      </select>
    </div>
    <div class="sh-row wrap sh-toolbar-row">
      <button type="button" class="sh-btn" data-testid="cf-manager-new" @click="editor = { row: null }">New Rule...</button>
      <button type="button" class="sh-btn" :disabled="index < 0 || !rows[index]?.rule" @click="editor = { row: rows[index] }">Edit Rule...</button>
      <button type="button" class="sh-btn" :disabled="index < 0" data-testid="cf-manager-delete" @click="remove">Delete Rule</button>
      <button type="button" class="sh-btn" :disabled="index < 0" @click="duplicate">Duplicate Rule</button>
      <button type="button" class="sh-btn icon" aria-label="Move up" :disabled="index <= 0" @click="move(-1)">▲</button>
      <button type="button" class="sh-btn icon" aria-label="Move down" :disabled="index < 0 || index >= rows.length - 1" @click="move(1)">▼</button>
    </div>
    <div class="sh-table" role="table" data-testid="cf-manager-rules">
      <div class="sh-table-head" role="row">
        <span>Rule (applied in order shown)</span>
        <span>Format</span>
        <span>Applies to</span>
        <span>Stop If True</span>
      </div>
      <div class="sh-table-body">
        <div v-for="r in rows" :key="r.key" role="row" :class="`sh-table-row${r.key === sel ? ' selected' : ''}`" @click="sel = r.key" @dblclick="r.rule && (editor = { row: r })">
          <span>{{ describe(r) }}</span>
          <span><RulePreview :rule="r.rule" /></span>
          <span>
            <input :value="r.text" aria-label="Applies to" @input="update(r.key, { text: ($event.target as HTMLInputElement).value })" />
          </span>
          <span>
            <input
              v-checked="r.stopIfTrue"
              type="checkbox"
              aria-label="Stop If True"
              :disabled="r.rule?.type === 'colorScale' || r.rule?.type === 'dataBar' || r.rule?.type === 'iconSet'"
              @change="update(r.key, { stopIfTrue: ($event.target as HTMLInputElement).checked })"
            />
          </span>
        </div>
        <div v-if="!rows.length" class="sh-table-empty">No rules.</div>
      </div>
    </div>
    <template #footer-left>
      <button type="button" class="sh-btn" :disabled="!dirty" @click="apply()">Apply</button>
    </template>
  </Dialog>
  <RuleEditorDialog v-if="editor" :title="editor.row ? 'Edit Formatting Rule' : 'New Formatting Rule'" :initial="editor.row?.rule ?? null" @close="editor = null" @save="save" />
</template>
