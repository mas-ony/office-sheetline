<script setup lang="ts">
/** Scenario Manager: add, edit (name, changing cells, values), delete, show and summarise scenarios. */
import { cellName, parseScenarioCells, rangeText, type Scenario, scenarioCellsText } from '@sheetline/core/views';
import { computed, ref, shallowRef } from 'vue';
import { useUI } from '../../context';
import Dialog from '../Dialog.vue';
import Field from '../Field.vue';
import ListBox from '../ListBox.vue';
import RangeField from '../RangeField.vue';
import ScenarioSummary from './ScenarioSummary.vue';

const emit = defineEmits<{ close: [] }>();

interface Edit {
  index: number | null;
  name: string;
  cells: string;
  comment: string;
  step: 1 | 2;
  values: string[];
}

const ui = useUI();
const list = shallowRef<Scenario[]>(ui.value.doc.sheet.scenarios?.list.map((s) => ({ ...s, cells: s.cells.map((c) => ({ ...c })), values: [...s.values] })) ?? []);
const sel = ref<number | null>(list.value.length ? 0 : null);
const edit = shallowRef<Edit | null>(null);
const summary = ref(false);
const indexes = computed(() => list.value.map((_, i) => i));
const keyOf = (i: number) => String(i);
const selected = computed(() => (sel.value !== null ? list.value[sel.value] : undefined));

const save = (next: Scenario[]) => {
  list.value = next;
  const { doc } = ui.value;
  doc.setScenarios({ list: next, current: doc.sheet.scenarios?.current });
};
const parseCells = (text: string) => {
  const { doc } = ui.value;
  return parseScenarioCells(doc, doc.sheet, text);
};
const show = (i: number) => ui.value.doc.showScenario(i);
const patch = (p: Partial<Edit>) => {
  if (edit.value) edit.value = { ...edit.value, ...p };
};

const add = () => {
  const { doc } = ui.value;
  edit.value = { index: null, name: '', cells: doc.selection.ranges.map((r) => rangeText(r, doc.sheet)).join(','), comment: `Created by ${doc.author} on ${new Date().toLocaleDateString()}`, step: 1, values: [] };
};
const remove = () => {
  const i = sel.value;
  if (i === null) return;
  const n = list.value.length;
  save(list.value.filter((_, k) => k !== i));
  sel.value = n > 1 ? 0 : null;
};
const editSelected = () => {
  const i = sel.value;
  if (i === null) return;
  const s = list.value[i];
  edit.value = { index: i, name: s.name, cells: scenarioCellsText(s), comment: s.comment ?? '', step: 1, values: s.values };
};
const toValues = () => {
  const ctx = ui.value;
  const e = edit.value;
  if (!e) return false;
  const cells = parseCells(e.cells);
  if (!e.name.trim()) {
    ctx.notify('Enter a scenario name.');
    return false;
  }
  if (!cells) {
    ctx.notify('Changing cells must be references on this sheet (up to 32 cells).');
    return false;
  }
  const values = cells.map((c, i) => e.values[i] ?? ctx.doc.editText(c.r, c.c));
  edit.value = { ...e, step: 2, values };
  return false;
};
const setValue = (i: number, v: string) => {
  const e = edit.value;
  if (!e) return;
  const values = e.values.slice();
  values[i] = v;
  edit.value = { ...e, values };
};
const finish = () => {
  const e = edit.value;
  if (!e) return false;
  const cells = parseCells(e.cells)!;
  const sc: Scenario = { name: e.name.trim(), cells, values: e.values, comment: e.comment || undefined, user: ui.value.doc.author, locked: true };
  const next = e.index === null ? [...list.value, sc] : list.value.map((x, i) => (i === e.index ? sc : x));
  save(next);
  sel.value = e.index ?? next.length - 1;
  edit.value = null;
  return false;
};
</script>

<template>
  <Dialog title="Scenario Manager" close-only :width="480" test-id="scenario-manager" @close="emit('close')">
    <div class="sh-two-col wide-left">
      <div class="sh-stack">
        <div class="sh-label">Scenarios:</div>
        <ListBox v-slot="{ item }" :items="indexes" :value="sel" :key-of="keyOf" :height="150" test-id="scenario-list" label="Scenarios" :on-activate="show" @change="sel = $event">{{ list[item].name }}</ListBox>
        <p v-if="!list.length" class="sh-hint">No Scenarios defined. Choose Add to add scenarios.</p>
      </div>
      <div class="sh-stack">
        <button type="button" class="sh-btn" data-testid="scenario-add" @click="add">Add...</button>
        <button type="button" class="sh-btn" :disabled="sel === null" @click="remove">Delete</button>
        <button type="button" class="sh-btn" :disabled="sel === null" @click="editSelected">Edit...</button>
        <button type="button" class="sh-btn" :disabled="!list.length" @click="summary = true">Summary...</button>
        <button type="button" class="sh-btn primary" :disabled="sel === null" data-testid="scenario-show" @click="sel !== null && show(sel)">Show</button>
      </div>
    </div>
    <div v-if="selected" class="sh-stack">
      <div class="sh-kv"><span>Changing cells:</span><b>{{ scenarioCellsText(selected) }}</b></div>
      <div class="sh-kv"><span>Comment:</span><span>{{ selected.comment }}</span></div>
    </div>
  </Dialog>
  <Dialog v-if="edit && edit.step === 1" :title="edit.index === null ? 'Add Scenario' : 'Edit Scenario'" :width="400" test-id="scenario-edit" :on-ok="toValues" @close="edit = null">
    <Field label="Scenario name:" wide>
      <input :value="edit.name" data-testid="scenario-name" @input="patch({ name: ($event.target as HTMLInputElement).value })" />
    </Field>
    <Field label="Changing cells:" wide>
      <RangeField :value="edit.cells" label="Changing cells" test-id="scenario-cells" multi @change="patch({ cells: $event })" />
    </Field>
    <Field label="Comment:" wide>
      <textarea :rows="3" :value="edit.comment" @input="patch({ comment: ($event.target as HTMLTextAreaElement).value })" />
    </Field>
  </Dialog>
  <Dialog v-if="edit && edit.step === 2" title="Scenario Values" :width="360" test-id="scenario-values" :on-ok="finish" @close="edit = null">
    <p class="sh-lead">Enter values for each of the changing cells.</p>
    <Field v-for="(c, i) in parseCells(edit.cells) ?? []" :key="i" :label="`${i + 1}: ${cellName(c.r, c.c, true, true)}`">
      <input :value="edit.values[i] ?? ''" :data-testid="`scenario-value-${i}`" @input="setValue(i, ($event.target as HTMLInputElement).value)" />
    </Field>
  </Dialog>
  <ScenarioSummary v-if="summary" @close="summary = false" />
</template>
