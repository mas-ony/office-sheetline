<script setup lang="ts">
/** Insert Function: a function found by a search or picked from a category. */
import { CATEGORY_LABELS, functionCatalog, type FunctionInfo, insertFunction, signature } from '@sheetline/core/views';
import { computed, ref, shallowRef } from 'vue';
import { useUI } from '../../context';
import Dialog from '../Dialog.vue';
import Field from '../Field.vue';
import ListBox from '../ListBox.vue';

const props = defineProps<{ category?: string }>();
const emit = defineEmits<{ close: [] }>();

const ui = useUI();
const all = [...functionCatalog().values()];
const categories = [...new Set(all.map((f) => f.category))].sort((a, b) => (CATEGORY_LABELS[a] ?? a).localeCompare(CATEGORY_LABELS[b] ?? b));
const cat = ref(props.category ?? 'All');
const query = ref('');
const list = computed(() => {
  const q = query.value.trim().toUpperCase();
  return all.filter((f) => (cat.value === 'All' || f.category === cat.value) && (!q || f.name.includes(q) || f.description.toUpperCase().includes(q)));
});
const sel = shallowRef<FunctionInfo | null>(list.value[0] ?? null);
const current = computed(() => (sel.value && list.value.includes(sel.value) ? sel.value : (list.value[0] ?? null)));
const nameOf = (f: FunctionInfo) => f.name;

const ok = () => {
  const f = current.value;
  if (!f) return false;
  const ctx = ui.value;
  // Opening the dialog ended any edit but a formula, which takes the function at its caret.
  // The dialog closes first so the cell editor keeps the focus.
  window.setTimeout(() => insertFunction(ctx, f.name), 0);
  return true;
};
const activate = (f: FunctionInfo) => {
  const ctx = ui.value;
  emit('close');
  window.setTimeout(() => insertFunction(ctx, f.name), 0);
};
</script>

<template>
  <Dialog title="Insert Function" :width="480" test-id="insert-function" :ok-disabled="!current" :on-ok="ok" @close="emit('close')">
    <Field label="Search for a function:" wide>
      <input :value="query" placeholder="Type a brief description of what you want to do" data-testid="insert-function-search" @input="query = ($event.target as HTMLInputElement).value" />
    </Field>
    <Field label="Or select a category:">
      <select :value="cat" data-testid="insert-function-category" @change="cat = ($event.target as HTMLSelectElement).value">
        <option value="All">All</option>
        <option v-for="c in categories" :key="c" :value="c">{{ CATEGORY_LABELS[c] ?? c }}</option>
      </select>
    </Field>
    <div class="sh-label">Select a function:</div>
    <ListBox v-slot="{ item }" :items="list" :value="current" :key-of="nameOf" :height="200" test-id="insert-function-list" label="Functions" :on-activate="activate" @change="sel = $event">{{ item.name }}</ListBox>
    <div v-if="current" class="sh-fn-help" data-testid="insert-function-help">
      <b>{{ signature(current) }}</b>
      <p>{{ current.description || 'No description available.' }}</p>
    </div>
  </Dialog>
</template>
