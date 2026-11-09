<script setup lang="ts">
/** The name of the table at the selection (Table Design). */
import { type Ctx, currentTable, renameTableTo } from '@sheetline/core/views';
import { computed, ref, watch } from 'vue';

const props = defineProps<{ ctx: Ctx }>();

// A new object on every change: the table is renamed in place.
const view = computed(() => ({ t: currentTable(props.ctx) }));
const draft = ref(view.value.t?.name ?? '');
watch(
  () => view.value.t?.name,
  (name) => (draft.value = name ?? ''),
);
const commit = () => {
  const { t } = view.value;
  if (t && draft.value !== t.name && !renameTableTo(props.ctx, t.id, draft.value)) draft.value = t.name;
};
const key = (e: KeyboardEvent) => {
  if (e.key === 'Enter') (e.target as HTMLInputElement).blur();
  if (e.key === 'Escape' && view.value.t) draft.value = view.value.t.name;
};
</script>

<template>
  <label v-if="view.t" class="sh-field-inline">
    <span>Table Name:</span>
    <input aria-label="Table Name" data-testid="table-name" :value="draft" @input="draft = ($event.target as HTMLInputElement).value" @blur="commit" @keydown="key" />
  </label>
</template>
