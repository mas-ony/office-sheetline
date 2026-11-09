<script setup lang="ts">
/** Unhide: one of the hidden sheets. */
import { computed, ref } from 'vue';
import { useUI } from '../../context';
import Dialog from '../Dialog.vue';
import ListBox from '../ListBox.vue';

const emit = defineEmits<{ close: [] }>();

const ui = useUI();
const hidden = computed(() =>
  ui.value.doc.wb.sheets
    .map((s, i) => ({ s, i }))
    .filter((x) => x.s.state === 'hidden')
    .map((x) => x.i),
);
const sel = ref<number | null>(hidden.value[0] ?? null);
const keyOf = (i: number) => String(i);
const ok = () => {
  if (sel.value !== null) ui.value.doc.setSheetHidden(sel.value, false);
};
</script>

<template>
  <Dialog title="Unhide" :width="300" test-id="unhide-sheet" :ok-disabled="sel === null" :on-ok="ok" @close="emit('close')">
    <div class="sh-label">Unhide sheet:</div>
    <ListBox v-slot="{ item }" :items="hidden" :value="sel" :key-of="keyOf" :height="160" test-id="unhide-list" label="Hidden sheets" @change="sel = $event">{{ ui.doc.wb.sheets[item].name }}</ListBox>
  </Dialog>
</template>
