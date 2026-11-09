<script setup lang="ts">
/** Move or Copy: the sheet to put the sheet (or its copy) before. */
import { commit } from '@sheetline/core/views';
import { computed, ref } from 'vue';
import { useUI } from '../../context';
import Check from '../Check.vue';
import Dialog from '../Dialog.vue';
import ListBox from '../ListBox.vue';

const props = defineProps<{ index: number }>();
const emit = defineEmits<{ close: [] }>();

const ui = useUI();
// A new object on every change: sheets are added and renamed in place.
const view = computed(() => {
  const sheets = ui.value.doc.wb.sheets;
  return { sheets, items: [...sheets.map((_, i) => i), sheets.length] };
});
const before = ref<number>(Math.min(props.index + 1, view.value.sheets.length));
const copy = ref(false);
const keyOf = (i: number) => String(i);

const ok = () => {
  const ctx = ui.value;
  const { doc } = ctx;
  const { index } = props;
  commit(ctx);
  if (copy.value) {
    const dup = doc.duplicateSheet(index);
    const from = dup ? doc.wb.sheets.indexOf(dup) : -1;
    // "before" indexes the list without the copy; the copy sits right after the original.
    if (from >= 0 && from !== before.value) doc.moveSheet(from, Math.min(doc.wb.sheets.length - 1, before.value));
    return true;
  }
  let to = before.value;
  if (to > index) to--;
  if (to !== index) doc.moveSheet(index, Math.max(0, Math.min(doc.wb.sheets.length - 1, to)));
  return true;
};
</script>

<template>
  <Dialog title="Move or Copy" :width="320" test-id="move-copy-sheet" :on-ok="ok" @close="emit('close')">
    <p class="sh-lead">Move selected sheets</p>
    <div class="sh-label">Before sheet:</div>
    <ListBox v-slot="{ item }" :items="view.items" :value="before" :key-of="keyOf" :height="180" test-id="move-copy-list" label="Before sheet" @change="before = $event">{{ item === view.sheets.length ? '(move to end)' : view.sheets[item].name }}</ListBox>
    <Check label="Create a copy" :checked="copy" test-id="move-copy-copy" @change="copy = $event" />
  </Dialog>
</template>
