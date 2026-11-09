<script setup lang="ts">
/** Symbol: a character to insert into the cell being edited (or to start editing with). */
import { SYMBOL_CHARS } from '@sheetline/core/views';
import { ref } from 'vue';
import { useUI } from '../../context';
import Dialog from '../Dialog.vue';

const emit = defineEmits<{ close: [] }>();

const ui = useUI();
const pick = ref<string | null>(null);
const insert = (ch: string) => {
  const { grid, doc } = ui.value;
  if (!grid) return;
  if (grid.isEditing()) {
    const { text, caret } = grid.editState;
    grid.replaceEditText(text.slice(0, caret) + ch + text.slice(caret), caret + ch.length);
  } else grid.startEdit(`${doc.editText(doc.selection.active.r, doc.selection.active.c)}${ch}`, 'edit');
};
const ok = () => {
  if (pick.value) insert(pick.value);
};
const activate = (ch: string) => {
  insert(ch);
  emit('close');
};
const code = (ch: string) => `U+${ch.codePointAt(0)!.toString(16).toUpperCase().padStart(4, '0')}`;
</script>

<template>
  <Dialog title="Symbol" :width="420" ok="Insert" :blocking="false" test-id="symbol-dialog" :on-ok="ok" @close="emit('close')">
    <div class="sh-symbols" role="listbox" aria-label="Symbols">
      <button v-for="ch in SYMBOL_CHARS" :key="ch" type="button" role="option" :aria-selected="pick === ch" :class="pick === ch ? 'selected' : ''" @click="pick = ch" @dblclick="activate(ch)">{{ ch }}</button>
    </div>
    <div class="sh-hint">Unicode: {{ pick ? code(pick) : '—' }}</div>
  </Dialog>
</template>
