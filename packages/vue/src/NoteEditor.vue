<script setup lang="ts">
/** The note of a cell, edited next to it; saved when closed. */
import { cellName } from '@sheetline/core/views';
import { ref } from 'vue';
import Anchored from './Anchored.vue';
import { useUI } from './context';
import { vFocus } from './directives';

const props = defineProps<{ r: number; c: number; rect: DOMRect }>();
const emit = defineEmits<{ close: [] }>();

const ui = useUI();
const existing = ui.value.doc.sheet.commentAt(props.r, props.c);
const text = ref(existing?.kind === 'note' ? existing.comments.map((x) => x.text).join('\n') : '');
const save = () => {
  ui.value.doc.setNote(props.r, props.c, text.value.trim());
  emit('close');
};
const key = (e: KeyboardEvent) => {
  if (e.key === 'Escape') {
    e.stopPropagation();
    emit('close');
  }
};
</script>

<template>
  <Anchored :rect="rect" box-class="sh-note" test-id="note-editor" :label="`Note ${cellName(r, c)}`" @close="save">
    <div class="sh-note-author">{{ ui.doc.author }}:</div>
    <textarea v-focus :value="text" aria-label="Note" data-testid="note-input" @input="text = ($event.target as HTMLTextAreaElement).value" @keydown="key" />
    <div class="sh-note-actions">
      <button type="button" class="sh-btn-primary" data-testid="note-save" @click="save">OK</button>
    </div>
  </Anchored>
</template>
