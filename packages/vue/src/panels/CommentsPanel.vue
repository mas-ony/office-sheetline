<script setup lang="ts">
/** Threaded Comments of the sheet. */
import { cellName, cmd } from '@sheetline/core/views';
import { computed, ref } from 'vue';
import CommentThreadCard from '../CommentThreadCard.vue';
import { useUI } from '../context';
import Icon from '../Icon.vue';
import SidePanel from './SidePanel.vue';

const ui = useUI();
const filter = ref<'all' | 'open' | 'resolved'>('all');
const threads = computed(() =>
  ui.value.doc.sheet.comments
    .filter((t) => t.kind === 'threaded')
    .filter((t) => (filter.value === 'all' ? true : filter.value === 'open' ? !t.resolved : t.resolved))
    .sort((a, b) => a.r - b.r || a.c - b.c),
);
const go = (r: number, c: number) => {
  ui.value.doc.selectCell(r, c);
  ui.value.grid?.scrollIntoView(r, c);
};
</script>

<template>
  <SidePanel title="Comments" test-id="comments-panel">
    <template #actions>
      <button type="button" class="sh-btn small" :disabled="ui.state.readOnly" data-testid="comments-panel-new" @click="cmd.newComment(ui)">
        <Icon name="add_comment" /> New</button>
    </template>
    <div class="sh-row">
      <span class="sh-hint grow">{{ threads.length }} items</span>
      <select :value="filter" aria-label="Filter comments" @change="filter = ($event.target as HTMLSelectElement).value as typeof filter">
        <option value="all">All</option>
        <option value="open">Active</option>
        <option value="resolved">Resolved</option>
      </select>
    </div>
    <div v-for="t in threads" :key="t.id" :class="`sh-panel-thread${ui.state.activeRef === cellName(t.r, t.c) ? ' active' : ''}`" @click="go(t.r, t.c)">
      <div class="sh-panel-thread-ref">{{ cellName(t.r, t.c) }}</div>
      <CommentThreadCard :r="t.r" :c="t.c" in-panel />
    </div>
    <p v-if="!threads.length" class="sh-hint">No comments yet. Select a cell and choose New to start a conversation.</p>
  </SidePanel>
</template>
