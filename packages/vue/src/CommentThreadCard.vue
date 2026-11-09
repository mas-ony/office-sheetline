<script setup lang="ts">
/** A threaded comment card: the thread, replies and the reply box (or a new comment box). */
import { cellName, commentStamp, type MenuItem } from '@sheetline/core/views';
import { computed, ref, shallowRef } from 'vue';
import Avatar from './Avatar.vue';
import { useUI } from './context';
import { vFocus } from './directives';
import Icon from './Icon.vue';
import Menu from './Menu.vue';

const props = defineProps<{ r: number; c: number; startNew?: boolean; inPanel?: boolean }>();
const emit = defineEmits<{ done: [] }>();

const ui = useUI();
// A new object on every change: the thread is updated in place.
const view = computed(() => {
  const { doc, state } = ui.value;
  const thread = doc.sheet.commentAt(props.r, props.c);
  return { doc, thread, ro: state.readOnly, isNew: !thread || thread.kind !== 'threaded' };
});
const text = ref('');
const editing = ref<{ id: string; text: string } | null>(null);
const menu = shallowRef<{ id: string; at: DOMRect } | null>(null);

const post = () => {
  const { doc, thread } = view.value;
  if (!text.value.trim()) return;
  if (thread?.kind === 'threaded') doc.reply(thread.comments[0].id, text.value.trim());
  else doc.addComment(props.r, props.c, text.value.trim());
  text.value = '';
  if (props.startNew && !thread) emit('done');
};
const key = (e: KeyboardEvent) => {
  if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) {
    e.preventDefault();
    post();
  }
};
const save = (id: string) => {
  view.value.doc.editComment(id, editing.value!.text);
  editing.value = null;
};
const cancel = () => {
  if (view.value.isNew) emit('done');
  else text.value = '';
};
const openMenu = (id: string, e: MouseEvent) => (menu.value = { id, at: (e.currentTarget as HTMLElement).getBoundingClientRect() });
const menuItems = computed<MenuItem[]>(() => {
  const { doc, thread } = view.value;
  const m = menu.value;
  if (!m || !thread) return [];
  return [
    { title: 'Edit Comment', icon: 'edit', onClick: () => (editing.value = { id: m.id, text: thread.comments.find((x) => x.id === m.id)?.text ?? '' }) },
    ...(m.id === thread.comments[0].id
      ? [
          { title: thread.resolved ? 'Reopen Thread' : 'Resolve Thread', icon: 'check_circle' as const, onClick: () => doc.setResolved(m.id, !thread.resolved) },
          { title: 'Delete Thread', icon: 'delete' as const, onClick: () => doc.deleteComment(m.id) },
        ]
      : [{ title: 'Delete Reply', icon: 'delete' as const, onClick: () => doc.deleteComment(m.id) }]),
  ];
});
</script>

<template>
  <div :class="`sh-thread${view.thread?.resolved ? ' resolved' : ''}`" data-testid="comment-thread">
    <div v-if="view.isNew" class="sh-thread-head">
      <Avatar :name="view.doc.author" />
      <span class="who">{{ view.doc.author }}</span>
      <span class="ref">{{ cellName(r, c) }}</span>
    </div>
    <template v-else>
      <div v-for="(cm, i) in view.thread!.comments" :key="cm.id" :class="`sh-comment${i ? ' reply' : ''}`" data-testid="comment">
        <div class="sh-thread-head">
          <Avatar :name="cm.author || 'Guest'" />
          <span class="who">{{ cm.author || 'Guest' }}</span>
          <span v-if="i === 0" class="ref">{{ cellName(r, c) }}</span>
          <button v-if="!view.ro" type="button" class="sh-icon-btn" aria-label="Thread actions" data-testid="comment-options" @click="openMenu(cm.id, $event)">
            <Icon :name="i === 0 ? 'settings' : 'more_horiz'" />
          </button>
        </div>
        <div v-if="editing?.id === cm.id" class="sh-thread-edit">
          <textarea class="sh-textarea" :value="editing.text" @input="editing = { ...editing!, text: ($event.target as HTMLTextAreaElement).value }" />
          <div class="row">
            <button type="button" class="sh-btn-primary" @click="save(cm.id)">Save</button>
            <button type="button" class="sh-btn-secondary" @click="editing = null">Cancel</button>
          </div>
        </div>
        <div v-else class="body">{{ cm.text }}</div>
        <div class="when">{{ commentStamp(cm.created) }}{{ i === 0 && view.thread!.resolved ? ' · Resolved' : '' }}</div>
      </div>
    </template>
    <div v-if="!view.ro && !view.thread?.resolved" class="sh-thread-input">
      <textarea
        v-focus="!inPanel"
        class="sh-textarea"
        :value="text"
        :placeholder="view.isNew ? 'Start a conversation' : '@mention or reply'"
        :aria-label="view.isNew ? 'Comment' : 'Reply'"
        data-testid="comment-input"
        :rows="view.isNew ? 3 : 1"
        @input="text = ($event.target as HTMLTextAreaElement).value"
        @keydown="key"
      />
      <div v-if="text || view.isNew" class="row">
        <button type="button" class="sh-send" aria-label="Post" title="Post (Ctrl+Enter)" :disabled="!text.trim()" data-testid="comment-save" @click="post">
          <Icon name="arrow_forward" />
        </button>
        <button type="button" class="sh-icon-btn" aria-label="Cancel" @click="cancel">
          <Icon name="close" />
        </button>
      </div>
    </div>
    <Menu v-if="menu && view.thread" :at="menu.at" :items="menuItems" @close="menu = null" />
  </div>
</template>
