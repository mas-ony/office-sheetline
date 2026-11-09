<script setup lang="ts">
/** Command Palette: every command and menu item, searched by name (Alt+Q). */
import { type PaletteEntry, paletteEntries } from '@sheetline/core/views';
import { computed, onMounted, ref } from 'vue';
import { usePersistent, useUI } from '../../context';
import Icon from '../../Icon.vue';
import Dialog from '../Dialog.vue';

const emit = defineEmits<{ close: [] }>();

const ui = useUI();
const query = ref('');
const recent = usePersistent<string[]>('recent-commands', []);
const index = ref(0);
const list = computed(() => paletteEntries(ui.value));
const results = computed(() => {
  const q = query.value.trim().toLowerCase();
  if (!q) {
    const rec = recent.value.map((id) => list.value.find((c) => c.id === id)).filter((c): c is PaletteEntry => !!c);
    // Suggestions: plain commands, not every menu item.
    return { recent: rec, rest: list.value.filter((c) => !recent.value.includes(c.id) && !c.id.includes('/')).slice(0, 40) };
  }
  const score = (c: PaletteEntry) => {
    const last = c.label.split(' › ').pop()!.toLowerCase();
    if (c.label.toLowerCase().startsWith(q) || last.startsWith(q)) return 0;
    if (c.hay.includes(q)) return 1;
    return q.split(/\s+/).every((w) => c.hay.includes(w)) ? 2 : 9;
  };
  return { recent: [] as PaletteEntry[], rest: list.value.map((c) => ({ c, s: score(c) })).filter((x) => x.s < 9).sort((a, b) => a.s - b.s).map((x) => x.c).slice(0, 60) };
});
const flat = computed(() => [...results.value.recent, ...results.value.rest]);
const input = ref<HTMLInputElement | null>(null);
onMounted(() => input.value?.focus());

const run = (c: PaletteEntry) => {
  emit('close');
  recent.value = [c.id, ...recent.value.filter((x) => x !== c.id)].slice(0, 8);
  window.setTimeout(() => c.run(), 0);
};
const typed = (e: Event) => {
  query.value = (e.target as HTMLInputElement).value;
  index.value = 0;
};
const key = (e: KeyboardEvent) => {
  if (e.key === 'ArrowDown') {
    e.preventDefault();
    index.value = Math.min(flat.value.length - 1, index.value + 1);
  } else if (e.key === 'ArrowUp') {
    e.preventDefault();
    index.value = Math.max(0, index.value - 1);
  } else if (e.key === 'Enter' && flat.value[index.value]) {
    e.preventDefault();
    run(flat.value[index.value]);
  }
};
</script>

<template>
  <Dialog title="Command Palette" close-only :width="520" test-id="command-palette" @close="emit('close')">
    <input ref="input" class="sh-palette-search" placeholder="Search for actions, text, and more" :value="query" data-testid="palette-search" @input="typed" @keydown="key" />
    <div class="sh-palette-list" role="listbox" aria-label="Commands">
      <div v-if="results.recent.length > 0" class="sh-panel-title">Recently Used Actions</div>
      <div
        v-for="(c, i) in results.recent"
        :key="c.id"
        role="option"
        :aria-selected="i === index"
        :class="`sh-palette-item${i === index ? ' selected' : ''}`"
        :data-testid="`palette-${c.id}`"
        @mouseenter="index = i"
        @click="run(c)"
      >
        <Icon v-if="c.icon" :name="c.icon" />
        <span v-else class="sh-icon-space" />
        <span class="grow">{{ c.label }}</span>
        <kbd v-if="c.shortcut">{{ c.shortcut }}</kbd>
      </div>
      <div v-if="!query" class="sh-panel-title">Suggested Actions</div>
      <div
        v-for="(c, i) in results.rest"
        :key="c.id"
        role="option"
        :aria-selected="i + results.recent.length === index"
        :class="`sh-palette-item${i + results.recent.length === index ? ' selected' : ''}`"
        :data-testid="`palette-${c.id}`"
        @mouseenter="index = i + results.recent.length"
        @click="run(c)"
      >
        <Icon v-if="c.icon" :name="c.icon" />
        <span v-else class="sh-icon-space" />
        <span class="grow">{{ c.label }}</span>
        <kbd v-if="c.shortcut">{{ c.shortcut }}</kbd>
      </div>
      <div v-if="!flat.length" class="sh-hint">No matching actions.</div>
    </div>
  </Dialog>
</template>
