<script setup lang="ts">
/**
 * The editor: ribbon (or one-row toolbar), formula bar, grid, sheet tabs and status bar, with the
 * File backstage, side panels (comments, selection, format), context menus, dialogs and the cell
 * pop-ups (comments, notes, filters).
 */
import { type SheetlineOptions, SheetlineUI, type DocumentInput, type FileActions, type Layout, rootClassOf, type SpreadsheetApi, type ThemeName } from '@sheetline/core/views';
import { onMounted, onUnmounted, type PropType, provide, ref, shallowRef, watch } from 'vue';
import { UI_KEY } from './context';
import Messages from './Messages.vue';
import WorkbookView from './WorkbookView.vue';

/** The complete editor. */
defineOptions({ name: 'SheetlineEditor' });

// The options of SheetlineOptions; its callbacks are the events below.
const props = defineProps({
  /** Opened at start, and again whenever it changes. */
  document: { type: null as unknown as PropType<DocumentInput>, default: undefined },
  /** Workbooks open view-only. */
  readOnly: { type: Boolean, default: undefined },
  /** Ribbon or the one-row toolbar (desktop); by default the one last chosen, else the ribbon. */
  layout: { type: String as PropType<Layout>, default: undefined },
  theme: { type: String as PropType<ThemeName>, default: undefined },
  /** Shown on comments and notes, and written as the author of new workbooks. */
  author: { type: String, default: undefined },
  /** Keeps the layout, theme, settings and recent colours in localStorage under `<storageKey>-…`. */
  storageKey: { type: String, default: undefined },
  /** Keeps unsaved changes in an IndexedDB database of this name and offers to recover them. */
  autosave: { type: String, default: undefined },
  /** With `autosave`: offer recovery when the editor starts without a document (default true). */
  offerRecovery: { type: Boolean, default: undefined },
  /** Replace some of the built-in file actions (New, Open, Save, Export, Print…). */
  files: { type: Object as PropType<Partial<FileActions>>, default: undefined },
  /** Handle the editor's shortcuts (Ctrl+1, Alt+Q, F9…) anywhere on the page, not only inside the editor. */
  globalShortcuts: { type: Boolean, default: undefined },
});

const emit = defineEmits<{
  /** The editor is ready and the initial document has opened. */
  ready: [api: SpreadsheetApi];
  /** The reader switched between the ribbon and the toolbar (`v-model:layout`). */
  'update:layout': [layout: Layout];
  /** The theme in use when the editor starts, and whenever the reader switches it (`v-model:theme`). */
  'update:theme': [theme: ThemeName];
}>();

function options(): SheetlineOptions {
  return {
    ...props,
    onReady: (api) => emit('ready', api),
    onLayoutChange: (layout) => emit('update:layout', layout),
    onThemeChange: (theme) => emit('update:theme', theme),
  };
}

// The first props create the interface; later ones are passed on.
const ui = new SheetlineUI(options());
watch(() => ({ ...props }), () => ui.setOptions(options()));

const state = shallowRef(ui.getState());
const ctx = shallowRef(ui.getCtx());
const off = ui.subscribe(() => {
  state.value = ui.getState();
  ctx.value = ui.getCtx();
});
provide(UI_KEY, ctx);

const root = ref<HTMLDivElement | null>(null);
onMounted(() => ui.attach(root.value!));
onUnmounted(() => {
  off();
  ui.detach();
});
</script>

<template>
  <div ref="root" :class="rootClassOf(state)" data-testid="editor" :data-layout="state.layout" :data-theme="state.theme">
    <WorkbookView :key="state.docKey" :state="state" />
    <Messages :state="state" />
  </div>
</template>
