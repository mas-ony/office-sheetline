<script setup lang="ts">
/**
 * Sheet tab strip under the grid: previous / next buttons, the list of all sheets, the tabs (click,
 * double click to rename, drag to reorder, right click or long press for the tab menu) and "+".
 */
import { cmd, type MenuItem } from '@sheetline/core/views';
import { computed, onMounted, ref, shallowRef, watch } from 'vue';
import { useUI } from './context';
import { vFocus } from './directives';
import Icon from './Icon.vue';
import Menu from './Menu.vue';

const emit = defineEmits<{ tabMenu: [index: number, at: { x: number; y: number }] }>();

const ui = useUI();
const editing = ref<number | null>(null);
const draft = ref('');
const drag = shallowRef<{ from: number; over: number; side: 'before' | 'after' } | null>(null);
const list = shallowRef<DOMRect | null>(null);
let press: { timer: number; index: number; x: number; y: number; moved: boolean; pointerId: number } | null = null;
const strip = ref<HTMLDivElement | null>(null);
const listBtn = ref<HTMLButtonElement | null>(null);
const locked = computed(() => ui.value.state.workbookProtected);

// Keep the active tab in view.
const showActive = () => strip.value?.querySelector<HTMLElement>('.sh-sheet-tab.active')?.scrollIntoView({ block: 'nearest', inline: 'nearest' });
onMounted(showActive);
watch(() => ui.value.state.activeIndex, showActive, { flush: 'post' });

const commitRename = () => {
  if (editing.value === null) return;
  const { doc } = ui.value;
  const index = editing.value;
  const name = draft.value.trim();
  editing.value = null;
  if (!name || name === doc.wb.sheets[index]?.name) return;
  const error = doc.renameSheet(index, name);
  if (error) ui.value.notify(error);
};

const step = (dir: 1 | -1) => {
  const { state, doc } = ui.value;
  const visible = state.sheets.filter((s) => s.state === 'visible').map((s) => s.index);
  const pos = visible.indexOf(state.activeIndex);
  const next = visible[pos + dir];
  if (next !== undefined) doc.setActiveSheet(next);
};

const down = (e: PointerEvent, index: number) => {
  if (e.button !== 0 || editing.value !== null) return;
  const timer = window.setTimeout(() => {
    if (press && !press.moved) {
      ui.value.doc.setActiveSheet(index);
      emit('tabMenu', index, { x: press.x, y: press.y });
      press = null;
    }
  }, 500);
  press = { timer, index, x: e.clientX, y: e.clientY, moved: false, pointerId: e.pointerId };
};

const move = (e: PointerEvent) => {
  const p = press;
  const d = drag.value;
  if (!p || p.pointerId !== e.pointerId) return;
  if (!p.moved && Math.hypot(e.clientX - p.x, e.clientY - p.y) > 8) {
    p.moved = true;
    window.clearTimeout(p.timer);
    if (e.pointerType === 'mouse' && !ui.value.state.readOnly && !locked.value) {
      (e.currentTarget as HTMLElement).setPointerCapture(e.pointerId);
      drag.value = { from: p.index, over: p.index, side: 'after' };
    }
  }
  if (d) {
    const el = document.elementFromPoint(e.clientX, e.clientY)?.closest<HTMLElement>('[data-tab-index]');
    if (el) {
      const over = Number(el.dataset.tabIndex);
      const rect = el.getBoundingClientRect();
      drag.value = { ...d, over, side: e.clientX < rect.left + rect.width / 2 ? 'before' : 'after' };
    }
  }
};

const up = (e: PointerEvent, index: number) => {
  const { doc, grid } = ui.value;
  const p = press;
  if (p) window.clearTimeout(p.timer);
  press = null;
  if (drag.value) {
    const { from, over, side } = drag.value;
    drag.value = null;
    let to = side === 'before' ? over : over + 1;
    if (from < to) to--;
    if (to !== from) doc.moveSheet(from, Math.max(0, Math.min(doc.wb.sheets.length - 1, to)));
    return;
  }
  if (p && !p.moved && e.button === 0) {
    if (grid?.isEditing() && !grid.editState.text.startsWith('=')) grid.commitEdit();
    doc.setActiveSheet(index);
  }
};

const cancel = () => {
  if (press) window.clearTimeout(press.timer);
  press = null;
  drag.value = null;
};

const rename = (index: number, name: string) => {
  if (ui.value.state.readOnly || locked.value) return;
  draft.value = name;
  editing.value = index;
};

const tabMenu = (e: MouseEvent, index: number) => {
  e.preventDefault();
  ui.value.doc.setActiveSheet(index);
  emit('tabMenu', index, { x: e.clientX, y: e.clientY });
};

const renameKey = (e: KeyboardEvent) => {
  if (e.key === 'Enter') commitRename();
  else if (e.key === 'Escape') editing.value = null;
};

const dropClass = (index: number) => {
  const d = drag.value;
  return d && d.over === index && d.from !== index ? ` drop-${d.side}` : '';
};

const listItems = computed<MenuItem[]>(() => {
  const { state, doc } = ui.value;
  return state.sheets.filter((s) => s.state === 'visible').map((s) => ({ title: s.name, checked: s.index === state.activeIndex, onClick: () => doc.setActiveSheet(s.index) }));
});
</script>

<template>
  <div class="sh-sheetbar" data-testid="sheet-tabs">
    <button type="button" class="sh-sheet-nav" aria-label="Previous sheet" data-testid="tab-prev" @click="step(-1)">
      <Icon name="chevron_left" />
    </button>
    <button type="button" class="sh-sheet-nav" aria-label="Next sheet" data-testid="tab-next" @click="step(1)">
      <Icon name="chevron_right" />
    </button>
    <button ref="listBtn" type="button" class="sh-sheet-nav" aria-label="All sheets" title="All sheets" data-testid="tab-list" @click="list = listBtn!.getBoundingClientRect()">
      <Icon name="menu" />
    </button>
    <div ref="strip" class="sh-sheet-tabs" role="tablist" aria-label="Worksheets">
      <template v-for="s in ui.state.sheets" :key="s.index">
        <template v-if="s.state === 'visible'">
          <input
            v-if="editing === s.index"
            v-focus
            class="sh-sheet-rename"
            aria-label="Sheet name"
            data-testid="tab-rename-input"
            :value="draft"
            @focus="($event.target as HTMLInputElement).select()"
            @input="draft = ($event.target as HTMLInputElement).value"
            @blur="commitRename"
            @keydown="renameKey"
          />
          <button
            v-else
            type="button"
            role="tab"
            :aria-selected="s.index === ui.state.activeIndex"
            :data-tab-index="s.index"
            :data-testid="`tab-${s.name}`"
            :class="`sh-sheet-tab${s.index === ui.state.activeIndex ? ' active' : ''}${dropClass(s.index)}`"
            :style="s.color ? { '--tab-color': s.color } : undefined"
            @pointerdown="down($event, s.index)"
            @pointermove="move"
            @pointerup="up($event, s.index)"
            @pointercancel="cancel"
            @dblclick="rename(s.index, s.name)"
            @contextmenu="tabMenu($event, s.index)"
          >{{ s.name }}</button>
        </template>
      </template>
    </div>
    <button
      v-if="!ui.state.readOnly && ui.view.newTab"
      type="button"
      class="sh-sheet-add"
      aria-label="New sheet"
      title="New sheet (Shift+F11)"
      data-testid="tab-add"
      :disabled="locked"
      @click="cmd.newWorksheet(ui)"
    >
      <Icon name="add" />
    </button>
    <div class="sh-sheetbar-rest" />
    <Menu v-if="list" :at="list" :sheet="ui.phone" title="Sheets" test-id="tab-list-menu" :items="listItems" @close="list = null" />
  </div>
</template>
