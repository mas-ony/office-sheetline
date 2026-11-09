<script setup lang="ts">
/**
 * Status bar: the cell mode (Ready, Enter, Edit), the summary of the selection (Average, Count,
 * Sum; more through the right-click menu) and the zoom slider.
 */
import { cmd, type MenuItem, selectionStats, sliderToZoom, STATUS_STAT_LABELS, type StatusStat, zoomToSlider } from '@sheetline/core/views';
import { computed, shallowRef } from 'vue';
import { usePersistent, useUI } from './context';
import Icon from './Icon.vue';
import Menu from './Menu.vue';

const STATS = Object.keys(STATUS_STAT_LABELS) as StatusStat[];
/** ": " as a text node of its own: the stats then read as three text nodes, as React renders them, which Chrome draws slightly differently from one. */
const Colon = () => ': ';

const ui = useUI();
const shown = usePersistent<StatusStat[]>('status-stats', ['average', 'count', 'sum']);
const menu = shallowRef<{ x: number; y: number } | null>(null);
const zoom = computed(() => ui.value.doc.sheet.zoom);

// Computed again only for a new document version or selection kind (the selection can be large).
let memo: { doc: unknown; version: number; multiCell: boolean; value: ReturnType<typeof selectionStats> } | null = null;
const stats = computed(() => {
  const { doc, state } = ui.value;
  if (memo && memo.doc === doc && memo.version === state.version && memo.multiCell === state.multiCell) return memo.value;
  const value = state.multiCell ? selectionStats(doc) : null;
  memo = { doc, version: state.version, multiCell: state.multiCell, value };
  return value;
});

const mode = computed(() => {
  const { edit } = ui.value;
  return edit.active ? (edit.mode === 'edit' ? 'Edit' : 'Enter') : 'Ready';
});
const menuItems = computed<MenuItem[]>(() =>
  STATS.map((k) => ({
    title: STATUS_STAT_LABELS[k],
    checked: shown.value.includes(k),
    keepOpen: true,
    onClick: () => (shown.value = shown.value.includes(k) ? shown.value.filter((x) => x !== k) : [...shown.value, k]),
  })),
);
const openMenu = (e: MouseEvent) => {
  e.preventDefault();
  menu.value = { x: e.clientX, y: e.clientY };
};
const zoomBy = (d: number) => cmd.zoomTo(ui.value, d < 0 ? Math.max(0.1, Math.round((zoom.value - 0.1) * 10) / 10) : Math.min(4, Math.round((zoom.value + 0.1) * 10) / 10));
const slide = (e: Event) => cmd.zoomTo(ui.value, sliderToZoom(Number((e.target as HTMLInputElement).value)));
</script>

<template>
  <div class="sh-statusbar" data-testid="status-bar" @contextmenu="openMenu">
    <span class="sh-status-mode" data-testid="status-mode">{{ mode }}</span>
    <button v-if="ui.doc.calcPending" type="button" class="sh-status-calc" @click="ui.doc.calculateNow()">Calculate</button>
    <span v-if="ui.state.sheetProtected" class="sh-status-flag" title="This sheet is protected"><Icon name="lock" /></span>
    <span class="sh-status-grow" />
    <span v-if="stats" class="sh-status-stats" data-testid="status-stats">
      <template v-for="k in STATS" :key="k">
        <span v-if="shown.includes(k) && stats[k] !== null" :data-testid="`status-${k}`">{{ STATUS_STAT_LABELS[k] }}<Colon />{{ stats[k] }}</span>
      </template>
    </span>
    <span v-if="!ui.phone" class="sh-zoom">
      <button type="button" aria-label="Zoom out" data-testid="zoom-out" @click="zoomBy(-1)">
        <Icon name="remove" />
      </button>
      <input type="range" :min="0" :max="100" :step="1" aria-label="Zoom" data-testid="zoom-slider" :value="zoomToSlider(zoom)" @input="slide" />
      <button type="button" aria-label="Zoom in" data-testid="zoom-in" @click="zoomBy(1)">
        <Icon name="add" />
      </button>
    </span>
    <button type="button" class="sh-zoom-value" aria-label="Zoom level" data-testid="zoom-value" @click="ui.openDialog({ kind: 'zoom' })">{{ Math.round(zoom * 100) }}%</button>
    <Menu v-if="menu" :items="menuItems" :at="menu" title="Customize Status Bar" @close="menu = null" />
  </div>
</template>
