<script setup lang="ts">
/** The Selection Pane: the sheet's objects, their order, visibility and names. */
import { computed, ref } from 'vue';
import { useUI } from '../context';
import { vFocus } from '../directives';
import Icon from '../Icon.vue';
import SidePanel from './SidePanel.vue';

const ui = useUI();
const rename = ref<{ id: string; name: string } | null>(null);
const objects = computed(() => {
  const sheet = ui.value.doc.sheet;
  return [
    ...sheet.charts.map((o) => ({ id: o.id, name: o.name, z: o.z, hidden: !!o.hidden, kind: 'chart' as const })),
    ...sheet.shapes.map((o) => ({ id: o.id, name: o.name, z: o.z, hidden: !!o.hidden, kind: 'shape' as const })),
    ...sheet.images.map((o, i) => ({ id: o.id, name: o.name, z: o.z ?? i, hidden: !!o.hidden, kind: 'image' as const })),
  ].sort((a, b) => b.z - a.z);
});
type Item = (typeof objects.value)[number];

const showAll = () => {
  const { doc } = ui.value;
  doc.transact('Show all', () => objects.value.forEach((o) => o.hidden && doc.setObjectProps(o.id, { hidden: false })));
};
const hideAll = () => {
  const { doc } = ui.value;
  ui.value.selectObject(null);
  doc.transact('Hide all', () => objects.value.forEach((o) => !o.hidden && doc.setObjectProps(o.id, { hidden: true })));
};
const reorder = (dir: 'forward' | 'backward') => {
  const sel = ui.value.selectedObject;
  if (sel) ui.value.doc.reorderObject(sel, dir);
};
const commitRename = (o: Item) => {
  const r = rename.value;
  if (r && r.name.trim()) ui.value.doc.setObjectProps(o.id, { name: r.name.trim() });
  rename.value = null;
};
const renameKey = (e: KeyboardEvent) => {
  if (e.key === 'Enter') (e.target as HTMLInputElement).blur();
  if (e.key === 'Escape') rename.value = null;
};
const toggleHidden = (o: Item) => {
  if (!o.hidden && ui.value.selectedObject === o.id) ui.value.selectObject(null);
  ui.value.doc.setObjectProps(o.id, { hidden: !o.hidden });
};
</script>

<template>
  <SidePanel title="Selection" test-id="selection-pane">
    <div class="sh-row wrap">
      <button type="button" class="sh-btn small" @click="showAll">Show All</button>
      <button type="button" class="sh-btn small" @click="hideAll">Hide All</button>
      <span class="grow" />
      <button type="button" class="sh-icon-btn" aria-label="Bring Forward" :disabled="!ui.selectedObject" @click="reorder('forward')"><Icon name="arrow_upward" /></button>
      <button type="button" class="sh-icon-btn" aria-label="Send Backward" :disabled="!ui.selectedObject" @click="reorder('backward')"><Icon name="arrow_downward" /></button>
    </div>
    <div class="sh-object-list" role="listbox" aria-label="Objects on this sheet">
      <div
        v-for="o in objects"
        :key="o.id"
        role="option"
        :aria-selected="ui.selectedObject === o.id"
        :class="`sh-object-row${ui.selectedObject === o.id ? ' selected' : ''}${o.hidden ? ' hidden' : ''}`"
        :data-testid="`object-row-${o.name}`"
        @click="!o.hidden && ui.selectObject(o.id)"
        @dblclick="rename = { id: o.id, name: o.name }"
      >
        <Icon :name="o.kind === 'chart' ? 'insert_chart' : o.kind === 'shape' ? 'interests' : 'image'" />
        <input
          v-if="rename?.id === o.id"
          v-focus
          :value="rename.name"
          aria-label="Object name"
          @input="rename = { ...rename!, name: ($event.target as HTMLInputElement).value }"
          @blur="commitRename(o)"
          @keydown="renameKey"
        />
        <span v-else class="grow">{{ o.name }}</span>
        <button type="button" class="sh-icon-btn" :aria-label="o.hidden ? 'Show' : 'Hide'" :data-testid="`object-eye-${o.name}`" @click.stop="toggleHidden(o)">
          <Icon :name="o.hidden ? 'visibility_off' : 'visibility'" />
        </button>
      </div>
      <p v-if="!objects.length" class="sh-hint">There are no shapes, charts or pictures on this sheet.</p>
    </div>
  </SidePanel>
</template>
