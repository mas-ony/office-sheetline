<script setup lang="ts">
/** The Format pane of the selected chart, shape or picture. */
import { computed } from 'vue';
import { useUI } from '../context';
import ChartFormat from './ChartFormat.vue';
import PaneSection from './PaneSection.vue';
import ShapeFormat from './ShapeFormat.vue';
import SidePanel from './SidePanel.vue';
import SizeSection from './SizeSection.vue';

const ui = useUI();
const hit = computed(() => {
  const id = ui.value.selectedObject;
  return { value: id ? ui.value.doc.objectById(id) : null };
});
</script>

<template>
  <SidePanel v-if="!hit.value" title="Format" test-id="format-pane">
    <p class="sh-hint">Select a chart, shape or picture to format it.</p>
  </SidePanel>
  <ChartFormat v-else-if="hit.value.kind === 'chart'" :chart="hit.value.obj" />
  <ShapeFormat v-else-if="hit.value.kind === 'shape'" :shape="hit.value.obj" />
  <SidePanel v-else title="Format Picture" test-id="format-pane">
    <SizeSection :id="hit.value.obj.id" />
    <PaneSection title="Alt Text">
      <button type="button" class="sh-btn" @click="ui.openDialog({ kind: 'alt-text', id: hit.value!.obj.id })">Edit Alt Text...</button>
    </PaneSection>
  </SidePanel>
</template>
