<script setup lang="ts">
/**
 * Side panels at the right of the grid: Threaded Comments, the Selection Pane (objects, show/hide,
 * order, names) and the Format pane of the selected chart, shape or picture.
 */
import { useUI } from '../context';
import CommentsPanel from './CommentsPanel.vue';
import FormatPane from './FormatPane.vue';
import SelectionPane from './SelectionPane.vue';

const ui = useUI();
</script>

<template>
  <CommentsPanel v-if="ui.panel === 'comments'" />
  <SelectionPane v-else-if="ui.panel === 'selection'" />
  <FormatPane v-else-if="ui.panel === 'format'" />
</template>
