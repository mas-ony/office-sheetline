<script setup lang="ts">
/** Messages of the editor itself: recovery, errors, dragged files and files being read. */
import type { UIState } from '@sheetline/core/views';
import { useUI } from './context';

defineProps<{ state: UIState }>();

const { ui, layer } = useUI().value;
const when = (savedAt: number) => new Date(savedAt).toLocaleString();
</script>

<template>
  <Teleport :to="layer">
    <div v-if="state.restore" class="sh-restore" role="status" data-testid="restore-banner">
      <span>Recover “{{ state.restore.fileName }}” from {{ when(state.restore.savedAt) }}?</span>
      <button type="button" class="sh-btn primary" data-testid="restore-recover" @click="ui.recover">Recover</button>
      <button type="button" class="sh-btn" data-testid="restore-discard" @click="ui.discardRecovery">Discard</button>
    </div>
    <div v-if="state.notice" class="sh-toast" role="alert" data-testid="notice">
      <span>{{ state.notice }}</span>
      <button type="button" @click="ui.notice(null)">OK</button>
    </div>
    <div v-if="state.dragging" class="sh-drop-hint">Drop a spreadsheet to open it, or a picture to insert it</div>
    <div v-if="state.busy" class="sh-busy" role="status" aria-live="polite" data-testid="busy"><span class="sh-spinner" aria-hidden="true" />{{ state.busy }}</div>
  </Teleport>
</template>
