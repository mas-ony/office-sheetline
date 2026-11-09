<script setup lang="ts">
/** Keyboard Shortcuts: the list of shortcut keys. */
import { KEYBOARD_SHORTCUTS } from '@sheetline/core/views';
import Dialog from '../Dialog.vue';

const emit = defineEmits<{ close: [] }>();
</script>

<template>
  <Dialog title="Keyboard Shortcuts" close-only :width="480" test-id="shortcuts" @close="emit('close')">
    <div class="sh-shortcuts">
      <div v-for="[k, v] in KEYBOARD_SHORTCUTS" :key="k" class="sh-kv">
        <kbd>{{ k }}</kbd>
        <span>{{ v }}</span>
      </div>
    </div>
  </Dialog>
</template>
