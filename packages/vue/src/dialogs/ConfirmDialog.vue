<script setup lang="ts">
/** A message box: OK / Cancel, or Yes / No / Cancel with `alt`. */
import Icon from '../Icon.vue';
import Modal from '../Modal.vue';

withDefaults(defineProps<{ message: string; title?: string; ok?: string; cancel?: string; alt?: string; danger?: boolean }>(), { ok: 'OK', cancel: 'Cancel' });
const emit = defineEmits<{ answer: [value: boolean, alt?: boolean] }>();
</script>

<template>
  <Modal :title="title ?? 'Spreadsheet'" :width="420" blocking test-id="confirm-dialog" @close="emit('answer', false)">
    <div class="sh-message">
      <Icon :name="danger ? 'warning' : 'info'" :class="danger ? 'warn' : 'info'" />
      <p>{{ message }}</p>
    </div>
    <template #footer>
      <button type="button" :class="`sh-btn-primary${danger ? ' danger' : ''}`" data-testid="confirm-ok" data-autofocus="true" @click="emit('answer', true)">{{ ok }}</button>
      <button v-if="alt" type="button" class="sh-btn-secondary" data-testid="confirm-alt" @click="emit('answer', false, true)">{{ alt }}</button>
      <button v-if="cancel" type="button" class="sh-btn-secondary" data-testid="confirm-cancel" @click="emit('answer', false)">{{ cancel }}</button>
    </template>
  </Modal>
</template>
