<script setup lang="ts">
/** A one-line prompt: Enter or OK answers with the text, Cancel with null. */
import { ref } from 'vue';
import Modal from '../Modal.vue';

const props = withDefaults(defineProps<{ message: string; value: string; title?: string; ok?: string }>(), { ok: 'OK' });
const emit = defineEmits<{ answer: [value: string | null] }>();
const text = ref(props.value);
</script>

<template>
  <Modal :title="title ?? message" :width="400" blocking test-id="prompt-dialog" @close="emit('answer', null)">
    <div class="sh-form">
      <label class="sh-field">
        <span>{{ message }}</span>
        <input
          class="sh-input"
          :value="text"
          data-testid="prompt-input"
          @focus="($event.target as HTMLInputElement).select()"
          @input="text = ($event.target as HTMLInputElement).value"
          @keydown="$event.key === 'Enter' && emit('answer', text)"
        />
      </label>
    </div>
    <template #footer>
      <button type="button" class="sh-btn-primary" data-testid="prompt-ok" @click="emit('answer', text)">{{ ok }}</button>
      <button type="button" class="sh-btn-secondary" @click="emit('answer', null)">Cancel</button>
    </template>
  </Modal>
</template>
