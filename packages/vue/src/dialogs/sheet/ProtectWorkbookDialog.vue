<script setup lang="ts">
/** Protect Structure and Windows, with an optional password. */
import { ref } from 'vue';
import { useUI } from '../../context';
import Check from '../Check.vue';
import Dialog from '../Dialog.vue';
import Field from '../Field.vue';

const emit = defineEmits<{ close: [] }>();

const ui = useUI();
const password = ref('');
const ok = () => {
  ui.value.doc.protectWorkbook(password.value);
};
</script>

<template>
  <Dialog title="Protect Structure and Windows" :width="340" test-id="protect-workbook" :on-ok="ok" @close="emit('close')">
    <div class="sh-label">Protect workbook for</div>
    <Check label="Structure" :checked="true" />
    <Field label="Password (optional):" wide>
      <input type="password" :value="password" data-testid="protect-workbook-password" @input="password = ($event.target as HTMLInputElement).value" />
    </Field>
  </Dialog>
</template>
