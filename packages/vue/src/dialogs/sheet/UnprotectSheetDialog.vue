<script setup lang="ts">
/** Unprotect Sheet: asks for the password when the sheet has one. */
import { computed, ref } from 'vue';
import { useUI } from '../../context';
import { vFocus } from '../../directives';
import Dialog from '../Dialog.vue';
import Field from '../Field.vue';

const emit = defineEmits<{ close: [] }>();

const ui = useUI();
const password = ref('');
// Sheets protected without a password unprotect at once.
const needs = computed(() => {
  const p = ui.value.doc.sheet.protection?.password;
  return !!p?.legacy || !!p?.hashValue;
});
const ok = () => {
  const ctx = ui.value;
  if (!needs.value) {
    ctx.doc.unprotectSheet('');
    return;
  }
  if (!ctx.doc.unprotectSheet(password.value)) {
    ctx.notify('The password you supplied is not correct. Verify that the CAPS LOCK key is off and be sure to use the correct capitalization.');
    return false;
  }
  return true;
};
</script>

<template>
  <Dialog title="Unprotect Sheet" :width="320" test-id="unprotect-sheet" :on-ok="ok" @close="emit('close')">
    <p v-if="!needs">Remove the protection of this sheet?</p>
    <Field v-else label="Password:" wide>
      <input v-focus type="password" :value="password" data-testid="unprotect-password" @input="password = ($event.target as HTMLInputElement).value" />
    </Field>
  </Dialog>
</template>
