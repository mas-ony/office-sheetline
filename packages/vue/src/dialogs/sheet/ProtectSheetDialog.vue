<script setup lang="ts">
/** Protect Sheet: what users may still do, an optional password (confirmed in a second step). */
import { commit, DEFAULT_PROTECTION, PROTECTION_LABELS, type ProtectionOptions } from '@sheetline/core/views';
import { ref, shallowRef } from 'vue';
import { useUI } from '../../context';
import { vFocus } from '../../directives';
import Check from '../Check.vue';
import Dialog from '../Dialog.vue';
import Field from '../Field.vue';

const emit = defineEmits<{ close: [] }>();

const ui = useUI();
const password = ref('');
const opts = shallowRef<ProtectionOptions>({ ...DEFAULT_PROTECTION });
const confirmPw = ref<string | null>(null);
const val = (e: Event) => (e.target as HTMLInputElement).value;

const ok = () => {
  if (password.value) {
    confirmPw.value = '';
    return false;
  }
  const ctx = ui.value;
  commit(ctx);
  ctx.doc.protectSheet(opts.value, '');
  return true;
};
const confirmOk = () => {
  const ctx = ui.value;
  if (confirmPw.value !== password.value) {
    ctx.notify('Confirmation password is not identical.');
    confirmPw.value = null;
    return false;
  }
  commit(ctx);
  ctx.doc.protectSheet(opts.value, password.value);
  emit('close');
  return false;
};
const close = () => {
  if (confirmPw.value !== null) confirmPw.value = null;
  else emit('close');
};
</script>

<template>
  <Dialog
    :title="confirmPw !== null ? 'Confirm Password' : 'Protect Sheet'"
    :width="confirmPw !== null ? 340 : 360"
    :test-id="confirmPw !== null ? 'protect-confirm' : 'protect-sheet'"
    :on-ok="confirmPw !== null ? confirmOk : ok"
    @close="close"
  >
    <template v-if="confirmPw !== null">
      <Field label="Reenter password to proceed." wide>
        <input v-focus type="password" :value="confirmPw" data-testid="protect-confirm-input" @input="confirmPw = val($event)" />
      </Field>
      <p class="sh-hint">Caution: If you lose or forget the password, it cannot be recovered.</p>
    </template>
    <template v-else>
      <Check label="Protect worksheet and contents of locked cells" :checked="true" />
      <Field label="Password to unprotect sheet:" wide>
        <input type="password" :value="password" data-testid="protect-password" @input="password = val($event)" />
      </Field>
      <div class="sh-label">Allow all users of this worksheet to:</div>
      <div class="sh-checklist" data-testid="protect-options">
        <Check v-for="p in PROTECTION_LABELS" :key="p.key" :label="p.label" :checked="opts[p.key]" @change="opts = { ...opts, [p.key]: $event }" />
      </div>
    </template>
  </Dialog>
</template>
