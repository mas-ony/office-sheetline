<script setup lang="ts">
/** Go To Special: comments, constants, formulas, blanks, differences and the like. */
import { GOTO_SPECIALS, type SpecialKind } from '@sheetline/core/views';
import { ref, shallowRef } from 'vue';
import { useUI } from '../../context';
import Check from '../Check.vue';
import Dialog from '../Dialog.vue';
import Radio from '../Radio.vue';

const emit = defineEmits<{ close: [] }>();

const ui = useUI();
const kind = ref<SpecialKind>('comments');
const types = shallowRef({ numbers: true, text: true, logicals: true, errors: true });
const ok = () => {
  const n = ui.value.doc.selectSpecial(kind.value, types.value);
  if (!n) {
    ui.value.notify('No cells were found.');
    return false;
  }
  return true;
};
</script>

<template>
  <Dialog title="Go To Special" :width="420" test-id="goto-special" :on-ok="ok" @close="emit('close')">
    <div class="sh-special-grid">
      <div v-for="s in GOTO_SPECIALS" :key="s.kind">
        <Radio name="goto-special" :label="s.label" :checked="kind === s.kind" :test-id="`special-${s.kind}`" @change="kind = s.kind" />
        <div v-if="(s.kind === 'formulas' || s.kind === 'constants') && kind === s.kind" class="sh-indent">
          <Check label="Numbers" :checked="types.numbers" @change="types = { ...types, numbers: $event }" />
          <Check label="Text" :checked="types.text" @change="types = { ...types, text: $event }" />
          <Check label="Logicals" :checked="types.logicals" @change="types = { ...types, logicals: $event }" />
          <Check label="Errors" :checked="types.errors" @change="types = { ...types, errors: $event }" />
        </div>
      </div>
    </div>
  </Dialog>
</template>
