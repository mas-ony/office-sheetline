<script setup lang="ts">
/** Custom Header / Custom Footer: the left, center and right sections, with codes to insert. */
import { hfJoin, hfSections } from '@sheetline/core/views';
import { ref, shallowRef } from 'vue';
import Dialog from '../Dialog.vue';

type Section = 'left' | 'center' | 'right';

const props = defineProps<{ title: string; code: string }>();
const emit = defineEmits<{ save: [code: string]; close: [] }>();

const CODES = [
  ['Page Number', '&P'],
  ['Number of Pages', '&N'],
  ['Date', '&D'],
  ['Time', '&T'],
  ['File Name', '&F'],
  ['Sheet Name', '&A'],
];
const SECTIONS: Section[] = ['left', 'center', 'right'];

const s = shallowRef(hfSections(props.code));
const focus = ref<Section>('center');
const insert = (token: string) => (s.value = { ...s.value, [focus.value]: s.value[focus.value] + token });
const ok = () => {
  emit('save', hfJoin(s.value));
};
</script>

<template>
  <Dialog :title="title" :width="600" test-id="header-footer-editor" :on-ok="ok" @close="emit('close')">
    <p class="sh-hint">To format text: select the text, then choose a code. Page number &amp;P, number of pages &amp;N, date &amp;D, time &amp;T, file &amp;F, sheet &amp;A.</p>
    <div class="sh-row wrap">
      <button v-for="[label, token] in CODES" :key="token" type="button" class="sh-btn" :title="`Insert ${label}`" @click="insert(token)">{{ label }}</button>
    </div>
    <div class="sh-hf-sections">
      <label v-for="k in SECTIONS" :key="k">
        <span>{{ k === 'left' ? 'Left section:' : k === 'center' ? 'Center section:' : 'Right section:' }}</span>
        <textarea :rows="5" :value="s[k]" :data-testid="`hf-${k}`" @focus="focus = k" @input="s = { ...s, [k]: ($event.target as HTMLTextAreaElement).value }" />
      </label>
    </div>
  </Dialog>
</template>
