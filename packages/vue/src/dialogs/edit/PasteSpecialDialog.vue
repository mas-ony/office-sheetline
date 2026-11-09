<script setup lang="ts">
/** Paste Special: what to paste, an operation, skip blanks and transpose; or Paste Link. */
import { cmd, PASTE_SPECIAL_WHAT, type PasteOptions, type PasteWhat } from '@sheetline/core/views';
import { computed, ref } from 'vue';
import { useUI } from '../../context';
import Check from '../Check.vue';
import Dialog from '../Dialog.vue';
import Radio from '../Radio.vue';

const emit = defineEmits<{ close: [] }>();

const OPS = ['none', 'add', 'subtract', 'multiply', 'divide'] as const;

const ui = useUI();
const what = ref<PasteWhat>('all');
const op = ref<NonNullable<PasteOptions['op']>>('none');
const skipBlanks = ref(false);
const transpose = ref(false);
const has = computed(() => ui.value.doc.hasClip());
const paste = (link = false) => {
  void cmd.paste(ui.value, { what: what.value, op: op.value, skipBlanks: skipBlanks.value, transpose: transpose.value, link });
};
const pasteLink = () => {
  paste(true);
  emit('close');
};
</script>

<template>
  <Dialog title="Paste Special" :width="460" test-id="paste-special" :ok-disabled="!has" :on-ok="() => paste()" @close="emit('close')">
    <template #footer-left>
      <button type="button" class="sh-btn" :disabled="!has" @click="pasteLink">Paste Link</button>
    </template>
    <p v-if="!has" class="sh-hint">Copy cells in this workbook first (Ctrl+C).</p>
    <div class="sh-row top wrap">
      <fieldset class="sh-group">
        <legend>Paste</legend>
        <div class="sh-two-col even">
          <Radio v-for="p in PASTE_SPECIAL_WHAT" :key="p.what" name="paste-what" :label="p.label" :checked="what === p.what" :test-id="`paste-what-${p.what}`" @change="what = p.what" />
        </div>
      </fieldset>
    </div>
    <fieldset class="sh-group">
      <legend>Operation</legend>
      <div class="sh-two-col even">
        <Radio v-for="o in OPS" :key="o" name="paste-op" :label="o === 'none' ? 'None' : o[0].toUpperCase() + o.slice(1)" :checked="op === o" :test-id="`paste-op-${o}`" @change="op = o" />
      </div>
    </fieldset>
    <div class="sh-row">
      <Check label="Skip blanks" :checked="skipBlanks" @change="skipBlanks = $event" />
      <Check label="Transpose" :checked="transpose" test-id="paste-transpose" @change="transpose = $event" />
    </div>
  </Dialog>
</template>
