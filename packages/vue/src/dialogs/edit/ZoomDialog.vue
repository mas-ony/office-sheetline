<script setup lang="ts">
/** Zoom: a preset, fit to the selection, or a custom magnification. */
import { cmd } from '@sheetline/core/views';
import { ref } from 'vue';
import { useUI } from '../../context';
import Dialog from '../Dialog.vue';
import NumberField from '../NumberField.vue';
import Radio from '../Radio.vue';

const emit = defineEmits<{ close: [] }>();

const PRESETS = [200, 100, 75, 50, 25];

const ui = useUI();
const current = Math.round(ui.value.doc.sheet.zoom * 100);
const choice = ref<number | 'fit' | 'custom'>(PRESETS.includes(current) ? current : 'custom');
const custom = ref<number | ''>(current);
const ok = () => {
  const ctx = ui.value;
  if (choice.value === 'fit') cmd.zoomToSelection(ctx);
  else {
    const z = choice.value === 'custom' ? Number(custom.value) : choice.value;
    if (!Number.isFinite(z) || z < 10 || z > 400) {
      ctx.notify('Enter a whole number between 10 and 400.');
      return false;
    }
    cmd.zoomTo(ctx, z / 100);
  }
  return true;
};
const setCustom = (v: number | '') => {
  custom.value = v;
  choice.value = 'custom';
};
</script>

<template>
  <Dialog title="Zoom" :width="260" test-id="zoom-dialog" :on-ok="ok" @close="emit('close')">
    <fieldset class="sh-group">
      <legend>Magnification</legend>
      <Radio v-for="p in PRESETS" :key="p" name="zoom" :label="`${p}%`" :checked="choice === p" :test-id="`zoom-${p}`" @change="choice = p" />
      <Radio name="zoom" label="Fit selection" :checked="choice === 'fit'" @change="choice = 'fit'" />
      <div class="sh-row">
        <Radio name="zoom" label="Custom:" :checked="choice === 'custom'" @change="choice = 'custom'" />
        <NumberField :value="custom" :min="10" :max="400" label="Custom zoom" test-id="zoom-custom" @change="setCustom" />
        <span>%</span>
      </div>
    </fieldset>
  </Dialog>
</template>
