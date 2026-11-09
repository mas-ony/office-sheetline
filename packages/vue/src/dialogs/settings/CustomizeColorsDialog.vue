<script setup lang="ts">
/** Create New Theme Colors: the twelve theme colours, with a sample. */
import { THEME_SLOT_LABELS, type ThemeSlot } from '@sheetline/core/views';
import { ref, shallowRef } from 'vue';
import ColorField from '../../ColorField.vue';
import { useUI } from '../../context';
import Dialog from '../Dialog.vue';
import Field from '../Field.vue';

const emit = defineEmits<{ close: [] }>();

const ACCENTS: ThemeSlot[] = ['accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6'];

const ui = useUI();
const theme = ui.value.doc.wb.theme;
const colors = shallowRef<Record<ThemeSlot, string>>({ ...theme.colors });
const name = ref('Custom 1');
const setColor = (slot: ThemeSlot, hex: string) => {
  if (hex) colors.value = { ...colors.value, [slot]: hex.replace('#', '').toUpperCase() };
};
const ok = () => ui.value.doc.applyTheme({ name: name.value, colors: colors.value });
</script>

<template>
  <Dialog title="Create New Theme Colors" :width="480" test-id="customize-colors" :on-ok="ok" @close="emit('close')">
    <template #footer-left>
      <button type="button" class="sh-btn" @click="colors = { ...theme.colors }">Reset</button>
    </template>
    <div class="sh-two-col wide-left">
      <div class="sh-stack tight">
        <div class="sh-label">Theme colors</div>
        <div v-for="s in THEME_SLOT_LABELS" :key="s.slot" class="sh-row">
          <span class="grow">{{ s.label }}</span>
          <ColorField :value="`#${colors[s.slot]}`" @change="setColor(s.slot, $event)" />
        </div>
      </div>
      <div class="sh-stack">
        <div class="sh-label">Sample</div>
        <div class="sh-theme-sample" :style="{ background: `#${colors.lt1}`, color: `#${colors.dk1}` }">
          <div :style="{ background: `#${colors.dk2}`, color: `#${colors.lt1}` }">Text</div>
          <div class="bars">
            <i v-for="k in ACCENTS" :key="k" :style="{ background: `#${colors[k]}` }" />
          </div>
          <u :style="{ color: `#${colors.hlink}` }">Hyperlink</u> <u :style="{ color: `#${colors.folHlink}` }">Followed</u>
        </div>
      </div>
    </div>
    <Field label="Name:" wide>
      <input :value="name" @input="name = ($event.target as HTMLInputElement).value" />
    </Field>
  </Dialog>
</template>
