<script setup lang="ts">
/** Sheet Settings: General, Gridlines, Headers and Sheet Tab. */
import type { ColorSpec } from '@sheetline/core/views';
import { ref, shallowRef } from 'vue';
import ColorField from '../../ColorField.vue';
import { useUI } from '../../context';
import Check from '../Check.vue';
import ColorButton from '../ColorButton.vue';
import Dialog from '../Dialog.vue';
import Field from '../Field.vue';
import Tabs from '../Tabs.vue';

const props = withDefaults(defineProps<{ tab?: string }>(), { tab: 'general' });
const emit = defineEmits<{ close: [] }>();

const TABS = [
  { id: 'general', label: 'General' },
  { id: 'gridlines', label: 'Gridlines' },
  { id: 'headers', label: 'Headers' },
  { id: 'sheettab', label: 'Sheet Tab' },
];

const ui = useUI();
const sheet = ui.value.doc.sheet;
const tab = ref(props.tab);
const frozenRows = ref(String(sheet.freeze.rows));
const frozenCols = ref(String(sheet.freeze.cols));
const zeros = ref(sheet.showZeros);
const gridH = ref(sheet.showGridLines && sheet.gridH);
const gridV = ref(sheet.showGridLines && sheet.gridV);
const gridColor = ref(sheet.gridColor ?? '');
const rowHeaders = ref(sheet.showRowHeaders);
const colHeaders = ref(sheet.showColHeaders);
const tabColor = shallowRef<ColorSpec | null>(sheet.tabColor ?? null);
const val = (e: Event) => (e.target as HTMLInputElement).value;

const ok = () => {
  const { doc } = ui.value;
  const sh = doc.sheet;
  const fr = Math.max(0, Number(frozenRows.value) || 0);
  const fc = Math.max(0, Number(frozenCols.value) || 0);
  if (fr !== sh.freeze.rows || fc !== sh.freeze.cols) doc.setFreeze(fr, fc);
  doc.setSheetView({ showZeros: zeros.value, gridH: gridH.value, gridV: gridV.value, showGridLines: gridH.value || gridV.value, gridColor: gridColor.value || undefined, showRowHeaders: rowHeaders.value, showColHeaders: colHeaders.value });
  if (JSON.stringify(tabColor.value) !== JSON.stringify(sh.tabColor ?? null)) doc.setTabColorSpec(doc.activeIndex, tabColor.value);
};
</script>

<template>
  <Dialog title="Sheet Settings" :width="420" test-id="sheet-settings" :on-ok="ok" @close="emit('close')">
    <Tabs :value="tab" :tabs="TABS" test-id="sheet-settings-tabs" @change="tab = $event" />
    <div class="sh-tab-body">
      <div v-if="tab === 'general'" class="sh-stack">
        <Field label="Frozen rows:">
          <input :value="frozenRows" inputmode="numeric" :style="{ width: '70px' }" data-testid="setting-frozen-rows" @input="frozenRows = val($event)" />
        </Field>
        <Field label="Frozen columns:">
          <input :value="frozenCols" inputmode="numeric" :style="{ width: '70px' }" @input="frozenCols = val($event)" />
        </Field>
        <Check label="Show zero values" :checked="zeros" test-id="setting-zeros" @change="zeros = $event" />
      </div>
      <div v-if="tab === 'gridlines'" class="sh-stack">
        <Check label="Show horizontal gridlines" :checked="gridH" @change="gridH = $event" />
        <Check label="Show vertical gridlines" :checked="gridV" @change="gridV = $event" />
        <Field label="Gridline color:">
          <ColorField :value="gridColor" placeholder="Automatic" @change="gridColor = $event" />
        </Field>
      </div>
      <div v-if="tab === 'headers'" class="sh-stack">
        <Check label="Show row header" :checked="rowHeaders" test-id="setting-row-header" @change="rowHeaders = $event" />
        <Check label="Show column header" :checked="colHeaders" @change="colHeaders = $event" />
      </div>
      <div v-if="tab === 'sheettab'" class="sh-stack">
        <Field label="Tab color:">
          <ColorButton :value="tabColor" none-label="No Color" test-id="setting-tab-color" @change="tabColor = $event" />
        </Field>
      </div>
    </div>
  </Dialog>
</template>
