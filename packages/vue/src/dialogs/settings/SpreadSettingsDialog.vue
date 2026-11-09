<script setup lang="ts">
/** Spread Settings: General, Scroll Bars, Calculation and Tab Strip. */
import { ref, shallowRef } from 'vue';
import { useUI } from '../../context';
import Check from '../Check.vue';
import Dialog from '../Dialog.vue';
import Field from '../Field.vue';
import Tabs from '../Tabs.vue';

const props = withDefaults(defineProps<{ tab?: string }>(), { tab: 'general' });
const emit = defineEmits<{ close: [] }>();

const TABS = [
  { id: 'general', label: 'General' },
  { id: 'scrollbars', label: 'Scroll Bars' },
  { id: 'calculation', label: 'Calculation' },
  { id: 'tabstrip', label: 'Tab Strip' },
];

const ui = useUI();
const { doc, settings, view } = ui.value;
const tab = ref(props.tab);
const s = shallowRef(settings);
const calc = ref(doc.wb.calcMode ?? 'auto');
const tabStrip = ref(view.tabStrip);
const newTab = ref(view.newTab);
const formulaBar = ref(view.formulaBar);

const ok = () => {
  const ctx = ui.value;
  ctx.setSettings(s.value);
  ctx.setView({ tabStrip: tabStrip.value, newTab: newTab.value, formulaBar: formulaBar.value });
  if (calc.value !== (ctx.doc.wb.calcMode ?? 'auto')) ctx.doc.setCalcMode(calc.value);
};
</script>

<template>
  <Dialog title="Spread Settings" :width="440" test-id="spread-settings" :on-ok="ok" @close="emit('close')">
    <Tabs :value="tab" :tabs="TABS" test-id="spread-settings-tabs" @change="tab = $event" />
    <div class="sh-tab-body">
      <div v-if="tab === 'general'" class="sh-stack">
        <Check label="Allow Drag and Fill" :checked="s.allowFill" test-id="setting-fill" @change="s = { ...s, allowFill: $event }" />
        <Check label="Allow Zoom" :checked="s.allowZoom" test-id="setting-zoom" @change="s = { ...s, allowZoom: $event }" />
        <Check label="Formula Format Hints" :checked="s.formulaHints" @change="s = { ...s, formulaHints: $event }" />
        <Check label="Show Formula Bar" :checked="formulaBar" @change="formulaBar = $event" />
      </div>
      <div v-if="tab === 'scrollbars'" class="sh-stack">
        <Check label="Show Horizontal Scroll Bar" :checked="s.showHScroll" test-id="setting-hscroll" @change="s = { ...s, showHScroll: $event }" />
        <Check label="Show Vertical Scroll Bar" :checked="s.showVScroll" @change="s = { ...s, showVScroll: $event }" />
      </div>
      <div v-if="tab === 'calculation'" class="sh-stack">
        <Field label="Calculation:">
          <select :value="calc" data-testid="setting-calc" @change="calc = ($event.target as HTMLSelectElement).value as typeof calc">
            <option value="auto">Automatic</option>
            <option value="autoNoTable">Automatic except for data tables</option>
            <option value="manual">Manual</option>
          </select>
        </Field>
        <p class="sh-hint">With Manual, formulas recalculate when you press F9 (Calculate Now) or Shift+F9 (Calculate Sheet).</p>
      </div>
      <div v-if="tab === 'tabstrip'" class="sh-stack">
        <Check label="Show Tab Strip" :checked="tabStrip" test-id="setting-tabstrip" @change="tabStrip = $event" />
        <Check label="Show New Tab" :checked="newTab" @change="newTab = $event" />
      </div>
    </div>
  </Dialog>
</template>
