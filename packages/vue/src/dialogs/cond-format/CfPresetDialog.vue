<script setup lang="ts">
/** The quick rules (Greater Than..., Top 10 Items...): one condition and a format for the selection. */
import { type CfFormat, CF_PRESET_TEXT, cfOperand, type CfPresetKind, type CfRule, commit, HIGHLIGHT_PRESETS, selectionAverage, TIME_PERIODS, type TimePeriod } from '@sheetline/core/views';
import { computed, ref, shallowRef } from 'vue';
import { useUI } from '../../context';
import Dialog from '../Dialog.vue';
import NumberField from '../NumberField.vue';
import RangeField from '../RangeField.vue';
import FormatChooser from './FormatChooser.vue';

const props = defineProps<{ preset: CfPresetKind }>();
const emit = defineEmits<{ close: [] }>();

const ui = useUI();
const avg = selectionAverage(ui.value.doc);
const v1 = ref(String(avg));
const v2 = ref(String(avg));
const text = ref('');
const period = ref<TimePeriod>('yesterday');
const dup = ref<'duplicateValues' | 'uniqueValues'>('duplicateValues');
const rank = ref<number | ''>(10);
const format = shallowRef<CfFormat>({ ...HIGHLIGHT_PRESETS[0].format });
const meta = computed(() => CF_PRESET_TEXT[props.preset]);

const ok = () => {
  const ctx = ui.value;
  const { preset } = props;
  commit(ctx);
  let rule: CfRule;
  switch (preset) {
    case 'greater':
    case 'less':
    case 'equal':
      rule = { type: 'cellIs', operator: preset === 'greater' ? 'greaterThan' : preset === 'less' ? 'lessThan' : 'equal', formula1: cfOperand(v1.value), format: format.value };
      break;
    case 'between':
      rule = { type: 'cellIs', operator: 'between', formula1: cfOperand(v1.value), formula2: cfOperand(v2.value), format: format.value };
      break;
    case 'text':
      if (!text.value) {
        ctx.notify('Enter the text to look for.');
        return false;
      }
      rule = { type: 'text', operator: 'containsText', text: text.value, format: format.value };
      break;
    case 'date':
      rule = { type: 'timePeriod', period: period.value, format: format.value };
      break;
    case 'duplicate':
      rule = { type: dup.value, format: format.value };
      break;
    case 'above':
    case 'below':
      rule = { type: 'aboveAverage', above: preset === 'above', equal: false, format: format.value };
      break;
    default: {
      const n = typeof rank.value === 'number' ? rank.value : 10;
      if (n < 1 || n > (preset.endsWith('pct') ? 100 : 1000)) {
        ctx.notify(preset.endsWith('pct') ? 'Enter a whole number between 1 and 100.' : 'Enter a whole number between 1 and 1000.');
        return false;
      }
      rule = { type: 'top10', bottom: preset.startsWith('bottom'), percent: preset.endsWith('pct'), rank: n, format: format.value };
    }
  }
  ctx.doc.addCondFormat(rule);
  return true;
};
</script>

<template>
  <Dialog :title="meta.title" :on-ok="ok" :width="460" :blocking="false" test-id="cf-preset" @close="emit('close')">
    <p class="sh-lead">{{ meta.lead }}</p>
    <div class="sh-row wrap">
      <RangeField v-if="preset === 'greater' || preset === 'less' || preset === 'equal' || preset === 'between'" :value="v1" label="Value" test-id="cf-value1" @change="v1 = $event" />
      <template v-if="preset === 'between'">
        <span>and</span>
        <RangeField :value="v2" label="Second value" test-id="cf-value2" @change="v2 = $event" />
      </template>
      <RangeField v-if="preset === 'text'" :value="text" label="Text" test-id="cf-text" @change="text = $event" />
      <select v-if="preset === 'date'" :value="period" data-testid="cf-period" @change="period = ($event.target as HTMLSelectElement).value as TimePeriod">
        <option v-for="p in TIME_PERIODS" :key="p.period" :value="p.period">{{ p.label }}</option>
      </select>
      <select v-if="preset === 'duplicate'" :value="dup" data-testid="cf-dup" @change="dup = ($event.target as HTMLSelectElement).value as typeof dup">
        <option value="duplicateValues">Duplicate</option>
        <option value="uniqueValues">Unique</option>
      </select>
      <template v-if="preset.startsWith('top') || preset.startsWith('bottom')">
        <NumberField :value="rank" :min="1" :max="preset.endsWith('pct') ? 100 : 1000" label="Rank" test-id="cf-rank" @change="rank = $event" />
        <span>{{ preset.endsWith('pct') ? '%' : '' }}</span>
      </template>
      <span v-if="preset === 'above' || preset === 'below'">for the selected range</span>
      <span>with</span>
      <FormatChooser :value="format" test-id="cf-format" @change="format = $event" />
    </div>
  </Dialog>
</template>
