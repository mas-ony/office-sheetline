<script setup lang="ts">
/** New / Edit Formatting Rule: the six rule types of Excel with their descriptions and the format. */
import {
  CELL_IS_OPERATORS, type CellIsOperator, type CfContainType, type CfFormat, cfHex, cfOperand, type CfRule, CF_RULE_KINDS, type CfRuleKind, cfRuleKindOf, cfSpec, type Cfvo,
  CFVO_TYPE_OPTIONS, type CfvoType, type ColorSpec, DATA_BAR_COLORS, defaultThresholds, ICON_SET_GROUPS, type IconSetName, TEXT_OPERATOR_OPTIONS, type TextOperator, TIME_PERIODS, type TimePeriod,
} from '@sheetline/core/views';
import { computed, ref, shallowRef } from 'vue';
import { useUI } from '../../context';
import Check from '../Check.vue';
import ColorButton from '../ColorButton.vue';
import Dialog from '../Dialog.vue';
import Field from '../Field.vue';
import NumberField from '../NumberField.vue';
import RangeField from '../RangeField.vue';
import CfvoEditor from './CfvoEditor.vue';
import FormatEditor from './FormatEditor.vue';
import RulePreview from './RulePreview.vue';

const props = defineProps<{ initial: CfRule | null; hint?: string; title: string }>();
const emit = defineEmits<{ save: [rule: CfRule]; close: [] }>();

type FormatStyle = '2scale' | '3scale' | 'dataBar' | 'iconSet';

/** "► " as a text node of its own, as React renders it before the label (Chrome draws one text node slightly differently). */
const Arrow = () => '► ';

const { initial, hint } = props;
const fmt0: CfFormat = initial && 'format' in initial ? { ...initial.format } : {};
const kind = ref<CfRuleKind>(cfRuleKindOf(initial, hint));
const format = shallowRef<CfFormat>(fmt0);
// Values
const formatStyle = ref<FormatStyle>(initial?.type === 'dataBar' || hint === 'dataBar' ? 'dataBar' : initial?.type === 'iconSet' || hint === 'iconSet' ? 'iconSet' : initial?.type === 'colorScale' && initial.points.length === 3 ? '3scale' : '2scale');
const scale0 = initial?.type === 'colorScale' ? initial.points : null;
const minV = shallowRef<Cfvo>(scale0?.[0]?.cfvo ?? (initial?.type === 'dataBar' ? initial.min : { type: 'min' }));
const midV = shallowRef<Cfvo>(scale0 && scale0.length === 3 ? scale0[1].cfvo : { type: 'percentile', value: '50' });
const maxV = shallowRef<Cfvo>(scale0?.[scale0.length - 1]?.cfvo ?? (initial?.type === 'dataBar' ? initial.max : { type: 'max' }));
const minC = ref(scale0?.[0]?.color ?? '#F8696B');
const midC = ref(scale0 && scale0.length === 3 ? scale0[1].color : '#FFEB84');
const maxC = ref(scale0?.[scale0.length - 1]?.color ?? '#63BE7B');
const barColor = ref(initial?.type === 'dataBar' ? initial.color : DATA_BAR_COLORS[0].color);
const gradient = ref(initial?.type === 'dataBar' ? initial.gradient : true);
const barOnly = ref(initial?.type === 'dataBar' ? !initial.showValue : false);
const iconSet = ref<IconSetName>(initial?.type === 'iconSet' ? initial.iconSet : '3TrafficLights1');
const reverse = ref(initial?.type === 'iconSet' ? initial.reverse : false);
const iconOnly = ref(initial?.type === 'iconSet' ? !initial.showValue : false);
const thresholds = shallowRef(initial?.type === 'iconSet' ? initial.thresholds : defaultThresholds('3TrafficLights1'));
// Contain
const containType = ref<CfContainType>(initial && ['cellIs', 'text', 'timePeriod', 'blanks', 'noBlanks', 'errors', 'noErrors'].includes(initial.type) ? (initial.type as CfContainType) : 'cellIs');
const op = ref<CellIsOperator>(initial?.type === 'cellIs' ? initial.operator : 'between');
const f1 = ref(initial?.type === 'cellIs' ? `=${initial.formula1}` : '');
const f2 = ref(initial?.type === 'cellIs' && initial.formula2 ? `=${initial.formula2}` : '');
const textOp = ref<TextOperator>(initial?.type === 'text' ? initial.operator : 'containsText');
const text = ref(initial?.type === 'text' ? initial.text : '');
const period = ref<TimePeriod>(initial?.type === 'timePeriod' ? initial.period : 'yesterday');
// Rank / average / unique / formula
const bottom = ref(initial?.type === 'top10' ? initial.bottom : false);
const percent = ref(initial?.type === 'top10' ? initial.percent : false);
const rank = ref<number | ''>(initial?.type === 'top10' ? initial.rank : 10);
const avgKind = ref(initial?.type === 'aboveAverage' ? `${initial.above ? 'above' : 'below'}${initial.equal ? '-eq' : ''}${initial.stdDev ? `-${initial.stdDev}` : ''}` : 'above');
const unique = ref<'duplicateValues' | 'uniqueValues'>(initial?.type === 'uniqueValues' ? 'uniqueValues' : 'duplicateValues');
const formula = ref(initial?.type === 'expression' ? `=${initial.formula}` : '');

const ui = useUI();
const hex = (c: ColorSpec) => cfHex(c, ui.value.doc.wb.theme)!;
const strip = (s: string) => cfOperand(s.trim());
const iconSets = ICON_SET_GROUPS.flatMap((g) => g.sets);
const thresholdTypes = CFVO_TYPE_OPTIONS.filter((c) => c.type !== 'min' && c.type !== 'max');
const iconPreview = computed<CfRule>(() => ({ type: 'iconSet', iconSet: iconSet.value, reverse: reverse.value, showValue: !iconOnly.value, thresholds: thresholds.value }));
/** The thresholds after the first, last first, with their index. */
const thresholdRows = computed(() =>
  thresholds.value
    .slice(1)
    .reverse()
    .map((t, k) => ({ t, k, i: thresholds.value.length - 1 - k })),
);
const setThreshold = (i: number, patch: (x: { cfvo: Cfvo; gte: boolean }) => { cfvo: Cfvo; gte: boolean }) => (thresholds.value = thresholds.value.map((x, j) => (j === i ? patch(x) : x)));
const pickIconSet = (e: Event) => {
  const s = (e.target as HTMLSelectElement).value as IconSetName;
  iconSet.value = s;
  thresholds.value = defaultThresholds(s);
};
const needsFormat = computed(() => kind.value !== 'values');

const build = (): CfRule | string => {
  switch (kind.value) {
    case 'values':
      if (formatStyle.value === 'dataBar') return { type: 'dataBar', min: minV.value, max: maxV.value, color: barColor.value, gradient: gradient.value, showValue: !barOnly.value };
      if (formatStyle.value === 'iconSet') return { type: 'iconSet', iconSet: iconSet.value, reverse: reverse.value, showValue: !iconOnly.value, thresholds: thresholds.value };
      return {
        type: 'colorScale',
        points:
          formatStyle.value === '3scale'
            ? [{ cfvo: minV.value, color: minC.value }, { cfvo: midV.value, color: midC.value }, { cfvo: maxV.value, color: maxC.value }]
            : [{ cfvo: minV.value, color: minC.value }, { cfvo: maxV.value, color: maxC.value }],
      };
    case 'contain':
      if (containType.value === 'cellIs') {
        if (!f1.value.trim()) return 'Enter a value.';
        const two = op.value === 'between' || op.value === 'notBetween';
        if (two && !f2.value.trim()) return 'Enter the second value.';
        return { type: 'cellIs', operator: op.value, formula1: strip(f1.value), formula2: two ? strip(f2.value) : undefined, format: format.value };
      }
      if (containType.value === 'text') return text.value ? { type: 'text', operator: textOp.value, text: text.value, format: format.value } : 'Enter the text.';
      if (containType.value === 'timePeriod') return { type: 'timePeriod', period: period.value, format: format.value };
      return { type: containType.value, format: format.value };
    case 'rank': {
      const n = typeof rank.value === 'number' ? rank.value : 10;
      return { type: 'top10', bottom: bottom.value, percent: percent.value, rank: Math.max(1, Math.min(percent.value ? 100 : 1000, n)), format: format.value };
    }
    case 'average': {
      const [dir, a, b] = avgKind.value.split('-');
      const equal = a === 'eq';
      const sd = Number(equal ? b : a) || undefined;
      return { type: 'aboveAverage', above: dir === 'above', equal, stdDev: sd, format: format.value };
    }
    case 'unique':
      return { type: unique.value, format: format.value };
    case 'formula':
      return formula.value.trim() ? { type: 'expression', formula: formula.value.trim().replace(/^=/, ''), format: format.value } : 'Enter a formula.';
  }
};

const ok = () => {
  const r = build();
  if (typeof r === 'string') {
    ui.value.notify(r);
    return false;
  }
  emit('save', r);
  return true;
};
</script>

<template>
  <Dialog :title="title" :width="560" test-id="cf-rule" :on-ok="ok" @close="emit('close')">
    <div class="sh-label">Select a Rule Type:</div>
    <div class="sh-listbox" role="listbox" aria-label="Rule type" :style="{ height: 'auto' }" data-testid="cf-rule-types">
      <div v-for="k in CF_RULE_KINDS" :key="k.kind" role="option" :aria-selected="kind === k.kind" :class="`sh-listbox-item${kind === k.kind ? ' selected' : ''}`" @click="kind = k.kind"><Arrow />{{ k.label }}</div>
    </div>
    <div class="sh-label">Edit the Rule Description:</div>
    <fieldset class="sh-group">
      <div v-if="kind === 'values'" class="sh-stack">
        <Field label="Format Style:">
          <select :value="formatStyle" data-testid="cf-style" @change="formatStyle = ($event.target as HTMLSelectElement).value as FormatStyle">
            <option value="2scale">2-Color Scale</option>
            <option value="3scale">3-Color Scale</option>
            <option value="dataBar">Data Bar</option>
            <option value="iconSet">Icon Sets</option>
          </select>
        </Field>
        <div v-if="formatStyle === '2scale' || formatStyle === '3scale' || formatStyle === 'dataBar'" class="sh-cfvo-grid">
          <div>
            <b>Minimum</b>
            <CfvoEditor :value="minV" which="min" @change="minV = $event" />
            <ColorButton v-if="formatStyle !== 'dataBar'" :value="cfSpec(minC)" none-label="—" @change="(c) => c && (minC = hex(c))" />
          </div>
          <div v-if="formatStyle === '3scale'">
            <b>Midpoint</b>
            <CfvoEditor :value="midV" which="mid" @change="midV = $event" />
            <ColorButton :value="cfSpec(midC)" none-label="—" @change="(c) => c && (midC = hex(c))" />
          </div>
          <div>
            <b>Maximum</b>
            <CfvoEditor :value="maxV" which="max" @change="maxV = $event" />
            <ColorButton v-if="formatStyle !== 'dataBar'" :value="cfSpec(maxC)" none-label="—" @change="(c) => c && (maxC = hex(c))" />
          </div>
        </div>
        <div v-if="formatStyle === 'dataBar'" class="sh-row wrap">
          <Field label="Bar color:">
            <ColorButton :value="cfSpec(barColor)" none-label="—" @change="(c) => c && (barColor = hex(c))" />
          </Field>
          <Field label="Fill:">
            <select :value="gradient ? 'gradient' : 'solid'" @change="gradient = ($event.target as HTMLSelectElement).value === 'gradient'">
              <option value="gradient">Gradient Fill</option>
              <option value="solid">Solid Fill</option>
            </select>
          </Field>
          <Check label="Show Bar Only" :checked="barOnly" @change="barOnly = $event" />
        </div>
        <div v-if="formatStyle === 'iconSet'" class="sh-stack">
          <div class="sh-row wrap">
            <Field label="Icon Style:">
              <select :value="iconSet" data-testid="cf-iconset" @change="pickIconSet">
                <option v-for="s in iconSets" :key="s.name" :value="s.name">{{ s.label }}</option>
              </select>
            </Field>
            <Check label="Reverse Icon Order" :checked="reverse" @change="reverse = $event" />
            <Check label="Show Icon Only" :checked="iconOnly" @change="iconOnly = $event" />
          </div>
          <RulePreview :rule="iconPreview" />
          <div v-for="{ t, k, i } in thresholdRows" :key="i" class="sh-row">
            <span>Icon {{ k + 1 }}: when value is</span>
            <select :value="t.gte ? '>=' : '>'" @change="setThreshold(i, (x) => ({ ...x, gte: ($event.target as HTMLSelectElement).value === '>=' }))">
              <option value=">=">&gt;=</option>
              <option value=">">&gt;</option>
            </select>
            <input :value="t.cfvo.value ?? ''" :style="{ width: '70px' }" @input="setThreshold(i, (x) => ({ ...x, cfvo: { ...x.cfvo, value: ($event.target as HTMLInputElement).value } }))" />
            <select :value="t.cfvo.type" @change="setThreshold(i, (x) => ({ ...x, cfvo: { ...x.cfvo, type: ($event.target as HTMLSelectElement).value as CfvoType } }))">
              <option v-for="c in thresholdTypes" :key="c.type" :value="c.type">{{ c.label }}</option>
            </select>
          </div>
        </div>
      </div>
      <div v-if="kind === 'contain'" class="sh-stack">
        <span>Format only cells with:</span>
        <div class="sh-row wrap">
          <select :value="containType" data-testid="cf-contain-type" @change="containType = ($event.target as HTMLSelectElement).value as CfContainType">
            <option value="cellIs">Cell Value</option>
            <option value="text">Specific Text</option>
            <option value="timePeriod">Dates Occurring</option>
            <option value="blanks">Blanks</option>
            <option value="noBlanks">No Blanks</option>
            <option value="errors">Errors</option>
            <option value="noErrors">No Errors</option>
          </select>
          <template v-if="containType === 'cellIs'">
            <select :value="op" data-testid="cf-operator" @change="op = ($event.target as HTMLSelectElement).value as CellIsOperator">
              <option v-for="o in CELL_IS_OPERATORS" :key="o.op" :value="o.op">{{ o.label }}</option>
            </select>
            <RangeField :value="f1" label="Value" test-id="cf-f1" @change="f1 = $event" />
            <template v-if="op === 'between' || op === 'notBetween'">
              <span>and</span>
              <RangeField :value="f2" label="Second value" test-id="cf-f2" @change="f2 = $event" />
            </template>
          </template>
          <template v-if="containType === 'text'">
            <select :value="textOp" @change="textOp = ($event.target as HTMLSelectElement).value as TextOperator">
              <option v-for="o in TEXT_OPERATOR_OPTIONS" :key="o.op" :value="o.op">{{ o.label }}</option>
            </select>
            <input :value="text" data-testid="cf-rule-text" @input="text = ($event.target as HTMLInputElement).value" />
          </template>
          <select v-if="containType === 'timePeriod'" :value="period" @change="period = ($event.target as HTMLSelectElement).value as TimePeriod">
            <option v-for="p in TIME_PERIODS" :key="p.period" :value="p.period">{{ p.label }}</option>
          </select>
        </div>
      </div>
      <div v-if="kind === 'rank'" class="sh-row wrap">
        <span>Format values that rank in the:</span>
        <select :value="bottom ? 'bottom' : 'top'" @change="bottom = ($event.target as HTMLSelectElement).value === 'bottom'">
          <option value="top">Top</option>
          <option value="bottom">Bottom</option>
        </select>
        <NumberField :value="rank" :min="1" :max="percent ? 100 : 1000" label="Rank" @change="rank = $event" />
        <Check label="% of the selected range" :checked="percent" @change="percent = $event" />
      </div>
      <div v-if="kind === 'average'" class="sh-row wrap">
        <span>Format values that are:</span>
        <select :value="avgKind" @change="avgKind = ($event.target as HTMLSelectElement).value">
          <option value="above">above</option>
          <option value="below">below</option>
          <option value="above-eq">equal or above</option>
          <option value="below-eq">equal or below</option>
          <option value="above-1">1 std dev above</option>
          <option value="below-1">1 std dev below</option>
          <option value="above-2">2 std dev above</option>
          <option value="below-2">2 std dev below</option>
          <option value="above-3">3 std dev above</option>
          <option value="below-3">3 std dev below</option>
        </select>
        <span>the average for the selected range</span>
      </div>
      <div v-if="kind === 'unique'" class="sh-row wrap">
        <span>Format all:</span>
        <select :value="unique" @change="unique = ($event.target as HTMLSelectElement).value as typeof unique">
          <option value="duplicateValues">duplicate</option>
          <option value="uniqueValues">unique</option>
        </select>
        <span>values in the selected range</span>
      </div>
      <Field v-if="kind === 'formula'" label="Format values where this formula is true:" wide>
        <RangeField :value="formula" label="Formula" test-id="cf-formula" @change="formula = $event" />
      </Field>
      <FormatEditor v-if="needsFormat" :value="format" @change="format = $event" />
    </fieldset>
  </Dialog>
</template>
