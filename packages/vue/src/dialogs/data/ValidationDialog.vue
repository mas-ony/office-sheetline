<script setup lang="ts">
/** Data Validation: criteria, input message and error alert for the selection. */
import { commit, type DataValidation, setSelectionValidation, VALIDATION_ALLOW_OPTIONS, VALIDATION_OPERATOR_OPTIONS, validationAt, validationFormula, type ValidationOperator } from '@sheetline/core/views';
import { computed, ref } from 'vue';
import { useUI } from '../../context';
import Check from '../Check.vue';
import Dialog from '../Dialog.vue';
import Field from '../Field.vue';
import RangeField from '../RangeField.vue';
import Tabs from '../Tabs.vue';

const emit = defineEmits<{ close: [] }>();

type Tab = 'settings' | 'input' | 'error';
const TABS: { id: Tab; label: string }[] = [
  { id: 'settings', label: 'Settings' },
  { id: 'input', label: 'Input Message' },
  { id: 'error', label: 'Error Alert' },
];

const ui = useUI();
const doc0 = ui.value.doc;
const existing = validationAt(doc0.sheet, doc0.selection.active.r, doc0.selection.active.c);
const tab = ref<Tab>('settings');
const type = ref(existing?.type ?? 'none');
const op = ref<ValidationOperator>((existing?.operator as ValidationOperator) ?? 'between');
const f1 = ref(existing?.formula1 ? (existing.type === 'list' && existing.formula1.startsWith('"') ? existing.formula1.slice(1, -1) : `=${existing.formula1}`) : '');
const f2 = ref(existing?.formula2 ? `=${existing.formula2}` : '');
const ignoreBlank = ref(existing?.allowBlank ?? true);
const dropDown = ref(!existing?.hideDropDown);
const showInput = ref(existing?.showInputMessage ?? true);
const promptTitle = ref(existing?.promptTitle ?? '');
const prompt = ref(existing?.prompt ?? '');
const showError = ref(existing?.showErrorMessage ?? true);
const errorStyle = ref<'stop' | 'warning' | 'information'>(existing?.errorStyle ?? 'stop');
const errorTitle = ref(existing?.errorTitle ?? '');
const error = ref(existing?.error ?? '');
const twoValues = computed(() => op.value === 'between' || op.value === 'notBetween');
const usesOp = computed(() => type.value !== 'none' && type.value !== 'list' && type.value !== 'custom');
const valueLabel = computed(() => (type.value === 'list' ? 'Source:' : type.value === 'custom' ? 'Formula:' : usesOp.value && twoValues.value ? 'Minimum:' : 'Value:'));
const val = (e: Event) => (e.target as HTMLInputElement | HTMLTextAreaElement | HTMLSelectElement).value;

const save = () => {
  const ctx = ui.value;
  const { doc } = ctx;
  const { r, c } = doc.selection.active;
  const toFormula = (text: string) => validationFormula(doc, r, c, type.value, text);
  const sel = doc.selection.ranges.map((x) => doc.bounded(x, doc.sheet) ?? x);
  let rule: DataValidation | null = null;
  if (type.value !== 'none' || prompt.value || error.value) {
    if (type.value !== 'none' && type.value !== 'custom' && !f1.value.trim()) {
      ctx.notify(type.value === 'list' ? 'Enter the list items or a reference (Source).' : 'Enter a value.');
      return false;
    }
    rule = {
      sqref: sel,
      type: type.value,
      operator: usesOp.value ? op.value : undefined,
      formula1: type.value === 'none' ? undefined : type.value === 'list' && !f1.value.trim().startsWith('=') ? `"${f1.value.replace(/"/g, '')}"` : toFormula(f1.value),
      formula2: usesOp.value && twoValues.value ? toFormula(f2.value) : undefined,
      allowBlank: ignoreBlank.value,
      hideDropDown: type.value === 'list' ? !dropDown.value : undefined,
      showInputMessage: showInput.value,
      promptTitle: promptTitle.value || undefined,
      prompt: prompt.value || undefined,
      showErrorMessage: showError.value,
      errorStyle: errorStyle.value,
      errorTitle: errorTitle.value || undefined,
      error: error.value || undefined,
    };
  }
  commit(ctx);
  setSelectionValidation(doc, sel, rule);
  return true;
};
const clearAll = () => {
  type.value = 'none';
  f1.value = '';
  f2.value = '';
  prompt.value = '';
  promptTitle.value = '';
  error.value = '';
  errorTitle.value = '';
};
</script>

<template>
  <Dialog title="Data Validation" :on-ok="save" :width="460" test-id="validation-dialog" @close="emit('close')">
    <Tabs :value="tab" :tabs="TABS" test-id="validation-tabs" @change="tab = $event" />
    <div class="sh-tab-body">
      <div v-if="tab === 'settings'" class="sh-stack">
        <div class="sh-label">Validation criteria</div>
        <div class="sh-row wrap">
          <Field label="Allow:">
            <select :value="type" data-testid="validation-allow" @change="type = val($event)">
              <option v-for="a in VALIDATION_ALLOW_OPTIONS" :key="a.value" :value="a.value">{{ a.label }}</option>
            </select>
          </Field>
          <Check label="Ignore blank" :checked="ignoreBlank" :disabled="type === 'none'" @change="ignoreBlank = $event" />
        </div>
        <Field v-if="usesOp" label="Data:">
          <select :value="op" data-testid="validation-operator" @change="op = val($event) as ValidationOperator">
            <option v-for="o in VALIDATION_OPERATOR_OPTIONS" :key="o.value" :value="o.value">{{ o.label }}</option>
          </select>
        </Field>
        <Check v-if="type === 'list'" label="In-cell dropdown" :checked="dropDown" @change="dropDown = $event" />
        <Field v-if="type !== 'none'" :label="valueLabel" wide>
          <RangeField :value="f1" label="Value" test-id="validation-f1" @change="f1 = $event" />
        </Field>
        <Field v-if="usesOp && twoValues" label="Maximum:" wide>
          <RangeField :value="f2" label="Maximum" test-id="validation-f2" @change="f2 = $event" />
        </Field>
        <p v-if="type === 'list'" class="sh-hint">Separate items with commas (Yes,No,Maybe) or refer to cells (=$A$1:$A$5).</p>
      </div>
      <div v-if="tab === 'input'" class="sh-stack">
        <Check label="Show input message when cell is selected" :checked="showInput" @change="showInput = $event" />
        <Field label="Title:" wide>
          <input :value="promptTitle" maxlength="32" @input="promptTitle = val($event)" />
        </Field>
        <Field label="Input message:" wide>
          <textarea :rows="5" :value="prompt" maxlength="255" data-testid="validation-prompt" @input="prompt = val($event)" />
        </Field>
      </div>
      <div v-if="tab === 'error'" class="sh-stack">
        <Check label="Show error alert after invalid data is entered" :checked="showError" @change="showError = $event" />
        <div class="sh-row wrap top">
          <Field label="Style:">
            <select :value="errorStyle" data-testid="validation-style" @change="errorStyle = val($event) as typeof errorStyle">
              <option value="stop">Stop</option>
              <option value="warning">Warning</option>
              <option value="information">Information</option>
            </select>
          </Field>
          <div class="sh-stack grow">
            <Field label="Title:" wide>
              <input :value="errorTitle" maxlength="32" @input="errorTitle = val($event)" />
            </Field>
            <Field label="Error message:" wide>
              <textarea :rows="4" :value="error" maxlength="255" data-testid="validation-error" @input="error = val($event)" />
            </Field>
          </div>
        </div>
      </div>
    </div>
    <template #footer-left>
      <button type="button" class="sh-btn" data-testid="validation-clear" @click="clearAll">Clear All</button>
    </template>
  </Dialog>
</template>
