<script setup lang="ts">
/** Scenario Summary: a report sheet of the scenarios' changing cells and the result cells. */
import { cellsOfRangeText, rangeText } from '@sheetline/core/views';
import { ref } from 'vue';
import { useUI } from '../../context';
import Dialog from '../Dialog.vue';
import Field from '../Field.vue';
import Radio from '../Radio.vue';
import RangeField from '../RangeField.vue';

const emit = defineEmits<{ close: [] }>();

const ui = useUI();
const result = ref(rangeText(ui.value.doc.range, ui.value.doc.sheet));
const ok = () => {
  const { doc } = ui.value;
  const cells = cellsOfRangeText(doc, result.value);
  const sheet = doc.sheet;
  const index = doc.scenarioSummary(cells.slice(0, 32), sheet);
  if (index >= 0) doc.setActiveSheet(index);
  return true;
};
</script>

<template>
  <Dialog title="Scenario Summary" :width="360" test-id="scenario-summary" :on-ok="ok" @close="emit('close')">
    <fieldset class="sh-group">
      <legend>Report type</legend>
      <Radio name="summary-type" label="Scenario summary" :checked="true" />
    </fieldset>
    <Field label="Result cells:" wide>
      <RangeField :value="result" label="Result cells" multi @change="result = $event" />
    </Field>
  </Dialog>
</template>
