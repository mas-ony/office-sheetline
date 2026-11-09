<script setup lang="ts">
/** New Name, or Edit Name of an existing name. */
import { rangeText } from '@sheetline/core/views';
import { ref } from 'vue';
import { useUI } from '../../context';
import { vFocus } from '../../directives';
import Dialog from '../Dialog.vue';
import Field from '../Field.vue';
import RangeField from '../RangeField.vue';

const props = defineProps<{ name?: string; scope?: number }>();
const emit = defineEmits<{ close: [] }>();

const ui = useUI();
const { doc } = ui.value;
const existing = props.name !== undefined ? doc.wb.names.find((n) => n.name === props.name && n.localSheet === props.scope) : undefined;
function suggestion(): string {
  if (existing) return existing.name;
  const { r, c } = doc.selection.active;
  // Excel suggests the text of the cell above or to the left.
  const near = [doc.displayText(r, c), r > 0 ? doc.displayText(r - 1, c) : '', c > 0 ? doc.displayText(r, c - 1) : ''].find((t) => /^[A-Za-z_][\w.]*$/.test(t.replace(/\s+/g, '_')) && Number.isNaN(Number(t)));
  return near ? near.replace(/\s+/g, '_') : '';
}
const n = ref(suggestion());
const sc = ref<number | -1>(props.scope ?? -1);
const comment = ref(existing?.comment ?? '');
const refText = ref(existing ? `=${existing.formula}` : `=${doc.selection.ranges.map((x) => `${doc.sheet.name.includes(' ') || /[^\w]/.test(doc.sheet.name) ? `'${doc.sheet.name.replace(/'/g, "''")}'` : doc.sheet.name}!${rangeText(doc.bounded(x, doc.sheet) ?? x, doc.sheet)}`).join(',')}`);

const ok = () => {
  const error = ui.value.doc.defineName({ name: n.value, formula: refText.value, localSheet: sc.value === -1 ? undefined : sc.value, comment: comment.value }, existing ? { name: existing.name, localSheet: existing.localSheet } : undefined);
  if (error) {
    ui.value.notify(error);
    return false;
  }
  return true;
};
</script>

<template>
  <Dialog :title="existing ? 'Edit Name' : 'New Name'" :width="420" test-id="new-name" :on-ok="ok" @close="emit('close')">
    <Field label="Name:" wide>
      <input v-focus :value="n" data-testid="new-name-name" @input="n = ($event.target as HTMLInputElement).value" />
    </Field>
    <Field label="Scope:">
      <select :value="sc" data-testid="new-name-scope" @change="sc = Number(($event.target as HTMLSelectElement).value)">
        <option :value="-1">Workbook</option>
        <option v-for="(s, i) in ui.doc.wb.sheets" :key="s.uid" :value="i">{{ s.name }}</option>
      </select>
    </Field>
    <Field label="Comment:" wide>
      <textarea :rows="3" :value="comment" @input="comment = ($event.target as HTMLTextAreaElement).value" />
    </Field>
    <Field label="Refers to:" wide>
      <RangeField :value="refText" label="Refers to" test-id="new-name-ref" multi @change="refText = $event.startsWith('=') ? $event : `=${$event}`" />
    </Field>
  </Dialog>
</template>
