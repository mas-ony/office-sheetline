<script setup lang="ts">
/** Edit Text (or Text Box) of a shape: the text, its size, bold and italic. */
import { computed, ref } from 'vue';
import { useUI } from '../../context';
import { vFocus } from '../../directives';
import Check from '../Check.vue';
import Dialog from '../Dialog.vue';
import Field from '../Field.vue';

const props = defineProps<{ id: string }>();
const emit = defineEmits<{ close: [] }>();

const ui = useUI();
const shape = computed(() => ui.value.doc.sheet.shapes.find((s) => s.id === props.id));
const start = shape.value;
const text = ref(start?.text ?? '');
const size = ref(String(start?.fontSize ?? 11));
const bold = ref(!!start?.bold);
const italic = ref(!!start?.italic);
const ok = () => {
  const s = shape.value;
  if (!s) return;
  const n = Number(size.value);
  ui.value.doc.updateShape(props.id, { text: text.value || undefined, fontSize: Number.isFinite(n) && n >= 1 && n <= 409 ? n : s.fontSize, bold: bold.value || undefined, italic: italic.value || undefined });
};
</script>

<template>
  <Dialog v-if="shape" :title="shape.textBox ? 'Text Box' : 'Edit Text'" :width="420" test-id="shape-text" :on-ok="ok" @close="emit('close')">
    <textarea v-focus :rows="6" :value="text" data-testid="shape-text-input" aria-label="Text" @input="text = ($event.target as HTMLTextAreaElement).value" />
    <div class="sh-row">
      <Field label="Size:">
        <input :value="size" inputmode="decimal" :style="{ width: '60px' }" @input="size = ($event.target as HTMLInputElement).value" />
      </Field>
      <Check label="Bold" :checked="bold" @change="bold = $event" />
      <Check label="Italic" :checked="italic" @change="italic = $event" />
    </div>
  </Dialog>
</template>
