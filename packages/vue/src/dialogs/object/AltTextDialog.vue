<script setup lang="ts">
/** Alt Text of a chart, picture or shape. */
import { computed, ref } from 'vue';
import { useUI } from '../../context';
import Dialog from '../Dialog.vue';
import Field from '../Field.vue';

const props = defineProps<{ id: string }>();
const emit = defineEmits<{ close: [] }>();

const ui = useUI();
const hit = computed(() => ui.value.doc.objectById(props.id));
const start = hit.value;
const title = ref(start && start.kind !== 'image' ? (start.obj.altTitle ?? '') : '');
const descr = ref(start ? (start.obj.descr ?? '') : '');
const ok = () => ui.value.doc.setAltText(props.id, title.value, descr.value);
</script>

<template>
  <Dialog v-if="hit" title="Alt Text" :width="400" test-id="alt-text" :on-ok="ok" @close="emit('close')">
    <p class="sh-hint">How would you describe this object and its context to someone who is blind or low vision?</p>
    <Field v-if="hit.kind !== 'image'" label="Title" wide>
      <input :value="title" data-testid="alt-text-title" @input="title = ($event.target as HTMLInputElement).value" />
    </Field>
    <Field label="Description" wide>
      <textarea :rows="5" :value="descr" data-testid="alt-text-description" @input="descr = ($event.target as HTMLTextAreaElement).value" />
    </Field>
  </Dialog>
</template>
