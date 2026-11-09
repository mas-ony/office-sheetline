<script setup lang="ts">
/** Dialog frame: title bar with close button, OK / Cancel at the bottom right. */
import type { IconName } from '@sheetline/core/views';
import Modal from '../Modal.vue';

const props = withDefaults(
  defineProps<{
    title: string;
    icon?: IconName;
    /** OK: return false to keep the dialog open. */
    onOk?: () => boolean | void | Promise<boolean | void>;
    ok?: string;
    cancel?: string;
    width?: number;
    /** Modal dialogs dim the page; small dialogs let the grid be used meanwhile. */
    blocking?: boolean;
    testId?: string;
    okDisabled?: boolean;
    bodyClass?: string;
    /** No OK button (Close only). */
    closeOnly?: boolean;
  }>(),
  { ok: 'OK', cancel: 'Cancel', width: 420, blocking: true },
);
const emit = defineEmits<{ close: [] }>();
defineSlots<{
  default?: () => unknown;
  /** Buttons at the left of the footer (Delete, Clear All...). */
  'footer-left'?: () => unknown;
}>();

let busy = false;
const submit = async () => {
  if (busy || props.okDisabled) return;
  busy = true;
  try {
    const keep = (await props.onOk?.()) === false;
    if (!keep) emit('close');
  } finally {
    busy = false;
  }
};
const frameWidth = () => Math.min(props.width, window.innerWidth - 16);
</script>

<template>
  <Modal :title="title" :icon="icon" :width="frameWidth()" :blocking="blocking" :test-id="testId" :body-class="`sh-dialog-body${bodyClass ? ` ${bodyClass}` : ''}`" @close="emit('close')">
    <form class="sh-dialog-form" @submit.prevent="submit">
      <slot />
      <button type="submit" hidden aria-hidden="true" tabindex="-1" />
    </form>
    <template #footer>
      <span v-if="$slots['footer-left']" class="sh-dialog-foot-left"><slot name="footer-left" /></span>
      <button v-if="!closeOnly" type="button" class="sh-btn primary" :data-testid="testId ? `${testId}-ok` : undefined" :disabled="okDisabled" @click="submit">{{ ok }}</button>
      <button type="button" class="sh-btn" :data-testid="testId ? `${testId}-cancel` : undefined" @click="emit('close')">{{ closeOnly ? 'Close' : cancel }}</button>
    </template>
  </Modal>
</template>
