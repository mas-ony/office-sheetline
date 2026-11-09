<script setup lang="ts">
/** The controls that do not fit the toolbar row. */
import type { Control } from '@sheetline/core/views';
import { computed, shallowRef } from 'vue';
import Dropdown from '../Dropdown.vue';
import ControlView from './ControlView.vue';
import type { ControlEnv, Open } from './env';

const props = defineProps<{ controls: Control[]; env: ControlEnv }>();

const inner = shallowRef<Open>(null);
const innerEnv = computed<ControlEnv>(() => ({
  ...props.env,
  open: inner.value,
  setOpen: (o) => (inner.value = o),
  compact: false,
  done: () => props.env.setOpen(null),
}));
const close = () => {
  if (!inner.value) props.env.setOpen(null);
};
</script>

<template>
  <Dropdown :anchor="env.open!.rect" :sheet="env.phone" extra-class="sh-group-popup" test-id="rb-toolbar-more-panel" @close="close">
    <div class="sh-group-popup-body column">
      <ControlView v-for="(c, i) in controls" :key="i" :c="c" :env="innerEnv" />
    </div>
  </Dropdown>
</template>
