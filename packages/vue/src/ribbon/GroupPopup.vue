<script setup lang="ts">
/** The controls of a collapsed group, in a drop-down (desktop) or a bottom sheet (phone). */
import { type Group, runRibbonCommand } from '@sheetline/core/views';
import { computed, shallowRef } from 'vue';
import Dropdown from '../Dropdown.vue';
import ControlView from './ControlView.vue';
import type { ControlEnv, Open } from './env';

const props = defineProps<{ g: Group; env: ControlEnv }>();

// Nested drop-downs open on top of this popup; keep their state separate.
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
const more = () => {
  props.env.setOpen(null);
  runRibbonCommand(props.g.launcher!, props.env.ctx);
};
</script>

<template>
  <Dropdown :anchor="env.open!.rect" :sheet="env.phone" :title="g.label" extra-class="sh-group-popup" :test-id="`rb-group-${g.id}-panel`" @close="close">
    <div :class="`sh-group-popup-body${env.phone ? ' touch' : ''}`">
      <ControlView v-for="(c, i) in g.controls" :key="i" :c="c" :env="innerEnv" />
      <button v-if="g.launcher" type="button" class="sh-group-more" @click="more">{{ g.launcher.label }}</button>
    </div>
  </Dropdown>
</template>
