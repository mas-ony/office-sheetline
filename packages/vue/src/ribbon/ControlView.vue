<script setup lang="ts">
/** One control of a ribbon group. */
import type { Control } from '@sheetline/core/views';
import CheckControl from './CheckControl.vue';
import ColorControl from './ColorControl.vue';
import ComboControl from './ComboControl.vue';
import CommandButton from './CommandButton.vue';
import type { ControlEnv } from './env';
import RibbonView from './RibbonView.vue';

defineProps<{ c: Control; env: ControlEnv }>();
</script>

<template>
  <CommandButton v-if="c.type === 'large'" :cmd="c.cmd" :env="env" size="large" :split="c.split" />
  <CommandButton v-else-if="c.type === 'small'" :cmd="c.cmd" :env="env" size="small" :split="c.split" :icon-only="c.iconOnly" />
  <div v-else-if="c.type === 'stack'" class="sh-stack">
    <ControlView v-for="(x, i) in c.items" :key="i" :c="x" :env="env" />
  </div>
  <div v-else-if="c.type === 'row'" class="sh-row">
    <ControlView v-for="(x, i) in c.items" :key="i" :c="x" :env="env" />
  </div>
  <ComboControl v-else-if="c.type === 'combo'" :c="c" :env="env" />
  <ColorControl v-else-if="c.type === 'color'" :c="c" :env="env" />
  <CheckControl v-else-if="c.type === 'check'" :cmd="c.cmd" :env="env" />
  <RibbonView v-else-if="c.type === 'view'" :spec="c.view" :ctx="env.ctx" />
</template>
