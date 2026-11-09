<script setup lang="ts">
/** A check box with a label (View tab). */
import { type Command, commandChecked, commandDisabled, commandTip, runRibbonCommand } from '@sheetline/core/views';
import type { ControlEnv } from './env';

defineProps<{ cmd: Command; env: ControlEnv }>();
</script>

<template>
  <label :class="`sh-check${commandDisabled(cmd, env.ctx) ? ' disabled' : ''}`" :title="commandTip(cmd)">
    <input
      type="checkbox"
      :checked="commandChecked(cmd, env.ctx)"
      :disabled="commandDisabled(cmd, env.ctx)"
      :data-testid="`rb-${cmd.id}`"
      :aria-label="cmd.tip ?? cmd.label"
      @change="runRibbonCommand(cmd, env.ctx)"
    />
    <span>{{ cmd.label }}</span>
  </label>
</template>
