<script setup lang="ts">
/** A ribbon button: runs its command, or opens its menu or panel (split buttons do both). */
import { type Command, commandChecked, commandDisabled, commandTip, runRibbonCommand } from '@sheetline/core/views';
import { computed, ref } from 'vue';
import Dropdown from '../Dropdown.vue';
import Icon from '../Icon.vue';
import Menu from '../Menu.vue';
import type { ControlEnv } from './env';
import RibbonPanel from './RibbonPanel.vue';

const props = withDefaults(defineProps<{ cmd: Command; env: ControlEnv; size: 'large' | 'small'; split?: boolean; iconOnly?: boolean }>(), {
  split: false,
  // Unset differs from false: compact buttons with an icon hide their label unless it is false.
  iconOnly: undefined,
});

const el = ref<HTMLDivElement | null>(null);
const hasDrop = computed(() => !!(props.cmd.menu || props.cmd.panel));
const disabled = computed(() => {
  const { cmd, env } = props;
  return commandDisabled(cmd, env.ctx) || (env.ctx.state.readOnly && !cmd.id.startsWith('view') && !cmd.id.startsWith('file'));
});
const checked = computed(() => commandChecked(props.cmd, props.env.ctx));
const opened = computed(() => props.env.open?.id === props.cmd.id);
/** A split button with a command runs it from its main part and opens its drop-down from the arrow. */
const isSplit = computed(() => props.split && hasDrop.value && !!props.cmd.run);
const dropOnMain = computed(() => hasDrop.value && !(props.split && props.cmd.run));
const large = computed(() => props.size === 'large' && !props.env.compact);
const showLabel = computed(() => !props.iconOnly && !(props.env.compact && props.size === 'small' && !!props.cmd.icon && props.iconOnly !== false));
const cls = computed(() => `sh-btn ${props.env.compact ? 'small' : props.size}${checked.value ? ' checked' : ''}${opened.value ? ' open' : ''}${isSplit.value ? ' split' : ''}`);

const openDrop = () => {
  const { cmd, env } = props;
  if (opened.value) {
    env.setOpen(null);
    return;
  }
  env.setOpen({ id: cmd.id, rect: el.value!.getBoundingClientRect(), kind: cmd.panel ? 'panel' : 'menu' });
};
const onMain = () => {
  const { cmd, env } = props;
  if (disabled.value) return;
  if (dropOnMain.value) openDrop();
  else {
    runRibbonCommand(cmd, env.ctx);
    if (!cmd.checked) env.done?.();
  }
};
const menuClosed = (picked?: boolean) => {
  const { env } = props;
  env.setOpen(null);
  if (picked) env.done?.();
  env.ctx.grid?.focusIdle();
};
const panelClosed = () => {
  const { env } = props;
  env.setOpen(null);
  env.done?.();
  env.ctx.grid?.focusIdle();
};
</script>

<template>
  <div ref="el" :class="cls" :title="commandTip(cmd)">
    <button
      type="button"
      class="sh-btn-main"
      :aria-label="cmd.label"
      :aria-pressed="cmd.checked ? checked : undefined"
      :aria-haspopup="dropOnMain ? 'menu' : undefined"
      :aria-expanded="dropOnMain ? opened : undefined"
      :disabled="disabled"
      :data-testid="`rb-${cmd.id}`"
      @mousedown.prevent
      @click="onMain"
    >
      <Icon v-if="cmd.icon" :name="cmd.icon" :color="disabled ? undefined : cmd.tint" :class="large ? 'large' : undefined" />
      <template v-if="large">
        <span v-if="!isSplit" class="sh-btn-label">{{ cmd.label }}<Icon v-if="hasDrop" name="arrow_drop_down" class="caret" /></span>
      </template>
      <span v-else-if="showLabel || !cmd.icon" class="sh-btn-label">{{ cmd.label }}</span>
      <Icon v-if="!large && dropOnMain" name="arrow_drop_down" class="caret" />
    </button>
    <button
      v-if="isSplit"
      type="button"
      class="sh-btn-arrow"
      :aria-label="`${cmd.label} options`"
      aria-haspopup="menu"
      :aria-expanded="opened"
      :disabled="disabled"
      :data-testid="`rb-${cmd.id}-more`"
      @mousedown.prevent
      @click="openDrop"
    >
      <span v-if="large" class="sh-btn-label">{{ cmd.label }}</span>
      <Icon name="arrow_drop_down" class="caret" />
    </button>
    <Menu
      v-if="opened && env.open!.kind === 'menu' && cmd.menu"
      :items="cmd.menu(env.ctx)"
      :at="env.open!.rect"
      :sheet="env.phone"
      :title="cmd.label"
      :test-id="`rb-${cmd.id}-menu`"
      @close="menuClosed"
    />
    <Dropdown v-if="opened && env.open!.kind === 'panel' && cmd.panel" :anchor="env.open!.rect" :sheet="env.phone" :title="cmd.label" :test-id="`rb-${cmd.id}-panel`" @close="env.setOpen(null)">
      <RibbonPanel :spec="cmd.panel" :ctx="env.ctx" @close="panelClosed" />
    </Dropdown>
  </div>
</template>
