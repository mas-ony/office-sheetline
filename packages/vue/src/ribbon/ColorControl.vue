<script setup lang="ts">
/** Split colour button: the icon applies the last colour, the arrow opens the palette. */
import { type ColorSpec, commandDisabled, type Control, resolveColor } from '@sheetline/core/views';
import { computed, ref, shallowRef } from 'vue';
import Dropdown from '../Dropdown.vue';
import Icon from '../Icon.vue';
import ThemePalette from '../ThemePalette.vue';
import type { ControlEnv } from './env';

const props = defineProps<{ c: Extract<Control, { type: 'color' }>; env: ControlEnv }>();

const last = shallowRef<{ spec: ColorSpec | null; css: string }>({ spec: null, css: props.c.initial });
const el = ref<HTMLDivElement | null>(null);
const opened = computed(() => props.env.open?.id === props.c.cmd.id);
const disabled = computed(() => commandDisabled(props.c.cmd, props.env.ctx) || props.env.ctx.state.readOnly);

const apply = (spec: ColorSpec | null) => {
  const { c, env } = props;
  const { ctx } = env;
  env.setOpen(null);
  if (ctx.grid?.isEditing()) ctx.grid.commitEdit();
  c.apply(ctx, spec);
  const css = spec ? `#${resolveColor(spec, ctx.doc.wb.theme, '000000')}` : c.noneLabel === 'No Fill' ? 'transparent' : c.initial;
  last.value = { spec, css };
  ctx.grid?.focusIdle();
};
const applyLast = () => {
  const { c, env } = props;
  const { ctx } = env;
  if (ctx.grid?.isEditing()) ctx.grid.commitEdit();
  c.apply(ctx, last.value.spec ?? (c.noneLabel === 'No Fill' ? { rgb: 'FFFFFF00' } : null));
};
const toggle = () => props.env.setOpen(opened.value ? null : { id: props.c.cmd.id, rect: el.value!.getBoundingClientRect(), kind: 'color' });
</script>

<template>
  <div ref="el" :class="`sh-btn small split color${opened ? ' open' : ''}`" :title="c.cmd.tip ?? c.cmd.label">
    <button type="button" class="sh-btn-main" :aria-label="c.cmd.label" :disabled="disabled" :data-testid="`rb-${c.cmd.id}`" @mousedown.prevent @click="applyLast">
      <Icon v-if="c.cmd.icon" :name="c.cmd.icon" />
      <span class="sh-color-bar" :style="{ background: last.css }" />
    </button>
    <button type="button" class="sh-btn-arrow" :aria-label="`${c.cmd.label} palette`" :disabled="disabled" :data-testid="`rb-${c.cmd.id}-more`" @mousedown.prevent @click="toggle">
      <Icon name="arrow_drop_down" class="caret" />
    </button>
    <Dropdown v-if="opened" :anchor="env.open!.rect" :sheet="env.phone" :title="c.cmd.label" :test-id="`rb-${c.cmd.id}-palette`" @close="env.setOpen(null)">
      <ThemePalette
        :theme="env.ctx.doc.wb.theme"
        :none-label="c.noneLabel"
        :none-swatch="c.noneLabel === 'Automatic' ? '#000000' : undefined"
        :test-id="`palette-${c.cmd.id}`"
        @pick="apply"
      />
    </Dropdown>
  </div>
</template>
