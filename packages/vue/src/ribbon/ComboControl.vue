<script setup lang="ts">
/** A ribbon combo box (font name, font size, number format): a text field with a list. */
import type { ComboOption, Control } from '@sheetline/core/views';
import { computed, ref, watch } from 'vue';
import Dropdown from '../Dropdown.vue';
import Icon from '../Icon.vue';
import type { ControlEnv } from './env';

const props = defineProps<{ c: Extract<Control, { type: 'combo' }>; env: ControlEnv }>();

const value = computed(() => props.c.value(props.env.ctx));
const text = ref(value.value);
const focused = ref(false);
const wrap = ref<HTMLDivElement | null>(null);
const opened = computed(() => props.env.open?.id === props.c.id);
watch([value, focused], ([v, f]) => {
  if (!f) text.value = v;
});
const disabled = computed(() => props.c.disabled?.(props.env.ctx) || props.env.ctx.state.readOnly);
const options = computed<ComboOption[]>(() => (opened.value ? props.c.options(props.env.ctx) : []));

const commit = (v: string) => {
  const { c, env } = props;
  focused.value = false;
  if (v && v !== value.value) c.onChange(env.ctx, v);
  env.ctx.grid?.focusIdle();
};
const toggle = () => props.env.setOpen(opened.value ? null : { id: props.c.id, rect: wrap.value!.getBoundingClientRect(), kind: 'menu' });
const focus = (e: FocusEvent) => {
  focused.value = true;
  (e.target as HTMLInputElement).select();
};
const down = (e: MouseEvent) => {
  if (!props.c.editable) {
    e.preventDefault();
    toggle();
  }
};
const key = (e: KeyboardEvent) => {
  if (e.key === 'Enter') {
    e.preventDefault();
    commit(text.value.trim());
  } else if (e.key === 'Escape') {
    e.preventDefault();
    text.value = value.value;
    props.env.ctx.grid?.focusIdle();
  }
};
const choose = (v: string) => {
  props.env.setOpen(null);
  commit(v);
};
</script>

<template>
  <div ref="wrap" :class="`sh-combo${opened ? ' open' : ''}`" :style="{ width: `${c.width}px` }" :title="c.label">
    <input
      :aria-label="c.label"
      :data-testid="`rb-${c.id}`"
      :value="text"
      :disabled="disabled"
      :readonly="!c.editable"
      spellcheck="false"
      @focus="focus"
      @blur="focused = false"
      @input="text = ($event.target as HTMLInputElement).value"
      @mousedown="down"
      @keydown="key"
    />
    <button type="button" class="sh-combo-arrow" :aria-label="`${c.label} list`" :disabled="disabled" :data-testid="`rb-${c.id}-list`" @mousedown.prevent @click="toggle">
      <Icon name="arrow_drop_down" />
    </button>
    <Dropdown v-if="opened" :anchor="env.open!.rect" :sheet="env.phone" :title="c.label" extra-class="sh-combo-list" :test-id="`rb-${c.id}-options`" @close="env.setOpen(null)">
      <div role="listbox" :aria-label="c.label">
        <button
          v-for="o in options"
          :key="o.value"
          type="button"
          role="option"
          :aria-selected="o.value === value"
          :class="`sh-combo-item${o.value === value ? ' selected' : ''}`"
          :style="o.font ? { fontFamily: o.font } : undefined"
          @mousedown.prevent
          @click="choose(o.value)"
        >{{ o.label ?? o.value }}</button>
      </div>
    </Dropdown>
  </div>
</template>
