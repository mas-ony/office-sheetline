<script setup lang="ts">
/** The toolbar layout: one row of small buttons; what does not fit goes in a "More commands" drop-down. */
import { type Tab, toolbarControls } from '@sheetline/core/views';
import { computed, onMounted, onUnmounted, onUpdated, ref, watch } from 'vue';
import Icon from '../Icon.vue';
import ControlView from './ControlView.vue';
import type { ControlEnv } from './env';
import OverflowPopup from './OverflowPopup.vue';

const props = defineProps<{ tab: Tab; env: ControlEnv }>();

// The controls follow the context, not the open drop-down.
const ctx = computed(() => props.env.ctx);
const controls = computed(() => toolbarControls(props.tab, ctx.value));
const row = ref<HTMLDivElement | null>(null);
const moreBtn = ref<HTMLButtonElement | null>(null);
const visible = ref(controls.value.length);
watch(controls, (c) => (visible.value = c.length));
const shrink = () => {
  const el = row.value;
  if (!el || props.env.phone) return;
  if (el.scrollWidth > el.clientWidth + 1 && visible.value > 3) visible.value--;
};
onMounted(shrink);
onUpdated(shrink);
const reset = () => (visible.value = controls.value.length);
onMounted(() => window.addEventListener('resize', reset));
onUnmounted(() => window.removeEventListener('resize', reset));

const compactEnv = computed<ControlEnv>(() => ({ ...props.env, compact: true }));
const hidden = computed(() => controls.value.slice(visible.value));
const opened = computed(() => props.env.open?.id === 'toolbar-more');
const toggle = () => props.env.setOpen(opened.value ? null : { id: 'toolbar-more', rect: moreBtn.value!.getBoundingClientRect(), kind: 'more' });
</script>

<template>
  <div ref="row" class="sh-toolbar-row">
    <ControlView v-for="(c, i) in controls.slice(0, visible)" :key="i" :c="c" :env="compactEnv" />
    <div v-if="hidden.length > 0" class="sh-btn small">
      <button ref="moreBtn" type="button" class="sh-btn-main" aria-label="More commands" data-testid="rb-toolbar-more" @mousedown.prevent @click="toggle">
        <Icon name="more_horiz" />
      </button>
      <OverflowPopup v-if="opened" :controls="hidden" :env="env" />
    </div>
  </div>
</template>
