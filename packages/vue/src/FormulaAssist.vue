<script setup lang="ts">
/** Function suggestions while a function name is typed, or the arguments of the function at the caret. */
import { highlightedArg, type UIContextValue } from '@sheetline/core/views';
import { computed, onMounted, onUpdated, ref, shallowRef, watch } from 'vue';
import { useUI } from './context';

const ui = useUI();
const box = ref<HTMLDivElement | null>(null);
const pos = shallowRef<{ left: number; top: number } | null>(null);

function anchorRect(ctx: UIContextValue): DOMRect | null {
  if (ctx.edit.origin === 'bar') return ctx.ui.rootElement?.querySelector<HTMLElement>('[data-testid="formula-input"]')?.getBoundingClientRect() ?? null;
  return ctx.grid?.editorClientRect() ?? null;
}

const place = () => {
  const ctx = ui.value;
  if (!ctx.ui.assist.view.visible) {
    if (pos.value) pos.value = null;
    return;
  }
  const a = anchorRect(ctx);
  const el = box.value;
  if (!a || !el) return;
  const vh = window.visualViewport?.height ?? window.innerHeight;
  const vw = window.innerWidth;
  const h = el.offsetHeight;
  const w = el.offsetWidth;
  let top = a.bottom + 2;
  if (top + h > vh - 4 && a.top - h - 2 > 4) top = a.top - h - 2;
  const left = Math.max(4, Math.min(a.left, vw - w - 4));
  if (!pos.value || pos.value.left !== left || pos.value.top !== top) pos.value = { left, top };
};
onMounted(place);
onUpdated(place);

const scrollActive = () => box.value?.querySelector('[aria-selected="true"]')?.scrollIntoView({ block: 'nearest' });
onMounted(scrollActive);
watch([() => ui.value.ui.assist.view.index, () => ui.value.ui.assist.view.showList], scrollActive, { flush: 'post' });

const boxStyle = computed(() => ({ position: 'fixed' as const, left: `${pos.value ? pos.value.left : -9999}px`, top: `${pos.value ? pos.value.top : 0}px` }));
</script>

<template>
  <Teleport v-if="ui.ui.assist.view.visible" :to="ui.layer">
    <div v-if="ui.ui.assist.view.showList" ref="box" class="sh-suggest" :style="boxStyle" role="listbox" aria-label="Functions" data-testid="formula-suggestions">
      <div
        v-for="(s, i) in ui.ui.assist.view.suggestions"
        :key="s.name"
        role="option"
        :aria-selected="i === ui.ui.assist.view.index"
        :class="`sh-suggest-item${i === ui.ui.assist.view.index ? ' active' : ''}`"
        :title="s.description"
        @mousedown.prevent
        @pointerdown.prevent
        @click="ui.ui.assist.apply(s)"
      >{{ s.name }}</div>
    </div>
    <div v-else-if="ui.ui.assist.view.call" ref="box" class="sh-arghint" :style="boxStyle" data-testid="formula-hint">
      <div class="sig">{{ ui.ui.assist.view.call.info.name }}(<span v-for="(a, i) in ui.ui.assist.view.call.info.args" :key="i">{{ i > 0 ? ', ' : '' }}<span :class="i === highlightedArg(ui.ui.assist.view.call.info, ui.ui.assist.view.call.argIndex) ? 'cur' : undefined">{{ a }}</span></span>)</div>
      <template v-if="ui.ui.assist.view.call.info.description">
        <div class="label">Summary</div>
        <div class="desc">{{ ui.ui.assist.view.call.info.description }}</div>
      </template>
    </div>
    <span v-else />
  </Teleport>
</template>
