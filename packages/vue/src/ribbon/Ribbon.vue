<script setup lang="ts">
/**
 * The ribbon: File button, tabs (with contextual tabs), and the active tab's groups. Groups that do
 * not fit collapse into group buttons (from the right), as on phones where every group is a
 * button opening a bottom sheet. The toolbar layout shows one row of small buttons instead.
 */
import type { Tab } from '@sheetline/core/views';
import { computed, onMounted, onUnmounted, onUpdated, ref, shallowRef, watch } from 'vue';
import { useUI } from '../context';
import Icon from '../Icon.vue';
import type { ControlEnv, Open } from './env';
import GroupView from './GroupView.vue';
import ToolbarRow from './ToolbarRow.vue';

const props = defineProps<{
  tabs: Tab[];
  active: string;
  layout: 'ribbon' | 'toolbar';
  collapsed: boolean;
}>();
const emit = defineEmits<{ tab: [id: string]; file: []; collapse: [value: boolean] }>();
defineSlots<{
  /** Right side of the tab row (theme and layout switches). */
  extra?: () => unknown;
}>();

const ui = useUI();
const phone = computed(() => ui.value.phone);
const open = shallowRef<Open>(null);
const setOpen = (o: Open) => (open.value = o);
const visibleTabs = computed(() => props.tabs.filter((t) => !t.visible || t.visible(ui.value)));
const tab = computed(() => visibleTabs.value.find((t) => t.id === props.active) ?? visibleTabs.value.find((t) => t.id === 'home') ?? visibleTabs.value[0]);
const groups = computed(() => tab.value.groups(ui.value));
const env = computed<ControlEnv>(() => ({ ctx: ui.value, open: open.value, setOpen, compact: false, phone: phone.value }));
const content = ref<HTMLDivElement | null>(null);
const collapsedCount = ref(phone.value ? groups.value.length : 0);
const peek = ref(false);

// Adaptive layout: collapse groups from the right until the tab fits.
const reset = () => (collapsedCount.value = phone.value ? groups.value.length : 0);
watch([() => tab.value.id, phone, () => groups.value.length], reset);
const fit = () => {
  const el = content.value;
  if (!el || phone.value || props.layout !== 'ribbon') return;
  if (el.scrollWidth > el.clientWidth + 1 && collapsedCount.value < groups.value.length) collapsedCount.value++;
};
onMounted(fit);
onUpdated(fit);
onMounted(() => window.addEventListener('resize', reset));
onUnmounted(() => window.removeEventListener('resize', reset));
watch([() => tab.value.id, () => props.layout], () => (open.value = null));

const clickTab = (t: Tab) => {
  if (props.collapsed) peek.value = t.id !== tab.value.id || !peek.value;
  emit('tab', t.id);
};
const leave = () => {
  if (props.collapsed && !open.value) peek.value = false;
};
const toggleCollapsed = () => {
  peek.value = false;
  emit('collapse', !props.collapsed);
};
</script>

<template>
  <div :class="`sh-ribbon ${layout}${collapsed ? ' collapsed' : ''}${phone ? ' phone' : ''}`" data-testid="ribbon" role="toolbar" aria-label="Ribbon">
    <div class="sh-tabrow">
      <button type="button" class="sh-file-btn" data-testid="file-button" @mousedown.prevent @click="emit('file')">File</button>
      <div class="sh-tabs" role="tablist">
        <button
          v-for="t in visibleTabs"
          :key="t.id"
          type="button"
          role="tab"
          :aria-selected="t.id === tab.id"
          :class="`sh-tab${t.id === tab.id ? ' active' : ''}${t.contextual ? ' contextual' : ''}`"
          :data-testid="`tab-${t.id}`"
          @mousedown.prevent
          @click="clickTab(t)"
          @dblclick="emit('collapse', !collapsed)"
        >{{ t.label }}</button>
      </div>
      <div class="sh-tabrow-extra"><slot name="extra" /></div>
    </div>
    <div v-if="!collapsed || peek" :class="`sh-ribbon-panel${collapsed ? ' peek' : ''}`" @mouseleave="leave">
      <ToolbarRow v-if="layout === 'toolbar'" :tab="tab" :env="env" />
      <div v-else ref="content" class="sh-ribbon-content" :data-testid="`ribbon-${tab.id}`">
        <GroupView v-for="(g, i) in groups" :key="g.id" :g="g" :env="env" :collapsed="i >= groups.length - collapsedCount" />
      </div>
      <button
        v-if="!phone"
        type="button"
        class="sh-collapse"
        :aria-label="collapsed ? 'Pin the ribbon' : 'Collapse the ribbon'"
        :title="collapsed ? 'Pin the ribbon' : 'Collapse the ribbon'"
        data-testid="ribbon-collapse"
        @mousedown.prevent
        @click="toggleCollapsed"
      >
        <Icon :name="collapsed ? 'push_pin' : 'expand_less'" />
      </button>
    </div>
  </div>
</template>
