<script setup lang="ts">
/** A ribbon group with its label and launcher, or the button of a collapsed group. */
import { type Group, runRibbonCommand } from '@sheetline/core/views';
import { computed, ref } from 'vue';
import Icon from '../Icon.vue';
import ControlView from './ControlView.vue';
import type { ControlEnv } from './env';
import GroupPopup from './GroupPopup.vue';

const props = defineProps<{ g: Group; env: ControlEnv; collapsed: boolean }>();

const el = ref<HTMLDivElement | null>(null);
const opened = computed(() => props.env.open?.id === `group:${props.g.id}`);
const toggle = () => props.env.setOpen(opened.value ? null : { id: `group:${props.g.id}`, rect: el.value!.getBoundingClientRect(), kind: 'group' });
</script>

<template>
  <div v-if="collapsed" ref="el" class="sh-group collapsed" :data-group="g.id">
    <div :class="`sh-btn large${opened ? ' open' : ''}`" :title="g.label">
      <button type="button" class="sh-btn-main" :aria-label="g.label" aria-haspopup="dialog" :aria-expanded="opened" :data-testid="`rb-group-${g.id}`" @mousedown.prevent @click="toggle">
        <Icon :name="g.icon" class="large" />
        <span class="sh-btn-label">{{ g.label }}<Icon name="arrow_drop_down" class="caret" /></span>
      </button>
    </div>
    <GroupPopup v-if="opened" :g="g" :env="env" />
  </div>
  <div v-else ref="el" class="sh-group" :data-group="g.id">
    <div class="sh-group-body">
      <ControlView v-for="(c, i) in g.controls" :key="i" :c="c" :env="env" />
    </div>
    <div class="sh-group-label">
      <span>{{ g.label }}</span>
      <button
        v-if="g.launcher"
        type="button"
        class="sh-launcher"
        :title="g.launcher.tip ?? g.launcher.label"
        :aria-label="g.launcher.tip ?? g.launcher.label"
        :data-testid="`rb-${g.launcher.id}`"
        @mousedown.prevent
        @click="runRibbonCommand(g.launcher!, env.ctx)"
      >
        <svg viewBox="0 0 8 8" width="8" height="8" aria-hidden="true"><path d="M0 0h1v7h7v1H0zm3 2 1-1 3 3V2h1v4H4V5h2z" fill="currentColor" /></svg>
      </button>
    </div>
  </div>
</template>
