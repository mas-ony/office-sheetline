<script setup lang="ts">
/** A section of the Format pane that opens and closes. */
import { ref } from 'vue';
import Icon from '../Icon.vue';

const props = withDefaults(defineProps<{ title: string; open?: boolean }>(), { open: true });
const isOpen = ref(props.open);
</script>

<template>
  <section :class="`sh-pane-section${isOpen ? ' open' : ''}`">
    <button type="button" class="sh-pane-section-head" :aria-expanded="isOpen" @click="isOpen = !isOpen"><Icon :name="isOpen ? 'expand_more' : 'chevron_right'" />{{ title }}</button>
    <div v-if="isOpen" class="sh-pane-section-body"><slot /></div>
  </section>
</template>
