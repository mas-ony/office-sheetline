<script setup lang="ts">
/** The initial of a comment's author on a colour of their name. */
import { avatarOf } from '@sheetline/core/views';
import { computed } from 'vue';

const props = defineProps<{ name: string }>();
const avatar = computed(() => avatarOf(props.name));
</script>

<template>
  <span class="sh-avatar" :style="{ background: avatar.color }" aria-hidden="true">{{ avatar.initial }}</span>
</template>
