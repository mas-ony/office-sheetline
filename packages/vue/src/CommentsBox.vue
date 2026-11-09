<script setup lang="ts">
/** The comment card of a cell, next to it. */
import { cellName } from '@sheetline/core/views';
import Anchored from './Anchored.vue';
import CommentThreadCard from './CommentThreadCard.vue';

defineProps<{ r: number; c: number; rect: DOMRect; startNew?: boolean }>();
const emit = defineEmits<{ close: [] }>();
</script>

<template>
  <Anchored :rect="rect" box-class="sh-comment-card" test-id="comments-box" :label="`Comment ${cellName(r, c)}`" @close="emit('close')">
    <CommentThreadCard :r="r" :c="c" :start-new="startNew" @done="emit('close')" />
  </Anchored>
</template>
