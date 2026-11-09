<script setup lang="ts">
/** The full-page app: the editor with the options and page connections of the shell. */
import { SheetlineEditor } from '@sheetline/vue';
import { connectShell, shellOptions } from '@sheetline/web-shared/shell';

const { onThemeChange, ...options } = shellOptions();
</script>

<template>
  <SheetlineEditor v-bind="options" @ready="connectShell" @update:theme="onThemeChange" />
</template>
