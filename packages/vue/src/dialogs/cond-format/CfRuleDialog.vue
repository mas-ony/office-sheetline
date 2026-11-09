<script setup lang="ts">
/** Home > Conditional Formatting > New Rule / More Rules: adds the rule to the selection. */
import { cfInitialRule, type CfRule, commit } from '@sheetline/core/views';
import { useUI } from '../../context';
import RuleEditorDialog from './RuleEditorDialog.vue';

const props = defineProps<{ hint?: string }>();
const emit = defineEmits<{ close: [] }>();

const ui = useUI();
const initial = cfInitialRule(props.hint);
const save = (rule: CfRule) => {
  const ctx = ui.value;
  commit(ctx);
  ctx.doc.addCondFormat(rule);
};
</script>

<template>
  <RuleEditorDialog title="New Formatting Rule" :initial="initial" :hint="hint" @close="emit('close')" @save="save" />
</template>
