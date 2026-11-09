<script setup lang="ts">
/** The Format pane of a shape: fill, line, text, rotation and size. */
import type { ColorSpec, SheetShape } from '@sheetline/core/views';
import { computed } from 'vue';
import { useUI } from '../context';
import Check from '../dialogs/Check.vue';
import ColorButton from '../dialogs/ColorButton.vue';
import Field from '../dialogs/Field.vue';
import { vChecked } from '../directives';
import { hexOf, specOf } from './format';
import LazyInput from './LazyInput.vue';
import LineEditor from './LineEditor.vue';
import PaneSection from './PaneSection.vue';
import SidePanel from './SidePanel.vue';
import SizeSection from './SizeSection.vue';

const props = defineProps<{ shape: SheetShape }>();

const ui = useUI();
// A new object on every change (the shape may be updated in place).
const view = computed(() => ({ ctx: ui.value, shape: props.shape }));
const update = (patch: Partial<SheetShape>) => ui.value.doc.updateShape(props.shape.id, patch);
const hex = (c: ColorSpec | null) => hexOf(ui.value, c);
const setSize = (v: string) => {
  if (Number(v) >= 1) update({ fontSize: Math.min(409, Number(v)) });
};
const setRotation = (v: string) => {
  if (Number.isFinite(Number(v))) update({ rotation: ((Number(v) % 360) + 360) % 360 || undefined });
};
</script>

<template>
  <SidePanel title="Format Shape" test-id="format-pane">
    <PaneSection title="Fill">
      <div class="sh-row">
        <label class="sh-check"><input v-checked="!view.shape.fill" type="radio" @change="update({ fill: null })" /> <span>No fill</span></label>
        <label class="sh-check"><input v-checked="!!view.shape.fill" type="radio" @change="update({ fill: view.shape.fill ?? '#5B9BD5' })" /> <span>Solid fill</span></label>
      </div>
      <Field v-if="view.shape.fill" label="Color:">
        <ColorButton :value="specOf(view.shape.fill)" none-label="No Fill" test-id="shape-fill-color" @change="update({ fill: hex($event) })" />
      </Field>
    </PaneSection>
    <PaneSection title="Line">
      <LineEditor :line="view.shape.line" @change="update({ line: $event })" />
    </PaneSection>
    <PaneSection title="Text">
      <Field label="Text:" wide>
        <LazyInput :value="view.shape.text ?? ''" label="Shape text" test-id="shape-text-pane" @commit="update({ text: $event || undefined })" />
      </Field>
      <Field label="Color:">
        <ColorButton :value="specOf(view.shape.fontColor)" none-label="Automatic" @change="update({ fontColor: hex($event) ?? '#000000' })" />
      </Field>
      <Field label="Size (pt):">
        <LazyInput :value="String(view.shape.fontSize ?? 11)" label="Font size" input-mode="decimal" @commit="setSize" />
      </Field>
      <div class="sh-row">
        <Check label="Bold" :checked="!!view.shape.bold" @change="update({ bold: $event || undefined })" />
        <Check label="Italic" :checked="!!view.shape.italic" @change="update({ italic: $event || undefined })" />
      </div>
    </PaneSection>
    <PaneSection title="Rotation">
      <Field label="Degrees:">
        <LazyInput :value="String(view.shape.rotation ?? 0)" label="Rotation" input-mode="decimal" @commit="setRotation" />
      </Field>
    </PaneSection>
    <SizeSection :id="view.shape.id" />
  </SidePanel>
</template>
