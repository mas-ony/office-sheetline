<script setup lang="ts">
/** Create Sparklines, Edit Sparklines and Edit Sparkline Data: the data and the location ranges. */
import { absRef, commit, localRef, makeRange, parseRangeText, rangeText, resolveRef, selectionText } from '@sheetline/core/views';
import { ref } from 'vue';
import { useUI } from '../../context';
import Dialog from '../Dialog.vue';
import Field from '../Field.vue';
import RangeField from '../RangeField.vue';

const props = defineProps<{ type: 'line' | 'column' | 'stacked'; groupId?: string; single?: boolean }>();
const emit = defineEmits<{ close: [] }>();

const ui = useUI();
const { doc } = ui.value;
const group = props.groupId ? doc.sheet.sparklineGroups.find((g) => g.id === props.groupId) : undefined;
const { r, c } = doc.selection.active;
const current = group?.sparklines.find((s) => s.r === r && s.c === c);
type Group = NonNullable<typeof group>;

function groupData(g: Group): string {
  const refs = g.sparklines.map((s) => resolveRef(doc, s.ref)).filter((x): x is NonNullable<ReturnType<typeof resolveRef>> => !!x);
  if (!refs.length) return '';
  const box = refs.reduce((b, x) => ({ r0: Math.min(b.r0, x.range.r0), c0: Math.min(b.c0, x.range.c0), r1: Math.max(b.r1, x.range.r1), c1: Math.max(b.c1, x.range.c1) }), refs[0].range);
  return rangeText(box, refs[0].sheet, doc.sheet);
}
function groupLocation(g: Group): string {
  const box = g.sparklines.reduce((b, s) => ({ r0: Math.min(b.r0, s.r), c0: Math.min(b.c0, s.c), r1: Math.max(b.r1, s.r), c1: Math.max(b.c1, s.c) }), { r0: Infinity, c0: Infinity, r1: -Infinity, c1: -Infinity });
  return rangeText(box, doc.sheet);
}
function defaultLocation(): string {
  const sel = doc.range;
  // Default location: the empty column right of the data.
  return rangeText(makeRange(sel.r0, sel.c1 + 1, sel.r1, sel.c1 + 1), doc.sheet);
}

const data = ref(props.single && current ? localRef(current.ref, doc.sheet) : group ? groupData(group) : selectionText(doc));
const location = ref(props.single && current ? rangeText(makeRange(current.r, current.c), doc.sheet) : group ? groupLocation(group) : defaultLocation());

const ok = () => {
  const ctx = ui.value;
  const { doc } = ctx;
  const d = parseRangeText(doc, data.value);
  if (!d) {
    ctx.notify('Data Range reference is not valid.');
    return false;
  }
  if (props.single && group && current) {
    doc.updateSparklineGroup(group.id, { sparklines: group.sparklines.map((s) => (s === current ? { ...s, ref: absRef(d.sheet, d.range) } : s)) });
    return true;
  }
  const loc = parseRangeText(doc, location.value);
  if (!loc || loc.sheet !== doc.sheet) {
    ctx.notify('Location reference is not valid. Select cells on this sheet.');
    return false;
  }
  commit(ctx);
  const error = doc.addSparklines(group?.type ?? props.type, d, loc.range);
  if (error) {
    ctx.notify(error);
    return false;
  }
  if (group) {
    // Keep the group's formatting on the rebuilt group.
    const added = doc.sheet.sparklineGroups[doc.sheet.sparklineGroups.length - 1];
    const { id: _id, sparklines: _s, ...look } = group;
    void _id;
    void _s;
    doc.updateSparklineGroup(added.id, look);
  }
  return true;
};
</script>

<template>
  <Dialog :title="group ? (single ? 'Edit Sparkline Data' : 'Edit Sparklines') : 'Create Sparklines'" :width="420" test-id="sparkline-dialog" :on-ok="ok" @close="emit('close')">
    <p class="sh-lead">Choose the data that you want</p>
    <Field label="Data Range:" wide>
      <RangeField :value="data" label="Data Range" test-id="sparkline-data" auto-focus @change="data = $event" />
    </Field>
    <template v-if="!single">
      <p class="sh-lead">Choose where you want the sparklines to be placed</p>
      <Field label="Location Range:" wide>
        <RangeField :value="location" label="Location Range" test-id="sparkline-location" @change="location = $event" />
      </Field>
    </template>
  </Dialog>
</template>
