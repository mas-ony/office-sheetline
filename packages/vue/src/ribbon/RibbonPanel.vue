<script setup lang="ts">
/** The content of a ribbon drop-down panel: the galleries of paste options, styles, shapes and themes. */
import {
  applyCellStyle,
  applyShapeStyle,
  CELL_STYLE_GROUPS,
  CELL_STYLES,
  type CellStyleDef,
  cellStylePreview,
  changeShapeTo,
  cmd,
  type Ctx,
  galleryStyles,
  insertShape,
  type PanelSpec,
  PASTE_GROUPS,
  SHAPE_STYLES,
  type ShapeGalleryItem,
  sparklineStyles,
  tableStyleLabel,
  tableStylePreviewSvg,
  themeTiles,
} from '@sheetline/core/views';
import Icon from '../Icon.vue';
import ShapesPanel from './ShapesPanel.vue';

const props = defineProps<{ spec: PanelSpec; ctx: Ctx }>();
const emit = defineEmits<{ close: [] }>();

const TABLE_CATEGORIES = ['Light', 'Medium', 'Dark'] as const;
type PasteOptions = Parameters<typeof cmd.paste>[1];
type ShapeStyle = (typeof SHAPE_STYLES)[number];

/** Closes the drop-down, then applies the pick. */
const pick = (apply: () => void) => {
  emit('close');
  apply();
};
const paste = (opts: PasteOptions) => pick(() => void cmd.paste(props.ctx, opts));
const pasteSpecial = () => pick(() => props.ctx.openDialog({ kind: 'paste-special' }));
const tableTiles = (cat: (typeof TABLE_CATEGORIES)[number], showNone?: boolean) => [
  ...(cat === 'Light' && showNone ? [{ name: '', label: 'None' }] : []),
  ...galleryStyles(cat).map((s) => ({ name: s.name, label: tableStyleLabel(s.name) })),
];
const pickTable = (name: string) => {
  const { spec, ctx } = props;
  if (spec.kind === 'table-styles') pick(() => spec.pick(ctx, name));
};
const cellStyleTestId = (name: string) => `cell-style-${name.replace(/[^A-Za-z0-9]+/g, '-').toLowerCase()}`;
const pickCellStyle = (s: CellStyleDef) => pick(() => applyCellStyle(props.ctx, s));
const pickShapeStyle = (st: ShapeStyle) => pick(() => applyShapeStyle(props.ctx, st));
const pickShape = (it: ShapeGalleryItem, change: boolean) => (change ? changeShapeTo(props.ctx, it) : insertShape(props.ctx, it));
</script>

<template>
  <div v-if="spec.kind === 'paste'" class="sh-paste-panel" data-testid="paste-panel">
    <div v-for="g in PASTE_GROUPS" :key="g.title">
      <div class="sh-panel-title">{{ g.title }}</div>
      <div class="sh-icon-grid">
        <button v-for="it in g.items" :key="it.id" type="button" class="sh-icon-tile" :title="it.label" :aria-label="it.label" :data-testid="`paste-${it.id}`" @click="paste(it.opts)">
          <Icon name="content_paste" class="base" />
          <Icon :name="it.icon" class="badge" />
        </button>
      </div>
    </div>
    <div class="sh-panel-sep" />
    <button type="button" class="sh-panel-link" :disabled="!ctx.doc.hasClip()" data-testid="paste-special" @click="pasteSpecial">Paste Special...</button>
  </div>
  <div v-else-if="spec.kind === 'table-styles'" class="sh-gallery-panel" data-testid="table-styles">
    <div v-for="cat in TABLE_CATEGORIES" :key="cat">
      <div class="sh-panel-title">{{ cat }}</div>
      <div class="sh-style-grid">
        <button
          v-for="t in tableTiles(cat, spec.showNone)"
          :key="t.name || 'none'"
          type="button"
          class="sh-style-tile"
          :title="t.label"
          :aria-label="t.label"
          :data-testid="`table-style-${t.name || 'none'}`"
          @click="pickTable(t.name)"
          v-html="tableStylePreviewSvg(t.name, ctx.doc.wb.theme, 58)"
        />
      </div>
    </div>
  </div>
  <div v-else-if="spec.kind === 'cell-styles'" class="sh-gallery-panel cell-styles" data-testid="cell-styles">
    <div v-for="g in CELL_STYLE_GROUPS" :key="g">
      <div class="sh-panel-title">{{ g }}</div>
      <div class="sh-cellstyle-grid">
        <button
          v-for="s in CELL_STYLES.filter((x) => x.group === g)"
          :key="s.name"
          type="button"
          class="sh-cellstyle"
          :style="cellStylePreview(ctx, s)"
          :title="s.name"
          :data-testid="cellStyleTestId(s.name)"
          @click="pickCellStyle(s)"
        >{{ s.name }}</button>
      </div>
    </div>
  </div>
  <ShapesPanel v-else-if="spec.kind === 'shapes'" test-id="shapes-panel" item-test-ids @close="emit('close')" @pick="pickShape($event, false)" />
  <ShapesPanel v-else-if="spec.kind === 'change-shape'" test-id="change-shape-panel" @close="emit('close')" @pick="pickShape($event, true)" />
  <div v-else-if="spec.kind === 'shape-styles'" class="sh-shape-styles" data-testid="shape-styles-panel">
    <button
      v-for="st in SHAPE_STYLES"
      :key="st.label"
      type="button"
      :title="st.label"
      class="sh-shape-style"
      :style="{ background: st.fill ?? 'transparent', borderColor: st.line, color: st.font }"
      @click="pickShapeStyle(st)"
    >Abc</button>
  </div>
  <div v-else-if="spec.kind === 'themes'" class="sh-themes-panel" data-testid="themes-panel">
    <div class="sh-panel-title">Built-In</div>
    <div class="sh-theme-grid">
      <button v-for="t in themeTiles(ctx)" :key="t.name" type="button" :class="`sh-theme-tile${t.active ? ' active' : ''}`" :title="t.name" :data-testid="t.testId" @click="pick(t.apply)">
        <span class="sh-theme-preview" :style="{ background: t.background, color: t.color }">
          <span class="sh-theme-aa" :style="{ fontFamily: t.font }">Aa</span>
          <span class="sh-theme-bars"><i v-for="(x, i) in t.accents" :key="i" :style="{ background: x }" /></span>
        </span>
        <span class="sh-theme-name">{{ t.name }}</span>
      </button>
    </div>
  </div>
  <div v-else-if="spec.kind === 'sparkline-styles'" class="sh-spark-styles" data-testid="sparkline-styles">
    <button v-for="st in sparklineStyles(ctx)" :key="st.key" type="button" :title="st.title" @click="pick(st.apply)">
      <svg width="44" height="22" viewBox="0 0 44 22"><polyline points="2,16 9,8 16,12 23,4 30,14 37,9 42,11" fill="none" :stroke="st.stroke" stroke-width="1.5" /></svg>
    </button>
  </div>
</template>
