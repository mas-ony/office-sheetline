<script setup lang="ts">
/**
 * Format Cells: Number (categories with options and a sample), Alignment, Font, Border (presets,
 * edges, line style and colour), Fill (colour, pattern) and Protection. Only what was changed is
 * applied to the selection.
 */
import {
  BORDER_LINE_STYLES,
  type BorderEdge,
  type BorderEdges,
  type BorderLine,
  type BorderLineStyle,
  type ColorSpec,
  commit,
  CURRENCY_SYMBOLS,
  CUSTOM_FORMATS,
  DATE_FORMAT_TYPES,
  FILL_PATTERNS,
  FONT_NAMES,
  FONT_SIZES,
  type FormatTab,
  formatValue,
  FRACTION_FORMAT_TYPES,
  H_ALIGN_OPTIONS,
  type HAlign,
  NUMBER_CATEGORIES,
  NUMBER_CATEGORY_DESCRIPTIONS,
  type NumberCategory,
  numberCode,
  type NumberFormatState,
  numberStateOf,
  resolveColor,
  SPECIAL_FORMAT_TYPES,
  type StylePatch,
  textRotationDegrees,
  TIME_FORMAT_TYPES,
  UNDERLINE_OPTIONS,
  type UnderlineType,
  V_ALIGN_OPTIONS,
  type VAlign,
} from '@sheetline/core/views';
import { computed, type CSSProperties, ref, shallowRef } from 'vue';
import { useUI } from '../../context';
import Check from '../Check.vue';
import ColorButton from '../ColorButton.vue';
import Dialog from '../Dialog.vue';
import Field from '../Field.vue';
import ListBox from '../ListBox.vue';
import NumberField from '../NumberField.vue';
import Tabs from '../Tabs.vue';
import LineStyleSvg from './LineStyleSvg.vue';

const props = withDefaults(defineProps<{ tab?: FormatTab }>(), { tab: 'number' });
const emit = defineEmits<{ close: [] }>();

const TABS: { id: FormatTab; label: string }[] = [
  { id: 'number', label: 'Number' },
  { id: 'alignment', label: 'Alignment' },
  { id: 'font', label: 'Font' },
  { id: 'border', label: 'Border' },
  { id: 'fill', label: 'Fill' },
  { id: 'protection', label: 'Protection' },
];
type FontStyle = 'Regular' | 'Italic' | 'Bold' | 'Bold Italic';
const FONT_STYLES: FontStyle[] = ['Regular', 'Italic', 'Bold', 'Bold Italic'];
const SIZES = FONT_SIZES.map(String);
const DECIMAL_CATEGORIES = ['Number', 'Currency', 'Accounting', 'Percentage', 'Scientific'];
const TOP_EDGES: BorderEdge[] = ['top', 'insideH', 'bottom'];
const BOTTOM_EDGES: BorderEdge[] = ['diagUp', 'left', 'insideV', 'right', 'diagDown'];
const DIAL_DOTS = [-90, -45, 0, 45, 90];

const ui = useUI();
const { doc } = ui.value;
const { r, c } = doc.selection.active;
const style = doc.styleAt(r, c);
const theme = doc.wb.theme;
const tab = ref<FormatTab>(props.tab);
const dirty = new Set<string>();
const touch = (k: string) => dirty.add(k);
const same = (x: string) => x;
const val = (e: Event) => (e.target as HTMLInputElement).value;

// Number
const num = shallowRef<NumberFormatState>(numberStateOf(style.numFmt));
const setN = (p: Partial<NumberFormatState>) => {
  num.value = { ...num.value, ...p };
  touch('numFmt');
};
const code = computed(() => numberCode(num.value));
const cellValue = doc.engine.cellValue(doc.sheet, r, c);
const sample = computed(() => {
  const { category } = num.value;
  const sampleValue = typeof cellValue === 'number' ? cellValue : category === 'Date' || category === 'Time' ? 40982.5632 : -1234.1;
  try {
    const v = category === 'Text' && typeof cellValue === 'string' ? cellValue : sampleValue;
    return formatValue(v, code.value, doc.wb.date1904);
  } catch {
    return { text: String(sampleValue), color: undefined };
  }
});
const pickCategory = (cat: NumberCategory) => {
  const next: Partial<NumberFormatState> = { category: cat };
  if (cat === 'Date') next.code = DATE_FORMAT_TYPES[0].code;
  if (cat === 'Time') next.code = TIME_FORMAT_TYPES[0].code;
  if (cat === 'Fraction') next.code = FRACTION_FORMAT_TYPES[0].code;
  if (cat === 'Special') next.code = SPECIAL_FORMAT_TYPES[0].code;
  if (cat === 'Custom') next.code = code.value;
  if (cat === 'Currency' || cat === 'Accounting') next.symbol = num.value.symbol || '"$"';
  setN(next);
};
const negatives = computed(() => [0, 1, 2, 3].map((i) => formatValue(-1234.1, numberCode({ ...num.value, negative: i }), doc.wb.date1904)));
const typeList = computed(() => {
  const { category } = num.value;
  return category === 'Date' ? DATE_FORMAT_TYPES : category === 'Time' ? TIME_FORMAT_TYPES : category === 'Fraction' ? FRACTION_FORMAT_TYPES : category === 'Special' ? SPECIAL_FORMAT_TYPES : null;
});
const customCodes = [...new Set([style.numFmt, ...CUSTOM_FORMATS])];
const codeOf = (t: { code: string }) => t.code;
const numberKey = (x: number) => String(x);

// Alignment
const align = style.align ?? {};
const hAlign = ref<HAlign>(align.horizontal ?? 'general');
const vAlign = ref<VAlign>(align.vertical ?? 'bottom');
const indent = ref<number | ''>(align.indent ?? 0);
const wrap = ref(!!align.wrap);
const shrink = ref(!!align.shrink);
const merged = ui.value.state.format.merged;
const merge = ref(merged);
const stacked = ref(align.rotation === 255);
const degrees = ref<number | ''>(textRotationDegrees(align.rotation));
const readingOrder = ref(align.readingOrder ?? 0);
const setWrap = (v: boolean) => {
  wrap.value = v;
  if (v) shrink.value = false;
  touch('align');
};
const dial = (e: MouseEvent) => {
  const rect = (e.currentTarget as SVGElement).getBoundingClientRect();
  const x = e.clientX - rect.left - 55;
  const y = rect.top + 55 - e.clientY;
  if (x < 0) return;
  degrees.value = Math.round((Math.atan2(y, x) * 180) / Math.PI);
  stacked.value = false;
  touch('align');
};
const dialAngle = computed(() => ((Number(degrees.value) || 0) * Math.PI) / 180);

// Font
const font = style.font;
const fontName = font.scheme === 'minor' ? theme.minorFont : font.scheme === 'major' ? theme.majorFont : font.name;
const fName = ref(fontName);
const fStyle = ref<FontStyle>(font.bold && font.italic ? 'Bold Italic' : font.bold ? 'Bold' : font.italic ? 'Italic' : 'Regular');
const fSize = ref(String(font.size));
const underline = ref<UnderlineType | 'none'>(font.underline ?? 'none');
const fColor = shallowRef<ColorSpec | null>(font.color ?? null);
const strike = ref(!!font.strike);
const vert = ref<'superscript' | 'subscript' | ''>(font.vertAlign ?? '');
const fonts = [...new Set([theme.majorFont, theme.minorFont, ...FONT_NAMES])];
const quoted = (x: string) => `"${x}"`;
const fontLabel = (x: string) => (x === theme.majorFont ? `${x} (Headings)` : x === theme.minorFont ? `${x} (Body)` : x);
const fontChanged = (apply: () => void) => {
  apply();
  touch('font');
};

// Border
const lineStyle = ref<BorderLineStyle>('thin');
const lineColor = shallowRef<ColorSpec | null>(null);
const multiRow = doc.range.r1 > doc.range.r0;
const multiCol = doc.range.c1 > doc.range.c0;
const own = style.border ?? {};
const edges = shallowRef<BorderEdges>({});
const edgeOn = (e: BorderEdge): boolean => {
  const x = edges.value;
  if (x[e] !== undefined) return !!x[e];
  if (e === 'top') return !!own.top;
  if (e === 'bottom') return !!own.bottom;
  if (e === 'left') return !!own.left;
  if (e === 'right') return !!own.right;
  if (e === 'diagUp') return !!own.diagonalUp;
  if (e === 'diagDown') return !!own.diagonalDown;
  return false;
};
const line = (): BorderLine => ({ style: lineStyle.value, color: lineColor.value ?? undefined });
const toggleEdge = (e: BorderEdge) => {
  const x = edges.value;
  edges.value = { ...x, [e]: edgeOn(e) && (x[e] === undefined || x[e]?.style === lineStyle.value) ? null : line() };
};
const preset = (p: 'none' | 'outline' | 'inside') => {
  const x = edges.value;
  if (p === 'none') edges.value = { top: null, bottom: null, left: null, right: null, insideH: null, insideV: null, diagUp: null, diagDown: null };
  else if (p === 'outline') edges.value = { ...x, top: line(), bottom: line(), left: line(), right: line() };
  else edges.value = { ...x, insideH: multiRow ? line() : x.insideH, insideV: multiCol ? line() : x.insideV };
};
const edgeGlyph = (e: BorderEdge) => (e === 'top' ? '▔' : e === 'bottom' ? '▁' : e === 'insideH' ? '═' : e === 'diagUp' ? '╱' : e === 'diagDown' ? '╲' : e === 'insideV' ? '║' : e === 'left' ? '▏' : '▕');

// Fill
const fill = style.fill && !('gradientXml' in style.fill) ? style.fill : undefined;
const bg = shallowRef<ColorSpec | null>(fill && fill.pattern !== 'none' ? (fill.pattern === 'solid' ? (fill.fg ?? null) : (fill.bg ?? null)) : null);
const patternColor = shallowRef<ColorSpec | null>(fill && fill.pattern !== 'solid' && fill.pattern !== 'none' ? (fill.fg ?? null) : null);
const pattern = ref<string>(fill && fill.pattern !== 'solid' ? fill.pattern : 'none');
const setBg = (v: ColorSpec | null) => {
  bg.value = v;
  if (!v && pattern.value === 'solid') pattern.value = 'none';
  touch('fill');
};
const patternLabel = (p: string) => (p === 'none' ? 'None' : p.replace(/([A-Z])/g, ' $1').replace(/^./, (m) => m.toUpperCase()));

// Protection
const locked = ref(style.locked !== false);
const hidden = ref(!!style.hidden);

const apply = () => {
  const ctx = ui.value;
  const { doc } = ctx;
  commit(ctx);
  const patch: StylePatch = {};
  if (dirty.has('numFmt')) patch.numFmt = code.value;
  if (dirty.has('align')) {
    const deg = degrees.value;
    const rot = stacked.value ? 255 : typeof deg === 'number' && deg !== 0 ? (deg > 0 ? Math.min(90, deg) : 90 + Math.min(90, -deg)) : null;
    const ind = indent.value;
    patch.align = {
      horizontal: hAlign.value === 'general' ? null : hAlign.value,
      vertical: vAlign.value === 'bottom' ? null : vAlign.value,
      indent: typeof ind === 'number' && ind > 0 && ['left', 'right', 'distributed'].includes(hAlign.value) ? ind : null,
      wrap: wrap.value || null,
      shrink: shrink.value || null,
      rotation: rot,
      readingOrder: readingOrder.value || null,
    };
  }
  if (dirty.has('font')) {
    const size = Number(fSize.value);
    patch.font = {
      name: fName.value,
      scheme: fName.value === fontName && font.scheme ? font.scheme : null,
      size: Number.isFinite(size) && size >= 1 && size <= 409 ? size : font.size,
      bold: fStyle.value.includes('Bold') || null,
      italic: fStyle.value.includes('Italic') || null,
      underline: underline.value === 'none' ? null : underline.value,
      color: fColor.value ?? { theme: 1 },
      strike: strike.value || null,
      vertAlign: vert.value || null,
    };
  }
  if (dirty.has('fill')) {
    if (!bg.value && pattern.value === 'none') patch.fill = null;
    else if (pattern.value === 'none' || pattern.value === 'solid') patch.fill = bg.value ? { pattern: 'solid', fg: bg.value, bg: { indexed: 64 } } : null;
    else patch.fill = { pattern: pattern.value, fg: patternColor.value ?? { rgb: 'FF000000' }, bg: bg.value ?? { rgb: 'FFFFFFFF' } };
  }
  if (dirty.has('protection')) {
    patch.locked = locked.value;
    patch.hidden = hidden.value;
  }
  const e = edges.value;
  const borderChanged = Object.keys(e).length > 0;
  doc.transact('Format Cells', () => {
    if (Object.keys(patch).length) doc.format(patch);
    if (dirty.has('align') && merge.value !== merged) {
      if (merge.value) doc.merge(undefined, undefined, 'all');
      else doc.unmerge();
    }
    if (borderChanged) {
      const set = (kind: 'top' | 'bottom' | 'left' | 'right' | 'horizontal' | 'vertical', v: BorderLine | null | undefined) => {
        if (v !== undefined) doc.setBorder(kind, v);
      };
      set('top', e.top);
      set('bottom', e.bottom);
      set('left', e.left);
      set('right', e.right);
      set('horizontal', e.insideH);
      set('vertical', e.insideV);
      if (e.diagUp !== undefined || e.diagDown !== undefined) {
        const up = e.diagUp !== undefined ? !!e.diagUp : !!own.diagonalUp;
        const down = e.diagDown !== undefined ? !!e.diagDown : !!own.diagonalDown;
        const diag = e.diagUp || e.diagDown || own.diagonal;
        doc.format({ border: { diagonal: up || down ? (diag ?? line()) : null, diagonalUp: up || null, diagonalDown: down || null } }, undefined, undefined, 'Borders');
      }
    }
  });
};

const css = (color: ColorSpec | null | undefined, fb: string) => `#${resolveColor(color ?? undefined, theme, fb)}`;
const previewEdge = (e: BorderEdge) => {
  const v = edges.value[e];
  const ln = v === undefined ? (e === 'top' ? own.top : e === 'bottom' ? own.bottom : e === 'left' ? own.left : e === 'right' ? own.right : e === 'diagUp' || e === 'diagDown' ? own.diagonal : undefined) : v;
  return ln && edgeOn(e) ? ln : null;
};
const edgeStroke = (e: BorderEdge) => {
  const ln = previewEdge(e);
  return ln ? { stroke: css(ln.color, '000000'), 'stroke-width': ln.style === 'thick' ? 3 : ln.style.startsWith('medium') ? 2 : 1, 'stroke-dasharray': /dash|dot|hair/i.test(ln.style) ? '4 2' : undefined } : { stroke: 'transparent' };
};
const previewFont = computed<CSSProperties>(() => ({
  fontFamily: quoted(fName.value),
  fontWeight: fStyle.value.includes('Bold') ? 700 : 400,
  fontStyle: fStyle.value.includes('Italic') ? 'italic' : 'normal',
  fontSize: `${Math.min(36, Number(fSize.value) || 11) * 1.333}px`,
  color: css(fColor.value, '000000'),
  textDecoration: `${underline.value !== 'none' ? 'underline' : ''} ${strike.value ? 'line-through' : ''}`.trim() || undefined,
  textDecorationStyle: underline.value.toLowerCase().includes('double') ? 'double' : undefined,
  verticalAlign: vert.value === 'superscript' ? 'super' : vert.value === 'subscript' ? 'sub' : undefined,
}));
const fillSample = computed(() => {
  const back = bg.value ? css(bg.value, 'FFFFFF') : '#fff';
  return pattern.value === 'none' || pattern.value === 'solid' ? back : `repeating-linear-gradient(45deg, ${css(patternColor.value, '000000')} 0 2px, ${back} 2px 5px)`;
});
</script>

<template>
  <Dialog title="Format Cells" :on-ok="apply" :width="620" test-id="format-cells" @close="emit('close')">
    <Tabs :value="tab" test-id="format-cells-tabs" :tabs="TABS" @change="tab = $event" />
    <div class="sh-tab-body" :style="{ minHeight: '330px' }">
      <div v-if="tab === 'number'" class="sh-two-col">
        <div>
          <div class="sh-label">Category:</div>
          <ListBox v-slot="{ item }" :items="NUMBER_CATEGORIES" :value="num.category" :key-of="same" :height="280" test-id="number-category" label="Category" @change="pickCategory">{{ item }}</ListBox>
        </div>
        <div class="sh-stack">
          <fieldset class="sh-group">
            <legend>Sample</legend>
            <div class="sh-sample" data-testid="number-sample" :style="{ color: sample.color }">{{ sample.text || ' ' }}</div>
          </fieldset>
          <Field v-if="DECIMAL_CATEGORIES.includes(num.category)" label="Decimal places:">
            <NumberField :value="num.decimals" :min="0" :max="30" test-id="number-decimals" label="Decimal places" @change="setN({ decimals: typeof $event === 'number' ? $event : 0 })" />
          </Field>
          <Check v-if="num.category === 'Number'" label="Use 1000 Separator (,)" :checked="num.separator" test-id="number-separator" @change="setN({ separator: $event })" />
          <Field v-if="num.category === 'Currency' || num.category === 'Accounting'" label="Symbol:">
            <select :value="num.symbol" data-testid="number-symbol" @change="setN({ symbol: val($event) })">
              <option v-for="s in CURRENCY_SYMBOLS" :key="s.label" :value="s.code">{{ s.label }}</option>
            </select>
          </Field>
          <template v-if="num.category === 'Number' || num.category === 'Currency'">
            <div class="sh-label">Negative numbers:</div>
            <ListBox v-slot="{ item }" :items="[0, 1, 2, 3]" :value="num.negative" :key-of="numberKey" :height="88" label="Negative numbers" @change="setN({ negative: $event })">
              <span :style="{ color: negatives[item].color }">{{ negatives[item].text }}</span>
            </ListBox>
          </template>
          <template v-if="typeList">
            <div class="sh-label">Type:</div>
            <ListBox v-slot="{ item }" :items="typeList" :value="typeList.find((t) => t.code === num.code) ?? null" :key-of="codeOf" :height="170" test-id="number-type" label="Type" @change="setN({ code: $event.code })">{{ item.label }}</ListBox>
          </template>
          <template v-if="num.category === 'Custom'">
            <Field label="Type:" wide>
              <input :value="num.code" spellcheck="false" data-testid="number-custom" @input="setN({ code: val($event) })" />
            </Field>
            <ListBox v-slot="{ item }" :items="customCodes" :value="num.code" :key-of="same" :height="150" label="Format codes" @change="setN({ code: $event })">{{ item }}</ListBox>
          </template>
          <p class="sh-hint">{{ NUMBER_CATEGORY_DESCRIPTIONS[num.category] }}</p>
        </div>
      </div>
      <div v-if="tab === 'alignment'" class="sh-two-col wide-left">
        <div class="sh-stack">
          <fieldset class="sh-group">
            <legend>Text alignment</legend>
            <Field label="Horizontal:">
              <select :value="hAlign" data-testid="align-horizontal" @change="hAlign = val($event) as HAlign; touch('align')">
                <option v-for="a in H_ALIGN_OPTIONS" :key="a.value" :value="a.value">{{ a.label }}</option>
              </select>
            </Field>
            <Field label="Indent:">
              <NumberField :value="indent" :min="0" :max="250" label="Indent" test-id="align-indent" @change="indent = $event; touch('align')" />
            </Field>
            <Field label="Vertical:">
              <select :value="vAlign" data-testid="align-vertical" @change="vAlign = val($event) as VAlign; touch('align')">
                <option v-for="a in V_ALIGN_OPTIONS" :key="a.value" :value="a.value">{{ a.label }}</option>
              </select>
            </Field>
          </fieldset>
          <fieldset class="sh-group">
            <legend>Text control</legend>
            <Check label="Wrap text" :checked="wrap" test-id="align-wrap" @change="setWrap" />
            <Check label="Shrink to fit" :checked="shrink" :disabled="wrap" test-id="align-shrink" @change="shrink = $event; touch('align')" />
            <Check label="Merge cells" :checked="merge" test-id="align-merge" @change="merge = $event; touch('align')" />
          </fieldset>
          <fieldset class="sh-group">
            <legend>Right-to-left</legend>
            <Field label="Text direction:">
              <select :value="readingOrder" @change="readingOrder = Number(val($event)); touch('align')">
                <option :value="0">Context</option>
                <option :value="1">Left-to-Right</option>
                <option :value="2">Right-to-Left</option>
              </select>
            </Field>
          </fieldset>
        </div>
        <fieldset class="sh-group sh-orientation">
          <legend>Orientation</legend>
          <div class="sh-orient-row">
            <button type="button" :class="`sh-orient-vertical${stacked ? ' active' : ''}`" :aria-pressed="stacked" title="Vertical text" @click="stacked = !stacked; touch('align')">
              <span v-for="(ch, i) in 'Text'.split('')" :key="i">{{ ch }}</span>
            </button>
            <svg class="sh-orient-dial" width="110" height="110" viewBox="-55 -55 110 110" @click="dial">
              <path d="M0 -50 A50 50 0 0 1 0 50" fill="none" stroke="#bbb" />
              <circle v-for="d in DIAL_DOTS" :key="d" :cx="Math.cos((d * Math.PI) / 180) * 50" :cy="-Math.sin((d * Math.PI) / 180) * 50" r="2.5" fill="#888" />
              <line v-if="!stacked" x1="0" y1="0" :x2="Math.cos(dialAngle) * 46" :y2="-Math.sin(dialAngle) * 46" stroke="#217346" stroke-width="2" />
              <text x="6" y="4" font-size="11" fill="#333" :transform="`rotate(${-(Number(degrees) || 0)} 0 0)`">Text ——</text>
            </svg>
          </div>
          <Field label="Degrees:">
            <NumberField :value="degrees" :min="-90" :max="90" label="Degrees" test-id="align-degrees" :disabled="stacked" @change="degrees = $event; touch('align')" />
          </Field>
        </fieldset>
      </div>
      <div v-if="tab === 'font'" class="sh-stack">
        <div class="sh-font-grid">
          <div>
            <Field label="Font:" wide>
              <input :value="fName" data-testid="font-name-input" @input="fontChanged(() => (fName = val($event)))" />
            </Field>
            <ListBox v-slot="{ item }" :items="fonts" :value="fName" :key-of="same" :height="140" label="Fonts" @change="fontChanged(() => (fName = $event))">
              <span :style="{ fontFamily: quoted(item) }">{{ fontLabel(item) }}</span>
            </ListBox>
          </div>
          <div>
            <Field label="Font style:" wide>
              <input :value="fStyle" readonly />
            </Field>
            <ListBox v-slot="{ item }" :items="FONT_STYLES" :value="fStyle" :key-of="same" :height="140" label="Font styles" test-id="font-style" @change="fontChanged(() => (fStyle = $event))">
              <span :style="{ fontWeight: item.includes('Bold') ? 700 : 400, fontStyle: item.includes('Italic') ? 'italic' : 'normal' }">{{ item }}</span>
            </ListBox>
          </div>
          <div>
            <Field label="Size:" wide>
              <input :value="fSize" inputmode="decimal" data-testid="font-size-input" @input="fontChanged(() => (fSize = val($event)))" />
            </Field>
            <ListBox v-slot="{ item }" :items="SIZES" :value="fSize" :key-of="same" :height="140" label="Sizes" @change="fontChanged(() => (fSize = $event))">{{ item }}</ListBox>
          </div>
        </div>
        <div class="sh-row">
          <Field label="Underline:">
            <select :value="underline" data-testid="font-underline" @change="fontChanged(() => (underline = val($event) as UnderlineType | 'none'))">
              <option v-for="u in UNDERLINE_OPTIONS" :key="u.value" :value="u.value">{{ u.label }}</option>
            </select>
          </Field>
          <Field label="Color:">
            <ColorButton :value="fColor" test-id="font-color" none-swatch="#000000" @change="fontChanged(() => (fColor = $event))" />
          </Field>
        </div>
        <div class="sh-row">
          <fieldset class="sh-group">
            <legend>Effects</legend>
            <Check label="Strikethrough" :checked="strike" test-id="font-strike" @change="fontChanged(() => (strike = $event))" />
            <Check label="Superscript" :checked="vert === 'superscript'" @change="fontChanged(() => (vert = $event ? 'superscript' : ''))" />
            <Check label="Subscript" :checked="vert === 'subscript'" @change="fontChanged(() => (vert = $event ? 'subscript' : ''))" />
          </fieldset>
          <fieldset class="sh-group grow">
            <legend>Preview</legend>
            <div class="sh-font-preview" :style="previewFont">AaBbCcYyZz</div>
          </fieldset>
        </div>
      </div>
      <div v-if="tab === 'border'" class="sh-two-col">
        <div class="sh-stack">
          <fieldset class="sh-group">
            <legend>Line</legend>
            <div class="sh-label">Style:</div>
            <div class="sh-line-styles" role="listbox" aria-label="Line style" data-testid="border-styles">
              <button v-for="ls in BORDER_LINE_STYLES" :key="ls" type="button" role="option" :aria-selected="lineStyle === ls" :class="lineStyle === ls ? 'selected' : ''" :title="ls" @click="lineStyle = ls">
                <LineStyleSvg :line-style="ls" :color="css(lineColor, '000000')" />
              </button>
            </div>
            <div class="sh-label">Color:</div>
            <ColorButton :value="lineColor" test-id="border-color" none-swatch="#000000" @change="lineColor = $event" />
          </fieldset>
        </div>
        <div class="sh-stack">
          <fieldset class="sh-group">
            <legend>Presets</legend>
            <div class="sh-border-presets">
              <button type="button" data-testid="border-preset-none" @click="preset('none')">
                <svg width="34" height="34"><rect x="4" y="4" width="26" height="26" fill="none" stroke="#c8c8c8" /></svg>
                <span>None</span>
              </button>
              <button type="button" data-testid="border-preset-outline" @click="preset('outline')">
                <svg width="34" height="34"><rect x="4" y="4" width="26" height="26" fill="none" stroke="#333" stroke-width="2" /></svg>
                <span>Outline</span>
              </button>
              <button type="button" data-testid="border-preset-inside" :disabled="!multiRow && !multiCol" @click="preset('inside')">
                <svg width="34" height="34"><rect x="4" y="4" width="26" height="26" fill="none" stroke="#c8c8c8" /><path d="M17 4 V30 M4 17 H30" stroke="#333" stroke-width="2" /></svg>
                <span>Inside</span>
              </button>
            </div>
          </fieldset>
          <fieldset class="sh-group">
            <legend>Border</legend>
            <div class="sh-border-editor">
              <div class="sh-border-buttons left">
                <button v-for="e in TOP_EDGES" :key="e" type="button" :aria-pressed="edgeOn(e)" :class="edgeOn(e) ? 'on' : ''" :disabled="e === 'insideH' && !multiRow" :title="e" :data-testid="`border-${e}`" @click="toggleEdge(e)">{{ edgeGlyph(e) }}</button>
              </div>
              <svg class="sh-border-preview" width="150" height="110" viewBox="0 0 150 110">
                <rect x="10" y="10" width="130" height="90" fill="#fff" />
                <text x="40" y="38" font-size="12" fill="#999">Text</text>
                <text v-if="multiCol" x="100" y="38" font-size="12" fill="#999">Text</text>
                <text v-if="multiRow" x="40" y="82" font-size="12" fill="#999">Text</text>
                <text v-if="multiRow && multiCol" x="100" y="82" font-size="12" fill="#999">Text</text>
                <line x1="10" y1="10" x2="140" y2="10" v-bind="edgeStroke('top')" />
                <line x1="10" y1="100" x2="140" y2="100" v-bind="edgeStroke('bottom')" />
                <line x1="10" y1="10" x2="10" y2="100" v-bind="edgeStroke('left')" />
                <line x1="140" y1="10" x2="140" y2="100" v-bind="edgeStroke('right')" />
                <line v-if="multiRow" x1="10" y1="55" x2="140" y2="55" v-bind="edgeStroke('insideH')" />
                <line v-if="multiCol" x1="75" y1="10" x2="75" y2="100" v-bind="edgeStroke('insideV')" />
                <line x1="10" y1="100" x2="140" y2="10" v-bind="edgeStroke('diagUp')" />
                <line x1="10" y1="10" x2="140" y2="100" v-bind="edgeStroke('diagDown')" />
                <rect x="10" y="0" width="130" height="14" fill="transparent" @click="toggleEdge('top')" />
                <rect x="10" y="96" width="130" height="14" fill="transparent" @click="toggleEdge('bottom')" />
                <rect x="0" y="10" width="14" height="90" fill="transparent" @click="toggleEdge('left')" />
                <rect x="136" y="10" width="14" height="90" fill="transparent" @click="toggleEdge('right')" />
              </svg>
              <div class="sh-border-buttons bottom">
                <button v-for="e in BOTTOM_EDGES" :key="e" type="button" :aria-pressed="edgeOn(e)" :class="edgeOn(e) ? 'on' : ''" :disabled="e === 'insideV' && !multiCol" :title="e" :data-testid="`border-${e}`" @click="toggleEdge(e)">{{ edgeGlyph(e) }}</button>
              </div>
            </div>
          </fieldset>
          <p class="sh-hint">The selected border style can be applied by clicking the presets, preview diagram or the buttons above.</p>
        </div>
      </div>
      <div v-if="tab === 'fill'" class="sh-two-col">
        <div class="sh-stack">
          <div class="sh-label">Background Color:</div>
          <ColorButton :value="bg" none-label="No Color" test-id="fill-color" @change="setBg" />
        </div>
        <div class="sh-stack">
          <Field label="Pattern Color:">
            <ColorButton :value="patternColor" none-label="Automatic" none-swatch="#000000" @change="patternColor = $event; touch('fill')" />
          </Field>
          <Field label="Pattern Style:">
            <select :value="pattern" data-testid="fill-pattern" @change="pattern = val($event); touch('fill')">
              <option v-for="p in FILL_PATTERNS" :key="p" :value="p">{{ patternLabel(p) }}</option>
            </select>
          </Field>
          <fieldset class="sh-group">
            <legend>Sample</legend>
            <div class="sh-fill-sample" data-testid="fill-sample" :style="{ background: fillSample }" />
          </fieldset>
        </div>
      </div>
      <div v-if="tab === 'protection'" class="sh-stack">
        <Check label="Locked" :checked="locked" test-id="protection-locked" @change="locked = $event; touch('protection')" />
        <Check label="Hidden" :checked="hidden" test-id="protection-hidden" @change="hidden = $event; touch('protection')" />
        <p class="sh-hint">Locking cells or hiding formulas has no effect until you protect the worksheet (Review tab, Protect group, Protect Sheet button).</p>
      </div>
    </div>
  </Dialog>
</template>
