/**
 * Built-in table styles (TableStyleLight1..21, TableStyleMedium1..28, TableStyleDark1..11) described
 * by family and colour, the way the Office gallery is organised: each family has one entry in the
 * dark text colour and one per accent. Elements follow the SpreadsheetML table style elements.
 * The looks are drawn for this project from the families' visible characteristics.
 */
import type { ColorSpec, Theme } from './model/color';
import { resolveColor } from './model/color';
import type { BorderLine } from './model/style';
import type { SheetTable } from './model/table';

export interface StyleElement {
  fill?: ColorSpec;
  color?: ColorSpec;
  bold?: boolean;
  top?: BorderLine;
  bottom?: BorderLine;
  left?: BorderLine;
  right?: BorderLine;
  /** Borders between rows / columns inside the element's area. */
  insideH?: BorderLine;
  insideV?: BorderLine;
}

export interface TableStyleDef {
  name: string;
  label: string;
  category: 'Light' | 'Medium' | 'Dark';
  wholeTable?: StyleElement;
  headerRow?: StyleElement;
  totalRow?: StyleElement;
  firstColumn?: StyleElement;
  lastColumn?: StyleElement;
  firstRowStripe?: StyleElement;
  firstColumnStripe?: StyleElement;
}

const WHITE: ColorSpec = { theme: 0 };
const DARK: ColorSpec = { theme: 1 };

/** Colour of the family member: 0 = dark text colour (its greys are chosen by the callers), 1..6 = accents. */
function base(i: number, tint?: number): ColorSpec {
  if (i === 0) return DARK;
  return tint === undefined ? { theme: 3 + i } : { theme: 3 + i, tint };
}

const line = (color: ColorSpec, style: BorderLine['style'] = 'thin'): BorderLine => ({ style, color });

function light1(i: number): Omit<TableStyleDef, 'name' | 'label' | 'category'> {
  const c = base(i);
  return {
    wholeTable: { top: line(c), bottom: line(c) },
    headerRow: { bold: true, bottom: line(c) },
    totalRow: { bold: true, top: line(c, 'double') },
    firstColumn: { bold: true },
    lastColumn: { bold: true },
    firstRowStripe: { fill: i ? base(i, 0.8) : { theme: 0, tint: -0.15 } },
    firstColumnStripe: { fill: i ? base(i, 0.8) : { theme: 0, tint: -0.15 } },
  };
}

function light2(i: number): Omit<TableStyleDef, 'name' | 'label' | 'category'> {
  const c = base(i);
  return {
    wholeTable: { top: line(c), bottom: line(c), left: line(c), right: line(c) },
    headerRow: { fill: c, color: WHITE, bold: true },
    totalRow: { bold: true, top: line(c, 'double') },
    firstColumn: { bold: true },
    lastColumn: { bold: true },
    firstRowStripe: { top: line(c), bottom: line(c) },
    firstColumnStripe: { left: line(c), right: line(c) },
  };
}

function light3(i: number): Omit<TableStyleDef, 'name' | 'label' | 'category'> {
  const c = base(i);
  return {
    wholeTable: { top: line(c), bottom: line(c), left: line(c), right: line(c), insideH: line(c), insideV: line(c) },
    headerRow: { bold: true, bottom: line(c) },
    totalRow: { bold: true, top: line(c, 'double') },
    firstColumn: { bold: true },
    lastColumn: { bold: true },
    firstRowStripe: { fill: i ? base(i, 0.8) : { theme: 0, tint: -0.15 } },
    firstColumnStripe: { fill: i ? base(i, 0.8) : { theme: 0, tint: -0.15 } },
  };
}

function medium1(i: number): Omit<TableStyleDef, 'name' | 'label' | 'category'> {
  const c = base(i);
  const soft = i ? base(i, 0.4) : { theme: 1, tint: 0.5 };
  return {
    wholeTable: { top: line(soft), bottom: line(soft), left: line(soft), right: line(soft), insideH: line(soft) },
    headerRow: { fill: c, color: WHITE, bold: true },
    totalRow: { bold: true, top: line(c, 'double') },
    firstColumn: { bold: true },
    lastColumn: { bold: true },
    firstRowStripe: { fill: i ? base(i, 0.8) : { theme: 0, tint: -0.15 } },
    firstColumnStripe: { fill: i ? base(i, 0.8) : { theme: 0, tint: -0.15 } },
  };
}

function medium2(i: number): Omit<TableStyleDef, 'name' | 'label' | 'category'> {
  const c = base(i);
  const white = line(WHITE);
  return {
    wholeTable: { fill: i ? base(i, 0.8) : { theme: 0, tint: -0.15 }, top: white, bottom: white, left: white, right: white, insideH: white, insideV: white },
    headerRow: { fill: c, color: WHITE, bold: true, bottom: line(WHITE, 'thick') },
    totalRow: { fill: c, color: WHITE, bold: true, top: line(WHITE, 'thick') },
    firstColumn: { fill: c, color: WHITE, bold: true },
    lastColumn: { fill: c, color: WHITE, bold: true },
    firstRowStripe: { fill: i ? base(i, 0.6) : { theme: 0, tint: -0.35 } },
    firstColumnStripe: { fill: i ? base(i, 0.6) : { theme: 0, tint: -0.35 } },
  };
}

function medium3(i: number): Omit<TableStyleDef, 'name' | 'label' | 'category'> {
  const c = base(i);
  return {
    wholeTable: { top: line(DARK, 'medium'), bottom: line(DARK, 'medium'), insideH: line(i ? base(i, 0.4) : { theme: 1, tint: 0.5 }) },
    headerRow: { fill: c, color: WHITE, bold: true, bottom: line(DARK, 'medium') },
    totalRow: { bold: true, top: line(DARK, 'double') },
    firstColumn: { bold: true },
    lastColumn: { bold: true },
    firstRowStripe: { fill: i ? base(i, 0.8) : { theme: 0, tint: -0.15 } },
    firstColumnStripe: { fill: i ? base(i, 0.8) : { theme: 0, tint: -0.15 } },
  };
}

function medium4(i: number): Omit<TableStyleDef, 'name' | 'label' | 'category'> {
  const soft = i ? base(i, 0.4) : { theme: 1, tint: 0.5 };
  const b = line(soft);
  return {
    wholeTable: { fill: i ? base(i, 0.8) : { theme: 0, tint: -0.15 }, top: b, bottom: b, left: b, right: b, insideH: b, insideV: b },
    headerRow: { bold: true },
    totalRow: { bold: true, top: line(base(i), 'double') },
    firstColumn: { bold: true },
    lastColumn: { bold: true },
    firstRowStripe: { fill: i ? base(i, 0.6) : { theme: 0, tint: -0.35 } },
    firstColumnStripe: { fill: i ? base(i, 0.6) : { theme: 0, tint: -0.35 } },
  };
}

function dark1(i: number): Omit<TableStyleDef, 'name' | 'label' | 'category'> {
  const body = i ? base(i, -0.25) : { theme: 1, tint: 0.35 };
  const stripe = i ? base(i, -0.5) : { theme: 1, tint: 0.25 };
  return {
    wholeTable: { fill: body, color: WHITE },
    headerRow: { fill: DARK, color: WHITE, bold: true, bottom: line(WHITE, 'medium') },
    totalRow: { fill: i ? base(i, -0.5) : DARK, color: WHITE, bold: true, top: line(WHITE, 'double') },
    firstColumn: { fill: stripe, color: WHITE, bold: true, right: line(WHITE, 'medium') },
    lastColumn: { fill: stripe, color: WHITE, bold: true, left: line(WHITE, 'medium') },
    firstRowStripe: { fill: stripe },
    firstColumnStripe: { fill: stripe },
  };
}

function dark2(i: number): Omit<TableStyleDef, 'name' | 'label' | 'category'> {
  // Dark 8..11 pair accents: black, accent 1/2, 3/4, 5/6.
  const a = i ? base(i * 2 - 1, 0.8) : { theme: 0, tint: -0.15 };
  const b = i ? base(i * 2 - 1, 0.6) : { theme: 0, tint: -0.35 };
  const head = i ? base(i * 2) : DARK;
  return {
    wholeTable: { fill: a },
    headerRow: { fill: head, color: WHITE, bold: true },
    totalRow: { bold: true, top: line(DARK, 'double') },
    firstColumn: { bold: true },
    lastColumn: { bold: true },
    firstRowStripe: { fill: b },
    firstColumnStripe: { fill: b },
  };
}

const COLOR_NAMES = ['Black', 'Accent 1', 'Accent 2', 'Accent 3', 'Accent 4', 'Accent 5', 'Accent 6'];

function build(): Map<string, TableStyleDef> {
  const out = new Map<string, TableStyleDef>();
  const add = (category: TableStyleDef['category'], n: number, def: Omit<TableStyleDef, 'name' | 'label' | 'category'>) => {
    const name = `TableStyle${category}${n}`;
    out.set(name, { name, label: `${category} ${n}`, category, ...def });
  };
  for (let i = 0; i < 7; i++) {
    add('Light', 1 + i, light1(i));
    add('Light', 8 + i, light2(i));
    add('Light', 15 + i, light3(i));
    add('Medium', 1 + i, medium1(i));
    add('Medium', 8 + i, medium2(i));
    add('Medium', 15 + i, medium3(i));
    add('Medium', 22 + i, medium4(i));
    add('Dark', 1 + i, dark1(i));
  }
  for (let i = 0; i < 4; i++) add('Dark', 8 + i, dark2(i));
  return out;
}

export const TABLE_STYLES = build();

export function tableStyleLabel(name: string): string {
  const def = TABLE_STYLES.get(name);
  if (!def) return name || 'None';
  const n = Number(/\d+$/.exec(name)?.[0] ?? 0);
  const color = def.category === 'Dark' && n >= 8 ? '' : COLOR_NAMES[(n - 1) % 7];
  return `${def.label}${color ? `, ${color}` : ''}`;
}

/** Styles in gallery order: Light 1..21, Medium 1..28, Dark 1..11 (rows of seven). */
export function galleryStyles(category: TableStyleDef['category']): TableStyleDef[] {
  return [...TABLE_STYLES.values()].filter((s) => s.category === category).sort((a, b) => Number(/\d+$/.exec(a.name)![0]) - Number(/\d+$/.exec(b.name)![0]));
}

// -----------------------------------------------------------------------------------------------
// Formatting of a cell inside a table

export interface CellDecor {
  fill?: string;
  color?: string;
  bold?: boolean;
  top?: BorderLine;
  bottom?: BorderLine;
  left?: BorderLine;
  right?: BorderLine;
}

function apply(out: CellDecor, el: StyleElement | undefined, edges: { top: boolean; bottom: boolean; left: boolean; right: boolean }, theme: Theme): void {
  if (!el) return;
  if (el.fill) out.fill = `#${resolveColor(el.fill, theme, 'FFFFFF')}`;
  if (el.color) out.color = `#${resolveColor(el.color, theme, '000000')}`;
  if (el.bold) out.bold = true;
  const pick = (outer: BorderLine | undefined, inner: BorderLine | undefined, isEdge: boolean) => (isEdge ? outer : inner);
  const t = pick(el.top, el.insideH, edges.top);
  const b = pick(el.bottom, el.insideH, edges.bottom);
  const l = pick(el.left, el.insideV, edges.left);
  const r = pick(el.right, el.insideV, edges.right);
  if (t) out.top = t;
  if (b) out.bottom = b;
  if (l) out.left = l;
  if (r) out.right = r;
}

/** The table style formatting of cell (r, c) of a table, or null outside the table / without a style. */
export function tableCellDecor(t: SheetTable, r: number, c: number, theme: Theme): CellDecor | null {
  const def = TABLE_STYLES.get(t.style.name);
  if (!def) return null;
  const { ref } = t;
  if (r < ref.r0 || r > ref.r1 || c < ref.c0 || c > ref.c1) return null;
  const out: CellDecor = {};
  const isHeader = t.headerRow && r === ref.r0;
  const isTotal = t.totalsRow && r === ref.r1;
  const dataR0 = ref.r0 + (t.headerRow ? 1 : 0);
  const dataR1 = ref.r1 - (t.totalsRow ? 1 : 0);
  apply(out, def.wholeTable, { top: r === ref.r0, bottom: r === ref.r1, left: c === ref.c0, right: c === ref.c1 }, theme);
  if (!isHeader && !isTotal) {
    const band = { top: true, bottom: true, left: c === ref.c0, right: c === ref.c1 };
    if (t.style.showColumnStripes && (c - ref.c0) % 2 === 0) apply(out, def.firstColumnStripe, { top: r === dataR0, bottom: r === dataR1, left: true, right: true }, theme);
    if (t.style.showRowStripes && (r - dataR0) % 2 === 0) apply(out, def.firstRowStripe, band, theme);
  }
  if (t.style.showFirstColumn && c === ref.c0 && !isHeader) apply(out, def.firstColumn, { top: r === dataR0, bottom: r === dataR1, left: true, right: true }, theme);
  if (t.style.showLastColumn && c === ref.c1 && !isHeader) apply(out, def.lastColumn, { top: r === dataR0, bottom: r === dataR1, left: true, right: true }, theme);
  if (isHeader) apply(out, def.headerRow, { top: true, bottom: true, left: c === ref.c0, right: c === ref.c1 }, theme);
  if (isTotal) apply(out, def.totalRow, { top: true, bottom: true, left: c === ref.c0, right: c === ref.c1 }, theme);
  return Object.keys(out).length ? out : null;
}

/** A small preview of a style (5 x 5 cells) as SVG markup, for the gallery. */
export function tableStylePreviewSvg(name: string, theme: Theme, size = 60): string {
  const t: SheetTable = {
    id: 0,
    name: 'P',
    ref: { r0: 0, c0: 0, r1: 4, c1: 4 },
    headerRow: true,
    totalsRow: false,
    columns: [],
    style: { name, showFirstColumn: false, showLastColumn: false, showRowStripes: true, showColumnStripes: false },
    autoFilter: false,
    filters: new Map(),
  };
  const cw = size / 5;
  const ch = (size * 0.75) / 5;
  const parts: string[] = [`<svg xmlns="http://www.w3.org/2000/svg" width="${size}" height="${size * 0.75}" viewBox="0 0 ${size} ${size * 0.75}">`, `<rect width="100%" height="100%" fill="#fff"/>`];
  const strokes: string[] = [];
  for (let r = 0; r < 5; r++) {
    for (let c = 0; c < 5; c++) {
      const d = tableCellDecor(t, r, c, theme) ?? {};
      const x = c * cw;
      const y = r * ch;
      if (d.fill) parts.push(`<rect x="${x}" y="${y}" width="${cw}" height="${ch}" fill="${d.fill}"/>`);
      const dash = `<line x1="${x + cw * 0.2}" y1="${y + ch / 2}" x2="${x + cw * 0.8}" y2="${y + ch / 2}" stroke="${d.color ?? '#000'}" stroke-width="${d.bold ? 1.6 : 1}"/>`;
      parts.push(dash);
      const ln = (l: BorderLine | undefined, x1: number, y1: number, x2: number, y2: number) => {
        if (!l) return;
        const w = l.style === 'thick' ? 2 : l.style === 'medium' ? 1.5 : 0.8;
        strokes.push(`<line x1="${x1}" y1="${y1}" x2="${x2}" y2="${y2}" stroke="#${resolveColor(l.color, theme, '000000')}" stroke-width="${w}"/>`);
      };
      ln(d.top, x, y, x + cw, y);
      ln(d.bottom, x, y + ch, x + cw, y + ch);
      ln(d.left, x, y, x, y + ch);
      ln(d.right, x + cw, y, x + cw, y + ch);
    }
  }
  parts.push(...strokes, '</svg>');
  return parts.join('');
}
