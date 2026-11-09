/**
 * Page setup of a sheet as stored in SpreadsheetML: <pageSetup>, <pageMargins>, <printOptions>,
 * <headerFooter>, <rowBreaks>/<colBreaks> and sheetPr/pageSetUpPr fitToPage. Print area and print
 * titles are the defined names _xlnm.Print_Area and _xlnm.Print_Titles.
 */
import type { XmlElement } from '../xml/xml';

export interface PageMargins {
  /** Inches, as in the file. */
  left: number;
  right: number;
  top: number;
  bottom: number;
  header: number;
  footer: number;
}

export interface HeaderFooterSet {
  oddHeader?: string;
  oddFooter?: string;
  evenHeader?: string;
  evenFooter?: string;
  firstHeader?: string;
  firstFooter?: string;
  differentOddEven?: boolean;
  differentFirst?: boolean;
}

export interface PageSetup {
  orientation: 'portrait' | 'landscape';
  /** Excel paper size code (1 Letter, 9 A4...). */
  paperSize: number;
  /** Percent, 10..400. */
  scale: number;
  /** Fit to pages instead of the scale. */
  fitToPage: boolean;
  /** Pages wide / tall (0 = automatic). */
  fitToWidth: number;
  fitToHeight: number;
  margins: PageMargins;
  centerH: boolean;
  centerV: boolean;
  gridLines: boolean;
  headings: boolean;
  blackAndWhite: boolean;
  pageOrder: 'downThenOver' | 'overThenDown';
  firstPageNumber?: number;
  printQuality?: number;
  headerFooter: HeaderFooterSet;
  /** Manual page breaks: a new page starts at these row / column indexes. */
  rowBreaks: number[];
  colBreaks: number[];
  /** Source elements (attributes the editor does not model are kept). */
  raw?: { pageSetup?: XmlElement; printOptions?: XmlElement; headerFooter?: XmlElement };
}

export const MARGIN_PRESETS: { name: 'normal' | 'wide' | 'narrow'; label: string; margins: PageMargins }[] = [
  { name: 'normal', label: 'Normal', margins: { top: 0.75, bottom: 0.75, left: 0.7, right: 0.7, header: 0.3, footer: 0.3 } },
  { name: 'wide', label: 'Wide', margins: { top: 1, bottom: 1, left: 1, right: 1, header: 0.5, footer: 0.5 } },
  { name: 'narrow', label: 'Narrow', margins: { top: 0.75, bottom: 0.75, left: 0.25, right: 0.25, header: 0.3, footer: 0.3 } },
];

/** Paper sizes of the Size menu: Excel code, label, size in inches. */
export const PAPER: { code: number; label: string; w: number; h: number; size: string }[] = [
  { code: 1, label: 'Letter', w: 8.5, h: 11, size: '8.5" x 11"' },
  { code: 3, label: 'Tabloid', w: 11, h: 17, size: '11" x 17"' },
  { code: 5, label: 'Legal', w: 8.5, h: 14, size: '8.5" x 14"' },
  { code: 6, label: 'Statement', w: 5.5, h: 8.5, size: '5.5" x 8.5"' },
  { code: 7, label: 'Executive', w: 7.25, h: 10.5, size: '7.25" x 10.5"' },
  { code: 8, label: 'A3', w: 11.69, h: 16.54, size: '11.69" x 16.54"' },
  { code: 9, label: 'A4', w: 8.27, h: 11.69, size: '8.27" x 11.69"' },
  { code: 11, label: 'A5', w: 5.83, h: 8.27, size: '5.83" x 8.27"' },
  { code: 12, label: 'B4 (JIS)', w: 10.12, h: 14.33, size: '10.12" x 14.33"' },
  { code: 13, label: 'B5 (JIS)', w: 7.17, h: 10.12, size: '7.17" x 10.12"' },
];

export function paperOf(code: number): (typeof PAPER)[number] {
  return PAPER.find((p) => p.code === code) ?? PAPER[0];
}

export function defaultPageSetup(): PageSetup {
  return {
    orientation: 'portrait',
    paperSize: 1,
    scale: 100,
    fitToPage: false,
    fitToWidth: 1,
    fitToHeight: 1,
    margins: { ...MARGIN_PRESETS[0].margins },
    centerH: false,
    centerV: false,
    gridLines: false,
    headings: false,
    blackAndWhite: false,
    pageOrder: 'downThenOver',
    headerFooter: {},
    rowBreaks: [],
    colBreaks: [],
  };
}

export function clonePageSetup(p: PageSetup): PageSetup {
  return { ...p, margins: { ...p.margins }, headerFooter: { ...p.headerFooter }, rowBreaks: [...p.rowBreaks], colBreaks: [...p.colBreaks] };
}

/** Header / footer presets of the Page Setup dialog (Excel codes). */
export const HEADER_PRESETS: { label: string; code: string }[] = [
  { label: '(none)', code: '' },
  { label: 'Page 1', code: '&CPage &P' },
  { label: 'Page 1 of ?', code: '&CPage &P of &N' },
  { label: 'Sheet1', code: '&C&A' },
  { label: 'Book1', code: '&C&F' },
  { label: 'Sheet1, Page 1', code: '&C&A&RPage &P' },
  { label: 'Prepared by, Date, Page', code: '&LPrepared by &D&RPage &P' },
  { label: 'Date and time', code: '&L&D&R&T' },
];

/** Splits an Excel header/footer string into left, center and right sections (codes kept). */
export function hfSections(code: string | undefined): { left: string; center: string; right: string } {
  const out = { left: '', center: '', right: '' };
  if (!code) return out;
  let cur: 'left' | 'center' | 'right' = 'center';
  let i = 0;
  while (i < code.length) {
    if (code[i] === '&' && /[LCR]/.test(code[i + 1] ?? '')) {
      cur = code[i + 1] === 'L' ? 'left' : code[i + 1] === 'C' ? 'center' : 'right';
      i += 2;
      continue;
    }
    if (code[i] === '&' && code[i + 1] === '&') {
      out[cur] += '&&';
      i += 2;
      continue;
    }
    out[cur] += code[i++];
  }
  return out;
}

export function hfJoin(s: { left: string; center: string; right: string }): string {
  return `${s.left ? `&L${s.left}` : ''}${s.center ? `&C${s.center}` : ''}${s.right ? `&R${s.right}` : ''}`;
}

/** The text of a header/footer section for a page (formatting codes dropped). */
export function hfText(section: string, ctx: { page: number; pages: number; file: string; sheet: string; now: Date }): string {
  return section
    .replace(/&"[^"]*"/g, '')
    .replace(/&\d+/g, '')
    .replace(/&[BIUSXYEOHK]/g, '')
    .replace(/&P(\+\d+|-\d+)?/g, (_m, off: string | undefined) => String(ctx.page + (off ? Number(off) : 0)))
    .replace(/&N/g, String(ctx.pages))
    .replace(/&D/g, ctx.now.toLocaleDateString())
    .replace(/&T/g, ctx.now.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }))
    .replace(/&F/g, ctx.file)
    .replace(/&A/g, ctx.sheet)
    .replace(/&Z/g, '')
    .replace(/&G/g, '')
    .replace(/&&/g, '&');
}
