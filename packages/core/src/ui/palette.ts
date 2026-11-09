/** Colour palettes and colour arithmetic of the palette and the colour picker. */
import { resolveColor, type ColorSpec, type Theme } from '../core/model/color';

/** Theme colour columns (SpreadsheetML theme index) and their tint rows, as in Office. */
export const THEME_COLUMNS: { theme: number; name: string; tints: number[] }[] = [
  { theme: 0, name: 'Background 1', tints: [-0.05, -0.15, -0.25, -0.35, -0.5] },
  { theme: 1, name: 'Text 1', tints: [0.5, 0.35, 0.25, 0.15, 0.05] },
  { theme: 2, name: 'Background 2', tints: [-0.1, -0.25, -0.5, -0.75, -0.9] },
  { theme: 3, name: 'Text 2', tints: [0.8, 0.6, 0.4, -0.25, -0.5] },
  ...[1, 2, 3, 4, 5, 6].map((i) => ({ theme: 3 + i, name: `Accent ${i}`, tints: [0.8, 0.6, 0.4, -0.25, -0.5] })),
];

export const STANDARD_COLORS: { hex: string; name: string }[] = [
  { hex: '#C00000', name: 'Dark Red' },
  { hex: '#FF0000', name: 'Red' },
  { hex: '#FFC000', name: 'Orange' },
  { hex: '#FFFF00', name: 'Yellow' },
  { hex: '#92D050', name: 'Light Green' },
  { hex: '#00B050', name: 'Green' },
  { hex: '#00B0F0', name: 'Light Blue' },
  { hex: '#0070C0', name: 'Blue' },
  { hex: '#002060', name: 'Dark Blue' },
  { hex: '#7030A0', name: 'Purple' },
];

export function tintLabel(t: number): string {
  return t > 0 ? `Lighter ${Math.round(t * 100)}%` : `Darker ${Math.round(-t * 100)}%`;
}

export function specToCss(spec: ColorSpec | null | undefined, theme: Theme, fallback = '000000'): string {
  return `#${resolveColor(spec ?? undefined, theme, fallback)}`;
}

/** Material palette (hue columns × shades 50..900). */
const HUES: string[][] = [
  ['#FFEBEE', '#FFCDD2', '#EF9A9A', '#E57373', '#EF5350', '#F44336', '#E53935', '#D32F2F', '#C62828', '#B71C1C'],
  ['#FCE4EC', '#F8BBD0', '#F48FB1', '#F06292', '#EC407A', '#E91E63', '#D81B60', '#C2185B', '#AD1457', '#880E4F'],
  ['#F3E5F5', '#E1BEE7', '#CE93D8', '#BA68C8', '#AB47BC', '#9C27B0', '#8E24AA', '#7B1FA2', '#6A1B9A', '#4A148C'],
  ['#EDE7F6', '#D1C4E9', '#B39DDB', '#9575CD', '#7E57C2', '#673AB7', '#5E35B1', '#512DA8', '#4527A0', '#311B92'],
  ['#E8EAF6', '#C5CAE9', '#9FA8DA', '#7986CB', '#5C6BC0', '#3F51B5', '#3949AB', '#303F9F', '#283593', '#1A237E'],
  ['#E3F2FD', '#BBDEFB', '#90CAF9', '#64B5F6', '#42A5F5', '#2196F3', '#1E88E5', '#1976D2', '#1565C0', '#0D47A1'],
  ['#E0F7FA', '#B2EBF2', '#80DEEA', '#4DD0E1', '#26C6DA', '#00BCD4', '#00ACC1', '#0097A7', '#00838F', '#006064'],
  ['#E0F2F1', '#B2DFDB', '#80CBC4', '#4DB6AC', '#26A69A', '#009688', '#00897B', '#00796B', '#00695C', '#004D40'],
  ['#E8F5E9', '#C8E6C9', '#A5D6A7', '#81C784', '#66BB6A', '#4CAF50', '#43A047', '#388E3C', '#2E7D32', '#1B5E20'],
  ['#F9FBE7', '#F0F4C3', '#E6EE9C', '#DCE775', '#D4E157', '#CDDC39', '#C0CA33', '#AFB42B', '#9E9D24', '#827717'],
  ['#FFFDE7', '#FFF9C4', '#FFF59D', '#FFF176', '#FFEE58', '#FFEB3B', '#FDD835', '#FBC02D', '#F9A825', '#F57F17'],
  ['#FFF3E0', '#FFE0B2', '#FFCC80', '#FFB74D', '#FFA726', '#FF9800', '#FB8C00', '#F57C00', '#EF6C00', '#E65100'],
  ['#EFEBE9', '#D7CCC8', '#BCAAA4', '#A1887F', '#8D6E63', '#795548', '#6D4C41', '#5D4037', '#4E342E', '#3E2723'],
  ['#FFFFFF', '#F5F5F5', '#EEEEEE', '#E0E0E0', '#BDBDBD', '#9E9E9E', '#757575', '#616161', '#424242', '#000000'],
];

export const PALETTE_ROWS: string[][] = HUES[0].map((_, shade) => HUES.map((hue) => hue[shade]));

export function hexToRgb(hex: string): [number, number, number] {
  const h = hex.replace('#', '');
  const n = parseInt(h.length === 3 ? h.split('').map((c) => c + c).join('') : h.slice(0, 6), 16) || 0;
  return [(n >> 16) & 255, (n >> 8) & 255, n & 255];
}

export function rgbToHex(r: number, g: number, b: number): string {
  return `#${[r, g, b].map((x) => Math.max(0, Math.min(255, Math.round(x))).toString(16).padStart(2, '0')).join('').toUpperCase()}`;
}

export function normalizeHex(text: string): string | null {
  const t = text.trim().replace(/^#/, '');
  if (/^[0-9a-f]{3}$/i.test(t)) return `#${t.split('').map((c) => c + c).join('').toUpperCase()}`;
  if (/^[0-9a-f]{6}$/i.test(t)) return `#${t.toUpperCase()}`;
  return null;
}

/** Paints the hue and shade spectrum of the picker on its canvas. */
export function paintSpectrum(c: HTMLCanvasElement): void {
  const ctx = c.getContext('2d');
  if (!ctx) return;
  const w = c.width;
  const h = c.height;
  const hue = ctx.createLinearGradient(0, 0, w, 0);
  for (let i = 0; i <= 6; i++) hue.addColorStop(i / 6, `hsl(${i * 60}, 100%, 50%)`);
  ctx.fillStyle = hue;
  ctx.fillRect(0, 0, w, h);
  const shade = ctx.createLinearGradient(0, 0, 0, h);
  shade.addColorStop(0, 'rgba(255,255,255,1)');
  shade.addColorStop(0.5, 'rgba(255,255,255,0)');
  shade.addColorStop(0.5, 'rgba(0,0,0,0)');
  shade.addColorStop(1, 'rgba(0,0,0,1)');
  ctx.fillStyle = shade;
  ctx.fillRect(0, 0, w, h);
}

/** The colour under a point of the spectrum canvas (client coordinates), and the point inside it. */
export function spectrumPick(c: HTMLCanvasElement, clientX: number, clientY: number): { x: number; y: number; hex: string | null } {
  const rect = c.getBoundingClientRect();
  const x = Math.max(0, Math.min(rect.width - 1, clientX - rect.left));
  const y = Math.max(0, Math.min(rect.height - 1, clientY - rect.top));
  const data = c.getContext('2d')?.getImageData(Math.floor((x / rect.width) * c.width), Math.floor((y / rect.height) * c.height), 1, 1).data;
  return { x, y, hex: data ? rgbToHex(data[0], data[1], data[2]) : null };
}

/** A list of recent colours with `hex` first. */
export function withRecent(list: string[], hex: string, max: number): string[] {
  return [hex, ...list.filter((c) => c !== hex)].slice(0, max);
}
