/**
 * A stand-in for the canvas API that records drawing as SVG, so that the canvas painters of cell
 * decorations (data bars, icon sets, sparklines) also draw the printed pages.
 */
import type { Paint2D } from '../grid/decor';

/** a b c d e f, as in DOMMatrix. */
type Matrix = [number, number, number, number, number, number];

const IDENTITY: Matrix = [1, 0, 0, 1, 0, 0];

function multiply(m: Matrix, n: Matrix): Matrix {
  return [
    m[0] * n[0] + m[2] * n[1],
    m[1] * n[0] + m[3] * n[1],
    m[0] * n[2] + m[2] * n[3],
    m[1] * n[2] + m[3] * n[3],
    m[0] * n[4] + m[2] * n[5] + m[4],
    m[1] * n[4] + m[3] * n[5] + m[5],
  ];
}

const num = (v: number) => String(Math.round(v * 100) / 100);
const esc = (s: string) => s.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;');

class SvgGradient implements CanvasGradient {
  readonly stops: [number, string][] = [];
  constructor(readonly x0: number, readonly y0: number, readonly x1: number, readonly y1: number) {}
  addColorStop(offset: number, color: string): void {
    this.stops.push([offset, color]);
  }
}

type Paint = string | CanvasGradient | CanvasPattern;

interface State {
  m: Matrix;
  fill: Paint;
  stroke: Paint;
  width: number;
  cap: CanvasLineCap;
  join: CanvasLineJoin;
  dash: number[];
  clip: string | null;
}

export class SvgCanvas implements Paint2D {
  private body: string[] = [];
  private defs: string[] = [];
  private path: string[] = [];
  private stack: State[] = [];
  private st: State = { m: IDENTITY, fill: '#000000', stroke: '#000000', width: 1, cap: 'butt', join: 'miter', dash: [], clip: null };
  private ids = 0;

  /** `idPrefix` keeps the ids of several recordings on one page apart. */
  constructor(readonly width: number, readonly height: number, private readonly idPrefix = 'p') {}

  get fillStyle(): Paint {
    return this.st.fill;
  }
  set fillStyle(v: Paint) {
    this.st.fill = v;
  }
  get strokeStyle(): Paint {
    return this.st.stroke;
  }
  set strokeStyle(v: Paint) {
    this.st.stroke = v;
  }
  get lineWidth(): number {
    return this.st.width;
  }
  set lineWidth(v: number) {
    this.st.width = v;
  }
  get lineCap(): CanvasLineCap {
    return this.st.cap;
  }
  set lineCap(v: CanvasLineCap) {
    this.st.cap = v;
  }
  get lineJoin(): CanvasLineJoin {
    return this.st.join;
  }
  set lineJoin(v: CanvasLineJoin) {
    this.st.join = v;
  }

  /** Whether anything was drawn. */
  get empty(): boolean {
    return this.body.length === 0;
  }

  save(): void {
    this.stack.push({ ...this.st, dash: [...this.st.dash] });
  }

  restore(): void {
    const s = this.stack.pop();
    if (s) this.st = s;
  }

  translate(x: number, y: number): void {
    this.st.m = multiply(this.st.m, [1, 0, 0, 1, x, y]);
  }

  rotate(angle: number): void {
    const c = Math.cos(angle);
    const s = Math.sin(angle);
    this.st.m = multiply(this.st.m, [c, s, -s, c, 0, 0]);
  }

  setLineDash(segments: number[]): void {
    this.st.dash = [...segments];
  }

  beginPath(): void {
    this.path = [];
  }

  private point(x: number, y: number): [number, number] {
    const m = this.st.m;
    return [m[0] * x + m[2] * y + m[4], m[1] * x + m[3] * y + m[5]];
  }

  /** Scale of lengths under the current transform (translations, rotations and uniform scales). */
  private get unit(): number {
    return Math.hypot(this.st.m[0], this.st.m[1]);
  }

  moveTo(x: number, y: number): void {
    const [px, py] = this.point(x, y);
    this.path.push(`M${num(px)} ${num(py)}`);
  }

  lineTo(x: number, y: number): void {
    const [px, py] = this.point(x, y);
    this.path.push(`${this.path.length ? 'L' : 'M'}${num(px)} ${num(py)}`);
  }

  closePath(): void {
    if (this.path.length) this.path.push('Z');
  }

  rect(x: number, y: number, w: number, h: number): void {
    this.moveTo(x, y);
    this.lineTo(x + w, y);
    this.lineTo(x + w, y + h);
    this.lineTo(x, y + h);
    this.closePath();
  }

  arc(x: number, y: number, radius: number, startAngle: number, endAngle: number, counterclockwise = false): void {
    const [cx, cy] = this.point(x, y);
    const r = radius * this.unit;
    const rot = Math.atan2(this.st.m[1], this.st.m[0]);
    const at = (a: number) => `${num(cx + r * Math.cos(a + rot))} ${num(cy + r * Math.sin(a + rot))}`;
    const full = Math.abs(endAngle - startAngle) >= Math.PI * 2 - 1e-9;
    let sweep = counterclockwise ? startAngle - endAngle : endAngle - startAngle;
    sweep = full ? Math.PI * 2 : ((sweep % (Math.PI * 2)) + Math.PI * 2) % (Math.PI * 2);
    const dir = counterclockwise ? -1 : 1;
    const flag = counterclockwise ? 0 : 1;
    this.path.push(`${this.path.length ? 'L' : 'M'}${at(startAngle)}`);
    if (full) {
      // An SVG arc cannot end where it starts: two halves.
      this.path.push(`A${num(r)} ${num(r)} 0 0 ${flag} ${at(startAngle + dir * Math.PI)}`, `A${num(r)} ${num(r)} 0 0 ${flag} ${at(startAngle)}`);
    } else if (sweep > 0) {
      this.path.push(`A${num(r)} ${num(r)} 0 ${sweep > Math.PI ? 1 : 0} ${flag} ${at(startAngle + dir * sweep)}`);
    }
  }

  private paint(p: Paint): string {
    if (typeof p === 'string') return p;
    if (p instanceof SvgGradient) {
      const id = `${this.idPrefix}g${++this.ids}`;
      const [x1, y1] = this.point(p.x0, p.y0);
      const [x2, y2] = this.point(p.x1, p.y1);
      const stops = p.stops.map(([o, c]) => `<stop offset="${num(o)}" stop-color="${esc(c)}"/>`).join('');
      this.defs.push(`<linearGradient id="${id}" gradientUnits="userSpaceOnUse" x1="${num(x1)}" y1="${num(y1)}" x2="${num(x2)}" y2="${num(y2)}">${stops}</linearGradient>`);
      return `url(#${id})`;
    }
    return '#000000';
  }

  private get clipAttr(): string {
    return this.st.clip ? ` clip-path="url(#${this.st.clip})"` : '';
  }

  fill(): void {
    if (!this.path.length) return;
    this.body.push(`<path d="${this.path.join('')}" fill="${this.paint(this.st.fill)}"${this.clipAttr}/>`);
  }

  stroke(): void {
    if (!this.path.length) return;
    const u = this.unit;
    const dash = this.st.dash.length ? ` stroke-dasharray="${this.st.dash.map((d) => num(d * u)).join(' ')}"` : '';
    this.body.push(
      `<path d="${this.path.join('')}" fill="none" stroke="${this.paint(this.st.stroke)}" stroke-width="${num(this.st.width * u)}" stroke-linecap="${this.st.cap}" stroke-linejoin="${this.st.join}"${dash}${this.clipAttr}/>`,
    );
  }

  fillRect(x: number, y: number, w: number, h: number): void {
    const saved = this.path;
    this.path = [];
    this.rect(x, y, w, h);
    this.fill();
    this.path = saved;
  }

  strokeRect(x: number, y: number, w: number, h: number): void {
    const saved = this.path;
    this.path = [];
    this.rect(x, y, w, h);
    this.stroke();
    this.path = saved;
  }

  clip(): void {
    const id = `${this.idPrefix}c${++this.ids}`;
    // A clip inside a clip: the clip path is clipped by the outer one.
    this.defs.push(`<clipPath id="${id}"${this.clipAttr}><path d="${this.path.join('') || 'M0 0'}"/></clipPath>`);
    this.st.clip = id;
  }

  createLinearGradient(x0: number, y0: number, x1: number, y1: number): CanvasGradient {
    return new SvgGradient(x0, y0, x1, y1);
  }

  /** The recording as an SVG document. */
  toSvg(): string {
    const defs = this.defs.length ? `<defs>${this.defs.join('')}</defs>` : '';
    return `<svg xmlns="http://www.w3.org/2000/svg" width="${num(this.width)}" height="${num(this.height)}" viewBox="0 0 ${num(this.width)} ${num(this.height)}">${defs}${this.body.join('')}</svg>`;
  }
}
