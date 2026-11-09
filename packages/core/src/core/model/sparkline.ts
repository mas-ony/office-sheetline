/**
 * Sparklines: small charts drawn inside cells (Insert > Sparklines), stored in groups that share
 * their type and formatting, as in Excel 2010's x14:sparklineGroups.
 */
import type { ColorSpec } from './color';

export type SparklineType = 'line' | 'column' | 'stacked';

export interface Sparkline {
  /** Cell the sparkline is drawn in. */
  r: number;
  c: number;
  /** Data range, sheet-qualified ('Sheet1'!A2:E2). */
  ref: string;
}

export type SparklineColors = 'colorSeries' | 'colorNegative' | 'colorAxis' | 'colorMarkers' | 'colorFirst' | 'colorLast' | 'colorHigh' | 'colorLow';

export interface SparklineGroup {
  id: string;
  type: SparklineType;
  sparklines: Sparkline[];
  colors: Record<SparklineColors, ColorSpec>;
  /** Line width in points. */
  lineWeight: number;
  markers: boolean;
  high: boolean;
  low: boolean;
  first: boolean;
  last: boolean;
  negative: boolean;
  displayXAxis: boolean;
  displayHidden: boolean;
  displayEmptyCellsAs: 'gap' | 'zero' | 'span';
  minAxisType: 'individual' | 'group' | 'custom';
  maxAxisType: 'individual' | 'group' | 'custom';
  manualMin?: number;
  manualMax?: number;
  rightToLeft: boolean;
}

/** Excel's default sparkline style (Accent 1). */
export function defaultSparklineColors(accent = 4): Record<SparklineColors, ColorSpec> {
  return {
    colorSeries: { theme: accent, tint: -0.499984740745262 },
    colorNegative: { theme: accent + 1 > 9 ? 5 : accent + 1 },
    colorAxis: { rgb: 'FF000000' },
    colorMarkers: { theme: accent, tint: -0.499984740745262 },
    colorFirst: { theme: accent, tint: 0.39997558519241921 },
    colorLast: { theme: accent, tint: 0.39997558519241921 },
    colorHigh: { theme: accent },
    colorLow: { theme: accent },
  };
}

let counter = 1;

export function newSparklineGroup(type: SparklineType, sparklines: Sparkline[]): SparklineGroup {
  return {
    id: `spk${Date.now().toString(36)}${(counter++).toString(36)}`,
    type,
    sparklines,
    colors: defaultSparklineColors(),
    lineWeight: 0.75,
    markers: false,
    high: false,
    low: false,
    first: false,
    last: false,
    negative: false,
    displayXAxis: false,
    displayHidden: false,
    displayEmptyCellsAs: 'gap',
    minAxisType: 'individual',
    maxAxisType: 'individual',
    rightToLeft: false,
  };
}

export function cloneSparklineGroup(g: SparklineGroup): SparklineGroup {
  return { ...g, sparklines: g.sparklines.map((s) => ({ ...s })), colors: { ...g.colors } };
}

/** Styles of the sparkline style gallery: the accent colour of each column, darker by row. */
export const SPARKLINE_STYLE_ACCENTS = [4, 5, 6, 7, 8, 9];
