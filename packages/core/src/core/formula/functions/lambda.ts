/**
 * LAMBDA helper functions (MAP, REDUCE, SCAN, BYROW, BYCOL, MAKEARRAY), the aggregation functions
 * GROUPBY and PIVOTBY, PERCENTOF and TRIMRANGE.
 */
import { CellError, Errors, type Scalar } from '../../values';
import { ArrayValue, RefValue, compareScalars, type Value } from '../runtime';
import { LambdaValue } from '../evaluator';
import { define, num, toArray, toScalar, type FnContext } from './helpers';

function lambdaArg(v: Value | undefined): LambdaValue | CellError {
  if (v instanceof CellError) return v;
  return v instanceof LambdaValue ? v : Errors.VALUE;
}

/** The single value a LAMBDA returned for one element (arrays are not allowed there). */
function single(v: Value, ctx: FnContext): Scalar {
  if (v instanceof LambdaValue) return Errors.CALC;
  if (v instanceof ArrayValue) return v.rows === 1 && v.cols === 1 ? v.get(0, 0) : Errors.CALC;
  if (v instanceof RefValue) return v.isCell() ? toScalar(v, ctx) : Errors.CALC;
  return v;
}

function arrayArg(v: Value, ctx: FnContext): ArrayValue | CellError {
  if (v instanceof CellError) return v;
  return toArray(v, ctx);
}

define('MAP', {
  min: 2,
  max: 254,
  impl: (args, ctx) => {
    const fn = lambdaArg(args[args.length - 1]);
    if (fn instanceof CellError) return fn;
    const arrays: ArrayValue[] = [];
    for (const a of args.slice(0, -1)) {
      const arr = arrayArg(a, ctx);
      if (arr instanceof CellError) return arr;
      arrays.push(arr);
    }
    const rows = Math.max(...arrays.map((a) => a.rows));
    const cols = Math.max(...arrays.map((a) => a.cols));
    const out: Scalar[][] = [];
    for (let r = 0; r < rows; r++) {
      const row: Scalar[] = [];
      for (let c = 0; c < cols; c++) {
        const vals = arrays.map((a) => (a.rows === 1 && a.cols === 1 ? a.get(0, 0) : r < a.rows && c < a.cols ? a.get(r, c) : Errors.NA));
        row.push(single(fn.invoke(vals), ctx));
      }
      out.push(row);
    }
    return new ArrayValue(out);
  },
});

function reduceLike(args: Value[], ctx: FnContext, scan: boolean): Value {
  const fn = lambdaArg(args[args.length - 1]);
  if (fn instanceof CellError) return fn;
  const arr = arrayArg(args[args.length - 2], ctx);
  if (arr instanceof CellError) return arr;
  let acc: Value = args.length === 3 ? args[0] : null;
  if (acc instanceof RefValue) acc = toScalar(acc, ctx);
  const out: Scalar[][] = [];
  for (let r = 0; r < arr.rows; r++) {
    const row: Scalar[] = [];
    for (let c = 0; c < arr.cols; c++) {
      acc = fn.invoke([acc, arr.get(r, c)]);
      if (scan) row.push(single(acc, ctx));
    }
    if (scan) out.push(row);
  }
  if (scan) return new ArrayValue(out);
  return acc instanceof LambdaValue ? Errors.CALC : acc;
}

define('REDUCE', { min: 2, max: 3, impl: (args, ctx) => reduceLike(args, ctx, false) });
define('SCAN', { min: 2, max: 3, impl: (args, ctx) => reduceLike(args, ctx, true) });

define('BYROW', {
  min: 2,
  max: 2,
  impl: (args, ctx) => {
    const arr = arrayArg(args[0], ctx);
    if (arr instanceof CellError) return arr;
    const fn = lambdaArg(args[1]);
    if (fn instanceof CellError) return fn;
    return new ArrayValue(arr.data.map((row) => [single(fn.invoke([new ArrayValue([row])]), ctx)]));
  },
});

define('BYCOL', {
  min: 2,
  max: 2,
  impl: (args, ctx) => {
    const arr = arrayArg(args[0], ctx);
    if (arr instanceof CellError) return arr;
    const fn = lambdaArg(args[1]);
    if (fn instanceof CellError) return fn;
    const row: Scalar[] = [];
    for (let c = 0; c < arr.cols; c++) row.push(single(fn.invoke([ArrayValue.column(arr.data.map((r) => r[c] ?? null))]), ctx));
    return new ArrayValue([row]);
  },
});

define('MAKEARRAY', {
  min: 3,
  max: 3,
  impl: (args, ctx) => {
    const rows = num(args[0], ctx);
    const cols = num(args[1], ctx);
    if (rows instanceof CellError) return rows;
    if (cols instanceof CellError) return cols;
    const fn = lambdaArg(args[2]);
    if (fn instanceof CellError) return fn;
    const R = Math.trunc(rows);
    const C = Math.trunc(cols);
    if (R < 1 || C < 1 || R * C > 1_048_576) return Errors.VALUE;
    const out: Scalar[][] = [];
    for (let r = 1; r <= R; r++) {
      const row: Scalar[] = [];
      for (let c = 1; c <= C; c++) row.push(single(fn.invoke([r, c]), ctx));
      out.push(row);
    }
    return new ArrayValue(out);
  },
});

// -----------------------------------------------------------------------------------------------
// GROUPBY / PIVOTBY

const keyOf = (vals: Scalar[]) => vals.map((v) => (typeof v === 'string' ? `s:${v.toLowerCase()}` : `${typeof v}:${String(v)}`)).join('\u0001');

function compareKeys(a: Scalar[], b: Scalar[]): number {
  for (let i = 0; i < Math.max(a.length, b.length); i++) {
    const x = a[i] ?? null;
    const y = b[i] ?? null;
    if (x === y) continue;
    if (x === null || x === '') return 1;
    if (y === null || y === '') return -1;
    const c = compareScalars(x, y);
    if (c) return c;
  }
  return 0;
}

/** Applies the aggregation (a LAMBDA or a function like SUM) to one column of values. */
function aggregate(fn: LambdaValue, values: Scalar[], ctx: FnContext): Scalar {
  return single(fn.invoke([ArrayValue.column(values.length ? values : [null])]), ctx);
}

interface Grouped {
  headers: { rows: Scalar[]; values: Scalar[] } | null;
  rows: Scalar[][];
  values: Scalar[][];
}

/** Rows of the fields and values (header row split off as asked), filtered by an optional mask. */
function prepare(rowsArg: Value, valuesArg: Value, headerMode: number | null, filterArg: Value | undefined, ctx: FnContext): Grouped | CellError {
  const rowsA = arrayArg(rowsArg, ctx);
  if (rowsA instanceof CellError) return rowsA;
  const valsA = arrayArg(valuesArg, ctx);
  if (valsA instanceof CellError) return valsA;
  if (rowsA.rows !== valsA.rows) return Errors.VALUE;
  let rows = rowsA.data.map((r) => r.slice());
  let values = valsA.data.map((r) => r.slice());
  // Automatic headers: text on top of a numeric first value column.
  const mode = headerMode ?? (typeof values[0]?.[0] === 'string' && values.slice(1).some((r) => typeof r[0] === 'number') ? 3 : 0);
  let headers: Grouped['headers'] = null;
  if (mode === 1 || mode === 3) {
    headers = { rows: rows[0] ?? [], values: values[0] ?? [] };
    rows = rows.slice(1);
    values = values.slice(1);
  }
  if (mode === 2) headers = { rows: rows[0]?.map((_, i) => `Row Field ${i + 1}`) ?? [], values: values[0]?.map((_, i) => `Value ${i + 1}`) ?? [] };
  if (mode === 1) headers = null;
  if (filterArg !== undefined && filterArg !== null) {
    const mask = arrayArg(filterArg, ctx);
    if (mask instanceof CellError) return mask;
    const offset = mode === 1 || mode === 3 ? 1 : 0;
    const keep = rows.map((_, i) => {
      const m = mask.get(i + (mask.rows === rowsA.rows ? offset : 0), 0);
      return m === true || (typeof m === 'number' && m !== 0);
    });
    rows = rows.filter((_, i) => keep[i]);
    values = values.filter((_, i) => keep[i]);
  }
  return { headers, rows, values };
}

define('GROUPBY', {
  min: 3,
  max: 8,
  impl: (args, ctx) => {
    const fn = lambdaArg(args[2]);
    if (fn instanceof CellError) return fn;
    const opt = (i: number): number | null | CellError => (args[i] === undefined || args[i] === null ? null : num(args[i], ctx));
    const headerMode = opt(3);
    const totalDepth = opt(4);
    const sortOrder = opt(5);
    for (const v of [headerMode, totalDepth, sortOrder]) if (v instanceof CellError) return v;
    const g = prepare(args[0], args[1], headerMode as number | null, args[6], ctx);
    if (g instanceof CellError) return g;
    const groups = new Map<string, { key: Scalar[]; values: Scalar[][] }>();
    g.rows.forEach((r, i) => {
      const k = keyOf(r);
      let grp = groups.get(k);
      if (!grp) groups.set(k, (grp = { key: r, values: [] }));
      grp.values.push(g.values[i]);
    });
    const nv = g.values[0]?.length ?? 1;
    const lines = [...groups.values()].map((grp) => {
      const aggs = Array.from({ length: nv }, (_, c) => aggregate(fn, grp.values.map((v) => v[c] ?? null), ctx));
      return { key: grp.key, aggs };
    });
    const so = (sortOrder as number | null) ?? 1;
    const col = Math.abs(so) - 1;
    const nk = g.rows[0]?.length ?? 1;
    lines.sort((a, b) => {
      const x = col < nk ? [a.key[col]] : [a.aggs[col - nk]];
      const y = col < nk ? [b.key[col]] : [b.aggs[col - nk]];
      const c = compareKeys(x, y) || compareKeys(a.key, b.key);
      return so < 0 ? -c : c;
    });
    const out: Scalar[][] = [];
    if (g.headers) out.push([...g.headers.rows, ...g.headers.values]);
    for (const l of lines) out.push([...l.key, ...l.aggs]);
    const depth = (totalDepth as number | null) ?? 1;
    if (depth !== 0) {
      const total = ['Total', ...Array.from({ length: nk - 1 }, () => '' as Scalar), ...Array.from({ length: nv }, (_, c) => aggregate(fn, g.values.map((v) => v[c] ?? null), ctx))];
      if (depth < 0) out.splice(g.headers ? 1 : 0, 0, total);
      else out.push(total);
    }
    return out.length ? new ArrayValue(out) : Errors.CALC;
  },
});

define('PIVOTBY', {
  min: 4,
  max: 11,
  impl: (args, ctx) => {
    const fn = lambdaArg(args[3]);
    if (fn instanceof CellError) return fn;
    const opt = (i: number): number | null | CellError => (args[i] === undefined || args[i] === null ? null : num(args[i], ctx));
    const headerMode = opt(4);
    const rowTotals = opt(5);
    const colTotals = opt(7);
    for (const v of [headerMode, rowTotals, colTotals]) if (v instanceof CellError) return v;
    const rowsG = prepare(args[0], args[2], headerMode as number | null, args[9], ctx);
    if (rowsG instanceof CellError) return rowsG;
    const colsG = prepare(args[1], args[2], headerMode as number | null, args[9], ctx);
    if (colsG instanceof CellError) return colsG;
    const values = rowsG.values.map((v) => v[0] ?? null);
    const rowKeys = new Map<string, Scalar[]>();
    const colKeys = new Map<string, Scalar[]>();
    rowsG.rows.forEach((r) => rowKeys.set(keyOf(r), r));
    colsG.rows.forEach((c) => colKeys.set(keyOf(c), c));
    const rk = [...rowKeys.values()].sort(compareKeys);
    const ck = [...colKeys.values()].sort(compareKeys);
    const cell = (rKey: Scalar[] | null, cKey: Scalar[] | null) =>
      aggregate(fn, values.filter((_, i) => (!rKey || keyOf(rowsG.rows[i]) === keyOf(rKey)) && (!cKey || keyOf(colsG.rows[i]) === keyOf(cKey))), ctx);
    const showRowTotal = ((rowTotals as number | null) ?? 1) !== 0;
    const showColTotal = ((colTotals as number | null) ?? 1) !== 0;
    const nr = rk[0]?.length ?? 1;
    const out: Scalar[][] = [];
    const head: Scalar[] = [...Array.from({ length: nr }, () => '' as Scalar), ...ck.map((k) => k.join(' ')), ...(showColTotal ? ['Total'] : [])];
    out.push(head);
    for (const r of rk) out.push([...r, ...ck.map((c) => cell(r, c)), ...(showColTotal ? [cell(r, null)] : [])]);
    if (showRowTotal) out.push(['Total', ...Array.from({ length: nr - 1 }, () => '' as Scalar), ...ck.map((c) => cell(null, c)), ...(showColTotal ? [cell(null, null)] : [])]);
    return new ArrayValue(out);
  },
});

// -----------------------------------------------------------------------------------------------
// PERCENTOF, TRIMRANGE

function sumOf(v: Value, ctx: FnContext): number | CellError {
  const arr = arrayArg(v, ctx);
  if (arr instanceof CellError) return arr;
  let total = 0;
  for (const row of arr.data) {
    for (const x of row) {
      if (x instanceof CellError) return x;
      if (typeof x === 'number') total += x;
    }
  }
  return total;
}

define('PERCENTOF', {
  min: 2,
  max: 2,
  impl: (args, ctx) => {
    const part = sumOf(args[0], ctx);
    const all = sumOf(args[1], ctx);
    if (part instanceof CellError) return part;
    if (all instanceof CellError) return all;
    return all === 0 ? Errors.DIV0 : part / all;
  },
});

define('TRIMRANGE', {
  min: 1,
  max: 3,
  impl: (args, ctx) => {
    const ref = args[0];
    if (ref instanceof CellError) return ref;
    if (!(ref instanceof RefValue) || ref.areas.length !== 1) return Errors.VALUE;
    const trimRows = args[1] === undefined || args[1] === null ? 3 : num(args[1], ctx);
    const trimCols = args[2] === undefined || args[2] === null ? 3 : num(args[2], ctx);
    if (trimRows instanceof CellError) return trimRows;
    if (trimCols instanceof CellError) return trimCols;
    const { sheet, range } = ref.areas[0];
    let r0 = Infinity;
    let r1 = -Infinity;
    let c0 = Infinity;
    let c1 = -Infinity;
    for (const [r, row] of sheet.cells) {
      if (r < range.r0 || r > range.r1) continue;
      for (const [c, cell] of row) {
        if (c < range.c0 || c > range.c1 || cell.v === null || cell.v === '') continue;
        r0 = Math.min(r0, r);
        r1 = Math.max(r1, r);
        c0 = Math.min(c0, c);
        c1 = Math.max(c1, c);
      }
    }
    if (!Number.isFinite(r0)) return Errors.REF;
    const tr = Math.trunc(trimRows);
    const tc = Math.trunc(trimCols);
    const out = {
      r0: tr === 1 || tr === 3 ? r0 : range.r0,
      r1: tr === 2 || tr === 3 ? r1 : range.r1,
      c0: tc === 1 || tc === 3 ? c0 : range.c0,
      c1: tc === 2 || tc === 3 ? c1 : range.c1,
    };
    return new RefValue([{ sheet, range: out }]);
  },
});
