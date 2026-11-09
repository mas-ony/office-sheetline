import { describe, expect, it } from 'vitest';
import { analyzeFormula, applySuggestion, highlightedArg } from '../../src/ui/formulaAssist';
import { functionCatalog, functionInfo, signature } from '../../src/ui/functionCatalog';

const names = (text: string, caret = text.length) => analyzeFormula(text, caret)?.suggestions.map((s) => s.name) ?? [];

describe('formula assist', () => {
  it('suggests functions while a name is typed, not inside a written call', () => {
    expect(names('=SU')).toEqual(expect.arrayContaining(['SUM', 'SUMIF', 'SUMPRODUCT']));
    expect(names('=A1+co')).toContain('COUNT');
    // The caret placed in an existing call shows no list.
    expect(names('=SUM(B5:E5)', 4)).toEqual([]);
    expect(names('=SUM(B5:E5)')).toEqual([]);
    // Not inside strings, after a sheet name or in numbers.
    expect(names('="SU')).toEqual([]);
    expect(names('=Sheet1!SU')).toEqual([]);
    expect(names('=1E')).toEqual([]);
  });

  it('follows the argument being typed', () => {
    const ctx = analyzeFormula('=SUMIF(A1:A3,">2",', 18)!;
    expect(ctx.call?.info.name).toBe('SUMIF');
    expect(ctx.call?.argIndex).toBe(2);
    expect(signature(ctx.call!.info)).toBe('SUMIF(range, criteria, [sum_range])');
    // Nested calls report the innermost one.
    expect(analyzeFormula('=ROUND(SUM(1,2', 14)!.call?.info.name).toBe('SUM');
    // Repeating arguments wrap around to their group.
    expect(highlightedArg(functionInfo('SUM')!, 5)).toBe(1);
  });

  it('replaces the typed name with the picked function', () => {
    const ctx = analyzeFormula('=1+su', 5)!;
    expect(applySuggestion('=1+su', ctx.ident!, functionInfo('SUM')!)).toEqual({ text: '=1+SUM(', caret: 7 });
  });

  it('describes every function with named arguments', () => {
    const all = [...functionCatalog().values()];
    expect(all.length).toBeGreaterThanOrEqual(499);
    expect(all.filter((f) => !f.description)).toEqual([]);
    expect(all.filter((f) => f.args.some((a) => /^\[?arg\d/.test(a)))).toEqual([]);
    expect(functionInfo('FILTERXML')?.category).toBe('Web');
  });
});
