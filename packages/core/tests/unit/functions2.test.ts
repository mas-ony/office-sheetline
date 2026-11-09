import { describe, expect, it } from 'vitest';
import { errCode, evalFormula, makeBook } from './helpers';

const close = (a: unknown, b: number, digits = 6) => expect(a as number).toBeCloseTo(b, digits);

describe('LAMBDA and its helpers', () => {
  it('calls LAMBDA values bound with LET and detects omitted parameters', () => {
    expect(evalFormula('=LET(f, LAMBDA(x, y, x*y), f(3, 4))')).toBe(12);
    expect(evalFormula('=LET(f, LAMBDA(x, y, IF(ISOMITTED(y), x, x+y)), f(5))')).toBe(5);
    expect(evalFormula('=LET(f, LAMBDA(x, y, IF(ISOMITTED(y), x, x+y)), f(5, 2))')).toBe(7);
    // A LAMBDA used as a value is #CALC!.
    expect(errCode(evalFormula('=LAMBDA(x, x+1)'))).toBe('#CALC!');
  });

  it('maps, reduces and scans arrays', () => {
    const b = makeBook();
    b.set({ A1: 1, A2: 2, A3: 3, B1: 10, B2: 20, B3: 30 });
    b.set({ D1: '=MAP(A1:A3, LAMBDA(v, v*10))' });
    expect([b.get('D1'), b.get('D2'), b.get('D3')]).toEqual([10, 20, 30]);
    b.set({ E1: '=MAP(A1:A3, B1:B3, LAMBDA(a, c, a+c))' });
    expect([b.get('E1'), b.get('E3')]).toEqual([11, 33]);
    b.set({ F1: '=REDUCE(0, A1:A3, LAMBDA(acc, v, acc+v*v))' });
    expect(b.get('F1')).toBe(14);
    b.set({ G1: '=SCAN(0, A1:A3, LAMBDA(acc, v, acc+v))' });
    expect([b.get('G1'), b.get('G2'), b.get('G3')]).toEqual([1, 3, 6]);
  });

  it('works by row and column and builds arrays', () => {
    const b = makeBook();
    b.set({ A1: 1, B1: 2, A2: 3, B2: 4 });
    b.set({ D1: '=BYROW(A1:B2, LAMBDA(r, SUM(r)))' });
    expect([b.get('D1'), b.get('D2')]).toEqual([3, 7]);
    b.set({ E1: '=BYCOL(A1:B2, LAMBDA(c, MAX(c)))' });
    expect([b.get('E1'), b.get('F1')]).toEqual([3, 4]);
    b.set({ H1: '=MAKEARRAY(2, 3, LAMBDA(r, c, r*c))' });
    expect([b.get('H1'), b.get('J1'), b.get('J2')]).toEqual([1, 3, 6]);
  });

  it('accepts function names as values (eta-reduced LAMBDAs)', () => {
    const b = makeBook();
    b.set({ A1: 1, A2: 2, A3: 3 });
    b.set({ C1: '=BYCOL(A1:A3, SUM)' });
    expect(b.get('C1')).toBe(6);
  });
});

describe('GROUPBY, PIVOTBY, PERCENTOF, TRIMRANGE', () => {
  it('groups values with a total row', () => {
    const b = makeBook();
    b.set({ A1: 'Region', B1: 'Sales', A2: 'East', B2: 10, A3: 'West', B3: 5, A4: 'East', B4: 7 });
    b.set({ D1: '=GROUPBY(A1:A4, B1:B4, SUM)' });
    expect([b.get('D1'), b.get('E1')]).toEqual(['Region', 'Sales']);
    expect([b.get('D2'), b.get('E2')]).toEqual(['East', 17]);
    expect([b.get('D3'), b.get('E3')]).toEqual(['West', 5]);
    expect([b.get('D4'), b.get('E4')]).toEqual(['Total', 22]);
  });

  it('pivots values by row and column fields', () => {
    const b = makeBook();
    b.set({ A1: 'East', B1: 'Q1', C1: 1, A2: 'East', B2: 'Q2', C2: 2, A3: 'West', B3: 'Q1', C3: 4 });
    b.set({ E1: '=PIVOTBY(A1:A3, B1:B3, C1:C3, SUM, 0)' });
    expect([b.get('F1'), b.get('G1'), b.get('H1')]).toEqual(['Q1', 'Q2', 'Total']);
    expect([b.get('E2'), b.get('F2'), b.get('G2'), b.get('H2')]).toEqual(['East', 1, 2, 3]);
    expect([b.get('E4'), b.get('H4')]).toEqual(['Total', 7]);
  });

  it('computes shares and trims empty edges of a range', () => {
    const b = makeBook();
    b.set({ A1: 1, A2: 3, B3: 'x' });
    b.set({ D1: '=PERCENTOF(A1, A1:A2)' });
    expect(b.get('D1')).toBe(0.25);
    b.set({ D2: '=ROWS(TRIMRANGE(A1:C10))', D3: '=COLUMNS(TRIMRANGE(A1:C10))' });
    expect([b.get('D2'), b.get('D3')]).toEqual([3, 2]);
  });
});

describe('financial and text additions', () => {
  it('ACCRINTM, AMORLINC, AMORDEGRC follow Excel', () => {
    // Excel examples: ACCRINTM(39539, 39614, 0.1, 1000, 3) = 20.54794521
    close(evalFormula('=ACCRINTM(DATE(2008,4,1),DATE(2008,6,15),0.1,1000,3)'), 20.5479452055, 6);
    // AMORLINC(2400, 39679, 39813, 300, 1, 0.15, 1) = 360
    close(evalFormula('=AMORLINC(2400,DATE(2008,8,19),DATE(2008,12,31),300,1,0.15,1)'), 360, 6);
    // AMORDEGRC(2400, 39679, 39813, 300, 1, 0.15, 1) = 776
    expect(evalFormula('=AMORDEGRC(2400,DATE(2008,8,19),DATE(2008,12,31),300,1,0.15,1)')).toBe(776);
  });

  it('prices bonds with odd periods', () => {
    // Excel example: ODDLPRICE(39217, 39340, 39126, 0.0375, 0.0405, 100, 2, 0) = 99.87828601
    close(evalFormula('=ODDLPRICE(DATE(2008,2,7),DATE(2008,6,15),DATE(2007,10,15),0.0375,0.0405,100,2,0)'), 99.87828601, 5);
    // ODDLYIELD(39196, 39340, 39126, 0.0375, 99.875, 100, 2, 0) = 0.04519
    close(evalFormula('=ODDLYIELD(DATE(2008,4,20),DATE(2008,6,15),DATE(2007,12,24),0.0375,99.875,100,2,0)'), 0.045192, 4);
    // ODDFPRICE example: 113.5977 (settlement 2008-11-11, maturity 2021-03-01, issue 2008-10-15, first 2009-03-01, 7.85%, 6.25%, 100, 2, 1)
    close(evalFormula('=ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,1)'), 113.5977, 2);
    const y = evalFormula('=ODDFYIELD(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0575,84.5,100,2,0)');
    close(y, 0.0772, 3);
  });

  it('converts euro legacy currencies', () => {
    expect(evalFormula('=EUROCONVERT(1.47,"FRF","DEM")')).toBe(0.44);
    close(evalFormula('=EUROCONVERT(1,"EUR","DEM",TRUE)'), 1.95583, 6);
  });

  it('spells baht amounts and encodes URLs', () => {
    expect(evalFormula('=BAHTTEXT(1234)')).toBe('หนึ่งพันสองร้อยสามสิบสี่บาทถ้วน');
    expect(evalFormula('=BAHTTEXT(21.5)')).toBe('ยี่สิบเอ็ดบาทห้าสิบสตางค์');
    expect(evalFormula('=ENCODEURL("a b/c?d=1")')).toBe('a%20b%2Fc%3Fd%3D1');
  });

  it('reads XML with FILTERXML', () => {
    expect(evalFormula('=FILTERXML("<r><v>7</v><v>9</v></r>","//v[2]")')).toBe(9);
  });
});
