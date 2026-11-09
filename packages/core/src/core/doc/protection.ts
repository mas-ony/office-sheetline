/** Sheet and workbook protection commands, and the checks commands use before changing protected cells. */
import { SpreadsheetDocument } from '../document';
import type { Sheet } from '../model/sheet';
import { DEFAULT_PROTECTION, makePassword, passwordMatches, type ProtectionOptions } from '../model/protection';
import type { Range } from '../refs';

export type ProtectedAction = 'edit' | 'formatCells' | 'formatColumns' | 'formatRows' | 'insertRows' | 'insertColumns' | 'deleteRows' | 'deleteColumns' | 'sort' | 'autoFilter' | 'objects' | 'insertHyperlinks' | 'structure';

export const PROTECTED_MESSAGE = 'The cell or chart you\'re trying to change is on a protected sheet. To make a change, unprotect the sheet. You might be requested to enter a password.';

declare module '../document' {
  interface SpreadsheetDocument {
    protectSheet(options?: Partial<ProtectionOptions>, password?: string, sheet?: Sheet): void;
    /** Returns false when the password does not match. */
    unprotectSheet(password?: string, sheet?: Sheet): boolean;
    protectWorkbook(password?: string): void;
    unprotectWorkbook(password?: string): boolean;
    isCellLocked(r: number, c: number, sheet?: Sheet): boolean;
    /** Null when allowed, else the message to show. */
    blockedBy(action: ProtectedAction, ranges?: Range[], sheet?: Sheet): string | null;
  }
}

const proto = SpreadsheetDocument.prototype;

proto.protectSheet = function (this: SpreadsheetDocument, options = {}, password = '', sheet = this.sheet) {
  this.transact('Protect sheet', (tx) => {
    tx.snapshotSheet(sheet);
    sheet.protection = { options: { ...DEFAULT_PROTECTION, ...options }, password: makePassword(password) };
  });
};

proto.unprotectSheet = function (this: SpreadsheetDocument, password = '', sheet = this.sheet) {
  const p = sheet.protection;
  if (!p) return true;
  if (!passwordMatches(p.password, password)) return false;
  this.transact('Unprotect sheet', (tx) => {
    tx.snapshotSheet(sheet);
    sheet.protection = undefined;
  });
  return true;
};

proto.protectWorkbook = function (this: SpreadsheetDocument, password = '') {
  this.transact('Protect workbook', (tx) => {
    tx.snapshotWorkbook(this.wb);
    this.wb.protection = { structure: true, windows: false, password: makePassword(password) };
  });
};

proto.unprotectWorkbook = function (this: SpreadsheetDocument, password = '') {
  const p = this.wb.protection;
  if (!p) return true;
  if (!passwordMatches(p.password, password)) return false;
  this.transact('Unprotect workbook', (tx) => {
    tx.snapshotWorkbook(this.wb);
    this.wb.protection = undefined;
  });
  return true;
};

proto.isCellLocked = function (this: SpreadsheetDocument, r, c, sheet = this.sheet) {
  return this.styleAt(r, c, sheet).locked !== false;
};

proto.blockedBy = function (this: SpreadsheetDocument, action, ranges = this.selectionOf(this.sheet).ranges, sheet = this.sheet) {
  if (action === 'structure') return this.wb.protection?.structure ? 'Workbook structure is protected and cannot be changed.' : null;
  const p = sheet.protection;
  if (!p) return null;
  if (action === 'edit') {
    for (const range of ranges) {
      const b = this.bounded(range, sheet) ?? range;
      if ((b.r1 - b.r0 + 1) * (b.c1 - b.c0 + 1) > 200000) return PROTECTED_MESSAGE;
      for (let r = b.r0; r <= b.r1; r++) for (let c = b.c0; c <= b.c1; c++) if (this.isCellLocked(r, c, sheet)) return PROTECTED_MESSAGE;
    }
    return null;
  }
  return p.options[action] ? null : PROTECTED_MESSAGE;
};

export {};
