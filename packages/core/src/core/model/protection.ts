/**
 * Sheet and workbook protection. Passwords are stored as Excel hashes: the legacy 16-bit hash is
 * written; SHA-512 hashes from newer files are checked with a JavaScript implementation (no secure
 * context needed).
 */
import type { XmlElement } from '../xml/xml';

/** What users may still do on a protected sheet (Excel's checkboxes; true = allowed). */
export interface ProtectionOptions {
  selectLockedCells: boolean;
  selectUnlockedCells: boolean;
  formatCells: boolean;
  formatColumns: boolean;
  formatRows: boolean;
  insertColumns: boolean;
  insertRows: boolean;
  insertHyperlinks: boolean;
  deleteColumns: boolean;
  deleteRows: boolean;
  sort: boolean;
  autoFilter: boolean;
  pivotTables: boolean;
  objects: boolean;
  scenarios: boolean;
}

export const DEFAULT_PROTECTION: ProtectionOptions = {
  selectLockedCells: true,
  selectUnlockedCells: true,
  formatCells: false,
  formatColumns: false,
  formatRows: false,
  insertColumns: false,
  insertRows: false,
  insertHyperlinks: false,
  deleteColumns: false,
  deleteRows: false,
  sort: false,
  autoFilter: false,
  pivotTables: false,
  objects: false,
  scenarios: false,
};

export const PROTECTION_LABELS: { key: keyof ProtectionOptions; label: string }[] = [
  { key: 'selectLockedCells', label: 'Select locked cells' },
  { key: 'selectUnlockedCells', label: 'Select unlocked cells' },
  { key: 'formatCells', label: 'Format cells' },
  { key: 'formatColumns', label: 'Format columns' },
  { key: 'formatRows', label: 'Format rows' },
  { key: 'insertColumns', label: 'Insert columns' },
  { key: 'insertRows', label: 'Insert rows' },
  { key: 'insertHyperlinks', label: 'Insert hyperlinks' },
  { key: 'deleteColumns', label: 'Delete columns' },
  { key: 'deleteRows', label: 'Delete rows' },
  { key: 'sort', label: 'Sort' },
  { key: 'autoFilter', label: 'Use AutoFilter' },
  { key: 'pivotTables', label: 'Use PivotTable & PivotChart' },
  { key: 'objects', label: 'Edit objects' },
  { key: 'scenarios', label: 'Edit scenarios' },
];

export interface PasswordHash {
  /** Legacy 16-bit hash as 4 hex digits. */
  legacy?: string;
  algorithm?: string;
  hashValue?: string;
  saltValue?: string;
  spinCount?: number;
}

export interface SheetProtection {
  options: ProtectionOptions;
  password?: PasswordHash;
  raw?: XmlElement;
}

export interface WorkbookProtection {
  structure: boolean;
  windows: boolean;
  password?: PasswordHash;
}

/** Excel's legacy password hash (as 4 upper-case hex digits). */
export function legacyHash(password: string): string {
  let hash = 0;
  const chars = [...password].map((c) => c.charCodeAt(0) & 0x7fff);
  for (let i = chars.length - 1; i >= 0; i--) {
    hash = ((hash >> 14) & 0x01) | ((hash << 1) & 0x7fff);
    hash ^= chars[i];
  }
  hash = ((hash >> 14) & 0x01) | ((hash << 1) & 0x7fff);
  hash ^= chars.length;
  hash ^= 0xce4b;
  return hash.toString(16).toUpperCase().padStart(4, '0');
}

// -----------------------------------------------------------------------------------------------
// SHA-512 (for passwords written by Excel 2013 and later)

const K = [
  '428a2f98d728ae22', '7137449123ef65cd', 'b5c0fbcfec4d3b2f', 'e9b5dba58189dbbc', '3956c25bf348b538', '59f111f1b605d019', '923f82a4af194f9b', 'ab1c5ed5da6d8118',
  'd807aa98a3030242', '12835b0145706fbe', '243185be4ee4b28c', '550c7dc3d5ffb4e2', '72be5d74f27b896f', '80deb1fe3b1696b1', '9bdc06a725c71235', 'c19bf174cf692694',
  'e49b69c19ef14ad2', 'efbe4786384f25e3', '0fc19dc68b8cd5b5', '240ca1cc77ac9c65', '2de92c6f592b0275', '4a7484aa6ea6e483', '5cb0a9dcbd41fbd4', '76f988da831153b5',
  '983e5152ee66dfab', 'a831c66d2db43210', 'b00327c898fb213f', 'bf597fc7beef0ee4', 'c6e00bf33da88fc2', 'd5a79147930aa725', '06ca6351e003826f', '142929670a0e6e70',
  '27b70a8546d22ffc', '2e1b21385c26c926', '4d2c6dfc5ac42aed', '53380d139d95b3df', '650a73548baf63de', '766a0abb3c77b2a8', '81c2c92e47edaee6', '92722c851482353b',
  'a2bfe8a14cf10364', 'a81a664bbc423001', 'c24b8b70d0f89791', 'c76c51a30654be30', 'd192e819d6ef5218', 'd69906245565a910', 'f40e35855771202a', '106aa07032bbd1b8',
  '19a4c116b8d2d0c8', '1e376c085141ab53', '2748774cdf8eeb99', '34b0bcb5e19b48a8', '391c0cb3c5c95a63', '4ed8aa4ae3418acb', '5b9cca4f7763e373', '682e6ff3d6b2b8a3',
  '748f82ee5defb2fc', '78a5636f43172f60', '84c87814a1f0ab72', '8cc702081a6439ec', '90befffa23631e28', 'a4506cebde82bde9', 'bef9a3f7b2c67915', 'c67178f2e372532b',
  'ca273eceea26619c', 'd186b8c721c0c207', 'eada7dd6cde0eb1e', 'f57d4f7fee6ed178', '06f067aa72176fba', '0a637dc5a2c898a6', '113f9804bef90dae', '1b710b35131c471b',
  '28db77f523047d84', '32caab7b40c72493', '3c9ebe0a15c9bebc', '431d67c49c100d4c', '4cc5d4becb3e42b6', '597f299cfc657e2a', '5fcb6fab3ad6faec', '6c44198c4a475817',
].map((h) => BigInt(`0x${h}`));

const MASK = (1n << 64n) - 1n;
const rotr = (x: bigint, n: bigint) => ((x >> n) | (x << (64n - n))) & MASK;

export function sha512(bytes: Uint8Array): Uint8Array {
  const H = ['6a09e667f3bcc908', 'bb67ae8584caa73b', '3c6ef372fe94f82b', 'a54ff53a5f1d36f1', '510e527fade682d1', '9b05688c2b3e6c1f', '1f83d9abfb41bd6b', '5be0cd19137e2179'].map((h) => BigInt(`0x${h}`));
  const len = bytes.length;
  const total = Math.ceil((len + 17) / 128) * 128;
  const msg = new Uint8Array(total);
  msg.set(bytes);
  msg[len] = 0x80;
  const bits = BigInt(len) * 8n;
  for (let i = 0; i < 16; i++) msg[total - 1 - i] = Number((bits >> BigInt(8 * i)) & 0xffn);
  const W = new Array<bigint>(80);
  for (let off = 0; off < total; off += 128) {
    for (let t = 0; t < 16; t++) {
      let v = 0n;
      for (let b = 0; b < 8; b++) v = (v << 8n) | BigInt(msg[off + t * 8 + b]);
      W[t] = v;
    }
    for (let t = 16; t < 80; t++) {
      const s0 = rotr(W[t - 15], 1n) ^ rotr(W[t - 15], 8n) ^ (W[t - 15] >> 7n);
      const s1 = rotr(W[t - 2], 19n) ^ rotr(W[t - 2], 61n) ^ (W[t - 2] >> 6n);
      W[t] = (W[t - 16] + s0 + W[t - 7] + s1) & MASK;
    }
    let [a, b, c, d, e, f, g, h] = H;
    for (let t = 0; t < 80; t++) {
      const S1 = rotr(e, 14n) ^ rotr(e, 18n) ^ rotr(e, 41n);
      const ch = (e & f) ^ (~e & MASK & g);
      const t1 = (h + S1 + ch + K[t] + W[t]) & MASK;
      const S0 = rotr(a, 28n) ^ rotr(a, 34n) ^ rotr(a, 39n);
      const maj = (a & b) ^ (a & c) ^ (b & c);
      const t2 = (S0 + maj) & MASK;
      h = g;
      g = f;
      f = e;
      e = (d + t1) & MASK;
      d = c;
      c = b;
      b = a;
      a = (t1 + t2) & MASK;
    }
    H[0] = (H[0] + a) & MASK;
    H[1] = (H[1] + b) & MASK;
    H[2] = (H[2] + c) & MASK;
    H[3] = (H[3] + d) & MASK;
    H[4] = (H[4] + e) & MASK;
    H[5] = (H[5] + f) & MASK;
    H[6] = (H[6] + g) & MASK;
    H[7] = (H[7] + h) & MASK;
  }
  const out = new Uint8Array(64);
  H.forEach((v, i) => {
    for (let b = 0; b < 8; b++) out[i * 8 + b] = Number((v >> BigInt(56 - 8 * b)) & 0xffn);
  });
  return out;
}

function fromBase64(s: string): Uint8Array {
  const bin = atob(s);
  return Uint8Array.from(bin, (ch) => ch.charCodeAt(0));
}

function toBase64(b: Uint8Array): string {
  let s = '';
  for (const x of b) s += String.fromCharCode(x);
  return btoa(s);
}

/** Excel's iterated SHA-512 of salt + UTF-16LE password. */
export function excelSha512(password: string, saltB64: string, spinCount: number): string {
  const salt = fromBase64(saltB64);
  const pw = new Uint8Array(password.length * 2);
  for (let i = 0; i < password.length; i++) {
    const code = password.charCodeAt(i);
    pw[i * 2] = code & 0xff;
    pw[i * 2 + 1] = code >> 8;
  }
  const first = new Uint8Array(salt.length + pw.length);
  first.set(salt);
  first.set(pw, salt.length);
  let hash = sha512(first);
  const buf = new Uint8Array(68);
  for (let i = 0; i < spinCount; i++) {
    buf.set(hash);
    buf[64] = i & 0xff;
    buf[65] = (i >> 8) & 0xff;
    buf[66] = (i >> 16) & 0xff;
    buf[67] = (i >> 24) & 0xff;
    hash = sha512(buf);
  }
  return toBase64(hash);
}

export function passwordMatches(hash: PasswordHash | undefined, password: string): boolean {
  if (!hash) return true;
  if (hash.legacy) return legacyHash(password) === hash.legacy.toUpperCase();
  if (hash.hashValue && hash.saltValue && /sha-?512/i.test(hash.algorithm ?? '')) {
    return excelSha512(password, hash.saltValue, hash.spinCount ?? 100000) === hash.hashValue;
  }
  // Unknown algorithm: cannot verify; Excel files rarely use others.
  return false;
}

export function makePassword(password: string): PasswordHash | undefined {
  return password ? { legacy: legacyHash(password) } : undefined;
}
