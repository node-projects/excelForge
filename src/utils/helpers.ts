/** Convert column index (1-based) to letter(s): 1→A, 27→AA */
const _colCache: string[] = [];
export function colIndexToLetter(n: number): string {
  if (_colCache[n]) return _colCache[n];
  let s = '', v = n;
  while (v > 0) {
    const r = (v - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    v = Math.floor((v - 1) / 26);
  }
  _colCache[n] = s;
  return s;
}

/** Convert column letter(s) to 1-based index: A→1, AA→27 */
export function colLetterToIndex(col: string): number {
  let n = 0;
  for (let i = 0; i < col.length; i++) {
    n = n * 26 + (col.charCodeAt(i) - 64);
  }
  return n;
}

/** "A1" → { row: 1, col: 1 } */
export function cellRefToIndices(ref: string): { row: number; col: number } {
  const m = ref.match(/^(\$?)([A-Z]+)(\$?)(\d+)$/);
  if (!m) throw new Error(`Invalid cell ref: ${ref}`);
  return { col: colLetterToIndex(m[2]), row: parseInt(m[4], 10) };
}

/** { row: 1, col: 1 } → "A1" */
export function indicesToCellRef(row: number, col: number, abs = false): string {
  const a = abs ? '$' : '';
  return `${a}${colIndexToLetter(col)}${a}${row}`;
}

/** "A1:C3" → { startRow, startCol, endRow, endCol } */
export function parseRange(range: string) {
  const [start, end] = range.split(':');
  const s = cellRefToIndices(start.replace(/\$/g, ''));
  const e = end ? cellRefToIndices(end.replace(/\$/g, '')) : s;
  return { startRow: s.row, startCol: s.col, endRow: e.row, endCol: e.col };
}

/** Escape XML entities (single-pass) */
const _xmlEsc: Record<string, string> = { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&apos;' };
const _xmlRe = /[&<>"']/g;
export function escapeXml(s: string): string {
  return _xmlRe.test(s) ? (_xmlRe.lastIndex = 0, s.replace(_xmlRe, ch => _xmlEsc[ch])) : s;
}

/** Build XML attribute string from object, skipping undefined */
export function xmlAttrs(attrs: Record<string, string | number | boolean | undefined>): string {
  return Object.entries(attrs)
    .filter(([, v]) => v !== undefined && v !== null)
    .map(([k, v]) => `${k}="${escapeXml(String(v))}"`)
    .join(' ');
}

/** Wrap content in an XML tag */
export function xmlTag(tag: string, attrs: Record<string, any>, content?: string): string {
  const a = xmlAttrs(attrs);
  const open = a ? `<${tag} ${a}` : `<${tag}`;
  if (content === undefined) return `${open}/>`;
  return `${open}>${content}</${tag}>`;
}

const enc = new TextEncoder();
export function strToBytes(s: string): Uint8Array { return enc.encode(s); }

/** Convert Date to Excel serial number */
export function dateToSerial(d: Date, date1904 = false): number {
  const epoch = date1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 30);
  const serial = (d.getTime() - epoch) / 86400000;
  // Excel wrongly treats 1900 as leap year; adjust
  return date1904 ? serial : serial >= 60 ? serial : serial;
}

/** Simple deep clone via JSON (for style objects) */
export function deepClone<T>(v: T): T { return JSON.parse(JSON.stringify(v)); }

/** Generate a unique ID-safe name */
let _id = 1;
export function uid(): string { return `ef_${_id++}`; }

/** Convert EMU to pixels (assuming 96 dpi) */
export function emuToPx(emu: number): number { return emu / 914400 * 96; }
export function pxToEmu(px: number): number { return Math.round(px * 914400 / 96); }

/** Convert column width (Excel units) to EMU */
export function colWidthToEmu(w: number, charWidth = 7): number {
  return pxToEmu(Math.round(w * charWidth));
}

/** Convert row height (pt) to EMU */
export function rowHeightToEmu(h: number): number {
  return pxToEmu(Math.round(h * 4 / 3));
}

/** Decode base64 string to Uint8Array */
export function base64ToBytes(b64: string): Uint8Array {
  const bin = atob(b64);
  const out = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
  return out;
}

/** Encode Uint8Array to base64 */
export function bytesToBase64(data: Uint8Array): string {
  // Process in chunks to avoid call stack limits on String.fromCharCode.apply
  const chunks: string[] = [];
  for (let i = 0; i < data.length; i += 8192) {
    chunks.push(String.fromCharCode.apply(null, data.subarray(i, i + 8192) as any));
  }
  return btoa(chunks.join(''));
}

// ─── R1C1 Reference Style ─────────────────────────────────────────────────────

/**
 * Convert an A1 reference to R1C1 notation (relative to a base cell).
 * e.g. a1ToR1C1("C3", 1, 1) → "R[2]C[2]"
 *      a1ToR1C1("$C$3", 1, 1) → "R3C3"
 */
export function a1ToR1C1(ref: string, baseRow: number, baseCol: number): string {
  const m = ref.match(/^(\$?)([A-Z]+)(\$?)(\d+)$/);
  if (!m) return ref;
  const colAbs = m[1] === '$', col = colLetterToIndex(m[2]);
  const rowAbs = m[3] === '$', row = parseInt(m[4], 10);
  const rPart = rowAbs ? `R${row}` : (row === baseRow ? 'R' : `R[${row - baseRow}]`);
  const cPart = colAbs ? `C${col}` : (col === baseCol ? 'C' : `C[${col - baseCol}]`);
  return rPart + cPart;
}

/**
 * Convert R1C1 notation back to A1 (relative to a base cell).
 * e.g. r1c1ToA1("R[2]C[2]", 1, 1) → "C3"
 *      r1c1ToA1("R3C3", 1, 1) → "$C$3"
 */
export function r1c1ToA1(ref: string, baseRow: number, baseCol: number): string {
  const m = ref.match(/^R(\[(-?\d+)\]|(\d+))?C(\[(-?\d+)\]|(\d+))?$/);
  if (!m) return ref;
  let row: number, rowAbs: boolean;
  let col: number, colAbs: boolean;
  if (m[3] !== undefined) { row = parseInt(m[3], 10); rowAbs = true; }
  else if (m[2] !== undefined) { row = baseRow + parseInt(m[2], 10); rowAbs = false; }
  else { row = baseRow; rowAbs = false; }
  if (m[6] !== undefined) { col = parseInt(m[6], 10); colAbs = true; }
  else if (m[5] !== undefined) { col = baseCol + parseInt(m[5], 10); colAbs = false; }
  else { col = baseCol; colAbs = false; }
  return `${colAbs ? '$' : ''}${colIndexToLetter(col)}${rowAbs ? '$' : ''}${row}`;
}

/**
 * Convert a formula's references from A1 to R1C1 notation.
 * baseRow/baseCol is the cell containing the formula.
 */
export function formulaToR1C1(formula: string, baseRow: number, baseCol: number): string {
  return formula.replace(/(\$?)([A-Z]+)(\$?)(\d+)/g, (m, d1, c, d2, r) =>
    a1ToR1C1(`${d1}${c}${d2}${r}`, baseRow, baseCol)
  );
}

/**
 * Convert a formula's references from R1C1 to A1 notation.
 */
export function formulaFromR1C1(formula: string, baseRow: number, baseCol: number): string {
  return formula.replace(/R(\[(-?\d+)\]|(\d+))?C(\[(-?\d+)\]|(\d+))?/g, (m) =>
    r1c1ToA1(m, baseRow, baseCol)
  );
}
