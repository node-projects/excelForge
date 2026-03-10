/** Convert column index (1-based) to letter(s): 1→A, 27→AA */
export function colIndexToLetter(n: number): string {
  let s = '';
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
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

/** Escape XML entities */
export function escapeXml(s: string): string {
  return s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
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
  let bin = '';
  for (let i = 0; i < data.length; i++) bin += String.fromCharCode(data[i]);
  return btoa(bin);
}
