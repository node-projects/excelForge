/**
 * ExcelForge — PDF Export Module (tree-shakeable, zero external dependencies).
 * Generates PDF 1.4 binary output from worksheets/workbooks with cell styling,
 * fonts, borders, fills, merged cells, number formatting, pagination,
 * page layout, headers/footers, images, and grid lines.
 */

import type { Worksheet } from '../core/Worksheet.js';
import type { Workbook } from '../core/Workbook.js';
import type {
  CellStyle, Font, Fill, PatternFill, Border, BorderSide, Alignment,
  Image, Sparkline, FormControl, CellImage, Chart, ChartSeries,
} from '../core/types.js';
import { colIndexToLetter, colLetterToIndex, base64ToBytes } from '../utils/helpers.js';
import { deflateRaw } from '../utils/zip.js';
import { FormulaEngine } from './FormulaEngine.js';

/* ═══════════════════════════════════════════════════════════════════════════ */
/*  Public Types                                                              */
/* ═══════════════════════════════════════════════════════════════════════════ */

export interface PdfExportOptions {
  /** Paper size (default 'a4') */
  paperSize?: 'letter' | 'legal' | 'a4' | 'a3' | 'tabloid';
  /** Page orientation (default 'portrait') */
  orientation?: 'portrait' | 'landscape';
  /** Margins in inches (default 0.75 all around) */
  margins?: { top?: number; bottom?: number; left?: number; right?: number };
  /** Scale content (0.1–2.0, default 1.0) */
  scale?: number;
  /** Automatically scale to fit content width on page (default true) */
  fitToWidth?: boolean;
  /** Draw cell grid lines (default true) */
  gridLines?: boolean;
  /** Print row/column headings (default false) */
  headings?: boolean;
  /** Restrict output to the worksheet's print area */
  printAreaOnly?: boolean;
  /** Skip hidden rows and columns */
  skipHidden?: boolean;
  /** Repeat header rows on each page (number of rows from top, default 0) */
  repeatRows?: number;
  /** Header text (center of top margin).  Use &P for page number, &N for total pages */
  headerText?: string;
  /** Footer text (center of bottom margin).  Use &P for page number, &N for total pages */
  footerText?: string;
  /** PDF document title metadata */
  title?: string;
  /** PDF document author metadata */
  author?: string;
  /** Default font size in points when cell has no explicit size (default 10) */
  defaultFontSize?: number;
  /** Evaluate formulas before export so calculated cells have values (default false) */
  evaluateFormulas?: boolean;
}

export interface PdfWorkbookOptions extends PdfExportOptions {
  /** Export specific sheets by name (default all) */
  sheets?: string[];
}

/* ═══════════════════════════════════════════════════════════════════════════ */
/*  Constants                                                                 */
/* ═══════════════════════════════════════════════════════════════════════════ */

/** Paper sizes in points (1 inch = 72pt) – [width, height] in portrait */
const PAPER_SIZES: Record<string, [number, number]> = {
  letter:  [612, 792],
  legal:   [612, 1008],
  a4:      [595.28, 841.89],
  a3:      [841.89, 1190.55],
  tabloid: [792, 1224],
};

/** Theme color defaults (Office standard) */
const THEME_COLORS = [
  '#000000', '#FFFFFF', '#44546A', '#E7E6E6', '#4472C4', '#ED7D31',
  '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47',
];

/* ══ Helvetica font metrics (char widths at 1000 units/em, ASCII 32–126) ═══ */

const HELV: number[] = [
  278,278,355,556,556,889,667,191,333,333,389,584,278,333,278,278,
  556,556,556,556,556,556,556,556,556,556,278,278,584,584,584,556,
  1015,667,667,722,722,667,611,778,722,278,500,667,556,833,722,778,
  667,778,722,667,611,722,667,944,667,667,611,278,278,278,469,556,
  333,556,556,500,556,556,278,556,556,222,222,500,222,833,556,556,
  556,556,333,500,278,556,500,722,500,500,500,334,260,334,584,
];

const HELV_B: number[] = [
  278,333,474,556,556,889,722,238,333,333,389,584,278,333,278,278,
  556,556,556,556,556,556,556,556,556,556,333,333,584,584,584,611,
  975,722,722,722,722,667,611,778,722,278,556,722,611,833,722,778,
  667,778,722,667,611,722,667,944,667,667,611,333,278,333,584,556,
  333,556,611,556,611,556,333,611,611,278,278,556,278,889,611,611,
  611,611,389,556,333,611,556,778,556,556,500,389,280,389,584,
];

/** Measure text width in points for a given font size using Helvetica metrics */
function textWidthPt(text: string, fontSize: number, bold: boolean): number {
  const tbl = bold ? HELV_B : HELV;
  let w = 0;
  for (let i = 0; i < text.length; i++) {
    const code = text.charCodeAt(i);
    w += (code >= 32 && code <= 126) ? tbl[code - 32] : 500;   // 500 fallback
  }
  return (w / 1000) * fontSize;
}

/* ═══════════════════════════════════════════════════════════════════════════ */
/*  Color helpers                                                             */
/* ═══════════════════════════════════════════════════════════════════════════ */

function colorToHex(c: string | undefined): string {
  if (!c) return '';
  if (c.startsWith('#')) return c;
  if (c.startsWith('theme:')) {
    const idx = parseInt(c.slice(6), 10);
    return THEME_COLORS[idx] ?? '#000000';
  }
  // AARRGGBB → #RRGGBB
  if (c.length === 8 && !c.startsWith('#')) return '#' + c.slice(2);
  return '#' + c;
}

/** Parse hex color to PDF RGB triplet (0–1 range) */
function hexToRgb(hex: string): [number, number, number] {
  const h = hex.replace('#', '');
  return [
    parseInt(h.slice(0, 2), 16) / 255,
    parseInt(h.slice(2, 4), 16) / 255,
    parseInt(h.slice(4, 6), 16) / 255,
  ];
}

function colorRgb(c: string | undefined): [number, number, number] | null {
  const h = colorToHex(c);
  return h ? hexToRgb(h) : null;
}

/* ═══════════════════════════════════════════════════════════════════════════ */
/*  Number formatting (subset from HtmlModule)                                */
/* ═══════════════════════════════════════════════════════════════════════════ */

function formatNumber(value: unknown, fmt: string | undefined): string {
  if (value == null) return '';
  if (!fmt || fmt === 'General') return String(value);
  const num = typeof value === 'number' ? value : parseFloat(String(value));
  if (isNaN(num)) return String(value);

  if (fmt.includes('%')) {
    const decimals = (fmt.match(/0\.(0+)%/) ?? [])[1]?.length ?? 0;
    return (num * 100).toFixed(decimals) + '%';
  }
  const currMatch = fmt.match(/[$€£¥]|"CHF"/);
  if (currMatch) {
    const sym = currMatch[0].replace(/"/g, '');
    const decimals = (fmt.match(/\.(0+)/) ?? [])[1]?.length ?? 2;
    const formatted = Math.abs(num).toFixed(decimals).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
    if (fmt.indexOf(currMatch[0]) < fmt.indexOf('0'))
      return (num < 0 ? '-' : '') + sym + formatted;
    return (num < 0 ? '-' : '') + formatted + ' ' + sym;
  }
  if (fmt.includes('#,##0') || fmt.includes('#,###')) {
    const decimals = (fmt.match(/\.(0+)/) ?? [])[1]?.length ?? 0;
    return num.toFixed(decimals).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  }
  const fixedMatch = fmt.match(/^0\.(0+)$/);
  if (fixedMatch) return num.toFixed(fixedMatch[1].length);
  if (/[ymdh]/i.test(fmt)) return formatDate(num, fmt);
  if (/0\.0+E\+0+/i.test(fmt)) {
    const decimals = (fmt.match(/0\.(0+)/) ?? [])[1]?.length ?? 2;
    return num.toExponential(decimals).toUpperCase();
  }
  return String(value);
}

function formatDate(serial: number, fmt: string): string {
  const epoch = new Date(1899, 11, 30);
  const d = new Date(epoch.getTime() + serial * 86400000);
  const Y = d.getFullYear(), M = d.getMonth() + 1, D = d.getDate();
  const h = d.getHours(), m = d.getMinutes(), s = d.getSeconds();
  return fmt
    .replace(/yyyy/gi, String(Y))
    .replace(/yy/gi, String(Y).slice(-2))
    .replace(/mmmm/gi, d.toLocaleDateString('en', { month: 'long' }))
    .replace(/mmm/gi, d.toLocaleDateString('en', { month: 'short' }))
    .replace(/mm/gi, String(M).padStart(2, '0'))
    .replace(/m/gi, String(M))
    .replace(/dd/gi, String(D).padStart(2, '0'))
    .replace(/d/gi, String(D))
    .replace(/hh/gi, String(h).padStart(2, '0'))
    .replace(/h/gi, String(h))
    .replace(/ss/gi, String(s).padStart(2, '0'))
    .replace(/nn|MM/g, String(m).padStart(2, '0'));
}

/* ═══════════════════════════════════════════════════════════════════════════ */
/*  PDF string/stream helpers                                                 */
/* ═══════════════════════════════════════════════════════════════════════════ */

/** Escape text for a PDF literal string: ( ) \ → escaped */
function pdfStr(s: string): string {
  return s.replace(/\\/g, '\\\\').replace(/\(/g, '\\(').replace(/\)/g, '\\)');
}

/** Encode a JS string to WinAnsiEncoding bytes for PDF text operators */
function encodeWinAnsi(s: string): Uint8Array {
  const out = new Uint8Array(s.length);
  for (let i = 0; i < s.length; i++) {
    const code = s.charCodeAt(i);
    out[i] = code < 256 ? code : 63; // '?' for non-Latin
  }
  return out;
}

/** Format a number for PDF operators (compact, no trailing zeros) */
function n(v: number): string {
  return +v.toFixed(4) + '';
}

const _enc = new TextEncoder();

/* ═══════════════════════════════════════════════════════════════════════════ */
/*  Low-level PDF Document Builder                                            */
/* ═══════════════════════════════════════════════════════════════════════════ */

class PdfDoc {
  private objects: (string | null)[] = [null]; // 1-based; index 0 unused
  private streams: (Uint8Array | null)[] = [null];

  /** Allocate an object ID (1-based) for forward references */
  alloc(): number {
    this.objects.push(null);
    this.streams.push(null);
    return this.objects.length - 1;
  }

  /** Define a previously allocated or new object */
  set(id: number, dict: string, stream?: Uint8Array): void {
    while (this.objects.length <= id) { this.objects.push(null); this.streams.push(null); }
    this.objects[id] = dict;
    this.streams[id] = stream ?? null;
  }

  /** Add a new object and return its ID */
  add(dict: string, stream?: Uint8Array): number {
    const id = this.alloc();
    this.set(id, dict, stream);
    return id;
  }

  /** Add a deflate-compressed stream object (zlib-wrapped for PDF /FlateDecode) */
  addDeflated(dict: string, data: Uint8Array): number {
    const raw = deflateRaw(data, 6);
    // PDF /FlateDecode expects zlib format: 2-byte header + raw deflate + 4-byte Adler32
    const zlib = new Uint8Array(2 + raw.length + 4);
    zlib[0] = 0x78; zlib[1] = 0x9C; // zlib header: deflate, default compression
    zlib.set(raw, 2);
    // Compute Adler32 of the original uncompressed data
    let a = 1, b = 0;
    for (let i = 0; i < data.length; i++) {
      a = (a + data[i]) % 65521;
      b = (b + a) % 65521;
    }
    const adler = ((b << 16) | a) >>> 0;
    const off = 2 + raw.length;
    zlib[off]     = (adler >>> 24) & 0xFF;
    zlib[off + 1] = (adler >>> 16) & 0xFF;
    zlib[off + 2] = (adler >>> 8) & 0xFF;
    zlib[off + 3] = adler & 0xFF;
    const full = `${dict}/Filter/FlateDecode/Length ${zlib.length}>>`;
    return this.add(full, zlib);
  }

  /** Build the complete PDF file as Uint8Array */
  build(rootId: number, infoId?: number): Uint8Array {
    const parts: Uint8Array[] = [];
    const push = (s: string) => parts.push(_enc.encode(s));
    const pushRaw = (b: Uint8Array) => parts.push(b);

    // Header
    push('%PDF-1.4\n%\xE2\xE3\xCF\xD3\n');

    // Calculate byte offsets
    let offset = 0;
    const offsets: number[] = new Array(this.objects.length).fill(0);
    // Pre-pass: compute header length
    for (const p of parts) offset += p.length;

    const objParts: Uint8Array[][] = [];
    for (let i = 1; i < this.objects.length; i++) {
      const dict = this.objects[i];
      if (!dict) continue;
      const subParts: Uint8Array[] = [];
      const stream = this.streams[i];
      offsets[i] = offset;
      const header = `${i} 0 obj\n`;
      subParts.push(_enc.encode(header));
      offset += header.length;
      if (stream) {
        const dictLine = dict + '\nstream\n';
        subParts.push(_enc.encode(dictLine));
        offset += dictLine.length;
        subParts.push(stream);
        offset += stream.length;
        const tail = '\nendstream\nendobj\n';
        subParts.push(_enc.encode(tail));
        offset += tail.length;
      } else {
        const body = dict + '\nendobj\n';
        subParts.push(_enc.encode(body));
        offset += body.length;
      }
      objParts.push(subParts);
    }

    // Write objects
    for (const sp of objParts) for (const p of sp) pushRaw(p);

    // Cross-reference table
    const xrefOffset = offset;
    push(`xref\n0 ${this.objects.length}\n`);
    push('0000000000 65535 f \n');
    for (let i = 1; i < this.objects.length; i++) {
      push(`${String(offsets[i]).padStart(10, '0')} 00000 n \n`);
    }

    // Trailer
    push(`trailer\n<</Size ${this.objects.length}/Root ${rootId} 0 R`);
    if (infoId) push(`/Info ${infoId} 0 R`);
    push(`>>\nstartxref\n${xrefOffset}\n%%EOF\n`);

    // Combine
    let total = 0;
    for (const p of parts) total += p.length;
    const result = new Uint8Array(total);
    let pos = 0;
    for (const p of parts) { result.set(p, pos); pos += p.length; }
    return result;
  }
}

/* ═══════════════════════════════════════════════════════════════════════════ */
/*  Content stream builder — generates PDF drawing operators                  */
/* ═══════════════════════════════════════════════════════════════════════════ */

class StreamBuilder {
  private _parts: string[] = [];

  /** Emit a raw PDF operator string */
  raw(s: string): this { this._parts.push(s); return this; }

  /** Save graphics state */
  gsave(): this { this._parts.push('q'); return this; }
  /** Restore graphics state */
  grestore(): this { this._parts.push('Q'); return this; }

  /** Set fill color (RGB 0–1) */
  fillColor(r: number, g: number, b: number): this {
    this._parts.push(`${n(r)} ${n(g)} ${n(b)} rg`);
    return this;
  }
  /** Set stroke color */
  strokeColor(r: number, g: number, b: number): this {
    this._parts.push(`${n(r)} ${n(g)} ${n(b)} RG`);
    return this;
  }
  /** Set line width */
  lineWidth(w: number): this { this._parts.push(`${n(w)} w`); return this; }
  /** Set dash pattern */
  dash(pattern: number[], phase: number): this {
    this._parts.push(`[${pattern.map(n).join(' ')}] ${n(phase)} d`);
    return this;
  }
  /** No dash (solid) */
  noDash(): this { this._parts.push('[] 0 d'); return this; }

  /** Draw a filled rectangle */
  fillRect(x: number, y: number, w: number, h: number): this {
    this._parts.push(`${n(x)} ${n(y)} ${n(w)} ${n(h)} re f`);
    return this;
  }
  /** Draw a stroked rectangle */
  strokeRect(x: number, y: number, w: number, h: number): this {
    this._parts.push(`${n(x)} ${n(y)} ${n(w)} ${n(h)} re S`);
    return this;
  }
  /** Draw a line */
  line(x1: number, y1: number, x2: number, y2: number): this {
    this._parts.push(`${n(x1)} ${n(y1)} m ${n(x2)} ${n(y2)} l S`);
    return this;
  }

  /** Begin text block */
  beginText(): this { this._parts.push('BT'); return this; }
  /** End text block */
  endText(): this { this._parts.push('ET'); return this; }
  /** Set font */
  font(name: string, size: number): this {
    this._parts.push(`/${name} ${n(size)} Tf`);
    return this;
  }
  /** Set text position */
  textPos(x: number, y: number): this {
    this._parts.push(`${n(x)} ${n(y)} Td`);
    return this;
  }
  /** Show text string */
  showText(s: string): this {
    this._parts.push(`(${pdfStr(s)}) Tj`);
    return this;
  }

  /** Draw an image XObject */
  drawImage(name: string, x: number, y: number, w: number, h: number): this {
    this.gsave();
    this._parts.push(`${n(w)} 0 0 ${n(h)} ${n(x)} ${n(y)} cm`);
    this._parts.push(`/${name} Do`);
    this.grestore();
    return this;
  }

  /** Clip to rectangle */
  clipRect(x: number, y: number, w: number, h: number): this {
    this._parts.push(`${n(x)} ${n(y)} ${n(w)} ${n(h)} re W n`);
    return this;
  }

  toBytes(): Uint8Array {
    return _enc.encode(this._parts.join('\n') + '\n');
  }
}

/* ═══════════════════════════════════════════════════════════════════════════ */
/*  Border style mapping                                                      */
/* ═══════════════════════════════════════════════════════════════════════════ */

function borderWidth(style: string | undefined): number {
  if (!style) return 0;
  const map: Record<string, number> = {
    thin: 0.5, medium: 1, thick: 1.5, dashed: 0.5, dotted: 0.5,
    double: 1.5, hair: 0.25, mediumDashed: 1, dashDot: 0.5,
    mediumDashDot: 1, dashDotDot: 0.5, mediumDashDotDot: 1, slantDashDot: 1,
  };
  return map[style] ?? 0.5;
}

function borderDash(style: string | undefined): number[] | null {
  if (!style) return null;
  const map: Record<string, number[]> = {
    dashed: [3, 2], dotted: [1, 1], mediumDashed: [4, 2],
    dashDot: [3, 1, 1, 1], mediumDashDot: [4, 1, 1, 1],
    dashDotDot: [3, 1, 1, 1, 1, 1], mediumDashDotDot: [4, 1, 1, 1, 1, 1],
    slantDashDot: [4, 1, 2, 1],
  };
  return map[style] ?? null;
}

/* ═══════════════════════════════════════════════════════════════════════════ */
/*  Image helpers                                                             */
/* ═══════════════════════════════════════════════════════════════════════════ */

interface PdfImage {
  id: number;       // PDF object ID
  name: string;     // resource name (Im1, Im2, ...)
  width: number;
  height: number;
}

/** Parse JPEG to get dimensions and return raw JPEG bytes for DCTDecode */
function parseJpeg(data: Uint8Array): { width: number; height: number } | null {
  if (data[0] !== 0xFF || data[1] !== 0xD8) return null; // not JPEG
  let pos = 2;
  while (pos < data.length - 1) {
    if (data[pos] !== 0xFF) break;
    const marker = data[pos + 1];
    if (marker === 0xD9) break; // EOI
    if (marker === 0xDA) break; // SOS - stop scanning
    const len = (data[pos + 2] << 8) | data[pos + 3];
    // SOF markers: C0, C1, C2
    if (marker === 0xC0 || marker === 0xC1 || marker === 0xC2) {
      const height = (data[pos + 5] << 8) | data[pos + 6];
      const width = (data[pos + 7] << 8) | data[pos + 8];
      return { width, height };
    }
    pos += 2 + len;
  }
  return null;
}

/* ═══════════════════════════════════════════════════════════════════════════ */
/*  Main export: worksheetToPdf                                               */
/* ═══════════════════════════════════════════════════════════════════════════ */

/**
 * Convert a worksheet to a PDF document (Uint8Array).
 *
 * Renders the cell grid with styles, borders, fills, merged cells,
 * number formatting, images, and automatic pagination.
 */
export function worksheetToPdf(ws: Worksheet, options: PdfExportOptions = {}): Uint8Array {
  const range = ws.getUsedRange();
  if (!range) return emptyPdf(options);

  let { startRow, startCol, endRow, endCol } = range;

  // Print area restriction
  if (options.printAreaOnly && ws.printArea) {
    const m = ws.printArea.match(/^'?[^']*'?!?\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)$/);
    if (m) {
      startCol = colLetterToIndex(m[1]); startRow = parseInt(m[2], 10);
      endCol = colLetterToIndex(m[3]); endRow = parseInt(m[4], 10);
    }
  }

  // Use worksheet page setup if available
  const wsSetup = ws.pageSetup;
  const wsMargins = ws.pageMargins;

  const paperKey = options.paperSize ?? (wsSetup?.paperSize ? paperSizeToKey(wsSetup.paperSize) : 'a4');
  const orientation = options.orientation ?? wsSetup?.orientation ?? 'portrait';
  const [pw, ph] = PAPER_SIZES[paperKey] ?? PAPER_SIZES.a4;
  const pageW = orientation === 'landscape' ? ph : pw;
  const pageH = orientation === 'landscape' ? pw : ph;

  const mar = {
    top:    (options.margins?.top ?? wsMargins?.top ?? 0.75) * 72,
    bottom: (options.margins?.bottom ?? wsMargins?.bottom ?? 0.75) * 72,
    left:   (options.margins?.left ?? wsMargins?.left ?? 0.7) * 72,
    right:  (options.margins?.right ?? wsMargins?.right ?? 0.7) * 72,
  };

  const contentW = pageW - mar.left - mar.right;
  const contentH = pageH - mar.top - mar.bottom;
  const defaultFontSize = options.defaultFontSize ?? 10;
  const drawGridLines = options.gridLines !== false;
  const drawHeadings = options.headings === true;
  const repeatRows = options.repeatRows ?? 0;
  const headerText = options.headerText ?? ws.headerFooter?.oddHeader;
  const footerText = options.footerText ?? ws.headerFooter?.oddFooter;

  // ── Gather column widths (in points) ─────────────────────────────────────

  const headingWidth = drawHeadings ? 30 : 0;
  const cols: number[] = []; // indices into sheet columns
  const colWidthsPt: number[] = [];

  for (let c = startCol; c <= endCol; c++) {
    const def = ws.getColumn(c);
    if (options.skipHidden && def?.hidden) continue;
    cols.push(c);
    // Excel default width ≈ 8.43 characters ≈ 64px ≈ 48pt
    const w = def?.width ? def.width * 6 : 48;  // Excel width units → approx points
    colWidthsPt.push(w);
  }

  // ── Calculate scale factor ───────────────────────────────────────────────

  const totalGridW = colWidthsPt.reduce((a, b) => a + b, 0) + headingWidth;
  let scale = options.scale ?? 1;
  if (options.fitToWidth !== false && totalGridW * scale > contentW) {
    scale = contentW / totalGridW;
  }
  scale = Math.max(0.1, Math.min(2, scale));

  // ── Gather row heights (in points) ───────────────────────────────────────

  const rows: number[] = [];
  const rowHeightsPt: number[] = [];
  for (let r = startRow; r <= endRow; r++) {
    const def = ws.getRow(r);
    if (options.skipHidden && def?.hidden) continue;
    rows.push(r);
    const h = def?.height ?? (defaultFontSize + 6); // default row height
    rowHeightsPt.push(h);
  }

  // ── Build merge map ──────────────────────────────────────────────────────

  const merges = ws.getMerges();
  const mergeMap = new Map<string, { rowSpan: number; colSpan: number } | 'skip'>();
  for (const m of merges) {
    const rs = m.endRow - m.startRow + 1;
    const cs = m.endCol - m.startCol + 1;
    mergeMap.set(`${m.startRow},${m.startCol}`, { rowSpan: rs, colSpan: cs });
    for (let r = m.startRow; r <= m.endRow; r++) {
      for (let c = m.startCol; c <= m.endCol; c++) {
        if (r !== m.startRow || c !== m.startCol) mergeMap.set(`${r},${c}`, 'skip');
      }
    }
  }

  // ── Build sparkline map (row,col → Sparkline) ─────────────────────────────

  const sparklineMap = new Map<string, Sparkline>();
  const sparklines = ws.getSparklines?.() ?? [];
  for (const sp of sparklines) {
    const lm = sp.location.match(/^([A-Z]+)(\d+)$/);
    if (lm) sparklineMap.set(`${parseInt(lm[2], 10)},${colLetterToIndex(lm[1])}`, sp);
  }

  // ── Build cell image map ──────────────────────────────────────────────────

  const cellImageMap = new Map<string, CellImage>();
  const cellImages = ws.getCellImages?.() ?? [];
  for (const ci of cellImages) cellImageMap.set(ci.cell, ci);

  // ── Evaluate formulas if requested ────────────────────────────────────────

  if (options.evaluateFormulas) {
    new FormulaEngine().calculateSheet(ws);
  }

  // ── Collect row breaks ────────────────────────────────────────────────────

  const rowBreakSet = new Set<number>();
  for (const brk of ws.getRowBreaks()) rowBreakSet.add(brk.id);

  // ── Pagination ────────────────────────────────────────────────────────────

  interface PageDef {
    rowStart: number;   // index into rows[]
    rowEnd: number;     // inclusive index into rows[]
  }

  const pages: PageDef[] = [];
  let pageRowStart = 0;
  let usedH = 0;
  const repeatH = rows.slice(0, repeatRows).reduce((s, _, i) => s + rowHeightsPt[i] * scale, 0);

  for (let ri = 0; ri < rows.length; ri++) {
    const rh = rowHeightsPt[ri] * scale;
    const effectiveContentH = contentH - (pages.length > 0 ? repeatH : 0);

    // Explicit page break?
    const explicitBreak = ri > pageRowStart && rowBreakSet.has(rows[ri]);

    if (usedH + rh > effectiveContentH && ri > pageRowStart || explicitBreak) {
      pages.push({ rowStart: pageRowStart, rowEnd: ri - 1 });
      pageRowStart = ri;
      usedH = repeatH + rh;
    } else {
      usedH += rh;
    }
  }
  if (pageRowStart < rows.length) {
    pages.push({ rowStart: pageRowStart, rowEnd: rows.length - 1 });
  }

  // ── Build PDF document ──────────────────────────────────────────────────

  const doc = new PdfDoc();
  const catalogId = doc.alloc();
  const pagesId = doc.alloc();

  // Fonts — use standard PDF fonts (no embedding needed)
  const fontReg = doc.add('<</Type/Font/Subtype/Type1/BaseFont/Helvetica/Encoding/WinAnsiEncoding>>');
  const fontBold = doc.add('<</Type/Font/Subtype/Type1/BaseFont/Helvetica-Bold/Encoding/WinAnsiEncoding>>');
  const fontItalic = doc.add('<</Type/Font/Subtype/Type1/BaseFont/Helvetica-Oblique/Encoding/WinAnsiEncoding>>');
  const fontBI = doc.add('<</Type/Font/Subtype/Type1/BaseFont/Helvetica-BoldOblique/Encoding/WinAnsiEncoding>>');

  // ── Process images ──────────────────────────────────────────────────────

  const images = ws.getImages?.() ?? [];
  const pdfImages: PdfImage[] = [];
  let imgCounter = 0;

  for (const img of images) {
    const data = typeof img.data === 'string' ? base64Decode(img.data) : img.data;
    if (img.format === 'jpeg' || img.format === 'png') {
      if (img.format === 'jpeg') {
        const info = parseJpeg(data);
        if (!info) continue;
        const name = `Im${++imgCounter}`;
        const objId = doc.add(
          `<</Type/XObject/Subtype/Image/Width ${info.width}/Height ${info.height}` +
          `/ColorSpace/DeviceRGB/BitsPerComponent 8/Filter/DCTDecode/Length ${data.length}>>`,
          data
        );
        pdfImages.push({ id: objId, name, width: info.width, height: info.height });
      }
      // PNG: parse and embed with FlateDecode + PNG predictor
      if (img.format === 'png') {
        const pngImg = parsePngForPdf(data, doc);
        if (pngImg) { pngImg.name = `Im${++imgCounter}`; pdfImages.push(pngImg); }
      }
    }
  }

  // Build image resource dict fragment (will be augmented by cell images per page)
  const cellImgResources: { name: string; id: number }[] = [];

  function buildImgResources(): string {
    const all = [...pdfImages.map(im => ({ name: im.name, id: im.id })), ...cellImgResources];
    if (!all.length) return '';
    return '/XObject<<' + all.map(im => `/${im.name} ${im.id} 0 R`).join('') + '>>';
  }

  // ── Info dictionary ─────────────────────────────────────────────────────

  let infoId: number | undefined;
  if (options.title || options.author) {
    const infoParts = ['/Producer(ExcelForge)'];
    if (options.title) infoParts.push(`/Title(${pdfStr(options.title)})`);
    if (options.author) infoParts.push(`/Author(${pdfStr(options.author)})`);
    infoId = doc.add(`<<${infoParts.join('')}>>`);
  }

  // ── Render pages ──────────────────────────────────────────────────────────

  const pageIds: number[] = [];

  for (let pi = 0; pi < pages.length; pi++) {
    const page = pages[pi];
    const sb = new StreamBuilder();

    // Apply scale
    if (scale !== 1) {
      sb.gsave();
      sb.raw(`${n(scale)} 0 0 ${n(scale)} ${n(mar.left * (1 - scale))} ${n(mar.bottom * (1 - scale))} cm`);
    }

    // ── Determine rows to render on this page ────────────────────────────

    const pageRows: { sheetRow: number; height: number; ri: number }[] = [];

    // Repeat header rows (pages after first)
    if (pi > 0 && repeatRows > 0) {
      for (let ri = 0; ri < Math.min(repeatRows, rows.length); ri++) {
        pageRows.push({ sheetRow: rows[ri], height: rowHeightsPt[ri], ri });
      }
    }

    for (let ri = page.rowStart; ri <= page.rowEnd; ri++) {
      pageRows.push({ sheetRow: rows[ri], height: rowHeightsPt[ri], ri });
    }

    // ── Draw cells ──────────────────────────────────────────────────────

    let curY = pageH - mar.top; // start from top (PDF Y is bottom-up)

    for (const pr of pageRows) {
      const rh = pr.height * scale;
      let curX = mar.left;

      // Row heading
      if (drawHeadings) {
        sb.gsave();
        sb.fillColor(0.9, 0.9, 0.9);
        sb.fillRect(curX, curY - rh, headingWidth * scale, rh);
        sb.strokeColor(0.7, 0.7, 0.7).lineWidth(0.25);
        sb.strokeRect(curX, curY - rh, headingWidth * scale, rh);
        sb.fillColor(0.3, 0.3, 0.3);
        sb.beginText().font('F1', 7 * scale);
        const label = String(pr.sheetRow);
        const lw = textWidthPt(label, 7, false) * scale;
        sb.textPos(curX + (headingWidth * scale - lw) / 2, curY - rh + (rh - 7 * scale) / 2 + 1);
        sb.showText(label).endText();
        sb.grestore();
        curX += headingWidth * scale;
      }

      // Column heading row (only at top of first page)
      // (We'll draw column headings as part of the first row processing)

      for (let ci = 0; ci < cols.length; ci++) {
        const sheetCol = cols[ci];
        const cw = colWidthsPt[ci] * scale;
        const key = `${pr.sheetRow},${sheetCol}`;
        const merge = mergeMap.get(key);
        if (merge === 'skip') { curX += cw; continue; }

        // Calculate cell dimensions (accounting for merged cells)
        let cellW = cw;
        let cellH = rh;
        if (merge && typeof merge !== 'string') {
          // Sum widths of merged columns
          cellW = 0;
          for (let mc = 0; mc < merge.colSpan; mc++) {
            const idx = cols.indexOf(sheetCol + mc);
            if (idx >= 0) cellW += colWidthsPt[idx] * scale;
          }
          // Sum heights of merged rows
          cellH = 0;
          for (let mr = 0; mr < merge.rowSpan; mr++) {
            const idx = rows.indexOf(pr.sheetRow + mr);
            if (idx >= 0) cellH += rowHeightsPt[idx] * scale;
          }
        }

        const cell = ws.getCell(pr.sheetRow, sheetCol);
        const style = cell.style;

        // ── Cell background fill ────────────────────────────────────────
        if (style?.fill && style.fill.type === 'pattern') {
          const pf = style.fill as PatternFill;
          if (pf.pattern === 'solid' && pf.fgColor) {
            const rgb = colorRgb(pf.fgColor);
            if (rgb) {
              sb.gsave();
              sb.fillColor(rgb[0], rgb[1], rgb[2]);
              sb.fillRect(curX, curY - cellH, cellW, cellH);
              sb.grestore();
            }
          }
        }

        // ── Grid lines ──────────────────────────────────────────────────
        if (drawGridLines && !style?.border) {
          sb.gsave();
          sb.strokeColor(0.82, 0.82, 0.82).lineWidth(0.25);
          sb.strokeRect(curX, curY - cellH, cellW, cellH);
          sb.grestore();
        }

        // ── Cell borders ────────────────────────────────────────────────
        if (style?.border) {
          sb.gsave();
          drawBorder(sb, style.border, curX, curY - cellH, cellW, cellH);
          sb.grestore();
        }

        // ── Cell text ───────────────────────────────────────────────────
        const textVal = getCellText(cell, style);
        if (textVal) {
          sb.gsave();
          // Clip to cell bounds
          sb.clipRect(curX + 1, curY - cellH, cellW - 2, cellH);

          const fontSize = (style?.font?.size ?? defaultFontSize) * scale;
          const bold = style?.font?.bold ?? false;
          const italic = style?.font?.italic ?? false;
          const fontName = bold && italic ? 'F4' : bold ? 'F2' : italic ? 'F3' : 'F1';

          const rgb = colorRgb(style?.font?.color) ?? [0, 0, 0];
          sb.fillColor(rgb[0], rgb[1], rgb[2]);

          // Calculate text position based on alignment
          const tw = textWidthPt(textVal, fontSize / scale, bold) * scale;
          const hAlign = style?.alignment?.horizontal ?? (typeof cell.value === 'number' ? 'right' : 'left');
          const vAlign = style?.alignment?.vertical ?? 'bottom';

          let tx: number;
          const pad = 2 * scale;
          switch (hAlign) {
            case 'center': case 'fill': case 'justify': case 'distributed':
              tx = curX + (cellW - tw) / 2;
              break;
            case 'right':
              tx = curX + cellW - tw - pad;
              break;
            default: // left
              tx = curX + pad + (style?.alignment?.indent ?? 0) * 6 * scale;
              break;
          }

          let ty: number;
          switch (vAlign) {
            case 'top':
              ty = curY - fontSize - pad;
              break;
            case 'center': case 'distributed':
              ty = curY - cellH / 2 - fontSize * 0.35;
              break;
            default: // bottom
              ty = curY - cellH + pad;
              break;
          }

          sb.beginText().font(fontName, fontSize).textPos(tx, ty).showText(textVal).endText();

          // Underline
          if (style?.font?.underline && style.font.underline !== 'none') {
            sb.strokeColor(rgb[0], rgb[1], rgb[2]);
            sb.lineWidth(fontSize * 0.05);
            sb.line(tx, ty - fontSize * 0.15, tx + tw, ty - fontSize * 0.15);
          }
          // Strikethrough
          if (style?.font?.strike) {
            sb.strokeColor(rgb[0], rgb[1], rgb[2]);
            sb.lineWidth(fontSize * 0.05);
            const sy = ty + fontSize * 0.3;
            sb.line(tx, sy, tx + tw, sy);
          }

          sb.grestore();
        }

        // ── Sparkline in cell ─────────────────────────────────────────────
        const spKey = `${pr.sheetRow},${sheetCol}`;
        const sp = sparklineMap.get(spKey);
        if (sp) {
          const spValues = resolveSparklineValues(ws, sp.dataRange);
          if (spValues.length) drawSparkline(sb, sp, spValues, curX, curY - cellH, cellW, cellH);
        }

        // ── Cell image ────────────────────────────────────────────────────
        const cellRef = `${colIndexToLetter(sheetCol)}${pr.sheetRow}`;
        const ciImg = cellImageMap.get(cellRef);
        if (ciImg) {
          const ciData = typeof ciImg.data === 'string' ? base64Decode(ciImg.data) : ciImg.data;
          if (ciImg.format === 'jpeg') {
            const info = parseJpeg(ciData);
            if (info) {
              const ciName = `Ci${pr.sheetRow}_${sheetCol}`;
              const ciObjId = doc.add(
                `<</Type/XObject/Subtype/Image/Width ${info.width}/Height ${info.height}` +
                `/ColorSpace/DeviceRGB/BitsPerComponent 8/Filter/DCTDecode/Length ${ciData.length}>>`,
                ciData
              );
              cellImgResources.push({ name: ciName, id: ciObjId });
              const aspect = info.width / info.height;
              let iw = cellW - 2, ih = cellH - 2;
              if (iw / ih > aspect) iw = ih * aspect; else ih = iw / aspect;
              sb.drawImage(ciName, curX + 1, curY - cellH + 1, iw, ih);
            }
          } else if (ciImg.format === 'png') {
            const pngImg = parsePngForPdf(ciData, doc);
            if (pngImg) {
              const ciName = `Ci${pr.sheetRow}_${sheetCol}`;
              pngImg.name = ciName;
              cellImgResources.push({ name: ciName, id: pngImg.id });
              const aspect = pngImg.width / pngImg.height;
              let iw = cellW - 2, ih = cellH - 2;
              if (iw / ih > aspect) iw = ih * aspect; else ih = iw / aspect;
              sb.drawImage(ciName, curX + 1, curY - cellH + 1, iw, ih);
            }
          }
        }

        curX += cw;
      }

      curY -= rh;
    }

    // ── Column headings ─────────────────────────────────────────────────

    if (drawHeadings) {
      const headY = pageH - mar.top;
      let headX = mar.left + headingWidth * scale;
      sb.gsave();
      for (let ci = 0; ci < cols.length; ci++) {
        const cw = colWidthsPt[ci] * scale;
        sb.fillColor(0.9, 0.9, 0.9);
        sb.fillRect(headX, headY, cw, 14 * scale);
        sb.strokeColor(0.7, 0.7, 0.7).lineWidth(0.25);
        sb.strokeRect(headX, headY, cw, 14 * scale);
        sb.fillColor(0.3, 0.3, 0.3);
        const label = colIndexToLetter(cols[ci]);
        const lw = textWidthPt(label, 7, false) * scale;
        sb.beginText().font('F1', 7 * scale);
        sb.textPos(headX + (cw - lw) / 2, headY + 3 * scale);
        sb.showText(label).endText();
        headX += cw;
      }
      sb.grestore();
    }

    // ── Images on this page ─────────────────────────────────────────────

    for (let ii = 0; ii < pdfImages.length; ii++) {
      const img = images[ii];
      const pImg = pdfImages[ii];
      if (!img || !pImg) continue;
      const pos = resolveImagePos(img, cols, rows, colWidthsPt, rowHeightsPt, scale, mar, pageH, startCol, startRow);
      if (pos) {
        sb.drawImage(pImg.name, pos.x, pos.y, pos.w, pos.h);
      }
    }

    // ── Form controls on this page ──────────────────────────────────────

    const fcs = ws.getFormControls?.() ?? [];
    for (const fc of fcs) {
      drawFormControl(sb, fc, cols, rows, colWidthsPt, rowHeightsPt, scale, mar, pageH);
    }

    // ── Header/footer ───────────────────────────────────────────────────

    if (headerText) {
      const text = replaceHFTokens(headerText, pi + 1, pages.length);
      sb.gsave();
      sb.fillColor(0.3, 0.3, 0.3);
      sb.beginText().font('F1', 8);
      const tw = textWidthPt(text, 8, false);
      sb.textPos(pageW / 2 - tw / 2, pageH - mar.top / 2);
      sb.showText(text).endText();
      sb.grestore();
    }
    if (footerText) {
      const text = replaceHFTokens(footerText, pi + 1, pages.length);
      sb.gsave();
      sb.fillColor(0.3, 0.3, 0.3);
      sb.beginText().font('F1', 8);
      const tw = textWidthPt(text, 8, false);
      sb.textPos(pageW / 2 - tw / 2, mar.bottom / 2);
      sb.showText(text).endText();
      sb.grestore();
    }

    if (scale !== 1) sb.grestore();

    // Create page content stream
    const streamData = sb.toBytes();
    const contentId = doc.addDeflated('<<', streamData);

    // Page object
    const pageId = doc.add(
      `<</Type/Page/Parent ${pagesId} 0 R/MediaBox[0 0 ${n(pageW)} ${n(pageH)}]` +
      `/Contents ${contentId} 0 R` +
      `/Resources<</Font<</F1 ${fontReg} 0 R/F2 ${fontBold} 0 R/F3 ${fontItalic} 0 R/F4 ${fontBI} 0 R>>` +
      buildImgResources() +
      `>>>>`
    );
    pageIds.push(pageId);
  }

  // ── Pages & Catalog ────────────────────────────────────────────────────

  doc.set(pagesId,
    `<</Type/Pages/Kids[${pageIds.map(id => `${id} 0 R`).join(' ')}]/Count ${pageIds.length}>>`
  );
  doc.set(catalogId, `<</Type/Catalog/Pages ${pagesId} 0 R>>`);

  return doc.build(catalogId, infoId);
}

/* ═══════════════════════════════════════════════════════════════════════════ */
/*  Workbook export                                                           */
/* ═══════════════════════════════════════════════════════════════════════════ */

/**
 * Export an entire workbook as a single PDF document with each sheet as pages.
 */
export function workbookToPdf(wb: Workbook, options: PdfWorkbookOptions = {}): Uint8Array {
  const sheets = wb.getSheets();
  const names = wb.getSheetNames();
  const selected = options.sheets ?? names;

  // Evaluate formulas across workbook if requested
  if (options.evaluateFormulas) {
    new FormulaEngine().calculateWorkbook(wb);
  }

  const filtered = sheets.filter((_, i) => selected.includes(names[i]));
  if (filtered.length === 0) return emptyPdf(options);
  if (filtered.length === 1 && !filtered[0]._isChartSheet && !filtered[0]._isDialogSheet)
    return worksheetToPdf(filtered[0], options);

  // Multi-sheet: build a single PDF with all pages
  const doc = new PdfDoc();
  const catalogId = doc.alloc();
  const pagesId = doc.alloc();

  const fontReg = doc.add('<</Type/Font/Subtype/Type1/BaseFont/Helvetica/Encoding/WinAnsiEncoding>>');
  const fontBold = doc.add('<</Type/Font/Subtype/Type1/BaseFont/Helvetica-Bold/Encoding/WinAnsiEncoding>>');
  const fontItalic = doc.add('<</Type/Font/Subtype/Type1/BaseFont/Helvetica-Oblique/Encoding/WinAnsiEncoding>>');
  const fontBI = doc.add('<</Type/Font/Subtype/Type1/BaseFont/Helvetica-BoldOblique/Encoding/WinAnsiEncoding>>');

  let infoId: number | undefined;
  if (options.title || options.author) {
    const infoParts = ['/Producer(ExcelForge)'];
    if (options.title) infoParts.push(`/Title(${pdfStr(options.title)})`);
    if (options.author) infoParts.push(`/Author(${pdfStr(options.author)})`);
    infoId = doc.add(`<<${infoParts.join('')}>>`);
  }

  const allPageIds: number[] = [];

  const paperKey = options.paperSize ?? 'a4';
  const orientation = options.orientation ?? 'portrait';
  const [pw, ph] = PAPER_SIZES[paperKey] ?? PAPER_SIZES.a4;
  const pageW = orientation === 'landscape' ? ph : pw;
  const pageH = orientation === 'landscape' ? pw : ph;

  for (const ws of filtered) {
    if (ws._isChartSheet) {
      // Render chart sheet as a full-page chart
      const charts = ws.getCharts();
      if (charts.length) {
        const sb = new StreamBuilder();
        drawChartOnPage(sb, charts[0], ws, 40, 40, pageW - 80, pageH - 80);
        const streamData = sb.toBytes();
        const contentId = doc.addDeflated('<<', streamData);
        const pageId = doc.add(
          `<</Type/Page/Parent ${pagesId} 0 R/MediaBox[0 0 ${n(pageW)} ${n(pageH)}]` +
          `/Contents ${contentId} 0 R` +
          `/Resources<</Font<</F1 ${fontReg} 0 R/F2 ${fontBold} 0 R>>>>>>`
        );
        allPageIds.push(pageId);
      }
    } else if (ws._isDialogSheet) {
      // Render dialog sheet form controls
      const fcs = ws.getFormControls?.() ?? [];
      if (fcs.length) {
        const sb = new StreamBuilder();
        const mar = { left: 50, top: 50 };
        const dummyCols = [1], dummyRows = [1], dummyColW = [48], dummyRowH = [16];
        for (const fc of fcs) {
          drawFormControl(sb, fc, dummyCols, dummyRows, dummyColW, dummyRowH, 1, mar, pageH);
        }
        const streamData = sb.toBytes();
        const contentId = doc.addDeflated('<<', streamData);
        const pageId = doc.add(
          `<</Type/Page/Parent ${pagesId} 0 R/MediaBox[0 0 ${n(pageW)} ${n(pageH)}]` +
          `/Contents ${contentId} 0 R` +
          `/Resources<</Font<</F1 ${fontReg} 0 R/F2 ${fontBold} 0 R>>>>>>`
        );
        allPageIds.push(pageId);
      }
    } else {
      const sheetPages = renderSheetPages(ws, options, doc, fontReg, fontBold, fontItalic, fontBI, pagesId);
      allPageIds.push(...sheetPages);
    }
  }

  doc.set(pagesId,
    `<</Type/Pages/Kids[${allPageIds.map(id => `${id} 0 R`).join(' ')}]/Count ${allPageIds.length}>>`
  );
  doc.set(catalogId, `<</Type/Catalog/Pages ${pagesId} 0 R>>`);

  return doc.build(catalogId, infoId);
}

/* ═══════════════════════════════════════════════════════════════════════════ */
/*  Internal: render one sheet's pages into the shared PdfDoc                 */
/* ═══════════════════════════════════════════════════════════════════════════ */

function renderSheetPages(
  ws: Worksheet, options: PdfExportOptions, doc: PdfDoc,
  fontReg: number, fontBold: number, fontItalic: number, fontBI: number,
  pagesId: number,
): number[] {
  const range = ws.getUsedRange();
  if (!range) return [];

  let { startRow, startCol, endRow, endCol } = range;
  if (options.printAreaOnly && ws.printArea) {
    const m = ws.printArea.match(/^'?[^']*'?!?\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)$/);
    if (m) {
      startCol = colLetterToIndex(m[1]); startRow = parseInt(m[2], 10);
      endCol = colLetterToIndex(m[3]); endRow = parseInt(m[4], 10);
    }
  }

  const wsSetup = ws.pageSetup;
  const wsMargins = ws.pageMargins;
  const paperKey = options.paperSize ?? (wsSetup?.paperSize ? paperSizeToKey(wsSetup.paperSize) : 'a4');
  const orientation = options.orientation ?? wsSetup?.orientation ?? 'portrait';
  const [pw, ph] = PAPER_SIZES[paperKey] ?? PAPER_SIZES.a4;
  const pageW = orientation === 'landscape' ? ph : pw;
  const pageH = orientation === 'landscape' ? pw : ph;
  const mar = {
    top:    (options.margins?.top ?? wsMargins?.top ?? 0.75) * 72,
    bottom: (options.margins?.bottom ?? wsMargins?.bottom ?? 0.75) * 72,
    left:   (options.margins?.left ?? wsMargins?.left ?? 0.7) * 72,
    right:  (options.margins?.right ?? wsMargins?.right ?? 0.7) * 72,
  };
  const contentW = pageW - mar.left - mar.right;
  const contentH = pageH - mar.top - mar.bottom;
  const defaultFontSize = options.defaultFontSize ?? 10;
  const drawGridLines = options.gridLines !== false;
  const repeatRows = options.repeatRows ?? 0;
  const headerText = options.headerText ?? ws.headerFooter?.oddHeader;
  const footerText = options.footerText ?? ws.headerFooter?.oddFooter;

  const cols: number[] = [];
  const colWidthsPt: number[] = [];
  for (let c = startCol; c <= endCol; c++) {
    const def = ws.getColumn(c);
    if (options.skipHidden && def?.hidden) continue;
    cols.push(c);
    colWidthsPt.push(def?.width ? def.width * 6 : 48);
  }

  const totalGridW = colWidthsPt.reduce((a, b) => a + b, 0);
  let scale = options.scale ?? 1;
  if (options.fitToWidth !== false && totalGridW * scale > contentW) {
    scale = contentW / totalGridW;
  }
  scale = Math.max(0.1, Math.min(2, scale));

  const rows: number[] = [];
  const rowHeightsPt: number[] = [];
  for (let r = startRow; r <= endRow; r++) {
    const def = ws.getRow(r);
    if (options.skipHidden && def?.hidden) continue;
    rows.push(r);
    rowHeightsPt.push(def?.height ?? (defaultFontSize + 6));
  }

  const merges = ws.getMerges();
  const mergeMap = new Map<string, { rowSpan: number; colSpan: number } | 'skip'>();
  for (const m of merges) {
    mergeMap.set(`${m.startRow},${m.startCol}`, { rowSpan: m.endRow - m.startRow + 1, colSpan: m.endCol - m.startCol + 1 });
    for (let r = m.startRow; r <= m.endRow; r++)
      for (let c = m.startCol; c <= m.endCol; c++)
        if (r !== m.startRow || c !== m.startCol) mergeMap.set(`${r},${c}`, 'skip');
  }

  // Build sparkline/cell-image maps
  const sparklineMap = new Map<string, Sparkline>();
  for (const sp of (ws.getSparklines?.() ?? [])) {
    const lm = sp.location.match(/^([A-Z]+)(\d+)$/);
    if (lm) sparklineMap.set(`${parseInt(lm[2], 10)},${colLetterToIndex(lm[1])}`, sp);
  }
  const cellImageMap = new Map<string, CellImage>();
  for (const ci of (ws.getCellImages?.() ?? [])) cellImageMap.set(ci.cell, ci);

  const rowBreakSet = new Set<number>();
  for (const brk of ws.getRowBreaks()) rowBreakSet.add(brk.id);

  interface PageDef { rowStart: number; rowEnd: number }
  const pages: PageDef[] = [];
  let pageRowStart = 0; let usedH = 0;
  const repeatH = rows.slice(0, repeatRows).reduce((s, _, i) => s + rowHeightsPt[i] * scale, 0);

  for (let ri = 0; ri < rows.length; ri++) {
    const rh = rowHeightsPt[ri] * scale;
    const effectiveContentH = contentH - (pages.length > 0 ? repeatH : 0);
    const explicitBreak = ri > pageRowStart && rowBreakSet.has(rows[ri]);
    if ((usedH + rh > effectiveContentH && ri > pageRowStart) || explicitBreak) {
      pages.push({ rowStart: pageRowStart, rowEnd: ri - 1 });
      pageRowStart = ri; usedH = repeatH + rh;
    } else { usedH += rh; }
  }
  if (pageRowStart < rows.length) pages.push({ rowStart: pageRowStart, rowEnd: rows.length - 1 });

  const pageIds: number[] = [];

  for (let pi = 0; pi < pages.length; pi++) {
    const page = pages[pi];
    const sb = new StreamBuilder();

    const pageRows: { sheetRow: number; height: number }[] = [];
    if (pi > 0 && repeatRows > 0) {
      for (let ri = 0; ri < Math.min(repeatRows, rows.length); ri++)
        pageRows.push({ sheetRow: rows[ri], height: rowHeightsPt[ri] });
    }
    for (let ri = page.rowStart; ri <= page.rowEnd; ri++)
      pageRows.push({ sheetRow: rows[ri], height: rowHeightsPt[ri] });

    let curY = pageH - mar.top;

    for (const pr of pageRows) {
      const rh = pr.height * scale;
      let curX = mar.left;

      for (let ci = 0; ci < cols.length; ci++) {
        const sheetCol = cols[ci];
        const cw = colWidthsPt[ci] * scale;
        const key = `${pr.sheetRow},${sheetCol}`;
        const merge = mergeMap.get(key);
        if (merge === 'skip') { curX += cw; continue; }

        let cellW = cw, cellH = rh;
        if (merge && typeof merge !== 'string') {
          cellW = 0;
          for (let mc = 0; mc < merge.colSpan; mc++) {
            const idx = cols.indexOf(sheetCol + mc);
            if (idx >= 0) cellW += colWidthsPt[idx] * scale;
          }
          cellH = 0;
          for (let mr = 0; mr < merge.rowSpan; mr++) {
            const idx = rows.indexOf(pr.sheetRow + mr);
            if (idx >= 0) cellH += rowHeightsPt[idx] * scale;
          }
        }

        const cell = ws.getCell(pr.sheetRow, sheetCol);
        const style = cell.style;

        if (style?.fill && style.fill.type === 'pattern') {
          const pf = style.fill as PatternFill;
          if (pf.pattern === 'solid' && pf.fgColor) {
            const rgb = colorRgb(pf.fgColor);
            if (rgb) { sb.gsave(); sb.fillColor(rgb[0], rgb[1], rgb[2]); sb.fillRect(curX, curY - cellH, cellW, cellH); sb.grestore(); }
          }
        }

        if (drawGridLines && !style?.border) {
          sb.gsave(); sb.strokeColor(0.82, 0.82, 0.82).lineWidth(0.25);
          sb.strokeRect(curX, curY - cellH, cellW, cellH); sb.grestore();
        }
        if (style?.border) { sb.gsave(); drawBorder(sb, style.border, curX, curY - cellH, cellW, cellH); sb.grestore(); }

        const textVal = getCellText(cell, style);
        if (textVal) {
          sb.gsave(); sb.clipRect(curX + 1, curY - cellH, cellW - 2, cellH);
          const fontSize = (style?.font?.size ?? defaultFontSize) * scale;
          const bold = style?.font?.bold ?? false;
          const italic = style?.font?.italic ?? false;
          const fontName = bold && italic ? 'F4' : bold ? 'F2' : italic ? 'F3' : 'F1';
          const rgb = colorRgb(style?.font?.color) ?? [0, 0, 0];
          sb.fillColor(rgb[0], rgb[1], rgb[2]);
          const tw = textWidthPt(textVal, fontSize / scale, bold) * scale;
          const hAlign = style?.alignment?.horizontal ?? (typeof cell.value === 'number' ? 'right' : 'left');
          const vAlign = style?.alignment?.vertical ?? 'bottom';
          const pad = 2 * scale;
          let tx: number;
          switch (hAlign) {
            case 'center': case 'fill': case 'justify': case 'distributed': tx = curX + (cellW - tw) / 2; break;
            case 'right': tx = curX + cellW - tw - pad; break;
            default: tx = curX + pad + (style?.alignment?.indent ?? 0) * 6 * scale; break;
          }
          let ty: number;
          switch (vAlign) {
            case 'top': ty = curY - fontSize - pad; break;
            case 'center': case 'distributed': ty = curY - cellH / 2 - fontSize * 0.35; break;
            default: ty = curY - cellH + pad; break;
          }
          sb.beginText().font(fontName, fontSize).textPos(tx, ty).showText(textVal).endText();
          sb.grestore();
        }

        // Sparkline in cell
        const spKey = `${pr.sheetRow},${sheetCol}`;
        const sp = sparklineMap.get(spKey);
        if (sp) {
          const spVals = resolveSparklineValues(ws, sp.dataRange);
          if (spVals.length) drawSparkline(sb, sp, spVals, curX, curY - cellH, cellW, cellH);
        }

        curX += cw;
      }
      curY -= rh;
    }

    // Form controls
    const fcs = ws.getFormControls?.() ?? [];
    for (const fc of fcs) {
      drawFormControl(sb, fc, cols, rows, colWidthsPt, rowHeightsPt, scale, mar, pageH);
    }

    // Header/footer
    if (headerText) {
      const text = replaceHFTokens(headerText, pi + 1, pages.length);
      sb.gsave().fillColor(0.3, 0.3, 0.3);
      const tw = textWidthPt(text, 8, false);
      sb.beginText().font('F1', 8).textPos(pageW / 2 - tw / 2, pageH - mar.top / 2).showText(text).endText();
      sb.grestore();
    }
    if (footerText) {
      const text = replaceHFTokens(footerText, pi + 1, pages.length);
      sb.gsave().fillColor(0.3, 0.3, 0.3);
      const tw = textWidthPt(text, 8, false);
      sb.beginText().font('F1', 8).textPos(pageW / 2 - tw / 2, mar.bottom / 2).showText(text).endText();
      sb.grestore();
    }

    const streamData = sb.toBytes();
    const contentId = doc.addDeflated('<<', streamData);
    const pageId = doc.add(
      `<</Type/Page/Parent ${pagesId} 0 R/MediaBox[0 0 ${n(pageW)} ${n(pageH)}]` +
      `/Contents ${contentId} 0 R` +
      `/Resources<</Font<</F1 ${fontReg} 0 R/F2 ${fontBold} 0 R/F3 ${fontItalic} 0 R/F4 ${fontBI} 0 R>>>>>>`
    );
    pageIds.push(pageId);
  }

  return pageIds;
}

/* ═══════════════════════════════════════════════════════════════════════════ */
/*  Helper functions                                                          */
/* ═══════════════════════════════════════════════════════════════════════════ */

/** Get display text for a cell */
function getCellText(cell: { value?: unknown; richText?: { text: string }[]; style?: CellStyle }, style: CellStyle | undefined): string {
  if (cell.richText) return cell.richText.map(r => r.text).join('');
  if (cell.value == null) return '';
  if (style?.numberFormat) return formatNumber(cell.value, style.numberFormat.formatCode);
  return String(cell.value);
}

/** Draw cell borders */
function drawBorder(sb: StreamBuilder, border: Border, x: number, y: number, w: number, h: number): void {
  const sides: [BorderSide | undefined, number, number, number, number][] = [
    [border.bottom, x, y, x + w, y],
    [border.top,    x, y + h, x + w, y + h],
    [border.left,   x, y, x, y + h],
    [border.right,  x + w, y, x + w, y + h],
  ];

  for (const [side, x1, y1, x2, y2] of sides) {
    if (!side?.style) continue;
    const bw = borderWidth(side.style);
    const rgb = colorRgb(side.color) ?? [0, 0, 0];
    sb.strokeColor(rgb[0], rgb[1], rgb[2]);
    sb.lineWidth(bw);
    const dp = borderDash(side.style);
    if (dp) sb.dash(dp, 0); else sb.noDash();
    sb.line(x1, y1, x2, y2);
  }
}

/** Map Excel PaperSize number to key */
function paperSizeToKey(ps: number): string {
  const map: Record<number, string> = {
    1: 'letter', 5: 'legal', 9: 'a4', 8: 'a3', 3: 'tabloid',
  };
  return map[ps] ?? 'a4';
}

/** Replace header/footer tokens (&P → page#, &N → total, &D → date) */
function replaceHFTokens(text: string, pageNum: number, totalPages: number): string {
  // Strip Excel section codes (&L, &C, &R, &"font,style", &size)
  let clean = text
    .replace(/&[LCR]/g, '')
    .replace(/&"[^"]*"/g, '')
    .replace(/&\d+/g, '');
  clean = clean
    .replace(/&P/gi, String(pageNum))
    .replace(/&N/gi, String(totalPages))
    .replace(/&D/gi, new Date().toLocaleDateString())
    .replace(/&T/gi, new Date().toLocaleTimeString())
    .replace(/&F/gi, '')
    .replace(/&A/gi, '');
  return clean.trim();
}

/** Resolve image position on page */
function resolveImagePos(
  img: Image,
  cols: number[], rows: number[],
  colWidthsPt: number[], rowHeightsPt: number[],
  scale: number, mar: { left: number; top: number },
  pageH: number, startCol: number, startRow: number,
): { x: number; y: number; w: number; h: number } | null {
  if (img.position) {
    // Absolute position (px → pt conversion)
    const x = mar.left + img.position.x * 0.75 * scale;
    const y = pageH - mar.top - img.position.y * 0.75 * scale;
    const w = (img.width ?? 100) * 0.75 * scale;
    const h = (img.height ?? 100) * 0.75 * scale;
    return { x, y: y - h, w, h };
  }
  if (img.from) {
    // Cell-anchored position
    let x = mar.left;
    for (let ci = 0; ci < cols.length; ci++) {
      if (cols[ci] >= img.from.col) break;
      x += colWidthsPt[ci] * scale;
    }
    let y = 0;
    for (let ri = 0; ri < rows.length; ri++) {
      if (rows[ri] >= img.from.row) break;
      y += rowHeightsPt[ri] * scale;
    }
    const w = (img.width ?? 100) * 0.75 * scale;
    const h = (img.height ?? 100) * 0.75 * scale;
    const pdfY = pageH - mar.top - y;
    return { x, y: pdfY - h, w, h };
  }
  return null;
}

/** Decode base64 to Uint8Array */
function base64Decode(b64: string): Uint8Array {
  return base64ToBytes(b64);
}

/** Parse PNG file and create PDF image XObject, returns ref info */
function parsePngForPdf(data: Uint8Array, doc: PdfDoc): PdfImage | null {
  // Validate PNG signature
  if (data[0] !== 0x89 || data[1] !== 0x50 || data[2] !== 0x4E || data[3] !== 0x47) return null;

  // Read IHDR
  let pos = 8;
  const readU32 = (p: number) => (data[p] << 24 | data[p+1] << 16 | data[p+2] << 8 | data[p+3]) >>> 0;

  const ihdrLen = readU32(pos); pos += 4;
  const ihdrType = String.fromCharCode(data[pos], data[pos+1], data[pos+2], data[pos+3]); pos += 4;
  if (ihdrType !== 'IHDR' || ihdrLen !== 13) return null;

  const width = readU32(pos); pos += 4;
  const height = readU32(pos); pos += 4;
  const bitDepth = data[pos++];
  const colorType = data[pos++];
  pos += 2 + 1 + 4; // compression, filter, interlace, CRC

  // Only support 8-bit RGB or RGBA non-interlaced
  if (bitDepth !== 8) return null;
  if (colorType !== 2 && colorType !== 6) return null;

  // Collect IDAT chunks
  const idatChunks: Uint8Array[] = [];
  while (pos < data.length - 4) {
    const chunkLen = readU32(pos); pos += 4;
    const chunkType = String.fromCharCode(data[pos], data[pos+1], data[pos+2], data[pos+3]); pos += 4;
    if (chunkType === 'IDAT') {
      idatChunks.push(data.subarray(pos, pos + chunkLen));
    }
    if (chunkType === 'IEND') break;
    pos += chunkLen + 4; // data + CRC
  }

  if (!idatChunks.length) return null;

  // Concatenate IDAT data (this is zlib-wrapped deflate)
  let totalLen = 0;
  for (const c of idatChunks) totalLen += c.length;
  const zlibData = new Uint8Array(totalLen);
  let off = 0;
  for (const c of idatChunks) { zlibData.set(c, off); off += c.length; }

  const colors = colorType === 6 ? 4 : 3;

  if (colorType === 2) {
    // RGB: embed directly with FlateDecode + PNG predictor
    const objId = doc.add(
      `<</Type/XObject/Subtype/Image/Width ${width}/Height ${height}` +
      `/ColorSpace/DeviceRGB/BitsPerComponent 8/Filter/FlateDecode` +
      `/DecodeParms<</Predictor 15/Colors 3/BitsPerComponent 8/Columns ${width}>>` +
      `/Length ${zlibData.length}>>`,
      zlibData
    );
    return { id: objId, name: '', width, height };
  }

  if (colorType === 6) {
    // RGBA: We need to separate RGB and alpha channels.
    // For simplicity, use the zlib data with all 4 channels and a SMask.
    // Actually, we can't split them without decompressing. Let's decompress.
    // For now, embed as RGB ignoring alpha (simple approach)
    // A full implementation would decompress, split channels, and re-compress.
    // We'll skip RGBA PNGs for now.
    return null;
  }

  return null;
}

/* ═══════════════════════════════════════════════════════════════════════════ */
/*  Sparkline PDF rendering                                                   */
/* ═══════════════════════════════════════════════════════════════════════════ */

function resolveSparklineValues(ws: Worksheet, dataRange: string): number[] {
  const vals: number[] = [];
  const ref = dataRange.includes('!') ? dataRange.split('!')[1] : dataRange;
  const m = ref.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
  if (m) {
    const c1 = colLetterToIndex(m[1]), r1 = parseInt(m[2], 10);
    const c2 = colLetterToIndex(m[3]), r2 = parseInt(m[4], 10);
    for (let r = r1; r <= r2; r++) {
      for (let c = c1; c <= c2; c++) {
        const cell = ws.getCell(r, c);
        if (typeof cell.value === 'number') vals.push(cell.value);
      }
    }
  }
  return vals;
}

function drawSparkline(sb: StreamBuilder, sp: Sparkline, values: number[],
  x: number, y: number, w: number, h: number): void {
  if (!values.length) return;
  const pad = 2;
  const sx = x + pad, sy = y + pad, sw = w - pad * 2, sh = h - pad * 2;
  const min = Math.min(...values), max = Math.max(...values);
  const range = max - min || 1;
  const rgb = colorRgb(sp.color as string | undefined) ?? [0.267, 0.447, 0.769];

  sb.gsave();
  sb.clipRect(x, y, w, h);

  if (sp.type === 'bar' || sp.type === 'stacked') {
    const bw = sw / values.length;
    for (let i = 0; i < values.length; i++) {
      const v = values[i];
      const barH = Math.max(1, ((v - min) / range) * sh);
      const bx = sx + i * bw + bw * 0.1;
      const by = sy + (sh - barH);
      let fill = rgb;
      if (v < 0 && sp.negativeColor) fill = colorRgb(sp.negativeColor as string) ?? rgb;
      sb.fillColor(fill[0], fill[1], fill[2]);
      sb.fillRect(bx, by, bw * 0.8, barH);
    }
  } else {
    sb.strokeColor(rgb[0], rgb[1], rgb[2]);
    sb.lineWidth(sp.lineWidth ?? 1);
    const pts: [number, number][] = values.map((v, i) => [
      sx + (i / (values.length - 1 || 1)) * sw,
      sy + ((v - min) / range) * sh,
    ]);
    if (pts.length >= 2) {
      let path = `${n(pts[0][0])} ${n(pts[0][1])} m`;
      for (let i = 1; i < pts.length; i++) path += ` ${n(pts[i][0])} ${n(pts[i][1])} l`;
      sb.raw(path + ' S');
    }
    if (sp.showMarkers && sp.markersColor) {
      const mc = colorRgb(sp.markersColor as string) ?? rgb;
      sb.fillColor(mc[0], mc[1], mc[2]);
      for (const [px, py] of pts) sb.fillRect(px - 1.5, py - 1.5, 3, 3);
    }
  }

  sb.grestore();
}

/* ═══════════════════════════════════════════════════════════════════════════ */
/*  Form Control PDF rendering                                                */
/* ═══════════════════════════════════════════════════════════════════════════ */

function drawFormControl(sb: StreamBuilder, fc: FormControl,
  cols: number[], rows: number[], colWidthsPt: number[], rowHeightsPt: number[],
  scale: number, mar: { left: number; top: number }, pageH: number): void {
  let fx = mar.left, fy = 0;
  for (let ci = 0; ci < cols.length; ci++) {
    if (cols[ci] >= fc.from.col + 1) break;
    fx += colWidthsPt[ci] * scale;
  }
  for (let ri = 0; ri < rows.length; ri++) {
    if (rows[ri] >= fc.from.row + 1) break;
    fy += rowHeightsPt[ri] * scale;
  }
  const fw = (fc.width ?? 80) * 0.75 * scale;
  const fh = (fc.height ?? 24) * 0.75 * scale;
  const pyTop = pageH - mar.top - fy;
  const pyBot = pyTop - fh;
  const text = fc.text ?? '';

  sb.gsave();

  switch (fc.type) {
    case 'button':
    case 'dialog':
      sb.fillColor(0.93, 0.93, 0.93);
      sb.fillRect(fx, pyBot, fw, fh);
      sb.strokeColor(0.6, 0.6, 0.6).lineWidth(0.5);
      sb.strokeRect(fx, pyBot, fw, fh);
      if (text) {
        sb.fillColor(0, 0, 0);
        const fs = Math.min(9 * scale, fh * 0.6);
        const tw = textWidthPt(text, fs / scale, false) * scale;
        sb.beginText().font('F1', fs)
          .textPos(fx + (fw - tw) / 2, pyBot + (fh - fs) / 2)
          .showText(text).endText();
      }
      break;

    case 'checkBox': {
      sb.strokeColor(0.4, 0.4, 0.4).lineWidth(0.5);
      const boxS = Math.min(fh - 2, 10 * scale);
      sb.strokeRect(fx + 2, pyBot + (fh - boxS) / 2, boxS, boxS);
      if (fc.checked === 'checked') {
        sb.strokeColor(0, 0, 0).lineWidth(1);
        const bx = fx + 2, by = pyBot + (fh - boxS) / 2;
        sb.line(bx + 2, by + boxS / 2, bx + boxS / 2, by + 2);
        sb.line(bx + boxS / 2, by + 2, bx + boxS - 2, by + boxS - 2);
      }
      if (text) {
        sb.fillColor(0, 0, 0);
        const fs = Math.min(8 * scale, fh * 0.7);
        sb.beginText().font('F1', fs)
          .textPos(fx + boxS + 6, pyBot + (fh - fs) / 2)
          .showText(text).endText();
      }
      break;
    }

    case 'optionButton': {
      const rad = Math.min(fh / 2 - 1, 5 * scale);
      const cx = fx + 2 + rad, cy = pyBot + fh / 2;
      const k = 0.5523;
      sb.strokeColor(0.4, 0.4, 0.4).lineWidth(0.5);
      sb.raw(`${n(cx + rad)} ${n(cy)} m ${n(cx + rad)} ${n(cy + rad * k)} ${n(cx + rad * k)} ${n(cy + rad)} ${n(cx)} ${n(cy + rad)} c`);
      sb.raw(`${n(cx - rad * k)} ${n(cy + rad)} ${n(cx - rad)} ${n(cy + rad * k)} ${n(cx - rad)} ${n(cy)} c`);
      sb.raw(`${n(cx - rad)} ${n(cy - rad * k)} ${n(cx - rad * k)} ${n(cy - rad)} ${n(cx)} ${n(cy - rad)} c`);
      sb.raw(`${n(cx + rad * k)} ${n(cy - rad)} ${n(cx + rad)} ${n(cy - rad * k)} ${n(cx + rad)} ${n(cy)} c S`);
      if (fc.checked === 'checked') {
        const ir = rad * 0.5;
        sb.fillColor(0, 0, 0);
        sb.raw(`${n(cx + ir)} ${n(cy)} m ${n(cx + ir)} ${n(cy + ir * k)} ${n(cx + ir * k)} ${n(cy + ir)} ${n(cx)} ${n(cy + ir)} c`);
        sb.raw(`${n(cx - ir * k)} ${n(cy + ir)} ${n(cx - ir)} ${n(cy + ir * k)} ${n(cx - ir)} ${n(cy)} c`);
        sb.raw(`${n(cx - ir)} ${n(cy - ir * k)} ${n(cx - ir * k)} ${n(cy - ir)} ${n(cx)} ${n(cy - ir)} c`);
        sb.raw(`${n(cx + ir * k)} ${n(cy - ir)} ${n(cx + ir)} ${n(cy - ir * k)} ${n(cx + ir)} ${n(cy)} c f`);
      }
      if (text) {
        sb.fillColor(0, 0, 0);
        const fs = Math.min(8 * scale, fh * 0.7);
        sb.beginText().font('F1', fs)
          .textPos(fx + rad * 2 + 6, pyBot + (fh - fs) / 2)
          .showText(text).endText();
      }
      break;
    }

    case 'label':
    case 'groupBox':
      if (text) {
        sb.fillColor(0, 0, 0);
        const fs = Math.min(9 * scale, fh * 0.7);
        sb.beginText().font('F1', fs)
          .textPos(fx + 2, pyBot + (fh - fs) / 2)
          .showText(text).endText();
      }
      if (fc.type === 'groupBox') {
        sb.strokeColor(0.7, 0.7, 0.7).lineWidth(0.5);
        sb.strokeRect(fx, pyBot, fw, fh);
      }
      break;

    case 'comboBox':
    case 'listBox':
      sb.fillColor(1, 1, 1);
      sb.fillRect(fx, pyBot, fw, fh);
      sb.strokeColor(0.7, 0.7, 0.7).lineWidth(0.5);
      sb.strokeRect(fx, pyBot, fw, fh);
      if (fc.type === 'comboBox') {
        const aw = Math.min(16 * scale, fw * 0.2);
        sb.strokeRect(fx + fw - aw, pyBot, aw, fh);
        const ax = fx + fw - aw / 2, ay = pyBot + fh / 2;
        sb.fillColor(0.3, 0.3, 0.3);
        sb.raw(`${n(ax - 3)} ${n(ay + 2)} m ${n(ax + 3)} ${n(ay + 2)} l ${n(ax)} ${n(ay - 2)} l f`);
      }
      break;

    case 'scrollBar':
    case 'spinner':
      sb.fillColor(0.92, 0.92, 0.92);
      sb.fillRect(fx, pyBot, fw, fh);
      sb.strokeColor(0.7, 0.7, 0.7).lineWidth(0.5);
      sb.strokeRect(fx, pyBot, fw, fh);
      break;

    default:
      sb.strokeColor(0.7, 0.7, 0.7).lineWidth(0.5);
      sb.strokeRect(fx, pyBot, fw, fh);
      if (text) {
        sb.fillColor(0, 0, 0);
        const fs = Math.min(8 * scale, fh * 0.7);
        sb.beginText().font('F1', fs)
          .textPos(fx + 2, pyBot + (fh - fs) / 2)
          .showText(text).endText();
      }
  }
  sb.grestore();
}

/* ═══════════════════════════════════════════════════════════════════════════ */
/*  Chart PDF rendering (simplified bar/line/pie)                             */
/* ═══════════════════════════════════════════════════════════════════════════ */

const PDF_CHART_PALETTE: [number, number, number][] = [
  [0.267, 0.447, 0.769], [0.929, 0.490, 0.192], [0.647, 0.647, 0.647],
  [1, 0.753, 0], [0.357, 0.608, 0.835], [0.439, 0.678, 0.278],
];

function drawChartOnPage(sb: StreamBuilder, chart: Chart, ws: Worksheet,
  x: number, y: number, w: number, h: number): void {
  const PAD = 10;
  const plotX = x + PAD + 30, plotY = y + PAD;
  const plotW = w - PAD * 2 - 40, plotH = h - PAD * 2 - 20;

  const allSeries: { name: string; values: number[]; color: [number, number, number] }[] = [];
  for (let si = 0; si < chart.series.length; si++) {
    const s = chart.series[si];
    const vals = resolveChartDataValues(ws, s.values);
    allSeries.push({
      name: s.name ?? `Series ${si + 1}`,
      values: vals,
      color: PDF_CHART_PALETTE[si % PDF_CHART_PALETTE.length],
    });
  }

  if (!allSeries.length || !allSeries[0].values.length) return;

  sb.gsave();
  if (chart.title) {
    sb.fillColor(0, 0, 0);
    const titleFs = 10;
    const tw = textWidthPt(chart.title, titleFs, true);
    sb.beginText().font('F2', titleFs)
      .textPos(x + (w - tw) / 2, y + h - 8)
      .showText(chart.title).endText();
  }

  const type = chart.type;
  const allVals = allSeries.flatMap(s => s.values);
  const dataMin = Math.min(0, ...allVals);
  const dataMax = Math.max(1, ...allVals);
  const dataRange = dataMax - dataMin || 1;

  sb.strokeColor(0.5, 0.5, 0.5).lineWidth(0.5);
  sb.line(plotX, plotY, plotX, plotY + plotH);
  sb.line(plotX, plotY, plotX + plotW, plotY);

  if (type === 'pie' || type === 'doughnut') {
    const cx = x + w / 2, cy = y + h / 2;
    const radius = Math.min(plotW, plotH) / 2 - 5;
    const vals = allSeries[0].values.filter(v => v > 0);
    const total = vals.reduce((s, v) => s + v, 0) || 1;
    let startAngle = 0;
    for (let i = 0; i < vals.length; i++) {
      const sweep = (vals[i] / total) * 2 * Math.PI;
      const c = PDF_CHART_PALETTE[i % PDF_CHART_PALETTE.length];
      sb.fillColor(c[0], c[1], c[2]);
      sb.raw(`${n(cx)} ${n(cy)} m`);
      const steps = Math.max(8, Math.ceil(sweep / 0.2));
      for (let s = 0; s <= steps; s++) {
        const a = startAngle + (sweep * s) / steps;
        sb.raw(`${n(cx + radius * Math.cos(a))} ${n(cy + radius * Math.sin(a))} l`);
      }
      sb.raw('f');
      startAngle += sweep;
    }
  } else if (type.startsWith('bar')) {
    const numCats = allSeries[0].values.length;
    const catH = plotH / numCats;
    for (let i = 0; i < numCats; i++) {
      for (let si = 0; si < allSeries.length; si++) {
        const v = allSeries[si].values[i] ?? 0;
        const barW = ((v - dataMin) / dataRange) * plotW;
        const by = plotY + i * catH + (catH * 0.1);
        const bh = (catH * 0.8) / allSeries.length;
        const c = allSeries[si].color;
        sb.fillColor(c[0], c[1], c[2]);
        sb.fillRect(plotX, by + si * bh, Math.max(0, barW), bh);
      }
    }
  } else if (type.startsWith('column') || type === 'stock') {
    const numCats = allSeries[0].values.length;
    const catW = plotW / numCats;
    for (let i = 0; i < numCats; i++) {
      for (let si = 0; si < allSeries.length; si++) {
        const v = allSeries[si].values[i] ?? 0;
        const barH = ((v - dataMin) / dataRange) * plotH;
        const bx = plotX + i * catW + (catW * 0.1);
        const bw = (catW * 0.8) / allSeries.length;
        const c = allSeries[si].color;
        sb.fillColor(c[0], c[1], c[2]);
        sb.fillRect(bx + si * bw, plotY, bw, Math.max(0, barH));
      }
    }
  } else {
    for (const series of allSeries) {
      const pts: [number, number][] = series.values.map((v, i) => [
        plotX + (i / (series.values.length - 1 || 1)) * plotW,
        plotY + ((v - dataMin) / dataRange) * plotH,
      ]);
      sb.strokeColor(series.color[0], series.color[1], series.color[2]).lineWidth(1.5);
      if (pts.length >= 2) {
        let path = `${n(pts[0][0])} ${n(pts[0][1])} m`;
        for (let i = 1; i < pts.length; i++) path += ` ${n(pts[i][0])} ${n(pts[i][1])} l`;
        if (type.startsWith('area')) {
          path += ` ${n(pts[pts.length - 1][0])} ${n(plotY)} l ${n(pts[0][0])} ${n(plotY)} l`;
          sb.fillColor(series.color[0], series.color[1], series.color[2]);
          sb.raw(path + ' f');
        } else {
          sb.raw(path + ' S');
        }
      }
      if (type.startsWith('scatter') || type === 'bubble') {
        sb.fillColor(series.color[0], series.color[1], series.color[2]);
        for (const [px, py] of pts) sb.fillRect(px - 2, py - 2, 4, 4);
      }
    }
  }

  sb.grestore();
}

function resolveChartDataValues(ws: Worksheet, ref: string): number[] {
  const vals: number[] = [];
  const part = ref.includes('!') ? ref.split('!')[1] : ref;
  const clean = part.replace(/\$/g, '');
  const m = clean.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
  if (!m) return vals;
  const c1 = colLetterToIndex(m[1]), r1 = parseInt(m[2], 10);
  const c2 = colLetterToIndex(m[3]), r2 = parseInt(m[4], 10);
  for (let r = r1; r <= r2; r++) {
    for (let c = c1; c <= c2; c++) {
      const cell = ws.getCell(r, c);
      vals.push(typeof cell.value === 'number' ? cell.value : 0);
    }
  }
  return vals;
}

/** Generate a minimal empty PDF */
function emptyPdf(options: PdfExportOptions): Uint8Array {
  const doc = new PdfDoc();
  const catalogId = doc.alloc();
  const pagesId = doc.alloc();
  const paperKey = options.paperSize ?? 'a4';
  const orientation = options.orientation ?? 'portrait';
  const [pw, ph] = PAPER_SIZES[paperKey] ?? PAPER_SIZES.a4;
  const pageW = orientation === 'landscape' ? ph : pw;
  const pageH = orientation === 'landscape' ? pw : ph;
  const fontId = doc.add('<</Type/Font/Subtype/Type1/BaseFont/Helvetica>>');
  const sb = new StreamBuilder();
  sb.beginText().font('F1', 12).fillColor(0.5, 0.5, 0.5);
  sb.textPos(pageW / 2 - 40, pageH / 2);
  sb.showText('Empty worksheet').endText();
  const streamData = sb.toBytes();
  const contentId = doc.addDeflated('<<', streamData);
  const pageId = doc.add(
    `<</Type/Page/Parent ${pagesId} 0 R/MediaBox[0 0 ${n(pageW)} ${n(pageH)}]` +
    `/Contents ${contentId} 0 R/Resources<</Font<</F1 ${fontId} 0 R>>>>>>`
  );
  doc.set(pagesId, `<</Type/Pages/Kids[${pageId} 0 R]/Count 1>>`);
  doc.set(catalogId, `<</Type/Catalog/Pages ${pagesId} 0 R>>`);
  return doc.build(catalogId);
}
