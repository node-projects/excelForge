/**
 * CSV read/write module for ExcelForge.
 * Tree-shakeable — only imported when used.
 */

import type { CellValue } from '../core/types.js';
import { CellError } from '../core/types.js';
import { Workbook } from '../core/Workbook.js';
import { Worksheet } from '../core/Worksheet.js';

export interface CsvOptions {
  /** Field delimiter (default: ',') */
  delimiter?: string;
  /** Row separator (default: '\r\n') */
  lineEnding?: string;
  /** Quote character (default: '"') */
  quote?: string;
  /** Sheet name when reading CSV to Workbook (default: 'Sheet1') */
  sheetName?: string;
  /** Whether to include header row (default: true for write) */
  includeHeaders?: boolean;
}

/**
 * Write a worksheet to CSV string.
 */
export function worksheetToCsv(ws: Worksheet, options: CsvOptions = {}): string {
  const delim = options.delimiter ?? ',';
  const eol = options.lineEnding ?? '\r\n';
  const q = options.quote ?? '"';
  const rows: string[] = [];
  const allCells = ws.readAllCells();

  let maxRow = 0, maxCol = 0;
  for (const { row, col } of allCells) {
    if (row > maxRow) maxRow = row;
    if (col > maxCol) maxCol = col;
  }

  const cellMap = new Map<string, CellValue>();
  for (const { row, col, cell } of allCells) {
    cellMap.set(`${row},${col}`, cell.value ?? null);
  }

  for (let r = 1; r <= maxRow; r++) {
    const fields: string[] = [];
    for (let c = 1; c <= maxCol; c++) {
      const v = cellMap.get(`${r},${c}`);
      const s = formatCsvValue(v);
      // Quote if contains delimiter, quote char, or newlines
      if (s.includes(delim) || s.includes(q) || s.includes('\n') || s.includes('\r')) {
        fields.push(q + s.replace(new RegExp(q.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'g'), q + q) + q);
      } else {
        fields.push(s);
      }
    }
    rows.push(fields.join(delim));
  }

  return rows.join(eol);
}

function formatCsvValue(v: CellValue): string {
  if (v === null || v === undefined) return '';
  if (v instanceof CellError) return v.error;
  if (v instanceof Date) return v.toISOString();
  return String(v);
}

/**
 * Parse CSV string into a Workbook.
 */
export function csvToWorkbook(csv: string, options: CsvOptions = {}): Workbook {
  const delim = options.delimiter ?? ',';
  const q = options.quote ?? '"';
  const sheetName = options.sheetName ?? 'Sheet1';

  const rows = parseCsv(csv, delim, q);
  const wb = new Workbook();
  const ws = wb.addSheet(sheetName);

  for (let r = 0; r < rows.length; r++) {
    for (let c = 0; c < rows[r].length; c++) {
      const raw = rows[r][c];
      if (raw === '') continue;
      // Try number
      const num = Number(raw);
      if (!isNaN(num) && raw.trim() !== '') {
        ws.setValue(r + 1, c + 1, num);
      } else if (raw.toLowerCase() === 'true') {
        ws.setValue(r + 1, c + 1, true);
      } else if (raw.toLowerCase() === 'false') {
        ws.setValue(r + 1, c + 1, false);
      } else {
        ws.setValue(r + 1, c + 1, raw);
      }
    }
  }

  return wb;
}

/**
 * Parse CSV string into 2D string array following RFC 4180.
 */
function parseCsv(csv: string, delim: string, quote: string): string[][] {
  const rows: string[][] = [];
  let row: string[] = [];
  let field = '';
  let inQuotes = false;
  let i = 0;
  const len = csv.length;

  while (i < len) {
    const ch = csv[i];
    if (inQuotes) {
      if (ch === quote) {
        if (i + 1 < len && csv[i + 1] === quote) {
          field += quote;
          i += 2;
        } else {
          inQuotes = false;
          i++;
        }
      } else {
        field += ch;
        i++;
      }
    } else {
      if (ch === quote) {
        inQuotes = true;
        i++;
      } else if (ch === delim) {
        row.push(field);
        field = '';
        i++;
      } else if (ch === '\r') {
        row.push(field);
        field = '';
        rows.push(row);
        row = [];
        i++;
        if (i < len && csv[i] === '\n') i++;
      } else if (ch === '\n') {
        row.push(field);
        field = '';
        rows.push(row);
        row = [];
        i++;
      } else {
        field += ch;
        i++;
      }
    }
  }

  // Last field
  if (field || row.length) {
    row.push(field);
    rows.push(row);
  }

  return rows;
}
