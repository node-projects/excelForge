/**
 * JSON export module for ExcelForge.
 * Tree-shakeable — only imported when used.
 */

import type { CellValue } from '../core/types.js';
import { CellError } from '../core/types.js';
import { Worksheet } from '../core/Worksheet.js';
import { Workbook } from '../core/Workbook.js';

export interface JsonExportOptions {
  /** Use first row as property names (default: true) */
  header?: boolean;
  /** Include empty cells as null (default: false) */
  includeEmpty?: boolean;
  /** Custom date format: 'iso' (default), 'epoch', 'serial' */
  dateFormat?: 'iso' | 'epoch' | 'serial';
  /** Range to export (e.g. "A1:D10"), or whole sheet if omitted */
  range?: string;
}

/**
 * Export a worksheet as an array of JSON objects (header=true)
 * or array of arrays (header=false).
 */
export function worksheetToJson(ws: Worksheet, options: JsonExportOptions = {}): any[] {
  const header = options.header !== false;
  const includeEmpty = options.includeEmpty ?? false;

  const range = options.range
    ? ws.readRange(options.range)
    : readAllAsArray(ws);

  if (range.length === 0) return [];

  if (header) {
    const headers = range[0].map(v => String(v ?? ''));
    return range.slice(1).map(row => {
      const obj: Record<string, any> = {};
      for (let c = 0; c < headers.length; c++) {
        const val = c < row.length ? row[c] : null;
        if (!includeEmpty && (val === null || val === undefined)) continue;
        obj[headers[c]] = convertValue(val, options);
      }
      return obj;
    });
  } else {
    return range.map(row => row.map(v => convertValue(v, options)));
  }
}

/**
 * Export all sheets in a workbook as a JSON object keyed by sheet name.
 */
export function workbookToJson(wb: Workbook, options: JsonExportOptions = {}): Record<string, any[]> {
  const result: Record<string, any[]> = {};
  for (const ws of wb.getSheets()) {
    result[ws.name] = worksheetToJson(ws, options);
  }
  return result;
}

function convertValue(v: CellValue, options: JsonExportOptions): any {
  if (v === null || v === undefined) return null;
  if (v instanceof CellError) return v.error;
  if (v instanceof Date) {
    switch (options.dateFormat) {
      case 'epoch': return v.getTime();
      case 'serial': return (v.getTime() - Date.UTC(1899, 11, 30)) / 86400000;
      default: return v.toISOString();
    }
  }
  return v;
}

function readAllAsArray(ws: Worksheet): CellValue[][] {
  const cells = ws.readAllCells();
  if (cells.length === 0) return [];

  let maxRow = 0, maxCol = 0;
  for (const { row, col } of cells) {
    if (row > maxRow) maxRow = row;
    if (col > maxCol) maxCol = col;
  }

  const grid: CellValue[][] = [];
  const cellMap = new Map<number, Map<number, CellValue>>();
  for (const { row, col, cell } of cells) {
    if (!cellMap.has(row)) cellMap.set(row, new Map());
    cellMap.get(row)!.set(col, cell.value ?? null);
  }

  for (let r = 1; r <= maxRow; r++) {
    const row: CellValue[] = [];
    const rm = cellMap.get(r);
    for (let c = 1; c <= maxCol; c++) {
      row.push(rm?.get(c) ?? null);
    }
    grid.push(row);
  }
  return grid;
}
