/**
 * ExcelForge — Formula Calculation Engine (tree-shakeable).
 *
 * Supports a subset of Excel functions for offline calculation.
 * Import only when needed — won't bloat bundles that don't need calculation.
 */

import type { Workbook } from '../core/Workbook.js';
import type { Worksheet } from '../core/Worksheet.js';
import type { CellValue } from '../core/types.js';
import {
  colLetterToIndex, colIndexToLetter, cellRefToIndices, parseRange,
} from '../utils/helpers.js';

type Value = number | string | boolean | null;

/** Extract numeric value from CellValue */
function toNum(v: CellValue): number {
  if (typeof v === 'number') return v;
  if (typeof v === 'boolean') return v ? 1 : 0;
  if (typeof v === 'string') { const n = Number(v); return isNaN(n) ? 0 : n; }
  return 0;
}

function isNum(v: CellValue): boolean {
  return typeof v === 'number' || (typeof v === 'string' && v !== '' && !isNaN(Number(v)));
}

/**
 * A simple formula engine that can evaluate a useful subset of Excel formulas.
 * Tree-shakeable: only imported when calculation is needed.
 */
export class FormulaEngine {
  private ws!: Worksheet;
  private calculating = new Set<string>(); // circular ref guard

  /** Calculate all formulas in all sheets of a workbook */
  calculateWorkbook(wb: Workbook): void {
    const sheets = wb.getSheets();
    for (const ws of sheets) {
      this.calculateSheet(ws);
    }
  }

  /** Calculate all formulas in a single worksheet */
  calculateSheet(ws: Worksheet): void {
    this.ws = ws;
    const cells = ws.readAllCells();
    for (const { row, col, cell } of cells) {
      if (cell.formula) {
        this.calculating.clear();
        const result = this.evaluate(cell.formula, row, col);
        cell.value = result as CellValue;
      }
    }
  }

  /** Get a cell value, calculating its formula if needed */
  private getCellValue(row: number, col: number): CellValue {
    const cell = this.ws.getCell(row, col);
    if (cell.formula) {
      const key = `${row},${col}`;
      if (this.calculating.has(key)) return 0; // circular ref
      this.calculating.add(key);
      const result = this.evaluate(cell.formula, row, col);
      cell.value = result as CellValue;
      this.calculating.delete(key);
    }
    return cell.value ?? null;
  }

  /** Resolve a range reference to an array of values */
  private resolveRange(ref: string): CellValue[] {
    // Handle sheet!ref
    const rangePart = ref.includes('!') ? ref.split('!')[1] : ref;
    const clean = rangePart.replace(/\$/g, '');

    if (clean.includes(':')) {
      const { startRow, startCol, endRow, endCol } = parseRange(clean);
      const values: CellValue[] = [];
      for (let r = startRow; r <= endRow; r++) {
        for (let c = startCol; c <= endCol; c++) {
          values.push(this.getCellValue(r, c));
        }
      }
      return values;
    }
    // Single cell
    const { row, col } = cellRefToIndices(clean);
    return [this.getCellValue(row, col)];
  }

  /** Evaluate a formula string and return the result */
  private evaluate(formula: string, row: number, col: number): Value {
    try {
      return this.parseExpression(formula.trim(), 0, row, col).value;
    } catch {
      return '#VALUE!';
    }
  }

  private parseExpression(expr: string, pos: number, row: number, col: number): { value: Value; pos: number } {
    return this.parseComparison(expr, pos, row, col);
  }

  private parseComparison(expr: string, pos: number, row: number, col: number): { value: Value; pos: number } {
    let result = this.parseAddSub(expr, pos, row, col);
    let { value, pos: p } = result;
    p = this.skipSpaces(expr, p);
    // Check for comparison operators
    const ops = ['>=', '<=', '<>', '>', '<', '='];
    for (const op of ops) {
      if (expr.slice(p, p + op.length) === op) {
        const r = this.parseAddSub(expr, p + op.length, row, col);
        const lv = value, rv = r.value;
        switch (op) {
          case '>=': value = toNum(lv as CellValue) >= toNum(rv as CellValue); break;
          case '<=': value = toNum(lv as CellValue) <= toNum(rv as CellValue); break;
          case '<>': value = lv !== rv; break;
          case '>':  value = toNum(lv as CellValue) > toNum(rv as CellValue); break;
          case '<':  value = toNum(lv as CellValue) < toNum(rv as CellValue); break;
          case '=':  value = lv === rv || toNum(lv as CellValue) === toNum(rv as CellValue); break;
        }
        p = r.pos;
        break;
      }
    }
    return { value, pos: p };
  }

  private parseAddSub(expr: string, pos: number, row: number, col: number): { value: Value; pos: number } {
    let result = this.parseMulDiv(expr, pos, row, col);
    let { value, pos: p } = result;
    while (p < expr.length) {
      p = this.skipSpaces(expr, p);
      if (expr[p] === '+') {
        const r = this.parseMulDiv(expr, p + 1, row, col);
        value = toNum(value as CellValue) + toNum(r.value as CellValue);
        p = r.pos;
      } else if (expr[p] === '-') {
        const r = this.parseMulDiv(expr, p + 1, row, col);
        value = toNum(value as CellValue) - toNum(r.value as CellValue);
        p = r.pos;
      } else if (expr[p] === '&') {
        // String concatenation
        const r = this.parseMulDiv(expr, p + 1, row, col);
        value = String(value ?? '') + String(r.value ?? '');
        p = r.pos;
      } else {
        break;
      }
    }
    return { value, pos: p };
  }

  private parseMulDiv(expr: string, pos: number, row: number, col: number): { value: Value; pos: number } {
    let result = this.parsePower(expr, pos, row, col);
    let { value, pos: p } = result;
    while (p < expr.length) {
      p = this.skipSpaces(expr, p);
      if (expr[p] === '*') {
        const r = this.parsePower(expr, p + 1, row, col);
        value = toNum(value as CellValue) * toNum(r.value as CellValue);
        p = r.pos;
      } else if (expr[p] === '/') {
        const r = this.parsePower(expr, p + 1, row, col);
        const d = toNum(r.value as CellValue);
        value = d === 0 ? '#DIV/0!' : toNum(value as CellValue) / d;
        p = r.pos;
      } else {
        break;
      }
    }
    return { value, pos: p };
  }

  private parsePower(expr: string, pos: number, row: number, col: number): { value: Value; pos: number } {
    let result = this.parseUnary(expr, pos, row, col);
    let { value, pos: p } = result;
    p = this.skipSpaces(expr, p);
    if (p < expr.length && expr[p] === '^') {
      const r = this.parseUnary(expr, p + 1, row, col);
      value = Math.pow(toNum(value as CellValue), toNum(r.value as CellValue));
      p = r.pos;
    }
    return { value, pos: p };
  }

  private parseUnary(expr: string, pos: number, row: number, col: number): { value: Value; pos: number } {
    pos = this.skipSpaces(expr, pos);
    if (expr[pos] === '-') {
      const r = this.parseAtom(expr, pos + 1, row, col);
      return { value: -toNum(r.value as CellValue), pos: r.pos };
    }
    if (expr[pos] === '+') {
      return this.parseAtom(expr, pos + 1, row, col);
    }
    return this.parseAtom(expr, pos, row, col);
  }

  private parseAtom(expr: string, pos: number, row: number, col: number): { value: Value; pos: number } {
    pos = this.skipSpaces(expr, pos);

    // Parenthesized expression
    if (expr[pos] === '(') {
      const r = this.parseExpression(expr, pos + 1, row, col);
      let p = this.skipSpaces(expr, r.pos);
      if (expr[p] === ')') p++;
      return { value: r.value, pos: p };
    }

    // String literal
    if (expr[pos] === '"') {
      let end = pos + 1;
      let s = '';
      while (end < expr.length) {
        if (expr[end] === '"') {
          if (expr[end + 1] === '"') { s += '"'; end += 2; continue; }
          break;
        }
        s += expr[end++];
      }
      return { value: s, pos: end + 1 };
    }

    // Number
    if (/[0-9.]/.test(expr[pos])) {
      let end = pos;
      while (end < expr.length && /[0-9.eE+-]/.test(expr[end]) && !(end > pos && (expr[end] === '+' || expr[end] === '-') && !/[eE]/.test(expr[end - 1]))) end++;
      return { value: parseFloat(expr.slice(pos, end)), pos: end };
    }

    // TRUE/FALSE
    if (expr.slice(pos, pos + 4).toUpperCase() === 'TRUE') {
      return { value: true, pos: pos + 4 };
    }
    if (expr.slice(pos, pos + 5).toUpperCase() === 'FALSE') {
      return { value: false, pos: pos + 5 };
    }

    // Function call or cell reference
    let end = pos;
    while (end < expr.length && /[A-Za-z0-9_$!:]/.test(expr[end])) end++;

    const token = expr.slice(pos, end);
    let p = this.skipSpaces(expr, end);

    // Function call
    if (expr[p] === '(') {
      const funcName = token.toUpperCase();
      const args = this.parseArgList(expr, p, row, col);
      return { value: this.callFunction(funcName, args.args, row, col), pos: args.pos };
    }

    // Cell reference or range
    if (token.includes(':') || /^[A-Z]+[0-9]+$/i.test(token.replace(/\$/g, '')) || token.includes('!')) {
      const values = this.resolveRange(token);
      return { value: (values.length === 1 ? values[0] : values[0]) as Value, pos: end };
    }

    return { value: token, pos: end };
  }

  private parseArgList(expr: string, pos: number, row: number, col: number): { args: Value[][]; pos: number } {
    pos++; // skip (
    const args: Value[][] = [];
    if (this.skipSpaces(expr, pos) < expr.length && expr[this.skipSpaces(expr, pos)] === ')') {
      return { args: [], pos: this.skipSpaces(expr, pos) + 1 };
    }

    while (pos < expr.length) {
      pos = this.skipSpaces(expr, pos);
      // Check if this arg is a range reference
      const rangeMatch = expr.slice(pos).match(/^([A-Z$]+[0-9$]+:[A-Z$]+[0-9$]+|[A-Za-z]+![A-Z$]+[0-9$]+:[A-Z$]+[0-9$]+)/i);
      if (rangeMatch) {
        const rangeRef = rangeMatch[1];
        args.push(this.resolveRange(rangeRef) as Value[]);
        pos += rangeRef.length;
      } else {
        const r = this.parseExpression(expr, pos, row, col);
        args.push([r.value]);
        pos = r.pos;
      }
      pos = this.skipSpaces(expr, pos);
      if (expr[pos] === ',') { pos++; continue; }
      if (expr[pos] === ')') { pos++; break; }
      break;
    }
    return { args, pos };
  }

  private skipSpaces(s: string, pos: number): number {
    while (pos < s.length && s[pos] === ' ') pos++;
    return pos;
  }

  /** Execute a built-in function */
  private callFunction(name: string, args: Value[][], row: number, col: number): Value {
    const flat = (a: Value[][]): Value[] => a.flat();
    const nums = (a: Value[][]): number[] => flat(a).filter(v => typeof v === 'number' || (typeof v === 'string' && v !== '' && !isNaN(Number(v)))).map(v => Number(v));

    switch (name) {
      // ── Math ──
      case 'SUM': return nums(args).reduce((a, b) => a + b, 0);
      case 'AVERAGE': { const n = nums(args); return n.length ? n.reduce((a, b) => a + b, 0) / n.length : 0; }
      case 'COUNT': return nums(args).length;
      case 'COUNTA': return flat(args).filter(v => v != null && v !== '').length;
      case 'COUNTBLANK': return flat(args).filter(v => v == null || v === '').length;
      case 'MAX': { const n = nums(args); return n.length ? Math.max(...n) : 0; }
      case 'MIN': { const n = nums(args); return n.length ? Math.min(...n) : 0; }
      case 'ABS': return Math.abs(toNum(flat(args)[0] as CellValue));
      case 'ROUND': { const v = toNum(flat(args)[0] as CellValue); const d = toNum(flat(args)[1] as CellValue ?? 0); const f = Math.pow(10, d); return Math.round(v * f) / f; }
      case 'ROUNDUP': { const v = toNum(flat(args)[0] as CellValue); const d = toNum(flat(args)[1] as CellValue ?? 0); const f = Math.pow(10, d); return Math.ceil(v * f) / f; }
      case 'ROUNDDOWN': { const v = toNum(flat(args)[0] as CellValue); const d = toNum(flat(args)[1] as CellValue ?? 0); const f = Math.pow(10, d); return Math.floor(v * f) / f; }
      case 'INT': return Math.floor(toNum(flat(args)[0] as CellValue));
      case 'MOD': return toNum(flat(args)[0] as CellValue) % toNum(flat(args)[1] as CellValue);
      case 'POWER': return Math.pow(toNum(flat(args)[0] as CellValue), toNum(flat(args)[1] as CellValue));
      case 'SQRT': return Math.sqrt(toNum(flat(args)[0] as CellValue));
      case 'PI': return Math.PI;
      case 'RAND': return Math.random();
      case 'RANDBETWEEN': { const lo = toNum(flat(args)[0] as CellValue); const hi = toNum(flat(args)[1] as CellValue); return Math.floor(Math.random() * (hi - lo + 1)) + lo; }
      case 'SUMPRODUCT': {
        if (args.length < 2) return 0;
        const len = args[0].length;
        let sum = 0;
        for (let i = 0; i < len; i++) {
          let prod = 1;
          for (const a of args) prod *= toNum((a[i] ?? 0) as CellValue);
          sum += prod;
        }
        return sum;
      }
      case 'SUMIF': {
        const range = args[0] ?? [];
        const criteria = String(flat(args.slice(1, 2))[0] ?? '');
        const sumRange = args[2] ?? range;
        let sum = 0;
        for (let i = 0; i < range.length; i++) {
          if (this.matchCriteria(range[i], criteria)) sum += toNum((sumRange[i] ?? 0) as CellValue);
        }
        return sum;
      }
      case 'COUNTIF': {
        const range = args[0] ?? [];
        const criteria = String(flat(args.slice(1, 2))[0] ?? '');
        return range.filter(v => this.matchCriteria(v, criteria)).length;
      }

      // ── Logical ──
      case 'IF': {
        const cond = flat(args.slice(0, 1))[0];
        const t = flat(args.slice(1, 2))[0] ?? true;
        const f = args.length > 2 ? flat(args.slice(2, 3))[0] ?? false : false;
        return cond ? t : f;
      }
      case 'AND': return flat(args).every(v => !!v);
      case 'OR': return flat(args).some(v => !!v);
      case 'NOT': return !flat(args)[0];
      case 'IFERROR': {
        const v = flat(args.slice(0, 1))[0];
        return typeof v === 'string' && v.startsWith('#') ? flat(args.slice(1, 2))[0] ?? '' : v;
      }
      case 'IFNA': {
        const v = flat(args.slice(0, 1))[0];
        return v === '#N/A' ? flat(args.slice(1, 2))[0] ?? '' : v;
      }

      // ── Text ──
      case 'CONCATENATE': case 'CONCAT': return flat(args).map(v => String(v ?? '')).join('');
      case 'TEXTJOIN': {
        const delim = String(flat(args.slice(0, 1))[0] ?? '');
        const ignoreEmpty = !!flat(args.slice(1, 2))[0];
        const vals = flat(args.slice(2));
        const filtered = ignoreEmpty ? vals.filter(v => v != null && v !== '') : vals;
        return filtered.map(v => String(v ?? '')).join(delim);
      }
      case 'LEN': return String(flat(args)[0] ?? '').length;
      case 'LEFT': return String(flat(args)[0] ?? '').slice(0, toNum(flat(args)[1] as CellValue) || 1);
      case 'RIGHT': { const s = String(flat(args)[0] ?? ''); const n = toNum(flat(args)[1] as CellValue) || 1; return s.slice(-n); }
      case 'MID': { const s = String(flat(args)[0] ?? ''); const start = toNum(flat(args)[1] as CellValue) - 1; const len = toNum(flat(args)[2] as CellValue); return s.slice(start, start + len); }
      case 'UPPER': return String(flat(args)[0] ?? '').toUpperCase();
      case 'LOWER': return String(flat(args)[0] ?? '').toLowerCase();
      case 'TRIM': return String(flat(args)[0] ?? '').trim().replace(/\s+/g, ' ');
      case 'SUBSTITUTE': {
        const s = String(flat(args)[0] ?? '');
        const old = String(flat(args)[1] ?? '');
        const rep = String(flat(args)[2] ?? '');
        return s.split(old).join(rep);
      }
      case 'FIND': case 'SEARCH': {
        const findText = String(flat(args)[0] ?? '');
        const within = String(flat(args)[1] ?? '');
        const start = toNum(flat(args)[2] as CellValue) || 1;
        const fn = name === 'SEARCH' ? (s: string) => s.toLowerCase() : (s: string) => s;
        const idx = fn(within).indexOf(fn(findText), start - 1);
        return idx >= 0 ? idx + 1 : '#VALUE!';
      }
      case 'REPLACE': {
        const s = String(flat(args)[0] ?? '');
        const start = toNum(flat(args)[1] as CellValue) - 1;
        const numChars = toNum(flat(args)[2] as CellValue);
        const rep = String(flat(args)[3] ?? '');
        return s.slice(0, start) + rep + s.slice(start + numChars);
      }
      case 'REPT': return String(flat(args)[0] ?? '').repeat(toNum(flat(args)[1] as CellValue));
      case 'TEXT': return String(flat(args)[0] ?? ''); // simplified
      case 'VALUE': return toNum(flat(args)[0] as CellValue);
      case 'EXACT': return String(flat(args)[0] ?? '') === String(flat(args)[1] ?? '');

      // ── Lookup ──
      case 'VLOOKUP': {
        const lookupVal = flat(args.slice(0, 1))[0];
        const table = args[1] ?? [];
        const colIdx = toNum(flat(args.slice(2, 3))[0] as CellValue) - 1;
        // We'd need 2D range for this - simplified version
        return '#N/A';
      }
      case 'INDEX': {
        const arr = args[0] ?? [];
        const rowIdx = toNum(flat(args.slice(1, 2))[0] as CellValue) - 1;
        return arr[rowIdx] ?? '#REF!';
      }
      case 'MATCH': {
        const val = flat(args.slice(0, 1))[0];
        const arr = args[1] ?? [];
        const idx = arr.findIndex(v => v === val || String(v) === String(val));
        return idx >= 0 ? idx + 1 : '#N/A';
      }

      // ── Info ──
      case 'ISBLANK': return flat(args)[0] == null || flat(args)[0] === '';
      case 'ISNUMBER': return typeof flat(args)[0] === 'number';
      case 'ISTEXT': return typeof flat(args)[0] === 'string';
      case 'ISLOGICAL': return typeof flat(args)[0] === 'boolean';
      case 'ISERROR': { const v = flat(args)[0]; return typeof v === 'string' && v.startsWith('#'); }
      case 'ISNA': return flat(args)[0] === '#N/A';
      case 'TYPE': {
        const v = flat(args)[0];
        if (typeof v === 'number') return 1;
        if (typeof v === 'string') return 2;
        if (typeof v === 'boolean') return 4;
        return 1;
      }
      case 'ROW': return row;
      case 'COLUMN': return col;
      case 'NA': return '#N/A';

      // ── Date ──
      case 'TODAY': return new Date().toISOString().slice(0, 10);
      case 'NOW': return new Date().toISOString();
      case 'YEAR': return new Date(String(flat(args)[0])).getFullYear();
      case 'MONTH': return new Date(String(flat(args)[0])).getMonth() + 1;
      case 'DAY': return new Date(String(flat(args)[0])).getDate();

      // ── Pivot ──
      case 'GETPIVOTDATA': {
        // GETPIVOTDATA(data_field, pivot_table, [field1, item1, ...])
        // Simplified: returns the data field name as we don't have a live pivot cache
        const dataField = String(flat(args.slice(0, 1))[0] ?? '');
        return dataField || '#REF!';
      }

      default: return '#NAME?';
    }
  }

  /** Simple criteria matching for SUMIF/COUNTIF (supports >, <, >=, <=, <>, =, plain value) */
  private matchCriteria(value: Value, criteria: string): boolean {
    const ops = ['>=', '<=', '<>', '>', '<', '='];
    for (const op of ops) {
      if (criteria.startsWith(op)) {
        const cv = criteria.slice(op.length);
        const nv = toNum(value as CellValue);
        const nc = Number(cv);
        if (!isNaN(nc)) {
          switch (op) {
            case '>=': return nv >= nc;
            case '<=': return nv <= nc;
            case '<>': return nv !== nc;
            case '>': return nv > nc;
            case '<': return nv < nc;
            case '=': return nv === nc;
          }
        }
        // String comparison
        switch (op) {
          case '<>': return String(value) !== cv;
          case '=': return String(value) === cv;
        }
      }
    }
    // Exact match
    if (typeof value === 'number') return value === Number(criteria);
    return String(value) === criteria;
  }
}
