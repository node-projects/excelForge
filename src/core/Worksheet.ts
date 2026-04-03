import {
  CellError,
} from '../core/types.js';
import type {
  Cell, CellValue, CellStyle, MergeRange, Image, CellImage, Chart,
  ConditionalFormat, Table, AutoFilter, FreezePane, SplitPane,
  SheetProtection, PageSetup, PageMargins, HeaderFooter, PrintOptions,
  SheetView, ColumnDef, RowDef, Sparkline, DataValidation,
  WorksheetOptions, PivotTable, PageBreak, FormControl,
  Shape, WordArt, QueryTable, TableSlicer, CFCustomIconSet,
} from '../core/types.js';
import type { SharedStrings } from '../core/SharedStrings.js';
import type { StyleRegistry } from '../styles/StyleRegistry.js';
import {
  colIndexToLetter, colLetterToIndex, indicesToCellRef,
  cellRefToIndices, parseRange, escapeXml, strToBytes, dateToSerial,
  pxToEmu, colWidthToEmu, rowHeightToEmu, base64ToBytes,
} from '../utils/helpers.js';

/** Two-level map: row → col → Cell.  Avoids V8's single-Map size limit (~16.7 M entries)
 *  and eliminates string key allocation/parsing overhead. */
type CellMap = Map<number, Map<number, Cell>>;

/** Emit a <color> element handling theme:X, #hex, and AARRGGBB formats */
function colorEl(c: string): string {
  if (c.startsWith('theme:')) return `<color theme="${c.slice(6)}"/>`;
  const rgb = c.startsWith('#') ? 'FF' + c.slice(1) : c;
  return `<color rgb="${rgb}"/>`;
}

// SUBTOTAL function numbers (101–110 ignore hidden rows, matching Excel table behaviour)
const SUBTOTAL_FN: Record<string, number> = {
  average: 101, count: 102, countNums: 103, max: 104, min: 105,
  stdDev: 107, sum: 109, var: 110, vars: 111,
};

export class Worksheet {
  readonly name: string;
  private cells: CellMap = new Map();
  private merges: MergeRange[] = [];
  private images: Image[] = [];
  private cellImages: CellImage[] = [];
  private charts: Chart[] = [];
  private conditionalFormats: ConditionalFormat[] = [];
  private tables: Table[] = [];
  private pivotTables: PivotTable[] = [];
  private sparklines: Sparkline[] = [];
  private formControls: FormControl[] = [];
  private shapes: Shape[] = [];
  private wordArt: WordArt[] = [];
  private queryTables: QueryTable[] = [];
  private tableSlicers: TableSlicer[] = [];
  private colDefs: Map<number, ColumnDef> = new Map();
  private rowDefs: Map<number, RowDef>    = new Map();
  private dataValidations: Map<string, DataValidation> = new Map();
  private rowBreaks: PageBreak[] = [];
  private colBreaks: PageBreak[] = [];
  /** Raw XML fragments for elements we don't parse */
  private preservedXml: string[] = [];
  /** Counter for shared formula indices */
  private _nextSharedIdx = 0;

  options: WorksheetOptions;
  view?: SheetView;
  freezePane?: FreezePane;
  splitPane?: SplitPane;
  protection?: SheetProtection;
  pageSetup?: PageSetup;
  pageMargins?: PageMargins;
  headerFooter?: HeaderFooter;
  printOptions?: PrintOptions;
  autoFilter?: AutoFilter;
  /** True when this sheet is a chart sheet (entire sheet is a single chart) */
  _isChartSheet = false;
  /** True when this sheet is a dialog sheet (Excel 5 dialog) */
  _isDialogSheet = false;

  // assigned by workbook
  sheetIndex = 0;
  rId = '';
  drawingRId = '';
  legacyDrawingRId = '';
  tableRIds: string[] = [];
  ctrlPropRIds: string[] = [];

  constructor(name: string, options: WorksheetOptions = {}) {
    this.name = name;
    this.options = { ...options, name };
  }

  // ─── Cell Access ─────────────────────────────────────────────────────────────

  getCell(row: number, col: number): Cell {
    let rowMap = this.cells.get(row);
    if (!rowMap) { rowMap = new Map(); this.cells.set(row, rowMap); }
    let cell = rowMap.get(col);
    if (!cell) { cell = {}; rowMap.set(col, cell); }
    return cell;
  }

  getCellByRef(ref: string): Cell {
    const { row, col } = cellRefToIndices(ref);
    return this.getCell(row, col);
  }

  setCell(row: number, col: number, cell: Cell): this {
    let rowMap = this.cells.get(row);
    if (!rowMap) { rowMap = new Map(); this.cells.set(row, rowMap); }
    rowMap.set(col, cell);
    return this;
  }

  setValue(row: number, col: number, value: CellValue): this {
    this.getCell(row, col).value = value;
    return this;
  }

  setFormula(row: number, col: number, formula: string): this {
    this.getCell(row, col).formula = formula;
    return this;
  }

  /** Set a dynamic array formula (spill formula) on a cell. The formula will spill results automatically. */
  setDynamicArrayFormula(row: number, col: number, formula: string): this {
    const cell = this.getCell(row, col);
    cell.arrayFormula = formula;
    (cell as any)._dynamic = true;
    return this;
  }

  /** Set a shared formula. The master cell defines the formula; dependent cells reference it by index. */
  setSharedFormula(masterRow: number, masterCol: number, formula: string, rangeRef: string): this {
    const si = this._nextSharedIdx++;
    const master = this.getCell(masterRow, masterCol);
    master.formula = formula;
    (master as any)._sharedRef = rangeRef;
    (master as any)._sharedIdx = si;
    // Set dependent cells
    const { startRow, startCol, endRow, endCol } = parseRange(rangeRef);
    for (let r = startRow; r <= endRow; r++) {
      for (let c = startCol; c <= endCol; c++) {
        if (r === masterRow && c === masterCol) continue;
        const dep = this.getCell(r, c);
        (dep as any)._sharedIdx = si;
      }
    }
    return this;
  }

  setStyle(row: number, col: number, style: CellStyle): this {
    this.getCell(row, col).style = style;
    return this;
  }

  /** Write a 2D array starting at (startRow, startCol) */
  writeArray(startRow: number, startCol: number, data: CellValue[][]): this {
    for (let r = 0; r < data.length; r++) {
      for (let c = 0; c < data[r].length; c++) {
        this.setValue(startRow + r, startCol + c, data[r][c]);
      }
    }
    return this;
  }

  /** Write a column of values */
  writeColumn(startRow: number, col: number, values: CellValue[]): this {
    values.forEach((v, i) => this.setValue(startRow + i, col, v));
    return this;
  }

  /** Write a row of values */
  writeRow(row: number, startCol: number, values: CellValue[]): this {
    values.forEach((v, i) => this.setValue(row, startCol + i, v));
    return this;
  }

  // ─── Columns & Rows ──────────────────────────────────────────────────────────

  setColumn(col: number, def: ColumnDef): this {
    this.colDefs.set(col, def);
    return this;
  }

  setColumnWidth(col: number, width: number): this {
    const d = this.colDefs.get(col) ?? {};
    this.colDefs.set(col, { ...d, width, customWidth: true });
    return this;
  }

  setRow(row: number, def: RowDef): this {
    this.rowDefs.set(row, def);
    return this;
  }

  setRowHeight(row: number, height: number): this {
    const d = this.rowDefs.get(row) ?? {};
    this.rowDefs.set(row, { ...d, height });
    return this;
  }

  // ─── Merges ──────────────────────────────────────────────────────────────────

  merge(startRow: number, startCol: number, endRow: number, endCol: number): this {
    this.merges.push({ startRow, startCol, endRow, endCol });
    return this;
  }

  mergeByRef(ref: string): this {
    const [s, e] = ref.split(':');
    const sc = cellRefToIndices(s); const ec = cellRefToIndices(e);
    return this.merge(sc.row, sc.col, ec.row, ec.col);
  }

  getMerges(): readonly MergeRange[] { return this.merges; }

  // ─── Images ──────────────────────────────────────────────────────────────────

  addImage(img: Image): this {
    this.images.push(img);
    return this;
  }

  getImages(): readonly Image[] { return this.images; }

  // ─── Cell Images (In-Cell Pictures) ───────────────────────────────────────

  addCellImage(img: CellImage): this {
    this.cellImages.push(img);
    return this;
  }

  getCellImages(): readonly CellImage[] { return this.cellImages; }

  /**
   * Map of cell ref → vm index (1-based) for cell images.
   * Set externally by Workbook during build to inject vm attributes into cell XML.
   */
  _cellImageVm: Map<string, number> = new Map();

  /** Return all cells that have a comment, keyed by "col,row" */
  getComments(): Array<{ row: number; col: number; comment: import('../core/types.js').Comment }> {
    const out: Array<{ row: number; col: number; comment: import('../core/types.js').Comment }> = [];
    for (const [r, rowMap] of this.cells) {
      for (const [c, cell] of rowMap) {
        if (cell.comment) {
          out.push({ row: r, col: c, comment: cell.comment });
        }
      }
    }
    return out;
  }

  // ─── Charts ──────────────────────────────────────────────────────────────────

  addChart(chart: Chart): this {
    this.charts.push(chart);
    return this;
  }

  getCharts(): readonly Chart[] { return this.charts; }

  // ─── Conditional Formatting ──────────────────────────────────────────────────

  addConditionalFormat(cf: ConditionalFormat): this {
    this.conditionalFormats.push(cf);
    return this;
  }

  getConditionalFormats(): readonly ConditionalFormat[] { return this.conditionalFormats; }

  getDataValidations(): ReadonlyMap<string, DataValidation> { return this.dataValidations; }

  // ─── Tables ──────────────────────────────────────────────────────────────────

  addTable(table: Table): this {
    this.tables.push(table);
    if (table.totalsRow && table.columns?.length) {
      const { startRow, startCol, endRow } = parseRange(table.ref);
      const dataStart = startRow + 1;   // first data row (after header)
      const dataEnd   = endRow - 1;     // last data row (before totals)
      table.columns.forEach((col, i) => {
        const colIdx = startCol + i;
        if (col.totalsRowLabel) {
          this.setValue(endRow, colIdx, col.totalsRowLabel);
        } else if (col.totalsRowFunction && col.totalsRowFunction !== 'none') {
          const fn = SUBTOTAL_FN[col.totalsRowFunction];
          if (fn !== undefined) {
            const letter = colIndexToLetter(colIdx);
            this.setFormula(endRow, colIdx, `SUBTOTAL(${fn},${letter}${dataStart}:${letter}${dataEnd})`);
          }
        }
      });
    }
    return this;
  }

  getTables(): readonly Table[] { return this.tables; }

  // ─── Pivot Tables ────────────────────────────────────────────────────────────

  addPivotTable(pt: PivotTable): this {
    this.pivotTables.push(pt);
    return this;
  }

  getPivotTables(): readonly PivotTable[] { return this.pivotTables; }

  /** Read all cell values from a range as a 2-D array (row-major). */
  readRange(ref: string): CellValue[][] {
    const { startRow, startCol, endRow, endCol } = parseRange(ref);
    const result: CellValue[][] = [];
    for (let r = startRow; r <= endRow; r++) {
      const row: CellValue[] = [];
      const rowMap = this.cells.get(r);
      for (let c = startCol; c <= endCol; c++) {
        const cell = rowMap?.get(c);
        row.push(cell?.value ?? null);
      }
      result.push(row);
    }
    return result;
  }

  /** Iterate all populated cells. */
  readAllCells(): Array<{ row: number; col: number; cell: Cell }> {
    const out: Array<{ row: number; col: number; cell: Cell }> = [];
    for (const [r, rowMap] of this.cells) {
      for (const [c, cell] of rowMap) {
        out.push({ row: r, col: c, cell });
      }
    }
    return out;
  }

  /** Get the used range dimensions. */
  getUsedRange(): { startRow: number; startCol: number; endRow: number; endCol: number } | null {
    let minR = Infinity, maxR = 0, minC = Infinity, maxC = 0;
    for (const [r, rowMap] of this.cells) {
      for (const [c] of rowMap) {
        if (r < minR) minR = r;
        if (r > maxR) maxR = r;
        if (c < minC) minC = c;
        if (c > maxC) maxC = c;
      }
    }
    return maxR === 0 ? null : { startRow: minR, startCol: minC, endRow: maxR, endCol: maxC };
  }

  /** Get column definition */
  getColumn(col: number): ColumnDef | undefined {
    return this.colDefs.get(col);
  }

  /** Get row definition */
  getRow(row: number): RowDef | undefined {
    return this.rowDefs.get(row);
  }

  // ─── Insert/Delete Rows & Columns ──────────────────────────────────────────

  /** Insert `count` empty rows at the given row index (1-based). Existing rows shift down. */
  insertRows(atRow: number, count: number): this {
    // Shift cell data
    const rows = [...this.cells.keys()].filter(r => r >= atRow).sort((a, b) => b - a);
    for (const r of rows) {
      const rowMap = this.cells.get(r)!;
      this.cells.delete(r);
      this.cells.set(r + count, rowMap);
    }
    // Shift row defs
    const rdKeys = [...this.rowDefs.keys()].filter(r => r >= atRow).sort((a, b) => b - a);
    for (const r of rdKeys) {
      const d = this.rowDefs.get(r)!;
      this.rowDefs.delete(r);
      this.rowDefs.set(r + count, d);
    }
    // Shift merges
    for (const m of this.merges) {
      if (m.startRow >= atRow) m.startRow += count;
      if (m.endRow >= atRow) m.endRow += count;
    }
    return this;
  }

  /** Delete `count` rows starting at the given row index (1-based). Rows below shift up. */
  deleteRows(atRow: number, count: number): this {
    for (let r = atRow; r < atRow + count; r++) {
      this.cells.delete(r);
      this.rowDefs.delete(r);
    }
    const rows = [...this.cells.keys()].filter(r => r >= atRow + count).sort((a, b) => a - b);
    for (const r of rows) {
      const rowMap = this.cells.get(r)!;
      this.cells.delete(r);
      this.cells.set(r - count, rowMap);
    }
    const rdKeys = [...this.rowDefs.keys()].filter(r => r >= atRow + count).sort((a, b) => a - b);
    for (const r of rdKeys) {
      const d = this.rowDefs.get(r)!;
      this.rowDefs.delete(r);
      this.rowDefs.set(r - count, d);
    }
    // Adjust merges
    this.merges = this.merges.filter(m => !(m.startRow >= atRow && m.endRow < atRow + count));
    for (const m of this.merges) {
      if (m.startRow >= atRow + count) m.startRow -= count;
      if (m.endRow >= atRow + count) m.endRow -= count;
    }
    return this;
  }

  /** Insert `count` empty columns at the given column index (1-based). Existing columns shift right. */
  insertColumns(atCol: number, count: number): this {
    for (const [, rowMap] of this.cells) {
      const cols = [...rowMap.keys()].filter(c => c >= atCol).sort((a, b) => b - a);
      for (const c of cols) {
        const cell = rowMap.get(c)!;
        rowMap.delete(c);
        rowMap.set(c + count, cell);
      }
    }
    const cdKeys = [...this.colDefs.keys()].filter(c => c >= atCol).sort((a, b) => b - a);
    for (const c of cdKeys) {
      const d = this.colDefs.get(c)!;
      this.colDefs.delete(c);
      this.colDefs.set(c + count, d);
    }
    for (const m of this.merges) {
      if (m.startCol >= atCol) m.startCol += count;
      if (m.endCol >= atCol) m.endCol += count;
    }
    return this;
  }

  // ─── Copy/Move Ranges ────────────────────────────────────────────────────────

  /** Copy cells from a source range to a target position. */
  copyRange(srcRef: string, targetRow: number, targetCol: number): this {
    const { startRow, startCol, endRow, endCol } = parseRange(srcRef);
    for (let r = startRow; r <= endRow; r++) {
      for (let c = startCol; c <= endCol; c++) {
        const src = this.getCell(r, c);
        const dr = targetRow + (r - startRow);
        const dc = targetCol + (c - startCol);
        if (src.value != null) this.setValue(dr, dc, src.value as any);
        if (src.formula) this.setFormula(dr, dc, src.formula);
        if (src.style) this.setStyle(dr, dc, { ...src.style });
      }
    }
    return this;
  }

  /** Move cells from a source range to a target position (clears source). */
  moveRange(srcRef: string, targetRow: number, targetCol: number): this {
    const { startRow, startCol, endRow, endCol } = parseRange(srcRef);
    // Copy first
    this.copyRange(srcRef, targetRow, targetCol);
    // Clear source (only cells not overlapping with target)
    const tEndRow = targetRow + (endRow - startRow);
    const tEndCol = targetCol + (endCol - startCol);
    for (let r = startRow; r <= endRow; r++) {
      for (let c = startCol; c <= endCol; c++) {
        const dr = targetRow + (r - startRow);
        const dc = targetCol + (c - startCol);
        if (dr === r && dc === c) continue; // same position
        const rowMap = this.cells.get(r);
        if (rowMap) rowMap.delete(c);
      }
    }
    return this;
  }

  // ─── Sort Ranges ─────────────────────────────────────────────────────────────

  /** Sort a range of cells by a column. `sortCol` is 1-based. */
  sortRange(ref: string, sortCol: number, order: 'asc' | 'desc' = 'asc'): this {
    const { startRow, startCol, endRow, endCol } = parseRange(ref);
    // Collect rows as arrays of cells
    const rows: Array<{ rowIdx: number; cells: Map<number, Cell> }> = [];
    for (let r = startRow; r <= endRow; r++) {
      const rowMap = this.cells.get(r);
      const subset = new Map<number, Cell>();
      for (let c = startCol; c <= endCol; c++) {
        const cell = rowMap?.get(c);
        if (cell) subset.set(c, { ...cell });
      }
      rows.push({ rowIdx: r, cells: subset });
    }
    // Sort by sortCol value
    rows.sort((a, b) => {
      const va = a.cells.get(sortCol)?.value;
      const vb = b.cells.get(sortCol)?.value;
      const na = typeof va === 'number' ? va : typeof va === 'string' ? va : '';
      const nb = typeof vb === 'number' ? vb : typeof vb === 'string' ? vb : '';
      let cmp = 0;
      if (typeof na === 'number' && typeof nb === 'number') cmp = na - nb;
      else cmp = String(na).localeCompare(String(nb));
      return order === 'desc' ? -cmp : cmp;
    });
    // Write sorted rows back
    for (let i = 0; i < rows.length; i++) {
      const r = startRow + i;
      for (let c = startCol; c <= endCol; c++) {
        const cell = rows[i].cells.get(c);
        const rowMap = this.cells.get(r) ?? new Map();
        if (!this.cells.has(r)) this.cells.set(r, rowMap);
        if (cell) rowMap.set(c, cell);
        else rowMap.delete(c);
      }
    }
    return this;
  }

  // ─── Fill Operations ──────────────────────────────────────────────────────────

  /** Fill a column with a numeric sequence. */
  fillNumber(startRow: number, col: number, count: number, startValue: number = 0, step: number = 1): this {
    for (let i = 0; i < count; i++) {
      this.setValue(startRow + i, col, startValue + i * step);
    }
    return this;
  }

  /** Fill a column with dates. */
  fillDate(startRow: number, col: number, count: number, startDate: Date, unit: 'day' | 'week' | 'month' | 'year' = 'day', step: number = 1): this {
    for (let i = 0; i < count; i++) {
      const d = new Date(startDate);
      switch (unit) {
        case 'day':   d.setDate(d.getDate() + i * step); break;
        case 'week':  d.setDate(d.getDate() + i * step * 7); break;
        case 'month': d.setMonth(d.getMonth() + i * step); break;
        case 'year':  d.setFullYear(d.getFullYear() + i * step); break;
      }
      this.setValue(startRow + i, col, d);
    }
    return this;
  }

  /** Fill a column by cycling through a list of values. */
  fillList(startRow: number, col: number, list: CellValue[], count: number): this {
    for (let i = 0; i < count; i++) {
      this.setValue(startRow + i, col, list[i % list.length]);
    }
    return this;
  }

  // ─── AutoFit Columns ─────────────────────────────────────────────────────────

  /** Approximate autofit based on character count (no font metrics, assumes ~1.2 chars/unit). */
  autoFitColumns(minWidth: number = 8, maxWidth: number = 60): this {
    const range = this.getUsedRange();
    if (!range) return this;
    for (let c = range.startCol; c <= range.endCol; c++) {
      let maxLen = 0;
      for (let r = range.startRow; r <= range.endRow; r++) {
        const cell = this.cells.get(r)?.get(c);
        if (cell?.value != null) {
          const s = String(cell.value);
          if (s.length > maxLen) maxLen = s.length;
        }
        if (cell?.richText) {
          const s = cell.richText.map(r => r.text).join('');
          if (s.length > maxLen) maxLen = s.length;
        }
      }
      if (maxLen > 0) {
        const w = Math.max(minWidth, Math.min(maxWidth, maxLen * 1.2 + 2));
        this.setColumn(c, { ...(this.colDefs.get(c) ?? {}), width: w, customWidth: true });
      }
    }
    return this;
  }

  // ─── Row Duplicate / Splice ───────────────────────────────────────────────────

  /** Duplicate a row and insert the copy at targetRow. */
  duplicateRow(sourceRow: number, targetRow: number): this {
    this.insertRows(targetRow, 1);
    const srcMap = this.cells.get(sourceRow >= targetRow ? sourceRow + 1 : sourceRow);
    if (srcMap) {
      const newMap = new Map<number, Cell>();
      for (const [c, cell] of srcMap) {
        newMap.set(c, { ...cell, style: cell.style ? { ...cell.style } : undefined });
      }
      this.cells.set(targetRow, newMap);
    }
    // Copy row def
    const srcDef = this.rowDefs.get(sourceRow >= targetRow ? sourceRow + 1 : sourceRow);
    if (srcDef) this.rowDefs.set(targetRow, { ...srcDef });
    return this;
  }

  /** Splice rows: delete `deleteCount` rows at `startRow`, then insert `newRows` data. */
  spliceRows(startRow: number, deleteCount: number, newRows?: CellValue[][]): this {
    if (deleteCount > 0) this.deleteRows(startRow, deleteCount);
    if (newRows && newRows.length > 0) {
      this.insertRows(startRow, newRows.length);
      for (let i = 0; i < newRows.length; i++) {
        this.writeRow(startRow + i, 1, newRows[i]);
      }
    }
    return this;
  }

  // ─── Advanced Auto Filters ────────────────────────────────────────────────────

  /** Set autoFilter with optional column filter criteria. */
  setAutoFilter(ref: string, opts?: {
    columns?: Array<{
      col: number;
      type: 'custom' | 'top10' | 'value' | 'dynamic';
      operator?: string;
      val?: string | number;
      top?: boolean;
      percent?: boolean;
      items?: string[];
      dynamicType?: string;
    }>;
  }): this {
    this.autoFilter = { ref };
    if (opts?.columns) {
      this._filterColumns = opts.columns;
    }
    return this;
  }

  /** Internal: advanced filter column definitions */
  _filterColumns?: Array<{
    col: number;
    type: string;
    operator?: string;
    val?: string | number;
    top?: boolean;
    percent?: boolean;
    items?: string[];
    dynamicType?: string;
  }>;

  // ─── Sparklines ──────────────────────────────────────────────────────────────

  addSparkline(s: Sparkline): this {
    this.sparklines.push(s);
    return this;
  }

  getSparklines(): readonly Sparkline[] { return this.sparklines; }

  // ─── Data Validation ─────────────────────────────────────────────────────────

  addDataValidation(sqref: string, dv: DataValidation): this {
    this.dataValidations.set(sqref, dv);
    return this;
  }

  // ─── Page Breaks ────────────────────────────────────────────────────────────

  addRowBreak(row: number, manual = true): this {
    this.rowBreaks.push({ id: row, manual });
    return this;
  }

  addColBreak(col: number, manual = true): this {
    this.colBreaks.push({ id: col, manual });
    return this;
  }

  getRowBreaks(): readonly PageBreak[] { return this.rowBreaks; }
  getColBreaks(): readonly PageBreak[] { return this.colBreaks; }

  // ─── Form Controls ─────────────────────────────────────────────────────────

  addFormControl(ctrl: FormControl): this {
    // Resolve 'to' from width/height if omitted (approx 64px/col, 20px/row)
    if (!ctrl.to && (ctrl.width || ctrl.height)) {
      const COL_PX = 64, ROW_PX = 20;
      const w = ctrl.width ?? 100, h = ctrl.height ?? 30;
      const endColFrac = ctrl.from.col + w / COL_PX;
      const endRowFrac = ctrl.from.row + h / ROW_PX;
      ctrl = { ...ctrl, to: {
        col: Math.floor(endColFrac),
        row: Math.floor(endRowFrac),
        colOff: Math.round((endColFrac % 1) * COL_PX),
        rowOff: Math.round((endRowFrac % 1) * ROW_PX),
      }};
    }
    this.formControls.push(ctrl);
    return this;
  }

  getFormControls(): FormControl[] { return this.formControls; }

  // ─── Shapes ──────────────────────────────────────────────────────────────────

  addShape(shape: Shape): this { this.shapes.push(shape); return this; }
  getShapes(): Shape[] { return this.shapes; }

  // ─── WordArt ─────────────────────────────────────────────────────────────────

  addWordArt(wa: WordArt): this { this.wordArt.push(wa); return this; }
  getWordArt(): WordArt[] { return this.wordArt; }

  // ─── Query Tables ────────────────────────────────────────────────────────────

  addQueryTable(qt: QueryTable): this { this.queryTables.push(qt); return this; }
  getQueryTables(): QueryTable[] { return this.queryTables; }

  // ─── Table Slicers ───────────────────────────────────────────────────────────

  addTableSlicer(slicer: TableSlicer): this { this.tableSlicers.push(slicer); return this; }
  getTableSlicers(): TableSlicer[] { return this.tableSlicers; }

  // ─── Print Area ──────────────────────────────────────────────────────────────

  /** Print area reference, e.g. "A1:D10" or "$A$1:$D$10".
   *  Converted to a _xlnm.Print_Area defined name at build time. */
  printArea?: string;

  // ─── Ignore Error Rules ──────────────────────────────────────────────────────

  private ignoreErrors: Array<{ sqref: string; numberStoredAsText?: boolean; formula?: boolean;
    formulaRange?: boolean; unlockedFormula?: boolean; evalError?: boolean;
    twoDigitTextYear?: boolean; emptyRef?: boolean; listDataValidation?: boolean;
    calculatedColumn?: boolean }> = [];

  /** Add an ignoredError rule to suppress green triangles for the given range. */
  addIgnoredError(sqref: string, opts: { numberStoredAsText?: boolean; formula?: boolean;
    formulaRange?: boolean; unlockedFormula?: boolean; evalError?: boolean;
    twoDigitTextYear?: boolean; emptyRef?: boolean; listDataValidation?: boolean;
    calculatedColumn?: boolean }): this {
    this.ignoreErrors.push({ sqref, ...opts });
    return this;
  }

  getIgnoredErrors() { return this.ignoreErrors; }

  // ─── Preserved XML (round-trip) ─────────────────────────────────────────────

  addPreservedXml(xml: string): this {
    this.preservedXml.push(xml);
    return this;
  }

  // ─── Freeze / Split ──────────────────────────────────────────────────────────

  freeze(row?: number, col?: number): this {
    this.freezePane = { row, col };
    return this;
  }

  // ─── XML Generation ──────────────────────────────────────────────────────────

  toXml(styles: StyleRegistry, shared: SharedStrings): string {
    // <sheetPr> — fitToPage lives here per OOXML spec §18.3.1.82
    const fitToPage = this.pageSetup?.fitToPage;
    const tabColor  = this.options?.tabColor;
    const sheetPrXml = (fitToPage || tabColor)
      ? `<sheetPr>${tabColor ? `<tabColor rgb="${tabColor}"/>` : ''}${fitToPage ? '<pageSetUpPr fitToPage="1"/>' : ''}</sheetPr>`
      : '';
    const sheetViewXml = this._sheetViewXml();
    const colsXml      = this._colsXml(styles);
    const sheetDataXml = this._sheetDataXml(styles, shared);
    const mergesXml    = this._mergesXml();
    const cfXml        = this._conditionalFormatXml(styles);
    const dvXml        = this._dataValidationsXml();
    // Skip worksheet-level autoFilter if a table covers the same range (table has its own)
    const autoFilterXml = this.autoFilter && !this.tables.some(t => t.ref === this.autoFilter!.ref)
      ? this._autoFilterXml() : '';
    const tablePartsXml = this.tables.length
      ? `<tableParts count="${this.tables.length}">${
          this.tableRIds.map(rId => `<tablePart r:id="${rId}"/>`).join('')
        }</tableParts>`
      : '';
    const drawingXml = (this.images.length || this.charts.length) && this.drawingRId
      ? `<drawing r:id="${this.drawingRId}"/>`
      : '';
    const legacyDrawingXml = this.legacyDrawingRId
      ? `<legacyDrawing r:id="${this.legacyDrawingRId}"/>`
      : '';
    const controlsXml = this._formControlsXml();
    const sparklineXml = this._sparklineXml();
    const customIconExtXml = this._customIconExtXml();
    const ignoredErrorsXml = this._ignoredErrorsXml();
    const protectionXml = this._protectionXml();
    const pageSetupXml  = this._pageSetupXml();
    const pageMarginsXml = this._pageMarginsXml();
    const headerFooterXml = this._headerFooterXml();
    const printOptionsXml = this._printOptionsXml();
    const rowBreaksXml = this._pageBreaksXml('rowBreaks', this.rowBreaks, 16383);
    const colBreaksXml = this._pageBreaksXml('colBreaks', this.colBreaks, 1048575);

    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
  xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision"
  xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
  xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing">
${sheetPrXml}
${sheetViewXml}
${colsXml}
${sheetDataXml}
${protectionXml}
${autoFilterXml}
${mergesXml}
${cfXml}
${dvXml}
${printOptionsXml}
${pageMarginsXml}
${pageSetupXml}
${headerFooterXml}
${rowBreaksXml}
${colBreaksXml}
${ignoredErrorsXml}
${drawingXml}
${legacyDrawingXml}
${controlsXml}
${sparklineXml}
${customIconExtXml}
${tablePartsXml}
${this.preservedXml.join('\n')}
</worksheet>`;
  }

  /** Generate XML for a chart sheet (dedicated sheet = entire chart). */
  toChartSheetXml(): string {
    const pageMarginsXml = this._pageMarginsXml();
    const pageSetupXml   = this._pageSetupXml();
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<chartsheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheetPr/>
<sheetViews><sheetView zoomScale="99" workbookViewId="0" zoomToFit="1"/></sheetViews>
${pageMarginsXml || '<pageMargins left="0.7" right="0.7" top="0.78740157499999996" bottom="0.78740157499999996" header="0.3" footer="0.3"/>'}
${pageSetupXml}
<drawing r:id="${this.drawingRId}"/>
</chartsheet>`;
  }

  /** Generate XML for a dialog sheet (Excel 5 dialog). */
  toDialogSheetXml(_styles: StyleRegistry, _shared: SharedStrings): string {
    const pageMarginsXml = this._pageMarginsXml();
    const legacyDrawingXml = this.legacyDrawingRId
      ? `<legacyDrawing r:id="${this.legacyDrawingRId}"/>` : '';
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<dialogsheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheetViews><sheetView showRowColHeaders="0" showZeros="0" showOutlineSymbols="0" workbookViewId="0"/></sheetViews>
<sheetFormatPr baseColWidth="10" defaultColWidth="1" defaultRowHeight="5.65" customHeight="1"/>
<sheetProtection sheet="1"/>
${pageMarginsXml || '<pageMargins left="0.7" right="0.7" top="0.78740157499999996" bottom="0.78740157499999996" header="0.3" footer="0.3"/>'}
${legacyDrawingXml}
</dialogsheet>`;
  }

  private _sheetViewXml(): string {
    const v = this.view ?? {};
    const attrs: string[] = [
      'workbookViewId="0"',
      v.showGridLines     === false ? 'showGridLines="0"'     : '',
      v.showRowColHeaders === false ? 'showRowColHeaders="0"' : '',
      v.zoomScale !== undefined     ? `zoomScale="${v.zoomScale}"` : '',
      v.rightToLeft ? 'rightToLeft="1"' : '',
      v.tabSelected ? 'tabSelected="1"' : '',
      v.view        ? `view="${v.view}"` : '',
    ].filter(Boolean);

    let paneXml = '';
    if (this.freezePane) {
      const { row = 0, col = 0 } = this.freezePane;
      const topLeft = row && col
        ? indicesToCellRef(row + 1, col + 1)
        : row ? indicesToCellRef(row + 1, 1)
               : indicesToCellRef(1, col + 1);
      const activePane = row && col ? 'bottomRight' : row ? 'bottomLeft' : 'topRight';
      paneXml = `<pane xSplit="${col}" ySplit="${row}" topLeftCell="${topLeft}" activePane="${activePane}" state="frozen"/>
<selection pane="${activePane}" activeCell="${topLeft}" sqref="${topLeft}"/>`;
    }

    return `<sheetViews><sheetView ${attrs.join(' ')}>${paneXml}</sheetView></sheetViews>`;
  }

  private _colsXml(styles: StyleRegistry): string {
    if (!this.colDefs.size) return '';
    const sorted = [...this.colDefs.entries()].sort((a, b) => a[0] - b[0]);
    const items = sorted.map(([idx, def]) => {
      const styleIdx = def.style ? styles.register(def.style) : 0;
      return `<col min="${idx}" max="${idx}"` +
        (def.width      ? ` width="${def.width}" customWidth="1"` : '') +
        (def.hidden     ? ` hidden="1"` : '') +
        (def.bestFit    ? ` bestFit="1"` : '') +
        (def.outlineLevel ? ` outlineLevel="${def.outlineLevel}"` : '') +
        (def.collapsed  ? ` collapsed="1"` : '') +
        (styleIdx       ? ` style="${styleIdx}"` : '') +
        '/>';
    });
    return `<cols>${items.join('')}</cols>`;
  }

  private _sheetDataXml(styles: StyleRegistry, shared: SharedStrings): string {
    const sortedRows = [...this.cells.entries()].sort((a, b) => a[0] - b[0]);
    const out: string[] = ['<sheetData>'];

    for (let ri = 0; ri < sortedRows.length; ri++) {
      const [rowIdx, colMap] = sortedRows[ri];
      const rowDef = this.rowDefs.get(rowIdx);
      const rowStyleIdx = rowDef?.style ? styles.register(rowDef.style) : 0;
      let attrs = `r="${rowIdx}"`;
      if (rowDef?.height)       attrs += ` ht="${rowDef.height}" customHeight="1"`;
      if (rowDef?.hidden)       attrs += ' hidden="1"';
      if (rowDef?.outlineLevel) attrs += ` outlineLevel="${rowDef.outlineLevel}"`;
      if (rowDef?.collapsed)    attrs += ' collapsed="1"';
      if (rowStyleIdx)          attrs += ` s="${rowStyleIdx}" customFormat="1"`;
      if (rowDef?.thickTop)     attrs += ' thickTop="1"';
      if (rowDef?.thickBot)     attrs += ' thickBot="1"';

      out.push(`<row ${attrs}>`);
      const sortedCells = [...colMap.entries()].sort((a, b) => a[0] - b[0]);
      for (let ci = 0; ci < sortedCells.length; ci++) {
        out.push(this._cellXml(rowIdx, sortedCells[ci][0], sortedCells[ci][1], styles, shared));
      }
      out.push('</row>');
    }

    out.push('</sheetData>');
    return out.join('');
  }

  private _cellXml(row: number, col: number, cell: Cell, styles: StyleRegistry, shared: SharedStrings): string {
    const ref = `${colIndexToLetter(col)}${row}`;
    const styleIdx = cell.style ? styles.register(cell.style) : 0;
    const sAttr = styleIdx ? ` s="${styleIdx}"` : '';
    const vmIdx = this._cellImageVm.get(ref);
    const vmAttr = vmIdx !== undefined ? ` vm="${vmIdx}"` : '';

    // Array formula (including dynamic)
    if (cell.arrayFormula) {
      const fml = `<f t="array" ref="${ref}">${escapeXml(cell.arrayFormula)}</f>`;
      return `<c r="${ref}"${sAttr}${vmAttr}>${fml}<v>0</v></c>`;
    }

    // Shared formula (master or dependent)
    if ((cell as any)._sharedIdx !== undefined) {
      const si = (cell as any)._sharedIdx;
      const sharedRef = (cell as any)._sharedRef;
      if (cell.formula && sharedRef) {
        // Master cell
        const fml = `<f t="shared" ref="${sharedRef}" si="${si}">${escapeXml(cell.formula)}</f>`;
        return `<c r="${ref}"${sAttr}${vmAttr}>${fml}</c>`;
      }
      // Dependent cell
      return `<c r="${ref}"${sAttr}${vmAttr}><f t="shared" si="${si}"/></c>`;
    }

    // Formula
    if (cell.formula) {
      const fml = `<f>${escapeXml(cell.formula)}</f>`;
      return `<c r="${ref}"${sAttr}${vmAttr}>${fml}</c>`;
    }

    // Rich text
    if (cell.richText) {
      const si = shared.internRichText(cell.richText);
      return `<c r="${ref}" t="s"${sAttr}${vmAttr}><v>${si}</v></c>`;
    }

    const v = cell.value;
    if (v === null || v === undefined) {
      // Cell image with no value — emit as error cell (Excel expects t="e" + #VALUE! for cell pictures)
      if (vmAttr) return `<c r="${ref}"${sAttr} t="e"${vmAttr}><v>#VALUE!</v></c>`;
      return styleIdx ? `<c r="${ref}"${sAttr}/>` : '';
    }

    if (v instanceof CellError) {
      return `<c r="${ref}" t="e"${sAttr}${vmAttr}><v>${escapeXml(v.error)}</v></c>`;
    }

    if (typeof v === 'boolean') {
      return `<c r="${ref}" t="b"${sAttr}${vmAttr}><v>${v ? 1 : 0}</v></c>`;
    }

    if (v instanceof Date) {
      const serial = dateToSerial(v);
      return `<c r="${ref}"${sAttr}${vmAttr}><v>${serial}</v></c>`;
    }

    if (typeof v === 'number') {
      return `<c r="${ref}"${sAttr}${vmAttr}><v>${v}</v></c>`;
    }

    const si = shared.intern(v as string);
    return `<c r="${ref}" t="s"${sAttr}${vmAttr}><v>${si}</v></c>`;
  }

  private _mergesXml(): string {
    if (!this.merges.length) return '';
    const items = this.merges.map(m => {
      const start = `${colIndexToLetter(m.startCol)}${m.startRow}`;
      const end   = `${colIndexToLetter(m.endCol)}${m.endRow}`;
      return `<mergeCell ref="${start}:${end}"/>`;
    });
    return `<mergeCells count="${this.merges.length}">${items.join('')}</mergeCells>`;
  }

  private _conditionalFormatXml(styles: StyleRegistry): string {
    return this.conditionalFormats.map(cf => {
      const dxfId = cf.style ? styles.registerDxf(cf.style) : undefined;
      let inner = '';

      if (cf.colorScale?.type === 'colorScale') {
        const cs = cf.colorScale;
        const cfvos = cs.cfvo.map(v => `<cfvo type="${v.type}"${v.val ? ` val="${v.val}"` : ''}/>`).join('');
        const colors = cs.color.map(c => colorEl(c)).join('');
        inner = `<colorScale>${cfvos}${colors}</colorScale>`;
      } else if (cf.dataBar?.type === 'dataBar') {
        const db = cf.dataBar;
        // OOXML requires cfvo min and max, then color element(s)
        const minCfvo = `<cfvo type="${db.minType ?? 'min'}"${db.minVal != null ? ` val="${db.minVal}"` : ''}/>`;
        const maxCfvo = `<cfvo type="${db.maxType ?? 'max'}"${db.maxVal != null ? ` val="${db.maxVal}"` : ''}/>`;
        const color   = db.color ?? db.minColor ?? 'FF638EC6';
        inner = `<dataBar${db.showValue === false ? ' showValue="0"' : ''}>${minCfvo}${maxCfvo}${colorEl(color)}</dataBar>`;
      } else if (cf.iconSet?.type === 'iconSet') {
        const is = cf.iconSet;
        const cfvos = is.cfvo.map(v => `<cfvo type="${v.type}"${v.val ? ` val="${v.val}"` : ''}/>`).join('');
        // Custom icon overrides go into extLst, standard iconSet stays in base CF
        inner = `<iconSet iconSet="${is.iconSet}"${is.showValue===false?' showValue="0"':''}${is.reverse?' reverse="1"':''}>${cfvos}</iconSet>`;
      }

      const ruleAttrs = [
        `type="${cf.type}"`,
        cf.operator ? `operator="${cf.operator}"` : '',
        dxfId !== undefined ? `dxfId="${dxfId}"` : '',
        `priority="${cf.priority ?? 1}"`,
        cf.aboveAverage === false ? `aboveAverage="0"` : '',
        cf.percent ? `percent="1"` : '',
        cf.rank    ? `rank="${cf.rank}"` : '',
        cf.timePeriod ? `timePeriod="${cf.timePeriod}"` : '',
        cf.text    ? `text="${escapeXml(cf.text)}"` : '',
      ].filter(Boolean).join(' ');

      const fml1 = cf.formula  ? `<formula>${escapeXml(cf.formula)}</formula>`  : '';
      const fml2 = cf.formula2 ? `<formula>${escapeXml(cf.formula2)}</formula>` : '';

      return `<conditionalFormatting sqref="${cf.sqref}"><cfRule ${ruleAttrs}>${fml1}${fml2}${inner}</cfRule></conditionalFormatting>`;
    }).join('');
  }

  private _dataValidationsXml(): string {
    if (!this.dataValidations.size) return '';
    const items = [...this.dataValidations.entries()].map(([sqref, dv]) => {
      const formula1 = dv.type === 'list' && dv.list
        ? `<formula1>"${dv.list.join(',')}"</formula1>`
        : dv.formula1 ? `<formula1>${escapeXml(dv.formula1)}</formula1>` : '';
      const formula2 = dv.formula2 ? `<formula2>${escapeXml(dv.formula2)}</formula2>` : '';
      const attrs = [
        `type="${dv.type}"`,
        dv.operator ? `operator="${dv.operator}"` : '',
        `sqref="${sqref}"`,
        dv.showDropDown !== false && dv.type === 'list' ? '' : 'showDropDown="1"',
        dv.allowBlank !== false ? 'allowBlank="1"' : '',
        dv.showErrorAlert ? 'showErrorMessage="1"' : '',
        dv.errorTitle ? `errorTitle="${escapeXml(dv.errorTitle)}"` : '',
        dv.error      ? `error="${escapeXml(dv.error)}"` : '',
        dv.showInputMessage ? 'showInputMessage="1"' : '',
        dv.promptTitle ? `promptTitle="${escapeXml(dv.promptTitle)}"` : '',
        dv.prompt      ? `prompt="${escapeXml(dv.prompt)}"` : '',
      ].filter(Boolean).join(' ');
      return `<dataValidation ${attrs}>${formula1}${formula2}</dataValidation>`;
    });
    return `<dataValidations count="${this.dataValidations.size}">${items.join('')}</dataValidations>`;
  }

  private _protectionXml(): string {
    const p = this.protection;
    if (!p) return '';
    const attrs = [
      p.sheet ? 'sheet="1"' : '',
      p.password ? `password="${hashPassword(p.password)}"` : '',
      p.selectLockedCells   === false ? 'selectLockedCells="0"' : '',
      p.selectUnlockedCells === false ? 'selectUnlockedCells="0"' : '',
      p.formatCells     ? 'formatCells="0"' : '',
      p.formatColumns   ? 'formatColumns="0"' : '',
      p.formatRows      ? 'formatRows="0"' : '',
      p.insertColumns   ? 'insertColumns="0"' : '',
      p.insertRows      ? 'insertRows="0"' : '',
      p.insertHyperlinks ? 'insertHyperlinks="0"' : '',
      p.deleteColumns   ? 'deleteColumns="0"' : '',
      p.deleteRows      ? 'deleteRows="0"' : '',
      p.sort            ? 'sort="0"' : '',
      p.autoFilter      ? 'autoFilter="0"' : '',
      p.pivotTables     ? 'pivotTables="0"' : '',
    ].filter(Boolean).join(' ');
    return `<sheetProtection ${attrs}/>`;
  }

  private _pageSetupXml(): string {
    const p = this.pageSetup;
    if (!p) return '';
    const attrs = [
      p.paperSize      ? `paperSize="${p.paperSize}"` : '',
      p.orientation    ? `orientation="${p.orientation}"` : '',
      p.fitToWidth  !== undefined ? `fitToWidth="${p.fitToWidth}"` : '',
      p.fitToHeight !== undefined ? `fitToHeight="${p.fitToHeight}"` : '',
      p.scale          ? `scale="${p.scale}"` : '',
      p.horizontalDpi  ? `horizontalDpi="${p.horizontalDpi}"` : '',
      p.verticalDpi    ? `verticalDpi="${p.verticalDpi}"` : '',
      p.firstPageNumber ? `firstPageNumber="${p.firstPageNumber}" useFirstPageNumber="1"` : '',
    ].filter(Boolean).join(' ');
    return `<pageSetup ${attrs}/>`;
  }

  private _pageMarginsXml(): string {
    const m = this.pageMargins;
    const defaults = { left:0.7, right:0.7, top:0.75, bottom:0.75, header:0.3, footer:0.3 };
    const d = { ...defaults, ...m };
    return `<pageMargins left="${d.left}" right="${d.right}" top="${d.top}" bottom="${d.bottom}" header="${d.header}" footer="${d.footer}"/>`;
  }

  private _headerFooterXml(): string {
    const hf = this.headerFooter;
    if (!hf) return '';
    const attrs = [
      hf.differentOddEven ? 'differentOddEven="1"' : '',
      hf.differentFirst   ? 'differentFirst="1"'   : '',
    ].filter(Boolean).join(' ');
    const oh = hf.oddHeader   ? `<oddHeader>${escapeXml(hf.oddHeader)}</oddHeader>` : '';
    const of_ = hf.oddFooter  ? `<oddFooter>${escapeXml(hf.oddFooter)}</oddFooter>` : '';
    const eh = hf.evenHeader  ? `<evenHeader>${escapeXml(hf.evenHeader)}</evenHeader>` : '';
    const ef = hf.evenFooter  ? `<evenFooter>${escapeXml(hf.evenFooter)}</evenFooter>` : '';
    const fh = hf.firstHeader ? `<firstHeader>${escapeXml(hf.firstHeader)}</firstHeader>` : '';
    const ff = hf.firstFooter ? `<firstFooter>${escapeXml(hf.firstFooter)}</firstFooter>` : '';
    return `<headerFooter${attrs ? ' '+attrs : ''}>${oh}${of_}${eh}${ef}${fh}${ff}</headerFooter>`;
  }

  private _printOptionsXml(): string {
    const p = this.printOptions;
    if (!p) return '';
    const attrs = [
      p.gridLines        ? 'gridLines="1"' : '',
      p.gridLinesSet     ? 'gridLinesSet="1"' : '',
      p.headings         ? 'headings="1"' : '',
      p.centerHorizontal ? 'horizontalCentered="1"' : '',
      p.centerVertical   ? 'verticalCentered="1"' : '',
    ].filter(Boolean).join(' ');
    return attrs ? `<printOptions ${attrs}/>` : '';
  }

  private _pageBreaksXml(tag: string, breaks: PageBreak[], maxVal: number): string {
    if (!breaks.length) return '';
    const manualCount = breaks.filter(b => b.manual !== false).length;
    const brks = breaks.map(b =>
      `<brk id="${b.id}" max="${maxVal}"${b.manual !== false ? ' man="1"' : ''}/>`
    ).join('');
    return `<${tag} count="${breaks.length}" manualBreakCount="${manualCount}">${brks}</${tag}>`;
  }

  private _formControlsXml(): string {
    if (!this.formControls.length || !this.ctrlPropRIds.length) return '';
    const baseShapeId = 1025 + this.sheetIndex * 1000;
    // Count comments to offset shape IDs
    let commentCount = 0;
    for (const rowMap of this.cells.values()) for (const c of rowMap.values()) if (c.comment) commentCount++;
    const controls = this.formControls.map((ctrl, i) => {
      const shapeId = ctrl._shapeId ?? (baseShapeId + commentCount + i);
      const ctrlPropRId = this.ctrlPropRIds[i];
      if (!ctrlPropRId) return '';
      const name = ctrl.text ?? `${ctrl.type} ${i + 1}`;
      return `<mc:AlternateContent><mc:Choice Requires="x14"><control shapeId="${shapeId}" r:id="${ctrlPropRId}" name="${escapeXml(name)}"><controlPr defaultSize="0" print="0" autoFill="0" autoPict="0"${ctrl.macro ? ` macro="${escapeXml(ctrl.macro)}"` : ''}><anchor moveWithCells="1"><from><xdr:col>${ctrl.from.col}</xdr:col><xdr:colOff>${ctrl.from.colOff ?? 0}</xdr:colOff><xdr:row>${ctrl.from.row}</xdr:row><xdr:rowOff>${ctrl.from.rowOff ?? 0}</xdr:rowOff></from><to><xdr:col>${ctrl.to.col}</xdr:col><xdr:colOff>${ctrl.to.colOff ?? 0}</xdr:colOff><xdr:row>${ctrl.to.row}</xdr:row><xdr:rowOff>${ctrl.to.rowOff ?? 0}</xdr:rowOff></to></anchor></controlPr></control></mc:Choice></mc:AlternateContent>`;
    }).join('');
    return `<mc:AlternateContent><mc:Choice Requires="x14"><controls>${controls}</controls></mc:Choice></mc:AlternateContent>`;
  }

  private _sparklineXml(): string {
    if (!this.sparklines.length) return '';
    const colorEl = (name: string, c?: string) => {
      if (!c) return '';
      // Normalize to 8-digit AARRGGBB; strip leading '#' if present
      const rgb = c.startsWith('#') ? 'FF' + c.slice(1) : c;
      return `<x14:${name} rgb="${rgb}"/>`;
    };
    const groups = this.sparklines.map(s => {
      const sparkType = s.type === 'bar' ? 'column' : s.type;
      const attrs = [
        `type="${sparkType}"`,
        s.lineWidth !== undefined ? `lineWeight="${s.lineWidth}"` : '',
        s.showMarkers  ? 'markers="1"'  : '',
        s.showFirst    ? 'first="1"'    : '',
        s.showLast     ? 'last="1"'     : '',
        s.showHigh     ? 'high="1"'     : '',
        s.showLow      ? 'low="1"'      : '',
        s.showNegative ? 'negative="1"' : '',
        s.minAxisType  ? `minAxisType="${s.minAxisType}"` : '',
        s.maxAxisType  ? `maxAxisType="${s.maxAxisType}"` : '',
      ].filter(Boolean).join(' ');
      const colors = [
        colorEl('colorSeries',  s.color),
        colorEl('colorHigh',    s.highColor),
        colorEl('colorLow',     s.lowColor),
        colorEl('colorFirst',   s.firstColor),
        colorEl('colorLast',    s.lastColor),
        colorEl('colorNegative',s.negativeColor),
        colorEl('colorMarkers', s.markersColor),
      ].join('');
      // xm:f requires a fully-qualified sheet reference
      const dataRef = s.dataRange.includes('!') ? s.dataRange : `${this.name}!${s.dataRange}`;
      const sparkline = `<x14:sparklines><x14:sparkline><xm:f>${escapeXml(dataRef)}</xm:f><xm:sqref>${s.location}</xm:sqref></x14:sparkline></x14:sparklines>`;
      return `<x14:sparklineGroup ${attrs}>${colors}${sparkline}</x14:sparklineGroup>`;
    });
    const inner = `<x14:sparklineGroups xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">${groups.join('')}</x14:sparklineGroups>`;
    return `<extLst><ext uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">${inner}</ext></extLst>`;
  }

  private _customIconExtXml(): string {
    // Custom icon set CF rules need to be wrapped in x14:conditionalFormattings extension
    const customCFs = this.conditionalFormats.filter(cf =>
      cf.iconSet?.type === 'iconSet' && 'custom' in cf.iconSet && cf.iconSet.custom?.length
    );
    if (!customCFs.length) return '';

    const rules = customCFs.map((cf, i) => {
      const is = cf.iconSet as CFCustomIconSet;
      const cfvos = is.cfvo.map(v => `<x14:cfvo type="${v.type}"${v.val ? ` val="${v.val}"` : ''}/>`).join('');
      const icons = is.custom!.map(ci => `<x14:cfIcon iconSet="${ci.iconSet}" iconId="${ci.iconId}"/>`).join('');
      return `<x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><x14:cfRule type="iconSet" id="{${this._uuid()}}"><x14:iconSet iconSet="${is.iconSet}" custom="1"${is.showValue===false?' showValue="0"':''}${is.reverse?' reverse="1"':''}>${cfvos}${icons}</x14:iconSet></x14:cfRule><xm:sqref>${cf.sqref}</xm:sqref></x14:conditionalFormatting>`;
    }).join('');

    return `<extLst><ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">${rules}</ext></extLst>`;
  }

  private _uuid(): string {
    const h = '0123456789ABCDEF';
    let u = '';
    for (let i = 0; i < 36; i++) {
      if (i === 8 || i === 13 || i === 18 || i === 23) u += '-';
      else if (i === 14) u += '4';
      else if (i === 19) u += h[(Math.random() * 4 | 0) + 8];
      else u += h[Math.random() * 16 | 0];
    }
    return u;
  }

  private _autoFilterXml(): string {
    if (!this.autoFilter) return '';
    const cols = this._filterColumns;
    if (!cols || cols.length === 0) return `<autoFilter ref="${this.autoFilter.ref}"/>`;
    const colXml = cols.map(fc => {
      const colId = fc.col - 1; // 0-based
      let inner = '';
      if (fc.type === 'custom') {
        const op = fc.operator ?? 'greaterThan';
        inner = `<customFilters><customFilter operator="${op}" val="${fc.val ?? ''}"/></customFilters>`;
      } else if (fc.type === 'top10') {
        const attrs = [
          fc.top === false ? 'top="0"' : '',
          fc.percent ? 'percent="1"' : '',
          `val="${fc.val ?? 10}"`,
        ].filter(Boolean).join(' ');
        inner = `<top10 ${attrs}/>`;
      } else if (fc.type === 'value' && fc.items) {
        inner = `<filters>${fc.items.map(v => `<filter val="${escapeXml(String(v))}"/>`).join('')}</filters>`;
      } else if (fc.type === 'dynamic') {
        inner = `<dynamicFilter type="${fc.dynamicType ?? 'aboveAverage'}"/>`;
      }
      return `<filterColumn colId="${colId}">${inner}</filterColumn>`;
    }).join('');
    return `<autoFilter ref="${this.autoFilter.ref}">${colXml}</autoFilter>`;
  }

  private _ignoredErrorsXml(): string {
    if (!this.ignoreErrors.length) return '';
    const items = this.ignoreErrors.map(ie => {
      const attrs: string[] = [`sqref="${ie.sqref}"`];
      if (ie.numberStoredAsText) attrs.push('numberStoredAsText="1"');
      if (ie.formula)            attrs.push('formula="1"');
      if (ie.formulaRange)       attrs.push('formulaRange="1"');
      if (ie.unlockedFormula)    attrs.push('unlockedFormula="1"');
      if (ie.evalError)          attrs.push('evalError="1"');
      if (ie.twoDigitTextYear)   attrs.push('twoDigitTextYear="1"');
      if (ie.emptyRef)           attrs.push('emptyRef="1"');
      if (ie.listDataValidation) attrs.push('listDataValidation="1"');
      if (ie.calculatedColumn)   attrs.push('calculatedColumn="1"');
      return `<ignoredError ${attrs.join(' ')}/>`;
    });
    return `<ignoredErrors>${items.join('')}</ignoredErrors>`;
  }

  /** Drawing XML (images + charts) — returned separately for the drawing part */
  toDrawingXml(imageRIds: string[], chartRIds: string[]): string {
    const parts: string[] = [];
    const EMU = pxToEmu;

    this.images.forEach((img, i) => {
      const rId = imageRIds[i];
      const w = EMU(img.width ?? 100);
      const h = EMU(img.height ?? 100);

      let anchor: string;
      let closeTag: string;
      if (img.position) {
        // Absolute positioning — no cell reference
        anchor = `<xdr:absoluteAnchor><xdr:pos x="${EMU(img.position.x)}" y="${EMU(img.position.y)}"/><xdr:ext cx="${w}" cy="${h}"/>`;
        closeTag = `</xdr:absoluteAnchor>`;
      } else {
        const from = img.from!;
        const to   = img.to;
        const fromXml = `<xdr:from><xdr:col>${from.col}</xdr:col><xdr:colOff>${from.colOff ?? 0}</xdr:colOff><xdr:row>${from.row}</xdr:row><xdr:rowOff>${from.rowOff ?? 0}</xdr:rowOff></xdr:from>`;
        if (to) {
          const toXml = `<xdr:to><xdr:col>${to.col}</xdr:col><xdr:colOff>${to.colOff ?? 0}</xdr:colOff><xdr:row>${to.row}</xdr:row><xdr:rowOff>${to.rowOff ?? 0}</xdr:rowOff></xdr:to>`;
          anchor = `<xdr:twoCellAnchor editAs="oneCell">${fromXml}${toXml}`;
          closeTag = `</xdr:twoCellAnchor>`;
        } else {
          anchor = `<xdr:oneCellAnchor>${fromXml}<xdr:ext cx="${w}" cy="${h}"/>`;
          closeTag = `</xdr:oneCellAnchor>`;
        }
      }

      const picXml = `<xdr:pic>
  <xdr:nvPicPr>
    <xdr:cNvPr id="${i + 2}" name="Image ${i + 1}"${img.altText ? ` descr="${escapeXml(img.altText)}"` : ''}/>
    <xdr:cNvPicPr><a:picLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/></xdr:cNvPicPr>
  </xdr:nvPicPr>
  <xdr:blipFill>
    <a:blip xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" r:embed="${rId}"/>
    <a:stretch xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:fillRect/></a:stretch>
  </xdr:blipFill>
  <xdr:spPr>
    <a:xfrm xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <a:off x="0" y="0"/><a:ext cx="${w}" cy="${h}"/>
    </a:xfrm>
    <a:prstGeom xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" prst="rect"><a:avLst/></a:prstGeom>
  </xdr:spPr>
</xdr:pic>`;

      parts.push(`${anchor}${picXml}<xdr:clientData/>${closeTag}`);
    });

    this.charts.forEach((chart, i) => {
      const rId = chartRIds[i];
      const from = chart.from;
      const to   = chart.to;
      const graphicXml = `<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
    <c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="${rId}"/>
  </a:graphicData>
</a:graphic>`;
      if (this._isChartSheet) {
        // Chart sheets use absoluteAnchor filling the entire sheet
        parts.push(`<xdr:absoluteAnchor><xdr:pos x="0" y="0"/><xdr:ext cx="9294091" cy="6003636"/><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${this.images.length + i + 2}" name="Chart ${i+1}"/><xdr:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noGrp="1"/></xdr:cNvGraphicFramePr></xdr:nvGraphicFramePr><xdr:xfrm><a:off xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" x="0" y="0"/><a:ext xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" cx="0" cy="0"/></xdr:xfrm>${graphicXml}</xdr:graphicFrame><xdr:clientData/></xdr:absoluteAnchor>`);
      } else {
        const fromXml = `<xdr:from><xdr:col>${from.col}</xdr:col><xdr:colOff>${from.colOff ?? 0}</xdr:colOff><xdr:row>${from.row}</xdr:row><xdr:rowOff>${from.rowOff ?? 0}</xdr:rowOff></xdr:from>`;
        const toXml   = `<xdr:to><xdr:col>${to.col}</xdr:col><xdr:colOff>${to.colOff ?? 0}</xdr:colOff><xdr:row>${to.row}</xdr:row><xdr:rowOff>${to.rowOff ?? 0}</xdr:rowOff></xdr:to>`;
        parts.push(`<xdr:twoCellAnchor editAs="oneCell">${fromXml}${toXml}<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${this.images.length + i + 2}" name="Chart ${i+1}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" x="0" y="0"/><a:ext xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" cx="0" cy="0"/></xdr:xfrm>${graphicXml}</xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>`);
      }
    });

    // ── Shapes ──────────────────────────────────────────────────────────────
    let shapeCounter = this.images.length + this.charts.length + 2;
    // Helper to normalize color to 6-char hex
    const toHex6 = (c: string) => { let h = c.replace(/^#/, ''); if (h.length === 8) h = h.substring(2); return h; };

    this.shapes.forEach(shape => {
      const id = shapeCounter++;
      const from = shape.from;
      const to   = shape.to;
      const fromXml = `<xdr:from><xdr:col>${from.col}</xdr:col><xdr:colOff>${from.colOff ?? 0}</xdr:colOff><xdr:row>${from.row}</xdr:row><xdr:rowOff>${from.rowOff ?? 0}</xdr:rowOff></xdr:from>`;
      const toXml   = `<xdr:to><xdr:col>${to.col}</xdr:col><xdr:colOff>${to.colOff ?? 0}</xdr:colOff><xdr:row>${to.row}</xdr:row><xdr:rowOff>${to.rowOff ?? 0}</xdr:rowOff></xdr:to>`;
      const fillXml = shape.fillColor
        ? `<a:solidFill><a:srgbClr val="${toHex6(shape.fillColor)}"/></a:solidFill>`
        : '<a:noFill/>';
      const lineXml = shape.lineColor
        ? `<a:ln${shape.lineWidth ? ` w="${shape.lineWidth * 12700}"` : ''}><a:solidFill><a:srgbClr val="${toHex6(shape.lineColor)}"/></a:solidFill></a:ln>`
        : '';
      const rotAttr = shape.rotation ? ` rot="${shape.rotation * 60000}"` : '';
      const textXml = shape.text ? `<xdr:txBody><a:bodyPr vertOverflow="clip" wrap="square" rtlCol="0" anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en-US"${shape.font?.bold ? ' b="1"' : ''}${shape.font?.size ? ` sz="${shape.font.size * 100}"` : ''}/><a:t>${escapeXml(shape.text)}</a:t></a:r></a:p></xdr:txBody>` : '';
      parts.push(`<xdr:twoCellAnchor editAs="oneCell">${fromXml}${toXml}<xdr:sp><xdr:nvSpPr><xdr:cNvPr id="${id}" name="Shape ${id}"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr><a:xfrm${rotAttr}><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm><a:prstGeom prst="${shape.type}"><a:avLst/></a:prstGeom>${fillXml}${lineXml}</xdr:spPr>${textXml}</xdr:sp><xdr:clientData/></xdr:twoCellAnchor>`);
    });

    // ── WordArt ─────────────────────────────────────────────────────────────
    this.wordArt.forEach(wa => {
      const id = shapeCounter++;
      const from = wa.from;
      const to   = wa.to;
      const fromXml = `<xdr:from><xdr:col>${from.col}</xdr:col><xdr:colOff>${from.colOff ?? 0}</xdr:colOff><xdr:row>${from.row}</xdr:row><xdr:rowOff>${from.rowOff ?? 0}</xdr:rowOff></xdr:from>`;
      const toXml   = `<xdr:to><xdr:col>${to.col}</xdr:col><xdr:colOff>${to.colOff ?? 0}</xdr:colOff><xdr:row>${to.row}</xdr:row><xdr:rowOff>${to.rowOff ?? 0}</xdr:rowOff></xdr:to>`;
      const preset = wa.preset ?? 'textPlain';
      const fillXml = wa.fillColor
        ? `<a:solidFill><a:srgbClr val="${toHex6(wa.fillColor)}"/></a:solidFill>`
        : '<a:solidFill><a:schemeClr val="tx1"/></a:solidFill>';
      const outlineXml = wa.outlineColor
        ? `<a:ln><a:solidFill><a:srgbClr val="${toHex6(wa.outlineColor)}"/></a:solidFill></a:ln>`
        : '';
      const fontAttrs = [
        'lang="en-US"',
        wa.font?.bold ? 'b="1"' : '',
        wa.font?.italic ? 'i="1"' : '',
        wa.font?.size ? `sz="${wa.font.size * 100}"` : 'sz="3600"',
      ].filter(Boolean).join(' ');
      parts.push(`<xdr:twoCellAnchor editAs="oneCell">${fromXml}${toXml}<xdr:sp><xdr:nvSpPr><xdr:cNvPr id="${id}" name="WordArt ${id}"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/>${outlineXml}</xdr:spPr><xdr:txBody><a:bodyPr wrap="none" lIns="91440" tIns="45720" rIns="91440" bIns="45720"><a:prstTxWarp prst="${preset}"><a:avLst/></a:prstTxWarp></a:bodyPr><a:lstStyle/><a:p><a:r><a:rPr ${fontAttrs}>${fillXml}</a:rPr><a:t>${escapeXml(wa.text)}</a:t></a:r></a:p></xdr:txBody></xdr:sp><xdr:clientData/></xdr:twoCellAnchor>`);
    });

    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
${parts.join('\n')}
</xdr:wsDr>`;
  }
}

/** Simple Excel password hash (XOR method) */
function hashPassword(pwd: string): string {
  let hash = 0;
  for (let i = pwd.length - 1; i >= 0; i--) {
    hash = ((hash >> 14) & 0x01) | ((hash << 1) & 0x7FFF);
    hash ^= pwd.charCodeAt(i);
  }
  hash = ((hash >> 14) & 0x01) | ((hash << 1) & 0x7FFF);
  hash ^= pwd.length;
  hash ^= 0xCE4B;
  return hash.toString(16).toUpperCase().padStart(4, '0');
}
