import type {
  Cell, CellValue, CellStyle, MergeRange, Image, Chart,
  ConditionalFormat, Table, AutoFilter, FreezePane, SplitPane,
  SheetProtection, PageSetup, PageMargins, HeaderFooter, PrintOptions,
  SheetView, ColumnDef, RowDef, Sparkline, DataValidation,
  WorksheetOptions,
} from '../core/types.js';
import type { SharedStrings } from '../core/SharedStrings.js';
import type { StyleRegistry } from '../styles/StyleRegistry.js';
import {
  colIndexToLetter, colLetterToIndex, indicesToCellRef,
  cellRefToIndices, parseRange, escapeXml, strToBytes, dateToSerial,
  pxToEmu, colWidthToEmu, rowHeightToEmu, base64ToBytes,
} from '../utils/helpers.js';

type CellMap = Map<string, Cell>;

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
  private charts: Chart[] = [];
  private conditionalFormats: ConditionalFormat[] = [];
  private tables: Table[] = [];
  private sparklines: Sparkline[] = [];
  private colDefs: Map<number, ColumnDef> = new Map();
  private rowDefs: Map<number, RowDef>    = new Map();
  private dataValidations: Map<string, DataValidation> = new Map();

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

  // assigned by workbook
  sheetIndex = 0;
  rId = '';
  drawingRId = '';
  legacyDrawingRId = '';
  tableRIds: string[] = [];

  constructor(name: string, options: WorksheetOptions = {}) {
    this.name = name;
    this.options = { ...options, name };
  }

  // ─── Cell Access ─────────────────────────────────────────────────────────────

  private key(row: number, col: number): string { return `${row},${col}`; }

  getCell(row: number, col: number): Cell {
    const k = this.key(row, col);
    if (!this.cells.has(k)) this.cells.set(k, {});
    return this.cells.get(k)!;
  }

  getCellByRef(ref: string): Cell {
    const { row, col } = cellRefToIndices(ref);
    return this.getCell(row, col);
  }

  setCell(row: number, col: number, cell: Cell): this {
    this.cells.set(this.key(row, col), cell);
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

  // ─── Images ──────────────────────────────────────────────────────────────────

  addImage(img: Image): this {
    this.images.push(img);
    return this;
  }

  getImages(): readonly Image[] { return this.images; }

  /** Return all cells that have a comment, keyed by "col,row" */
  getComments(): Array<{ row: number; col: number; comment: import('../core/types.js').Comment }> {
    const out: Array<{ row: number; col: number; comment: import('../core/types.js').Comment }> = [];
    for (const [key, cell] of this.cells) {
      if (cell.comment) {
        const [r, c] = key.split(',').map(Number);
        out.push({ row: r, col: c, comment: cell.comment });
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

  // ─── Sparklines ──────────────────────────────────────────────────────────────

  addSparkline(s: Sparkline): this {
    this.sparklines.push(s);
    return this;
  }

  // ─── Data Validation ─────────────────────────────────────────────────────────

  addDataValidation(sqref: string, dv: DataValidation): this {
    this.dataValidations.set(sqref, dv);
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
      ? `<sheetPr>${tabColor ? `<tabColor rgb="${tabColor}"/>` : ''}${fitToPage ? '<pageSetPr fitToPage="1"/>' : ''}</sheetPr>`
      : '';
    const sheetViewXml = this._sheetViewXml();
    const colsXml      = this._colsXml(styles);
    const sheetDataXml = this._sheetDataXml(styles, shared);
    const mergesXml    = this._mergesXml();
    const cfXml        = this._conditionalFormatXml(styles);
    const dvXml        = this._dataValidationsXml();
    const autoFilterXml = this.autoFilter ? `<autoFilter ref="${this.autoFilter.ref}"/>` : '';
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
    const sparklineXml = this._sparklineXml();
    const protectionXml = this._protectionXml();
    const pageSetupXml  = this._pageSetupXml();
    const pageMarginsXml = this._pageMarginsXml();
    const headerFooterXml = this._headerFooterXml();
    const printOptionsXml = this._printOptionsXml();

    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
  xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision">
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
${drawingXml}
${legacyDrawingXml}
${sparklineXml}
${tablePartsXml}
</worksheet>`;
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
    // Group cells by row
    const rows = new Map<number, Array<[number, Cell]>>();
    for (const [key, cell] of this.cells) {
      const [r, c] = key.split(',').map(Number);
      if (!rows.has(r)) rows.set(r, []);
      rows.get(r)!.push([c, cell]);
    }

    const sortedRows = [...rows.entries()].sort((a, b) => a[0] - b[0]);
    const rowsXml = sortedRows.map(([rowIdx, cells]) => {
      const rowDef = this.rowDefs.get(rowIdx);
      const rowStyleIdx = rowDef?.style ? styles.register(rowDef.style) : 0;
      const rowAttrs = [
        `r="${rowIdx}"`,
        rowDef?.height    ? `ht="${rowDef.height}" customHeight="1"` : '',
        rowDef?.hidden    ? `hidden="1"` : '',
        rowDef?.outlineLevel ? `outlineLevel="${rowDef.outlineLevel}"` : '',
        rowDef?.collapsed ? `collapsed="1"` : '',
        rowStyleIdx       ? `s="${rowStyleIdx}" customFormat="1"` : '',
        rowDef?.thickTop  ? `thickTop="1"` : '',
        rowDef?.thickBot  ? `thickBot="1"` : '',
      ].filter(Boolean).join(' ');

      const sortedCells = cells.sort((a, b) => a[0] - b[0]);
      const cellsXml = sortedCells.map(([colIdx, cell]) =>
        this._cellXml(rowIdx, colIdx, cell, styles, shared)
      ).join('');

      return `<row ${rowAttrs}>${cellsXml}</row>`;
    });

    return `<sheetData>${rowsXml.join('')}</sheetData>`;
  }

  private _cellXml(row: number, col: number, cell: Cell, styles: StyleRegistry, shared: SharedStrings): string {
    const ref = `${colIndexToLetter(col)}${row}`;
    const styleIdx = cell.style ? styles.register(cell.style) : 0;
    const sAttr = styleIdx ? ` s="${styleIdx}"` : '';

    // Array formula
    if (cell.arrayFormula) {
      const fml = `<f t="array" ref="${ref}">${escapeXml(cell.arrayFormula)}</f>`;
      return `<c r="${ref}"${sAttr}>${fml}<v>0</v></c>`;
    }

    // Formula
    if (cell.formula) {
      const fml = `<f>${escapeXml(cell.formula)}</f>`;
      return `<c r="${ref}"${sAttr}>${fml}</c>`;
    }

    // Rich text
    if (cell.richText) {
      const si = shared.internRichText(cell.richText);
      return `<c r="${ref}" t="s"${sAttr}><v>${si}</v></c>`;
    }

    const v = cell.value;
    if (v === null || v === undefined) {
      return styleIdx ? `<c r="${ref}"${sAttr}/>` : '';
    }

    if (typeof v === 'boolean') {
      return `<c r="${ref}" t="b"${sAttr}><v>${v ? 1 : 0}</v></c>`;
    }

    if (v instanceof Date) {
      const serial = dateToSerial(v);
      return `<c r="${ref}"${sAttr}><v>${serial}</v></c>`;
    }

    if (typeof v === 'number') {
      return `<c r="${ref}"${sAttr}><v>${v}</v></c>`;
    }

    // String — check if it starts with '=' (formula string)
    if (typeof v === 'string' && v.startsWith('=')) {
      return `<c r="${ref}"${sAttr}><f>${escapeXml(v.slice(1))}</f></c>`;
    }

    const si = shared.intern(v as string);
    return `<c r="${ref}" t="s"${sAttr}><v>${si}</v></c>`;
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
        const colors = cs.color.map(c => `<color rgb="${c.startsWith('#') ? 'FF'+c.slice(1) : c}"/>`).join('');
        inner = `<colorScale>${cfvos}${colors}</colorScale>`;
      } else if (cf.dataBar?.type === 'dataBar') {
        const db = cf.dataBar;
        // OOXML requires cfvo min and max, then color element(s)
        const minCfvo = `<cfvo type="${db.minType ?? 'min'}"${db.minVal != null ? ` val="${db.minVal}"` : ''}/>`;
        const maxCfvo = `<cfvo type="${db.maxType ?? 'max'}"${db.maxVal != null ? ` val="${db.maxVal}"` : ''}/>`;
        const color   = db.color ?? db.minColor ?? 'FF638EC6';
        const rgb     = color.startsWith('#') ? 'FF'+color.slice(1) : color;
        inner = `<dataBar${db.showValue === false ? ' showValue="0"' : ''}>${minCfvo}${maxCfvo}<color rgb="${rgb}"/></dataBar>`;
      } else if (cf.iconSet?.type === 'iconSet') {
        const is = cf.iconSet;
        const cfvos = is.cfvo.map(v => `<cfvo type="${v.type}"${v.val ? ` val="${v.val}"` : ''}/>`).join('');
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

  /** Drawing XML (images + charts) — returned separately for the drawing part */
  toDrawingXml(imageRIds: string[], chartRIds: string[]): string {
    const parts: string[] = [];
    const EMU = pxToEmu;

    this.images.forEach((img, i) => {
      const rId = imageRIds[i];
      const from = img.from;
      const to   = img.to;
      const fromXml = `<xdr:from><xdr:col>${from.col}</xdr:col><xdr:colOff>${from.colOff ?? 0}</xdr:colOff><xdr:row>${from.row}</xdr:row><xdr:rowOff>${from.rowOff ?? 0}</xdr:rowOff></xdr:from>`;

      let anchor: string;
      if (to) {
        const toXml = `<xdr:to><xdr:col>${to.col}</xdr:col><xdr:colOff>${to.colOff ?? 0}</xdr:colOff><xdr:row>${to.row}</xdr:row><xdr:rowOff>${to.rowOff ?? 0}</xdr:rowOff></xdr:to>`;
        anchor = `<xdr:twoCellAnchor editAs="oneCell">${fromXml}${toXml}`;
      } else {
        const w = EMU(img.width ?? 100);
        const h = EMU(img.height ?? 100);
        anchor = `<xdr:oneCellAnchor>${fromXml}<xdr:ext cx="${w}" cy="${h}"/>`;
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
      <a:off x="0" y="0"/><a:ext cx="${EMU(img.width ?? 100)}" cy="${EMU(img.height ?? 100)}"/>
    </a:xfrm>
    <a:prstGeom xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" prst="rect"><a:avLst/></a:prstGeom>
  </xdr:spPr>
</xdr:pic>`;

      const closeTag = to ? `</xdr:twoCellAnchor>` : `</xdr:oneCellAnchor>`;
      parts.push(`${anchor}${picXml}<xdr:clientData/>${closeTag}`);
    });

    this.charts.forEach((chart, i) => {
      const rId = chartRIds[i];
      const from = chart.from;
      const to   = chart.to;
      const fromXml = `<xdr:from><xdr:col>${from.col}</xdr:col><xdr:colOff>${from.colOff ?? 0}</xdr:colOff><xdr:row>${from.row}</xdr:row><xdr:rowOff>${from.rowOff ?? 0}</xdr:rowOff></xdr:from>`;
      const toXml   = `<xdr:to><xdr:col>${to.col}</xdr:col><xdr:colOff>${to.colOff ?? 0}</xdr:colOff><xdr:row>${to.row}</xdr:row><xdr:rowOff>${to.rowOff ?? 0}</xdr:rowOff></xdr:to>`;
      const graphicXml = `<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
    <c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="${rId}"/>
  </a:graphicData>
</a:graphic>`;
      parts.push(`<xdr:twoCellAnchor editAs="oneCell">${fromXml}${toXml}<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${this.images.length + i + 2}" name="Chart ${i+1}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" x="0" y="0"/><a:ext xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" cx="0" cy="0"/></xdr:xfrm>${graphicXml}</xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>`);
    });

    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
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
