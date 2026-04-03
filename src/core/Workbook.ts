import type {
  WorkbookProperties, NamedRange, WorksheetOptions, Image, CellImage, Comment, PivotTable,
  Connection, PowerQuery, ConnectionType,
  Theme, ExternalLink, CustomPivotStyle, LocaleSettings, PivotSlicer,
} from '../core/types.js';
import { Worksheet } from './Worksheet.js';
import { StyleRegistry } from '../styles/StyleRegistry.js';
import { SharedStrings } from './SharedStrings.js';
import { buildChartXml } from '../features/ChartBuilder.js';
import { buildTableXml } from '../features/TableBuilder.js';
import { buildPivotTableFiles } from '../features/PivotTableBuilder.js';
import { buildCtrlPropXml, buildFormControlVmlShape, buildVmlWithControls } from '../features/FormControlBuilder.js';
import { VbaProject } from '../vba/VbaProject.js';
import { buildZip, type ZipEntry, type ZipOptions } from '../utils/zip.js';
import { strToBytes, base64ToBytes, escapeXml, colIndexToLetter } from '../utils/helpers.js';
import { readWorkbook, connTypeToNum, cmdTypeToNum, type ReadResult } from './WorkbookReader.js';
import {
  buildCoreXml, buildAppXml, buildCustomXml,
  type CoreProperties, type ExtendedProperties, type CustomProperty,
} from './properties.js';

export class Workbook {
  private sheets: Worksheet[] = [];
  private namedRanges: NamedRange[] = [];
  private connections: Connection[] = [];
  private powerQueries: PowerQuery[] = [];
  private externalLinks: ExternalLink[] = [];
  private customPivotStyles: CustomPivotStyle[] = [];
  private pivotSlicers: PivotSlicer[] = [];
  /** Theme definition (colors, fonts) */
  theme?: Theme;
  /** Locale settings for number/date formatting */
  locale?: LocaleSettings;
  properties: WorkbookProperties = {};

  /**
   * Compression level for the output ZIP (DEFLATE).
   * 0 = no compression (STORE — fastest, largest file)
   * 1 = fastest compression
   * 6 = default (recommended — good balance of speed and size)
   * 9 = maximum compression (slowest, smallest file)
   * Default: 6
   */
  compressionLevel: number = 6;

  // ─── Extended / custom properties ─────────────────────────────────────────

  /** Core document properties (dc/cp namespace — docProps/core.xml) */
  coreProperties: CoreProperties = {};

  /** Extended application properties (docProps/app.xml) */
  extendedProperties: ExtendedProperties = {};

  /** Custom document properties (docProps/custom.xml) */
  customProperties: CustomProperty[] = [];

  /** VBA macro project (set to enable .xlsm output) */
  vbaProject?: VbaProject;

  /** Save as .xltx template (changes the content type) */
  isTemplate = false;

  // ─── Internal state for round-trip patching ────────────────────────────────

  private _readResult?: ReadResult;
  private _dirtySheets = new Set<number>();

  /** Mark a sheet as modified so it will be re-serialised on write */
  markDirty(sheetIndexOrName: number | string): void {
    if (typeof sheetIndexOrName === 'string') {
      const idx = this.sheets.findIndex(s => s.name === sheetIndexOrName);
      if (idx >= 0) this._dirtySheets.add(idx);
    } else {
      this._dirtySheets.add(sheetIndexOrName);
    }
  }

  // ─── Static factory methods ────────────────────────────────────────────────

  /** Load an existing .xlsx from a Uint8Array (works in browser + Node.js + Deno) */
  static async fromBytes(data: Uint8Array): Promise<Workbook> {
    const wb = new Workbook();
    const result = await readWorkbook(data);
    wb._readResult = result;

    wb.coreProperties     = result.core;
    wb.extendedProperties = result.extended;
    wb.customProperties   = result.custom;

    // Back-compat: mirror into legacy .properties
    wb.properties = {
      title:          result.core.title,
      author:         result.core.creator,
      subject:        result.core.subject,
      description:    result.core.description,
      keywords:       result.core.keywords,
      company:        result.extended.company,
      lastModifiedBy: result.core.lastModifiedBy,
      created:        result.core.created,
      modified:       result.core.modified,
      category:       result.core.category,
      status:         result.core.contentStatus,
    };

    wb.sheets = result.sheets.map(s => s.ws);
    wb.namedRanges = result.namedRanges;
    wb.connections = result.connections;
    wb.powerQueries = result.powerQueries;

    // Parse VBA project if present
    const vbaData = result.unknownParts.get('xl/vbaProject.bin');
    if (vbaData) {
      try { wb.vbaProject = VbaProject.fromBytes(vbaData); } catch { /* not fatal */ }
    }

    return wb;
  }

  /** Load from a base64-encoded .xlsx string */
  static async fromBase64(b64: string): Promise<Workbook> {
    return Workbook.fromBytes(base64ToBytes(b64));
  }

  /** Load from the filesystem (Node.js / Deno / Bun) */
  static async fromFile(path: string): Promise<Workbook> {
    // @ts-ignore
    const fs = await import('fs/promises');
    const buf = await fs.readFile(path);
    return Workbook.fromBytes(new Uint8Array(buf));
  }

  /** Load from a File or Blob (browser) */
  static async fromBlob(blob: Blob): Promise<Workbook> {
    return Workbook.fromBytes(new Uint8Array(await blob.arrayBuffer()));
  }

  // ─── Sheet management ──────────────────────────────────────────────────────

  addSheet(name: string, options: WorksheetOptions = {}): Worksheet {
    const ws = new Worksheet(name, options);
    ws.sheetIndex = this.sheets.length + 1;
    this.sheets.push(ws);
    this._dirtySheets.add(ws.sheetIndex - 1);
    return ws;
  }

  getSheet(name: string): Worksheet | undefined {
    return this.sheets.find(s => s.name === name);
  }

  getSheetByIndex(idx: number): Worksheet | undefined {
    return this.sheets[idx];
  }

  getSheetNames(): string[] {
    return this.sheets.map(s => s.name);
  }

  getSheets(): readonly Worksheet[] {
    return this.sheets;
  }

  removeSheet(name: string): this {
    this.sheets = this.sheets.filter(s => s.name !== name);
    return this;
  }

  /**
   * Add a chart sheet — a dedicated sheet that is entirely a chart.
   * The chart fills the whole sheet area.
   */
  addChartSheet(name: string, chart: import('./types.js').Chart): Worksheet {
    const ws = this.addSheet(name);
    ws._isChartSheet = true;
    ws.addChart(chart);
    return ws;
  }

  /**
   * Add a dialog sheet (Excel 5 dialog).
   * Dialog sheets can contain form controls and legacy dialog elements.
   */
  addDialogSheet(name: string): Worksheet {
    const ws = this.addSheet(name);
    ws._isDialogSheet = true;
    return ws;
  }

  /**
   * Copy an existing worksheet (with all cell data, styles, merges) to a new sheet.
   */
  copySheet(sourceName: string, newName: string): Worksheet {
    const src = this.getSheet(sourceName);
    if (!src) throw new Error(`Sheet "${sourceName}" not found`);
    const ws = this.addSheet(newName);
    // Copy cell data
    const cells = src.readAllCells();
    for (const { row, col, cell } of cells) {
      const target = ws.getCell(row, col);
      if (cell.value != null) target.value = cell.value;
      if (cell.formula)       target.formula = cell.formula;
      if (cell.arrayFormula)  target.arrayFormula = cell.arrayFormula;
      if (cell.richText)      target.richText = cell.richText.map(r => ({ ...r, font: r.font ? { ...r.font } : undefined }));
      if (cell.style)         target.style = { ...cell.style };
      if (cell.comment)       target.comment = { ...cell.comment };
      if (cell.hyperlink)     target.hyperlink = { ...cell.hyperlink };
    }
    // Copy merges
    for (const m of src.getMerges()) {
      ws.merge(m.startRow, m.startCol, m.endRow, m.endCol);
    }
    // Copy column defs
    const range = src.getUsedRange();
    if (range) {
      for (let c = range.startCol; c <= range.endCol; c++) {
        const cd = src.getColumn(c);
        if (cd) ws.setColumn(c, { ...cd });
      }
    }
    // Copy page setup and other properties
    if (src.pageSetup) ws.pageSetup = { ...src.pageSetup };
    if (src.printArea) ws.printArea = src.printArea;
    return ws;
  }

  /** Custom table styles for tables */
  private _customTableStyles: Map<string, {
    headerRow?: import('./types.js').CellStyle;
    dataRow1?: import('./types.js').CellStyle;
    dataRow2?: import('./types.js').CellStyle;
    totalRow?: import('./types.js').CellStyle;
  }> = new Map();

  /** Register a custom table style that can be referenced by tables. */
  registerTableStyle(name: string, def: {
    headerRow?: import('./types.js').CellStyle;
    dataRow1?: import('./types.js').CellStyle;
    dataRow2?: import('./types.js').CellStyle;
    totalRow?: import('./types.js').CellStyle;
  }): this {
    this._customTableStyles.set(name, def);
    return this;
  }

  // ─── Named ranges ──────────────────────────────────────────────────────────

  addNamedRange(nr: NamedRange): this {
    this.namedRanges.push(nr);
    return this;
  }

  getNamedRanges(): readonly NamedRange[] {
    return this.namedRanges;
  }

  getNamedRange(name: string): NamedRange | undefined {
    return this.namedRanges.find(nr => nr.name === name);
  }

  removeNamedRange(name: string): this {
    this.namedRanges = this.namedRanges.filter(nr => nr.name !== name);
    return this;
  }

  // ─── Connections ──────────────────────────────────────────────────────────

  addConnection(conn: Connection): this {
    this.connections.push(conn);
    return this;
  }

  getConnections(): readonly Connection[] {
    return this.connections;
  }

  getConnection(name: string): Connection | undefined {
    return this.connections.find(c => c.name === name);
  }

  removeConnection(name: string): this {
    this.connections = this.connections.filter(c => c.name !== name);
    return this;
  }

  // ─── Power Query ────────────────────────────────────────────────────────

  getPowerQueries(): readonly PowerQuery[] {
    return this.powerQueries;
  }

  getPowerQuery(name: string): PowerQuery | undefined {
    return this.powerQueries.find(q => q.name === name);
  }

  // ─── Custom property helpers ───────────────────────────────────────────────

  getCustomProperty(name: string): CustomProperty | undefined {
    return this.customProperties.find(p => p.name === name);
  }

  setCustomProperty(name: string, value: CustomProperty['value']): this {
    const idx = this.customProperties.findIndex(p => p.name === name);
    if (idx >= 0) this.customProperties[idx] = { name, value };
    else          this.customProperties.push({ name, value });
    return this;
  }

  removeCustomProperty(name: string): this {
    this.customProperties = this.customProperties.filter(p => p.name !== name);
    return this;
  }

  // ─── External Links ────────────────────────────────────────────────────────

  addExternalLink(link: ExternalLink): this {
    this.externalLinks.push(link);
    return this;
  }

  getExternalLinks(): readonly ExternalLink[] {
    return this.externalLinks;
  }

  // ─── Custom Pivot Styles ───────────────────────────────────────────────────

  registerPivotStyle(style: CustomPivotStyle): this {
    this.customPivotStyles.push(style);
    return this;
  }

  // ─── Pivot Slicers ─────────────────────────────────────────────────────────

  addPivotSlicer(slicer: PivotSlicer): this {
    this.pivotSlicers.push(slicer);
    return this;
  }

  getPivotSlicers(): readonly PivotSlicer[] {
    return this.pivotSlicers;
  }

  // ─── Build ─────────────────────────────────────────────────────────────────

  /**
   * Build the final XLSX Uint8Array.
   *
   * • If loaded from an existing file → patch-only mode:
   *   – Sheets marked dirty (via markDirty()) are re-serialised
   *   – All unknown parts (pivot tables, VBA, drawings, macros…) are kept verbatim
   *   – Properties are patched on top of the original XML
   *
   * • If created from scratch → full build mode
   */
  async build(): Promise<Uint8Array> {
    this._syncLegacyProperties();
    return this._readResult ? this._buildPatched() : this._buildFresh();
  }

  // ── Patch-mode ────────────────────────────────────────────────────────────

  private async _buildPatched(): Promise<Uint8Array> {
    const rr = this._readResult!;
    const entries: ZipEntry[] = [];
    const hasDirty = this._dirtySheets.size > 0;

    const styles = new StyleRegistry();
    const shared = new SharedStrings();

    // When any sheet is dirty we must re-serialise ALL sheets so that every
    // sheet uses the same fresh style-registry / shared-strings indices.
    // When nothing is dirty we preserve the original styles & shared strings.
    const sheetXmls = new Map<number, string>();
    if (hasDirty) {
      // Preserve original dxf entries so table dataDxfId references remain valid.
      // Extract raw <dxf>...</dxf> inner content from the original styles XML.
      const dxfRe = /<dxf>([\s\S]*?)<\/dxf>|<dxf\/>/g;
      const rawDxfs: string[] = [];
      let m: RegExpExecArray | null;
      while ((m = dxfRe.exec(rr.stylesXml)) !== null) rawDxfs.push(m[1] ?? '');
      if (rawDxfs.length) styles.prependRawDxfs(rawDxfs);

      for (let i = 0; i < this.sheets.length; i++) {
        sheetXmls.set(i, this.sheets[i].toXml(styles, shared));
      }
    }

    // ── Core properties ────────────────────────────────────────────────────
    entries.push({
      name: 'docProps/core.xml',
      data: strToBytes(buildCoreXml({ ...rr.core, ...this.coreProperties, modified: new Date() })),
    });

    // ── Extended properties ────────────────────────────────────────────────
    entries.push({
      name: 'docProps/app.xml',
      data: strToBytes(buildAppXml({
        ...rr.extended,
        ...this.extendedProperties,
        titlesOfParts: this.sheets.map(s => s.name),
        headingPairs:  this._headingPairs(),
      }, rr.extendedUnknownRaw)),
    });

    // ── Custom properties ──────────────────────────────────────────────────
    const customProps = this.customProperties.length > 0
      ? this.customProperties
      : (rr.hasCustomProps ? rr.custom : null);

    if (customProps && customProps.length > 0) {
      entries.push({ name: 'docProps/custom.xml', data: strToBytes(buildCustomXml(customProps)) });
    }

    // ── Workbook XML (patch sheet names) ───────────────────────────────────
    entries.push({ name: 'xl/workbook.xml', data: strToBytes(this._patchWorkbookXml(rr.workbookXml)) });

    // ── Connections ─────────────────────────────────────────────────────────
    const connectionsXml = this._connectionsXml(rr.connectionsXml);
    if (connectionsXml) {
      entries.push({ name: 'xl/connections.xml', data: strToBytes(connectionsXml) });
    }

    // ── Styles & shared strings ────────────────────────────────────────────
    if (hasDirty) {
      // All sheets re-serialised → use fresh registries
      entries.push({ name: 'xl/styles.xml',        data: strToBytes(styles.toXml()) });
      entries.push({ name: 'xl/sharedStrings.xml', data: strToBytes(shared.toXml()) });
    } else {
      // No sheets modified → preserve originals so indices remain valid
      entries.push({ name: 'xl/styles.xml',        data: strToBytes(rr.stylesXml) });
      entries.push({ name: 'xl/sharedStrings.xml', data: strToBytes(rr.sharedXml) });
    }

    // ── Sheets ────────────────────────────────────────────────────────────
    for (let i = 0; i < this.sheets.length; i++) {
      const ws = this.sheets[i];
      const folder = ws._isChartSheet ? 'chartsheets' : ws._isDialogSheet ? 'dialogsheets' : 'worksheets';
      const path = `xl/${folder}/sheet${i + 1}.xml`;
      if (hasDirty) {
        if (ws._isChartSheet) {
          entries.push({ name: path, data: strToBytes(ws.toChartSheetXml()) });
        } else if (ws._isDialogSheet) {
          entries.push({ name: path, data: strToBytes(ws.toDialogSheetXml(styles, shared)) });
        } else {
          entries.push({ name: path, data: strToBytes(sheetXmls.get(i) ?? '') });
        }
      } else {
        // Preserve original sheet verbatim (unknownParts are already in originalXml)
        entries.push({ name: path, data: strToBytes(rr.sheets[i]?.originalXml ?? '') });
      }
    }

    // ── Table XMLs — preserve originals verbatim, only regenerate truly new tables ──
    const allTablePaths = new Set<string>();
    for (let i = 0; i < this.sheets.length; i++) {
      const ws = this.sheets[i];
      const tables = ws.getTables();
      const paths = rr.sheets[i]?.tablePaths ?? [];
      const xmls = rr.sheets[i]?.tableXmls ?? [];
      for (let j = 0; j < tables.length; j++) {
        const tblPath = paths[j];
        if (tblPath) {
          allTablePaths.add(tblPath);
          if (j < xmls.length) {
            // Preserve original table XML — update only the ref attribute if it changed
            let xml = xmls[j];
            const origRefMatch = xml.match(/\bref="([^"]+)"/);
            if (origRefMatch && origRefMatch[1] !== tables[j].ref) {
              xml = xml.replace(`ref="${origRefMatch[1]}"`, `ref="${tables[j].ref}"`);
            }
            entries.push({ name: tblPath, data: strToBytes(xml) });
          } else {
            // New table without original XML — generate from scratch
            const idMatch = tblPath.match(/table(\d+)\.xml$/);
            const tableId = idMatch ? parseInt(idMatch[1], 10) : j + 1;
            entries.push({ name: tblPath, data: strToBytes(buildTableXml(tables[j], tableId)) });
          }
        }
      }
    }

    // ── Unknown parts — verbatim ──────────────────────────────────────────
    for (const [path, data] of rr.unknownParts) {
      // Skip table files already emitted above
      if (allTablePaths.has(path)) continue;
      // Skip vbaProject.bin if we're managing VBA ourselves
      if (path === 'xl/vbaProject.bin' && this.vbaProject) continue;
      // Skip rels files already emitted from allRels
      if (rr.allRels.has(path)) continue;
      // Drop calcChain.xml when sheets are dirty (Excel rebuilds it)
      if (hasDirty && path === 'xl/calcChain.xml') continue;
      entries.push({ name: path, data });
    }

    // ── VBA project ─────────────────────────────────────────────────────
    if (this.vbaProject) {
      this._ensureVbaSheetModules();
      entries.push({ name: 'xl/vbaProject.bin', data: this.vbaProject.build() });
    }

    // ── Rels ──────────────────────────────────────────────────────────────
    entries.push({ name: '_rels/.rels',                data: strToBytes(this._buildRootRels(customProps != null && customProps.length > 0)) });
    entries.push({ name: 'xl/_rels/workbook.xml.rels', data: strToBytes(this._buildWorkbookRels(rr, hasDirty)) });

    for (const [relPath, relMap] of rr.allRels) {
      if (relPath === 'xl/_rels/workbook.xml.rels' || relPath === '_rels/.rels') continue;
      entries.push({ name: relPath, data: strToBytes(this._relsToXml(relMap)) });
    }

    // ── Content types ──────────────────────────────────────────────────────
    entries.push({
      name: '[Content_Types].xml',
      data: strToBytes(this._patchContentTypes(rr.contentTypesXml, customProps != null && customProps.length > 0, hasDirty)),
    });

    return buildZip(entries, { level: this.compressionLevel });
  }

  // ── Fresh build ───────────────────────────────────────────────────────────

  private async _buildFresh(): Promise<Uint8Array> {
    const styles = new StyleRegistry();
    const shared = new SharedStrings();
    const entries: ZipEntry[] = [];

    // Register custom table styles
    for (const [name, def] of this._customTableStyles) {
      styles.registerTableStyle(name, def);
    }

    let globalRId = 1;
    for (const ws of this.sheets) ws.rId = `rId${globalRId++}`;

    const allImages:  Array<{ ws: Worksheet; img: Image; ext: string; idx: number }> = [];
    const allCharts:  Array<{ ws: Worksheet; chartIdx: number; globalIdx: number }> = [];
    const allTables:  Array<{ ws: Worksheet; tableIdx: number; globalTableId: number }> = [];
    const allPivotTables: Array<{ ws: Worksheet; pt: PivotTable; pivotIdx: number; cacheId: number; pivotRId: string; cacheRId: string }> = [];
    const sheetImageRIds  = new Map<Worksheet, string[]>();
    const sheetChartRIds  = new Map<Worksheet, string[]>();
    const sheetTableRIds  = new Map<Worksheet, string[]>();
    const sheetPivotRIds  = new Map<Worksheet, string[]>();
    let imgCtr = 1, chartCtr = 1, tableCtr = 1, vmlCtr = 1, pivotCtr = 1, pivotCacheIdCtr = 0, ctrlPropGlobal = 0;

    for (const ws of this.sheets) {
      const imgs = ws.getImages() as Image[];
      const charts = ws.getCharts();
      const tables = ws.getTables();
      const imgRIds: string[] = [], chartRIds: string[] = [], tblRIds: string[] = [];

      if (imgs.length || charts.length || ws.getShapes().length || ws.getWordArt().length || ws.getMathEquations().length || ws.getTableSlicers().length) ws.drawingRId = `rId${globalRId++}`;
      const controls = ws.getFormControls();
      // legacyDrawing needed for comments OR form controls (they share VML)
      if (ws.getComments().length || controls.length) ws.legacyDrawingRId = `rId${globalRId++}`;

      for (const img of imgs)    { const r = `rId${globalRId++}`; imgRIds.push(r);   allImages.push({ ws, img, ext: imageExt(img.format), idx: imgCtr++ }); }
      for (let i=0;i<charts.length;i++) { const r = `rId${globalRId++}`; chartRIds.push(r); allCharts.push({ ws, chartIdx: i, globalIdx: chartCtr++ }); }
      for (let i=0;i<tables.length;i++) { const r = `rId${globalRId++}`; tblRIds.push(r);   allTables.push({ ws, tableIdx: i, globalTableId: tableCtr++ }); }

      // Allocate ctrlProp rIds for each form control
      const ctrlPropRIds: string[] = [];
      for (let ci = 0; ci < controls.length; ci++) ctrlPropRIds.push(`rId${globalRId++}`);
      ws.ctrlPropRIds = ctrlPropRIds;

      sheetImageRIds.set(ws, imgRIds);
      sheetChartRIds.set(ws, chartRIds);
      sheetTableRIds.set(ws, tblRIds);
      ws.tableRIds = tblRIds;

      const ptRIds: string[] = [];
      for (const pt of ws.getPivotTables()) {
        const pivotRId = `rId${globalRId++}`;
        const cacheRId = `rId${globalRId++}`;
        ptRIds.push(pivotRId);
        allPivotTables.push({ ws, pt, pivotIdx: pivotCtr++, cacheId: pivotCacheIdCtr++, pivotRId, cacheRId });
      }
      sheetPivotRIds.set(ws, ptRIds);
    }

    // ── Cell images (in-cell pictures via richData) ──────────────────────
    const allCellImages: Array<{ img: CellImage; ext: string; idx: number }> = [];
    let cellImgCtr = imgCtr;   // continue numbering from floating images
    let vmCounter = 1;         // 1-based vm index for metadata
    for (const ws of this.sheets) {
      ws._cellImageVm = new Map();
      for (const ci of ws.getCellImages()) {
        const ext = imageExt(ci.format);
        allCellImages.push({ img: ci, ext, idx: cellImgCtr++ });
        ws._cellImageVm.set(ci.cell, vmCounter++);
        // Ensure the cell exists in the cells map so it gets emitted with vm attr
        ws.getCellByRef(ci.cell);
      }
    }
    const hasCellImages = allCellImages.length > 0;

    // ── Slicer info collection ───────────────────────────────────────────
    type SlicerCacheEntry = {
      name: string; sourceName: string; type: 'table' | 'pivot';
      rId: string; idx: number;
      tableId?: number; columnIndex?: number; sortOrder?: string;
      pivotTableName?: string; pivotCacheId?: number; tabId?: number; items?: string[];
    };
    type SheetSlicerEntry = {
      tableSlicers: Array<{ slicer: import('../core/types.js').TableSlicer; tableId: number; columnIndex: number }>;
      pivotSlicers: Array<{ slicer: PivotSlicer; pivotCacheId: number }>;
      slicerDefRId: string; slicerDefIdx: number;
    };
    const sheetSlicerMap = new Map<Worksheet, SheetSlicerEntry>();
    const allSlicerCaches: SlicerCacheEntry[] = [];
    let slicerDefCtr = 0, slicerCacheCtr = 0;

    // Collect table slicers per sheet
    for (const ws of this.sheets) {
      const tSlicers = ws.getTableSlicers();
      if (tSlicers.length) {
        sheetSlicerMap.set(ws, { tableSlicers: [], pivotSlicers: [], slicerDefRId: '', slicerDefIdx: 0 });
      }
    }

    // Map pivot slicers to sheets via pivot table name
    for (const ps of this.pivotSlicers) {
      for (const ws of this.sheets) {
        if (ws.getPivotTables().some(p => p.name === ps.pivotTableName)) {
          if (!sheetSlicerMap.has(ws)) {
            sheetSlicerMap.set(ws, { tableSlicers: [], pivotSlicers: [], slicerDefRId: '', slicerDefIdx: 0 });
          }
          break;
        }
      }
    }

    // Allocate rIds and populate slicer cache info
    for (const [ws, info] of sheetSlicerMap) {
      if (!ws.drawingRId) ws.drawingRId = `rId${globalRId++}`;
      info.slicerDefRId = `rId${globalRId++}`;
      info.slicerDefIdx = ++slicerDefCtr;
      ws.slicerRId = info.slicerDefRId;

      const drawingInfo: Array<{ name: string; cell?: string }> = [];

      // Table slicers
      for (const s of ws.getTableSlicers()) {
        const table = ws.getTables().find(t => t.name === s.tableName);
        const tableEntry = allTables.find(t => t.ws === ws && ws.getTables()[t.tableIdx] === table);
        const tableId = tableEntry?.globalTableId ?? 1;
        const columnIndex = table ? (table.columns?.findIndex(c => c.name === s.columnName) ?? 0) + 1 : 1;
        info.tableSlicers.push({ slicer: s, tableId, columnIndex });
        drawingInfo.push({ name: s.name, cell: s.cell });
        allSlicerCaches.push({
          name: s.name + '_cache', sourceName: s.columnName, type: 'table',
          rId: `rId${globalRId++}`, idx: ++slicerCacheCtr,
          tableId, columnIndex, sortOrder: s.sortOrder ?? 'ascending',
        });
      }

      // Pivot slicers on this sheet
      for (const ps of this.pivotSlicers) {
        const pt = ws.getPivotTables().find(p => p.name === ps.pivotTableName);
        if (!pt) continue;
        const ptEntry = allPivotTables.find(p => p.ws === ws && p.pt === pt);
        const sheetIdx = this.sheets.indexOf(ws) + 1;

        // Get unique values for the slicer field from source data
        let items: string[] = [];
        const sourceWs = this.sheets.find(s => s.name === pt.sourceSheet);
        if (sourceWs) {
          const sourceData = sourceWs.readRange(pt.sourceRef);
          const headers = (sourceData[0] ?? []).map(v => String(v ?? ''));
          const fieldIdx = headers.indexOf(ps.fieldName);
          if (fieldIdx >= 0) {
            const uniqueSet = new Set<string>();
            for (let r = 1; r < sourceData.length; r++) uniqueSet.add(String(sourceData[r][fieldIdx] ?? ''));
            items = [...uniqueSet];
          }
        }
        info.pivotSlicers.push({ slicer: ps, pivotCacheId: ptEntry?.cacheId ?? 0 });
        drawingInfo.push({ name: ps.name, cell: ps.cell });
        allSlicerCaches.push({
          name: ps.name + '_cache', sourceName: ps.fieldName, type: 'pivot',
          rId: `rId${globalRId++}`, idx: ++slicerCacheCtr,
          pivotTableName: ps.pivotTableName, pivotCacheId: ptEntry?.cacheId ?? 0,
          tabId: sheetIdx, items,
        });
      }

      ws._slicerDrawingInfo = drawingInfo;
    }
    const hasSlicers = sheetSlicerMap.size > 0;

    const hasCustom = this.customProperties.length > 0;
    const hasVba    = !!this.vbaProject;

    // Content types
    const imgCTs = new Set<string>();
    for (const { ext } of [...allImages, ...allCellImages]) {
      const ct = imageContentType(ext);
      imgCTs.add(`<Default Extension="${ext}" ContentType="${ct}"/>`);
    }
    const sheetsWithComments = this.sheets.filter(ws => ws.getComments().length);
    const sheetsWithVml = this.sheets.filter(ws => ws.getComments().length || ws.getFormControls().length);
    const vmlCT  = sheetsWithVml.length ? '<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>' : '';
    let vmlIdx = 0;
    const commentsCTs = sheetsWithComments.map(() =>
      `<Override PartName="/xl/comments${++vmlIdx}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>`
    ).join('');
    // ctrlProp content types — skip dialog sheets (they use pure VML)
    let ctrlPropCtr = 0;
    const ctrlPropCTs: string[] = [];
    for (const ws of this.sheets) {
      if (ws._isDialogSheet) continue;
      for (let ci = 0; ci < ws.getFormControls().length; ci++) {
        ctrlPropCTs.push(`<Override PartName="/xl/ctrlProps/ctrlProp${++ctrlPropCtr}.xml" ContentType="application/vnd.ms-excel.controlproperties+xml"/>`);
      }
    }
    entries.push({ name: '[Content_Types].xml', data: strToBytes(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
${vmlCT}
${[...imgCTs].join('')}
<Override PartName="/xl/workbook.xml" ContentType="${hasVba ? 'application/vnd.ms-excel.sheet.macroEnabled.main+xml' : this.isTemplate ? 'application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml' : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'}"/>
${hasVba ? '<Override PartName="/xl/vbaProject.bin" ContentType="application/vnd.ms-office.vbaProject"/>' : ''}
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
<Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
${this.sheets.filter(ws => !ws._isChartSheet && !ws._isDialogSheet).map(ws => { const idx = this.sheets.indexOf(ws); return `<Override PartName="/xl/worksheets/sheet${idx+1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`; }).join('')}
${this.sheets.filter(ws => ws._isChartSheet).map(ws => { const idx = this.sheets.indexOf(ws); return `<Override PartName="/xl/chartsheets/sheet${idx+1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml"/>`; }).join('')}
${this.sheets.filter(ws => ws._isDialogSheet).map(ws => { const idx = this.sheets.indexOf(ws); return `<Override PartName="/xl/dialogsheets/sheet${idx+1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.dialogsheet+xml"/>`; }).join('')}
${this.sheets.filter(ws=>ws.drawingRId).map((_,i) => `<Override PartName="/xl/drawings/drawing${i+1}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>`).join('')}
${allCharts.map(({globalIdx}) => `<Override PartName="/xl/charts/chart${globalIdx}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>`).join('')}
${allTables.map(({globalTableId}) => `<Override PartName="/xl/tables/table${globalTableId}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>`).join('')}
${allPivotTables.map(p => `<Override PartName="/xl/pivotTables/pivotTable${p.pivotIdx}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml"/>`).join('\n')}
${allPivotTables.map(p => `<Override PartName="/xl/pivotCache/pivotCacheDefinition${p.pivotIdx}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml"/>`).join('\n')}
${allPivotTables.map(p => `<Override PartName="/xl/pivotCache/pivotCacheRecords${p.pivotIdx}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml"/>`).join('\n')}
${commentsCTs}
${ctrlPropCTs.join('')}
<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
${hasCustom ? '<Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/>' : ''}
${this.connections.length ? '<Override PartName="/xl/connections.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml"/>' : ''}
${this.externalLinks.map((_,i) => `<Override PartName="/xl/externalLinks/externalLink${i+1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"/>`).join('\n')}
${[...sheetSlicerMap.values()].map(info => `<Override PartName="/xl/slicers/slicer${info.slicerDefIdx}.xml" ContentType="application/vnd.ms-excel.slicer+xml"/>`).join('\n')}
${allSlicerCaches.map(sc => `<Override PartName="/xl/slicerCaches/slicerCache${sc.idx}.xml" ContentType="application/vnd.ms-excel.slicerCache+xml"/>`).join('\n')}
${this.sheets.flatMap(ws => ws.getQueryTables()).map((_,i) => `<Override PartName="/xl/queryTables/queryTable${i+1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.queryTable+xml"/>`).join('\n')}
${hasCellImages ? `<Override PartName="/xl/metadata.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"/>
<Override PartName="/xl/richData/rdrichvalue.xml" ContentType="application/vnd.ms-excel.rdrichvalue+xml"/>
<Override PartName="/xl/richData/rdRichValueStructure.xml" ContentType="application/vnd.ms-excel.rdrichvaluestructure+xml"/>
<Override PartName="/xl/richData/richValueRel.xml" ContentType="application/vnd.ms-excel.richvaluerel+xml"/>
<Override PartName="/xl/richData/rdRichValueTypes.xml" ContentType="application/vnd.ms-excel.rdrichvaluetypes+xml"/>
<Override PartName="/xl/richData/rdarray.xml" ContentType="application/vnd.ms-excel.rdarray+xml"/>` : ''}
</Types>`) });

    entries.push({ name: '_rels/.rels', data: strToBytes(this._buildRootRels(hasCustom)) });

    entries.push({ name: 'xl/_rels/workbook.xml.rels', data: strToBytes(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
${this.sheets.map((ws,i) => {
  const type = ws._isChartSheet ? 'chartsheet' : ws._isDialogSheet ? 'dialogsheet' : 'worksheet';
  const folder = ws._isChartSheet ? 'chartsheets' : ws._isDialogSheet ? 'dialogsheets' : 'worksheets';
  return `<Relationship Id="${ws.rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${type}" Target="${folder}/sheet${i+1}.xml"/>`;
}).join('')}
<Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rIdShared" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
<Relationship Id="rIdTheme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
${allPivotTables.map(p => `<Relationship Id="${p.cacheRId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="pivotCache/pivotCacheDefinition${p.pivotIdx}.xml"/>`).join('\n')}
${allSlicerCaches.map(sc => `<Relationship Id="${sc.rId}" Type="http://schemas.microsoft.com/office/2007/relationships/slicerCache" Target="slicerCaches/slicerCache${sc.idx}.xml"/>`).join('\n')}
${hasVba ? '<Relationship Id="rIdVBA" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/>' : ''}
${this.connections.length ? '<Relationship Id="rIdConns" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/connections" Target="connections.xml"/>' : ''}
${this.externalLinks.map((_,i) => `<Relationship Id="rIdExtLink${i+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink" Target="externalLinks/externalLink${i+1}.xml"/>`).join('\n')}
${hasCellImages ? '<Relationship Id="rIdMeta" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata" Target="metadata.xml"/>' : ''}
${hasCellImages ? '<Relationship Id="rIdRichValueRel" Type="http://schemas.microsoft.com/office/2022/10/relationships/richValueRel" Target="richData/richValueRel.xml"/>' : ''}
${hasCellImages ? '<Relationship Id="rIdRichValue" Type="http://schemas.microsoft.com/office/2017/06/relationships/rdRichValue" Target="richData/rdrichvalue.xml"/>' : ''}
${hasCellImages ? '<Relationship Id="rIdRichValueStruct" Type="http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueStructure" Target="richData/rdRichValueStructure.xml"/>' : ''}
${hasCellImages ? '<Relationship Id="rIdRichValueTypes" Type="http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueTypes" Target="richData/rdRichValueTypes.xml"/>' : ''}
${hasCellImages ? '<Relationship Id="rIdRdArray" Type="http://schemas.microsoft.com/office/2017/06/relationships/rdArray" Target="richData/rdarray.xml"/>' : ''}
</Relationships>`) });

    // ── VBA project binary ──────────────────────────────────────────────
    if (hasVba) {
      this._ensureVbaSheetModules();
      entries.push({ name: 'xl/vbaProject.bin', data: this.vbaProject!.build() });
    }

    const wbPrAttrs = [
      this.properties.date1904 ? 'date1904="1"' : '',
      hasVba ? 'codeName="ThisWorkbook"' : '',
    ].filter(Boolean).join(' ');
    const date1904 = `<workbookPr${wbPrAttrs ? ' ' + wbPrAttrs : ''}/>`;
    const namedRangesXml = this._definedNamesXml();
    const pivotCachesXml = allPivotTables.length
      ? `<pivotCaches>${allPivotTables.map(p => `<pivotCache cacheId="${p.cacheId}" r:id="${p.cacheRId}"/>`).join('')}</pivotCaches>`
      : '';

    entries.push({ name: 'xl/workbook.xml', data: strToBytes(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
${date1904}
<bookViews><workbookView xWindow="0" yWindow="0" windowWidth="14400" windowHeight="8260"/></bookViews>
<sheets>${this.sheets.map((ws,i) => `<sheet name="${escapeXml(ws.name)}" sheetId="${i+1}" r:id="${ws.rId}"${ws.options?.state && ws.options.state !== 'visible' ? ` state="${ws.options.state}"` : ''}/>`).join('')}</sheets>
${namedRangesXml}
<calcPr calcId="191028"/>
${pivotCachesXml}
${hasSlicers ? `<extLst><ext uri="{BBE1A952-AA13-448e-AADC-164F8A28A991}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:slicerCaches>${allSlicerCaches.map(sc => `<x14:slicerCache r:id="${sc.rId}"/>`).join('')}</x14:slicerCaches></ext></extLst>` : ''}
</workbook>`) });

    // ── Connections ─────────────────────────────────────────────────────────
    if (this.connections.length) {
      entries.push({ name: 'xl/connections.xml', data: strToBytes(this._connectionsXml()) });
    }

    // Per-sheet
    for (let i = 0; i < this.sheets.length; i++) {
      const ws       = this.sheets[i];
      const imgRIds  = sheetImageRIds.get(ws) ?? [];
      const cRIds    = sheetChartRIds.get(ws) ?? [];
      const tblEntries = allTables.filter(t => t.ws === ws);
      const tblRIds_ = sheetTableRIds.get(ws) ?? [];

      // Determine sheet path based on type
      const sheetFolder = ws._isChartSheet ? 'chartsheets' : ws._isDialogSheet ? 'dialogsheets' : 'worksheets';
      const sheetPath = `xl/${sheetFolder}/sheet${i+1}.xml`;

      // Generate appropriate XML based on sheet type
      if (ws._isChartSheet) {
        entries.push({ name: sheetPath, data: strToBytes(ws.toChartSheetXml()) });
      } else if (ws._isDialogSheet) {
        entries.push({ name: sheetPath, data: strToBytes(ws.toDialogSheetXml(styles, shared)) });
      } else {
        entries.push({ name: sheetPath, data: strToBytes(ws.toXml(styles, shared)) });
      }

      const wsRels: string[] = [];
      if (ws.drawingRId) {
        const dIdx = this.sheets.filter((s,j)=>j<=i&&s.drawingRId).length;
        wsRels.push(`<Relationship Id="${ws.drawingRId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing${dIdx}.xml"/>`);
      }
      for (let j=0;j<(ws.getImages() as Image[]).length;j++) {
        const g = allImages.filter(x=>x.ws===ws)[j];
        if (g) wsRels.push(`<Relationship Id="${imgRIds[j]}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image${g.idx}.${g.ext}"/>`);
      }
      for (let j=0;j<tblEntries.length;j++) {
        wsRels.push(`<Relationship Id="${tblRIds_[j]}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table${tblEntries[j].globalTableId}.xml"/>`);
      }
      const ptRIds_ = sheetPivotRIds.get(ws) ?? [];
      const ptEntries = allPivotTables.filter(p => p.ws === ws);
      for (let j = 0; j < ptEntries.length; j++) {
        wsRels.push(`<Relationship Id="${ptRIds_[j]}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable" Target="../pivotTables/pivotTable${ptEntries[j].pivotIdx}.xml"/>`);
      }
      // Slicer definition relationship
      const slicerInfo = sheetSlicerMap.get(ws);
      if (slicerInfo) {
        wsRels.push(`<Relationship Id="${slicerInfo.slicerDefRId}" Type="http://schemas.microsoft.com/office/2007/relationships/slicer" Target="../slicers/slicer${slicerInfo.slicerDefIdx}.xml"/>`);
      }
      const sheetComments = ws.getComments();
      const sheetControls = ws.getFormControls();
      if ((sheetComments.length || sheetControls.length) && ws.legacyDrawingRId) {
        const vIdx = vmlCtr++;
        wsRels.push(`<Relationship Id="${ws.legacyDrawingRId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/vmlDrawing${vIdx}.vml"/>`);

        if (sheetComments.length) {
          const commRId = `rId${globalRId++}`;
          wsRels.push(`<Relationship Id="${commRId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments${vIdx}.xml"/>`);
          entries.push({ name: `xl/comments${vIdx}.xml`, data: strToBytes(this._buildCommentsXml(sheetComments)) });
        }

        // Build comment VML shapes
        const commentShapes = sheetComments.map(({ row, col }, ci) => {
          const left  = (col + 1) * 64;
          const top   = (row - 1) * 20;
          const sid = 1025 + i * 1000 + ci;
          return `<v:shape id="_x0000_s${sid}" type="#_x0000_t202" style="position:absolute;margin-left:${left}pt;margin-top:${top}pt;width:108pt;height:59.25pt;z-index:${ci + 1};visibility:hidden" fillcolor="#ffffe1" o:insetmode="auto">
<v:fill color2="#ffffe1"/>
<v:shadow color="black" obscured="t"/>
<v:path o:connecttype="none"/>
<v:textbox style="mso-direction-alt:auto"><div style="text-align:left"/></v:textbox>
<x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/><x:Anchor>${col + 1},15,${row - 1},10,${col + 3},15,${row + 4},4</x:Anchor><x:AutoFill>False</x:AutoFill><x:Row>${row - 1}</x:Row><x:Column>${col - 1}</x:Column></x:ClientData>
</v:shape>`;
        });

        // Build form control VML shapes — IDs must match <control shapeId> in Worksheet._formControlsXml()
        const ctrlBaseId = 1025 + ws.sheetIndex * 1000 + sheetComments.length;
        const controlShapes = sheetControls.map((ctrl, ci) =>
          buildFormControlVmlShape(ctrl, ctrlBaseId + ci)
        );

        entries.push({ name: `xl/drawings/vmlDrawing${vIdx}.vml`, data: strToBytes(buildVmlWithControls(commentShapes, controlShapes)) });

        // ctrlProp rels and files — skip for dialog sheets (dialog controls are purely VML-based)
        if (!ws._isDialogSheet) {
          const ctrlRIds = ws.ctrlPropRIds;
          for (let ci = 0; ci < sheetControls.length; ci++) {
            const ctrlPropIdx = ++ctrlPropGlobal;
            wsRels.push(`<Relationship Id="${ctrlRIds[ci]}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/ctrlProp" Target="../ctrlProps/ctrlProp${ctrlPropIdx}.xml"/>`);
            entries.push({ name: `xl/ctrlProps/ctrlProp${ctrlPropIdx}.xml`, data: strToBytes(buildCtrlPropXml(sheetControls[ci])) });
          }
        }
      }
      if (wsRels.length) {
        entries.push({ name: `xl/${sheetFolder}/_rels/sheet${i+1}.xml.rels`, data: strToBytes(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
${wsRels.join('\n')}
</Relationships>`) });
      }

      if (ws.drawingRId) {
        const dIdx = this.sheets.filter((s,j)=>j<=i&&s.drawingRId).length;
        entries.push({ name: `xl/drawings/drawing${dIdx}.xml`, data: strToBytes(ws.toDrawingXml(imgRIds, cRIds)) });
        const dRels: string[] = [];
        for (let j=0;j<(ws.getImages() as Image[]).length;j++) {
          const g = allImages.filter(x=>x.ws===ws)[j];
          if (g) dRels.push(`<Relationship Id="${imgRIds[j]}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image${g.idx}.${g.ext}"/>`);
        }
        for (let j=0;j<ws.getCharts().length;j++) {
          const g = allCharts.filter(x=>x.ws===ws)[j];
          if (g) dRels.push(`<Relationship Id="${cRIds[j]}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart${g.globalIdx}.xml"/>`);
        }
        if (dRels.length) {
          entries.push({ name: `xl/drawings/_rels/drawing${dIdx}.xml.rels`, data: strToBytes(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
${dRels.join('\n')}
</Relationships>`) });
        }
      }
    }

    for (const { img, ext, idx } of allImages) {
      entries.push({ name: `xl/media/image${idx}.${ext}`, data: typeof img.data === 'string' ? base64ToBytes(img.data) : img.data });
    }

    // ── Cell image media + richData files ────────────────────────────────
    if (hasCellImages) {
      const cellImgRIds: string[] = [];
      const cellImgRels: string[] = [];
      for (let ci = 0; ci < allCellImages.length; ci++) {
        const { img, ext, idx } = allCellImages[ci];
        const rId = `rId${ci + 1}`;
        cellImgRIds.push(rId);
        cellImgRels.push(`<Relationship Id="${rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="/xl/media/image${idx}.${ext}"/>`);
        entries.push({ name: `xl/media/image${idx}.${ext}`, data: typeof img.data === 'string' ? base64ToBytes(img.data) : img.data });
      }

      entries.push({ name: 'xl/metadata.xml',                        data: strToBytes(buildMetadataXml(allCellImages.length)) });
      entries.push({ name: 'xl/richData/rdrichvalue.xml',             data: strToBytes(buildRichValueXml(allCellImages.length)) });
      entries.push({ name: 'xl/richData/richValueRel.xml',            data: strToBytes(buildRichValueRelXml(cellImgRIds)) });
      entries.push({ name: 'xl/richData/rdRichValueStructure.xml',    data: strToBytes(buildRichValueStructureXml()) });
      entries.push({ name: 'xl/richData/rdRichValueTypes.xml',        data: strToBytes(buildRichValueTypesXml()) });
      entries.push({ name: 'xl/richData/rdarray.xml',                data: strToBytes(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><arrayData xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2" count="0"></arrayData>`) });
      entries.push({ name: 'xl/richData/_rels/richValueRel.xml.rels', data: strToBytes(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
${cellImgRels.join('\n')}
</Relationships>`) });
    }
    for (const { ws, chartIdx, globalIdx } of allCharts) {
      entries.push({ name: `xl/charts/chart${globalIdx}.xml`, data: strToBytes(buildChartXml(ws.getCharts()[chartIdx])) });
    }
    for (const { ws, tableIdx, globalTableId } of allTables) {
      entries.push({ name: `xl/tables/table${globalTableId}.xml`, data: strToBytes(buildTableXml(ws.getTables()[tableIdx], globalTableId)) });
    }

    for (const { ws, pt, pivotIdx, cacheId: cId } of allPivotTables) {
      const sourceWs   = this.sheets.find(s => s.name === pt.sourceSheet);
      const sourceData = sourceWs ? sourceWs.readRange(pt.sourceRef) : [[]];
      const { pivotTableXml, cacheDefXml, cacheRecordsXml } = buildPivotTableFiles(pt, sourceData, pivotIdx, cId);
      const wbRel = (type: string, target: string) =>
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${type}" Target="${target}"/>\n</Relationships>`;

      entries.push({ name: `xl/pivotTables/pivotTable${pivotIdx}.xml`,                               data: strToBytes(pivotTableXml) });
      entries.push({ name: `xl/pivotTables/_rels/pivotTable${pivotIdx}.xml.rels`,                    data: strToBytes(wbRel('pivotCacheDefinition', `../pivotCache/pivotCacheDefinition${pivotIdx}.xml`)) });
      entries.push({ name: `xl/pivotCache/pivotCacheDefinition${pivotIdx}.xml`,                      data: strToBytes(cacheDefXml) });
      entries.push({ name: `xl/pivotCache/_rels/pivotCacheDefinition${pivotIdx}.xml.rels`,           data: strToBytes(wbRel('pivotCacheRecords', `pivotCacheRecords${pivotIdx}.xml`)) });
      entries.push({ name: `xl/pivotCache/pivotCacheRecords${pivotIdx}.xml`,                         data: strToBytes(cacheRecordsXml) });
    }

    entries.push({ name: 'xl/styles.xml',        data: strToBytes(styles.toXml()) });
    entries.push({ name: 'xl/sharedStrings.xml', data: strToBytes(shared.toXml()) });

    // ── Theme ──────────────────────────────────────────────────────────────
    entries.push({ name: 'xl/theme/theme1.xml', data: strToBytes(this._buildThemeXml()) });

    // ── External Links ──────────────────────────────────────────────────────
    for (let i = 0; i < this.externalLinks.length; i++) {
      const link = this.externalLinks[i];
      const idx = i + 1;
      const sheetsXml = link.sheets.map(s => {
        const dNames = s.definedNames?.map(d =>
          `<definedName name="${escapeXml(d.name)}" refersTo="${escapeXml(d.ref)}"/>`
        ).join('') ?? '';
        return `<sheetName val="${escapeXml(s.name)}"/>${dNames ? `<sheetDataSet>${dNames}</sheetDataSet>` : ''}`;
      }).join('');
      entries.push({ name: `xl/externalLinks/externalLink${idx}.xml`, data: strToBytes(
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<externalBook r:id="rId1"><sheetNames>${sheetsXml}</sheetNames></externalBook>
</externalLink>`) });
      entries.push({ name: `xl/externalLinks/_rels/externalLink${idx}.xml.rels`, data: strToBytes(
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath" Target="${escapeXml(link.target)}" TargetMode="External"/>
</Relationships>`) });
    }

    // ── Slicers (per-sheet definitions + caches) ───────────────────────────
    for (const [ws, info] of sheetSlicerMap) {
      const allSheetSlicerItems: string[] = [];
      // Table slicer items
      for (const ts of info.tableSlicers) {
        const s = ts.slicer;
        allSheetSlicerItems.push(`<slicer name="${escapeXml(s.name)}" cache="${escapeXml(s.name + '_cache')}" caption="${escapeXml(s.caption ?? s.columnName)}" rowHeight="241300" columnCount="${s.columnCount ?? 1}" style="${s.style ?? 'SlicerStyleLight1'}"/>`);
      }
      // Pivot slicer items
      for (const ps of info.pivotSlicers) {
        const s = ps.slicer;
        allSheetSlicerItems.push(`<slicer name="${escapeXml(s.name)}" cache="${escapeXml(s.name + '_cache')}" caption="${escapeXml(s.caption ?? s.fieldName)}" rowHeight="241300" columnCount="${s.columnCount ?? 1}" style="${s.style ?? 'SlicerStyleLight1'}"/>`);
      }
      entries.push({ name: `xl/slicers/slicer${info.slicerDefIdx}.xml`, data: strToBytes(
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<slicers xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main" mc:Ignorable="x">
${allSheetSlicerItems.join('\n')}
</slicers>`) });
    }

    // Slicer caches (all types unified)
    for (const sc of allSlicerCaches) {
      let cacheBody: string;
      if (sc.type === 'table') {
        cacheBody = `<extLst><ext xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" uri="{2F2917AC-EB37-4324-AD4E-5DD8C200BD13}" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"><x15:tableSlicerCache tableId="${sc.tableId}" column="${sc.columnIndex}" sortOrder="${sc.sortOrder ?? 'ascending'}"/></ext></extLst>`;
      } else {
        const itemsXml = (sc.items ?? []).map((_, xi) => `<i x="${xi}" s="1"/>`).join('');
        cacheBody = `<pivotTables><pivotTable tabId="${sc.tabId}" name="${escapeXml(sc.pivotTableName ?? '')}"/></pivotTables>` +
          (sc.items?.length ? `<data><tabular pivotCacheId="${sc.pivotCacheId}"><items count="${sc.items.length}">${itemsXml}</items></tabular></data>` : '');
      }
      entries.push({ name: `xl/slicerCaches/slicerCache${sc.idx}.xml`, data: strToBytes(
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<slicerCacheDefinition xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" name="${escapeXml(sc.name)}" sourceName="${escapeXml(sc.sourceName)}">
${cacheBody}
</slicerCacheDefinition>`) });
    }

    // ── Query Tables ────────────────────────────────────────────────────────
    const allQueryTables = this.sheets.flatMap(ws => ws.getQueryTables());
    for (let i = 0; i < allQueryTables.length; i++) {
      const qt = allQueryTables[i];
      entries.push({ name: `xl/queryTables/queryTable${i + 1}.xml`, data: strToBytes(
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<queryTable xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="${escapeXml(qt.name)}" connectionId="${qt.connectionId}" autoFormatId="16" applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0" applyPatternFormats="0" applyAlignmentFormats="0" applyWidthHeightFormats="0">
<queryTableRefresh nextId="${(qt.columns?.length ?? 0) + 1}">
<queryTableFields count="${qt.columns?.length ?? 0}">
${(qt.columns ?? []).map((c,ci) => `<queryTableField id="${ci+1}" name="${escapeXml(c)}" tableColumnId="${ci+1}"/>`).join('\n')}
</queryTableFields>
</queryTableRefresh>
</queryTable>`) });
    }

    const cp = { ...this.coreProperties, created: this.coreProperties.created ?? new Date(), modified: new Date() };
    if (!cp.creator && this.properties.author) cp.creator = this.properties.author;
    entries.push({ name: 'docProps/core.xml', data: strToBytes(buildCoreXml(cp)) });

    entries.push({ name: 'docProps/app.xml', data: strToBytes(buildAppXml({
      ...this.extendedProperties,
      application: this.extendedProperties.application ?? 'ExcelForge',
      company:     this.extendedProperties.company ?? this.properties.company,
      titlesOfParts: this.sheets.map(s => s.name),
      headingPairs:  this._headingPairs(),
    })) });

    if (hasCustom) entries.push({ name: 'docProps/custom.xml', data: strToBytes(buildCustomXml(this.customProperties)) });

    return buildZip(entries, { level: this.compressionLevel });
  }

  // ─── Internal helpers ──────────────────────────────────────────────────────

  private _headingPairs(): Array<{ name: string; count: number }> {
    const normalCount = this.sheets.filter(ws => !ws._isChartSheet && !ws._isDialogSheet).length;
    const chartCount  = this.sheets.filter(ws => ws._isChartSheet).length;
    const dialogCount = this.sheets.filter(ws => ws._isDialogSheet).length;
    const pairs: Array<{ name: string; count: number }> = [];
    if (normalCount) pairs.push({ name: 'Worksheets', count: normalCount });
    if (chartCount)  pairs.push({ name: 'Charts', count: chartCount });
    if (dialogCount) pairs.push({ name: 'Dialogs', count: dialogCount });
    if (!pairs.length) pairs.push({ name: 'Worksheets', count: 0 });
    return pairs;
  }

  private _syncLegacyProperties(): void {
    const p = this.properties;
    if (p.title)          this.coreProperties.title          ??= p.title;
    if (p.author)         this.coreProperties.creator        ??= p.author;
    if (p.subject)        this.coreProperties.subject        ??= p.subject;
    if (p.description)    this.coreProperties.description    ??= p.description;
    if (p.keywords)       this.coreProperties.keywords       ??= p.keywords;
    if (p.company)        this.extendedProperties.company    ??= p.company;
    if (p.lastModifiedBy) this.coreProperties.lastModifiedBy ??= p.lastModifiedBy;
    if (p.created)        this.coreProperties.created        ??= p.created;
    if (p.category)       this.coreProperties.category       ??= p.category;
    if (p.status)         this.coreProperties.contentStatus  ??= p.status;
  }

  /** Ensure the VBA project has a document module for each worksheet. */
  private _ensureVbaSheetModules(): void {
    if (!this.vbaProject) return;
    // If the VBA project already has enough document modules (from an existing file),
    // don't add more — the existing code names may differ from display names.
    const existingDocModules = this.vbaProject.modules.filter(
      m => m.type === 'document' && m.name !== 'ThisWorkbook');
    if (existingDocModules.length >= this.sheets.length) return;
    for (const ws of this.sheets) {
      const sheetCodeName = ws.name.replace(/[^A-Za-z0-9_]/g, '_');
      if (!this.vbaProject.getModule(sheetCodeName)) {
        this.vbaProject.addModule({ name: sheetCodeName, type: 'document', code: '' });
      }
    }
  }

  private _patchWorkbookXml(originalXml: string): string {
    let xml = originalXml;
    for (let i = 0; i < this.sheets.length; i++) {
      xml = xml.replace(
        new RegExp(`(<sheet[^>]+sheetId="${i+1}"[^>]+)name="[^"]*"`),
        `$1name="${escapeXml(this.sheets[i].name)}"`
      );
    }
    // Ensure codeName on workbookPr when VBA is present
    if (this.vbaProject && !xml.includes('codeName=')) {
      xml = xml.replace('<workbookPr', '<workbookPr codeName="ThisWorkbook"');
      // If there's no workbookPr at all, add one before bookViews
      if (!xml.includes('<workbookPr')) {
        xml = xml.replace('<bookViews', '<workbookPr codeName="ThisWorkbook"/><bookViews');
      }
    }
    // Replace <definedNames> section with current named ranges
    const dnXml = this._definedNamesXml();
    if (xml.includes('<definedNames')) {
      xml = xml.replace(/<definedNames[\s\S]*?<\/definedNames>/, dnXml);
    } else if (dnXml) {
      // Insert after </sheets>
      xml = xml.replace('</sheets>', `</sheets>${dnXml}`);
    }
    return xml;
  }

  private _definedNamesXml(): string {
    // Collect print areas from sheets
    const printAreaNames: NamedRange[] = [];
    for (const ws of this.sheets) {
      if (ws.printArea) {
        const ref = ws.printArea.includes('!') ? ws.printArea : `'${ws.name}'!${ws.printArea}`;
        printAreaNames.push({ name: '_xlnm.Print_Area', ref, scope: ws.name });
      }
    }
    const all = [...this.namedRanges, ...printAreaNames];
    if (!all.length) return '';
    return `<definedNames>${all.map(nr => {
      let attrs = `name="${escapeXml(nr.name)}"`;
      if (nr.scope) {
        const idx = this.sheets.findIndex(s => s.name === nr.scope);
        if (idx >= 0) attrs += ` localSheetId="${idx}"`;
      }
      if (nr.comment) attrs += ` comment="${escapeXml(nr.comment)}"`;
      return `<definedName ${attrs}>${escapeXml(nr.ref)}</definedName>`;
    }).join('')}</definedNames>`;
  }

  /**
   * Build or patch connections.xml.
   * If originalXml is provided and no new connections were added, preserve original.
   * Otherwise generate fresh XML from connections array.
   */
  private _connectionsXml(originalXml?: string): string {
    if (!this.connections.length) return '';
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<connections xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${this.connections.map(c => {
      // Use preserved raw XML for round-tripped connections
      if (c._rawXml) return c._rawXml;
      // Generate fresh XML for new connections
      let attrs = ` id="${c.id}" name="${escapeXml(c.name)}" type="${connTypeToNum(c.type)}" refreshedVersion="6"`;
      if (c.description) attrs += ` description="${escapeXml(c.description)}"`;
      if (c.refreshOnLoad) attrs += ' refreshOnLoad="1"';
      if (c.background) attrs += ' background="1"';
      if (c.saveData) attrs += ' saveData="1"';
      if (c.keepAlive) attrs += ' keepAlive="1"';
      if (c.interval) attrs += ` interval="${c.interval}"`;
      const dbPr = c.connectionString || c.command
        ? `<dbPr${c.connectionString ? ` connection="${escapeXml(c.connectionString)}"` : ''}${c.command ? ` command="${escapeXml(c.command)}"` : ''}${c.commandType ? ` commandType="${cmdTypeToNum(c.commandType)}"` : ''}/>`
        : '';
      return `<connection${attrs}>${dbPr}</connection>`;
    }).join('')}</connections>`;
  }

  private _buildWorkbookRels(rr: ReadResult, dropCalcChain = false): string {
    const rels = [...rr.workbookRels.entries()]
      .filter(([_, rel]) => !(dropCalcChain && rel.type.includes('/calcChain')))
      .map(([id, rel]) =>
      `<Relationship Id="${id}" Type="${rel.type}" Target="${rel.target}"/>`
    );
    if (![...rr.workbookRels.values()].some(r => r.type.includes('/styles')))
      rels.push(`<Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>`);
    if (![...rr.workbookRels.values()].some(r => r.type.includes('/sharedStrings')))
      rels.push(`<Relationship Id="rIdShared" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>`);
    if (this.vbaProject && ![...rr.workbookRels.values()].some(r => r.type.includes('vbaProject')))
      rels.push(`<Relationship Id="rIdVBA" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/>`);
    if (this.connections.length && ![...rr.workbookRels.values()].some(r => r.type.includes('/connections')))
      rels.push(`<Relationship Id="rIdConns" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/connections" Target="connections.xml"/>`);
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
${rels.join('\n')}
</Relationships>`;
  }

  private _relsToXml(relMap: Map<string, { type: string; target: string; targetMode?: string }>): string {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
${[...relMap.entries()].map(([id,r]) => `<Relationship Id="${escapeXml(id)}" Type="${escapeXml(r.type)}" Target="${escapeXml(r.target)}"${r.targetMode ? ` TargetMode="${escapeXml(r.targetMode)}"` : ''}/>`).join('\n')}
</Relationships>`;
  }

  private _buildThemeXml(): string {
    const t = this.theme;
    const majorFont = t?.majorFont ?? 'Calibri Light';
    const minorFont = t?.minorFont ?? 'Calibri';
    const defaultColors = [
      { name: 'dk1', color: '000000' }, { name: 'lt1', color: 'FFFFFF' },
      { name: 'dk2', color: '44546A' }, { name: 'lt2', color: 'E7E6E6' },
      { name: 'accent1', color: '4472C4' }, { name: 'accent2', color: 'ED7D31' },
      { name: 'accent3', color: 'A5A5A5' }, { name: 'accent4', color: 'FFC000' },
      { name: 'accent5', color: '5B9BD5' }, { name: 'accent6', color: '70AD47' },
      { name: 'hlink', color: '0563C1' }, { name: 'folHlink', color: '954F72' },
    ];
    const colors = t?.colors?.map(c => {
      let hex = c.color.replace(/^#/, '');
      if (hex.length === 8) hex = hex.substring(2); // strip alpha prefix like FF
      return { name: c.name, color: hex };
    }) ?? defaultColors;
    const colorElements = colors.map(c => {
      if (c.name === 'dk1' || c.name === 'lt1') {
        return `<a:${c.name}><a:sysClr val="${c.name === 'dk1' ? 'windowText' : 'window'}" lastClr="${c.color}"/></a:${c.name}>`;
      }
      return `<a:${c.name}><a:srgbClr val="${c.color}"/></a:${c.name}>`;
    }).join('');

    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="${escapeXml(t?.name ?? 'Office Theme')}">
<a:themeElements>
<a:clrScheme name="Office">${colorElements}</a:clrScheme>
<a:fontScheme name="Office">
<a:majorFont><a:latin typeface="${escapeXml(majorFont)}"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>
<a:minorFont><a:latin typeface="${escapeXml(minorFont)}"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>
</a:fontScheme>
<a:fmtScheme name="Office">
<a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:lumMod val="110000"/><a:satMod val="105000"/><a:tint val="67000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="103000"/><a:tint val="73000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="109000"/><a:tint val="81000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:satMod val="103000"/><a:lumMod val="102000"/><a:tint val="94000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:satMod val="110000"/><a:lumMod val="100000"/><a:shade val="100000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="99000"/><a:satMod val="120000"/><a:shade val="78000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></a:fillStyleLst>
<a:lnStyleLst><a:ln w="6350" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="19050" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln></a:lnStyleLst>
<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="63000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle></a:effectStyleLst>
<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/><a:satMod val="170000"/></a:schemeClr></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="93000"/><a:satMod val="150000"/><a:shade val="98000"/><a:lumMod val="102000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:tint val="98000"/><a:satMod val="130000"/><a:shade val="90000"/><a:lumMod val="103000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="63000"/><a:satMod val="120000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></a:bgFillStyleLst>
</a:fmtScheme>
</a:themeElements>
<a:objectDefaults/>
<a:extraClrSchemeLst/>
</a:theme>`;
  }

  private _buildRootRels(hasCustom: boolean): string {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
${hasCustom ? `<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties" Target="docProps/custom.xml"/>` : ''}
</Relationships>`;
  }

  private _patchContentTypes(originalXml: string, addCustom: boolean, dropCalcChain = false): string {
    let xml = originalXml;
    if (dropCalcChain)
      xml = xml.replace(/<Override[^>]*calcChain[^>]*\/>/g, '');
    if (!xml.includes('sharedStrings'))
      xml = xml.replace('</Types>', `<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>\n</Types>`);
    if (addCustom && !xml.includes('custom-properties'))
      xml = xml.replace('</Types>', `<Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/>\n</Types>`);
    if (this.vbaProject) {
      // Add vbaProject.bin content type override if missing
      if (!xml.includes('vbaProject.bin'))
        xml = xml.replace('</Types>', `<Override PartName="/xl/vbaProject.bin" ContentType="application/vnd.ms-office.vbaProject"/>\n</Types>`);
      // Switch workbook content type to macro-enabled
      xml = xml.replace(
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
        'application/vnd.ms-excel.sheet.macroEnabled.main+xml'
      );
    }
    if (this.isTemplate && !this.vbaProject) {
      xml = xml.replace(
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
        'application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml'
      );
    }
    if (this.connections.length && !xml.includes('connections.xml'))
      xml = xml.replace('</Types>', `<Override PartName="/xl/connections.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml"/>\n</Types>`);
    return xml;
  }

  // ─── Output ────────────────────────────────────────────────────────────────

  async buildBase64(): Promise<string> {
    const bytes = await this.build();
    let bin = '';
    for (let i = 0; i < bytes.length; i++) bin += String.fromCharCode(bytes[i]);
    return btoa(bin);
  }

  async writeFile(path: string): Promise<void> {
    const bytes = await this.build();
    // @ts-ignore
    const fs = await import('fs/promises');
    await fs.writeFile(path, bytes);
  }

  // ─── Comments helpers ──────────────────────────────────────────────────────

  private _buildCommentsXml(comments: Array<{ row: number; col: number; comment: Comment }>): string {
    const authors = [...new Set(comments.map(c => c.comment.author ?? ''))];
    const authorsXml = authors.map(a => `<author>${escapeXml(a)}</author>`).join('');
    const commentsXml = comments.map(({ row, col, comment }) => {
      const ref = `${colIndexToLetter(col)}${row}`;
      const authorIdx = authors.indexOf(comment.author ?? '');
      let textXml: string;
      if (comment.richText && comment.richText.length > 0) {
        textXml = comment.richText.map(run => {
          let rPr = '';
          if (run.font) {
            const f = run.font;
            if (f.bold)   rPr += '<b/>';
            if (f.italic) rPr += '<i/>';
            if (f.underline && f.underline !== 'none') rPr += `<u val="${f.underline === 'single' ? 'single' : f.underline}"/>`;
            if (f.strike) rPr += '<strike/>';
            if (f.size)   rPr += `<sz val="${f.size}"/>`;
            if (f.color)  rPr += `<color rgb="${f.color}"/>`;
            if (f.name)   rPr += `<rFont val="${escapeXml(f.name)}"/>`;
            if (f.family != null) rPr += `<family val="${f.family}"/>`;
          }
          const rPrTag = rPr ? `<rPr>${rPr}</rPr>` : '';
          return `<r>${rPrTag}<t xml:space="preserve">${escapeXml(run.text)}</t></r>`;
        }).join('');
      } else {
        textXml = `<r><t>${escapeXml(comment.text)}</t></r>`;
      }
      return `<comment ref="${ref}" authorId="${authorIdx}"><text>${textXml}</text></comment>`;
    }).join('');
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<authors>${authorsXml}</authors>
<commentList>${commentsXml}</commentList>
</comments>`;
  }

  private _buildVmlXml(comments: Array<{ row: number; col: number; comment: Comment }>, sheetIdx: number): string {
    const shapes = comments.map(({ row, col }, i) => {
      // Position the comment box roughly 2 columns right and 0 rows above the cell
      const left  = (col + 1) * 64;
      const top   = (row - 1) * 20;
      return `<v:shape id="_x0000_s${1025 + sheetIdx * 1000 + i}" type="#_x0000_t202" style="position:absolute;margin-left:${left}pt;margin-top:${top}pt;width:108pt;height:59.25pt;z-index:${i + 1};visibility:hidden" fillcolor="#ffffe1" o:insetmode="auto">
<v:fill color2="#ffffe1"/>
<v:shadow color="black" obscured="t"/>
<v:path o:connecttype="none"/>
<v:textbox style="mso-direction-alt:auto"><div style="text-align:left"/></v:textbox>
<x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/><x:Anchor>${col + 1},15,${row - 1},10,${col + 3},15,${row + 4},4</x:Anchor><x:AutoFill>False</x:AutoFill><x:Row>${row - 1}</x:Row><x:Column>${col - 1}</x:Column></x:ClientData>
</v:shape>`;
    }).join('\n');
    return `<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">
<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout>
<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>
${shapes}
</xml>`;
  }

  async download(filename = 'workbook.xlsx'): Promise<void> {
    const bytes = await this.build();
    const blob = new Blob([bytes.buffer as ArrayBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url  = URL.createObjectURL(blob);
    const a    = document.createElement('a');
    a.href = url; a.download = filename; a.click();
    URL.revokeObjectURL(url);
  }
}

/** Map image file extension to MIME content type */
function imageContentType(ext: string): string {
  switch (ext) {
    case 'jpg':  return 'image/jpeg';
    case 'png':  return 'image/png';
    case 'gif':  return 'image/gif';
    case 'bmp':  return 'image/bmp';
    case 'tiff': return 'image/tiff';
    case 'emf':  return 'image/x-emf';
    case 'wmf':  return 'image/x-wmf';
    case 'svg':  return 'image/svg+xml';
    case 'ico':  return 'image/x-icon';
    case 'webp': return 'image/webp';
    default:     return `image/${ext}`;
  }
}

/** Map ImageFormat to file extension used in the ZIP */
function imageExt(format: string): string {
  return format === 'jpeg' ? 'jpg' : format;
}

// ─── Cell Image (richData) XML builders ─────────────────────────────────────

function buildMetadataXml(count: number): string {
  const bks = Array.from({ length: count }, (_, i) =>
    `<bk><extLst><ext uri="{3e2802c4-a4d2-4d8b-9148-e3be6c30e623}"><xlrd:rvb i="${i}"/></ext></extLst></bk>`
  ).join('');
  const cellBks = Array.from({ length: count }, (_, i) =>
    `<bk><rc t="1" v="${i}"/></bk>`
  ).join('');
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"` +
    ` xmlns:xlrd="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">` +
    `<metadataTypes count="1">` +
    `<metadataType name="XLRICHVALUE" minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1"/>` +
    `</metadataTypes>` +
    `<futureMetadata name="XLRICHVALUE" count="${count}">${bks}</futureMetadata>` +
    `<valueMetadata count="${count}">${cellBks}</valueMetadata>` +
    `</metadata>`;
}

function buildRichValueXml(count: number): string {
  const rvs = Array.from({ length: count }, (_, i) =>
    `<rv s="0"><v>${i}</v><v>5</v><v></v></rv>`
  ).join('');
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<rvData xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata" count="${count}">${rvs}</rvData>`;
}

function buildRichValueRelXml(rIds: string[]): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<richValueRels xmlns="http://schemas.microsoft.com/office/spreadsheetml/2022/richvaluerel"` +
    ` xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
    rIds.map(id => `<rel r:id="${id}"/>`).join('') +
    `</richValueRels>`;
}

function buildRichValueStructureXml(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<rvStructures xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata" count="1">` +
    `<s t="_localImage"><k n="_rvRel:LocalImageIdentifier" t="i"/><k n="CalcOrigin" t="i"/><k n="Text" t="s"/></s>` +
    `</rvStructures>`;
}

function buildRichValueTypesXml(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<rvTypesInfo xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2"` +
    ` xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"` +
    ` mc:Ignorable="x"` +
    ` xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">` +
    `<global><keyFlags>` +
    `<key name="_Self"><flag name="ExcludeFromFile" value="1"/><flag name="ExcludeFromCalcComparison" value="1"/></key>` +
    `<key name="_DisplayString"><flag name="ExcludeFromCalcComparison" value="1"/></key>` +
    `<key name="_Flags"><flag name="ExcludeFromCalcComparison" value="1"/></key>` +
    `<key name="_Format"><flag name="ExcludeFromCalcComparison" value="1"/></key>` +
    `<key name="_SubLabel"><flag name="ExcludeFromCalcComparison" value="1"/></key>` +
    `<key name="_Attribution"><flag name="ExcludeFromCalcComparison" value="1"/></key>` +
    `<key name="_Icon"><flag name="ExcludeFromCalcComparison" value="1"/></key>` +
    `<key name="_Display"><flag name="ExcludeFromCalcComparison" value="1"/></key>` +
    `<key name="_CanonicalPropertyNames"><flag name="ExcludeFromCalcComparison" value="1"/></key>` +
    `<key name="_ClassificationId"><flag name="ExcludeFromCalcComparison" value="1"/></key>` +
    `</keyFlags></global>` +
    `<types></types>` +
    `</rvTypesInfo>`;
}
