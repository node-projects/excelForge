import type {
  WorkbookProperties, NamedRange, WorksheetOptions, Image
} from '../core/types.js';
import { Worksheet } from './Worksheet.js';
import { StyleRegistry } from '../styles/StyleRegistry.js';
import { SharedStrings } from './SharedStrings.js';
import { buildChartXml } from '../features/ChartBuilder.js';
import { buildTableXml } from '../features/TableBuilder.js';
import { buildZip, type ZipEntry, type ZipOptions } from '../utils/zip.js';
import { strToBytes, base64ToBytes, escapeXml } from '../utils/helpers.js';
import { readWorkbook, type ReadResult } from './WorkbookReader.js';
import {
  buildCoreXml, buildAppXml, buildCustomXml,
  type CoreProperties, type ExtendedProperties, type CustomProperty,
} from './properties.js';

export class Workbook {
  private sheets: Worksheet[] = [];
  private namedRanges: NamedRange[] = [];
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

  removeSheet(name: string): this {
    this.sheets = this.sheets.filter(s => s.name !== name);
    return this;
  }

  // ─── Named ranges ──────────────────────────────────────────────────────────

  addNamedRange(nr: NamedRange): this {
    this.namedRanges.push(nr);
    return this;
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

    const styles = new StyleRegistry();
    const shared = new SharedStrings();

    // Re-serialise dirty sheets (collecting styles & shared strings as we go)
    const sheetXmls = new Map<number, string>();
    for (let i = 0; i < this.sheets.length; i++) {
      if (this._dirtySheets.has(i)) {
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
        headingPairs:  [{ name: 'Worksheets', count: this.sheets.length }],
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

    // ── Styles (always rewrite — may have new xfs) ─────────────────────────
    entries.push({ name: 'xl/styles.xml',        data: strToBytes(styles.toXml()) });
    entries.push({ name: 'xl/sharedStrings.xml', data: strToBytes(shared.toXml()) });

    // ── Sheets ────────────────────────────────────────────────────────────
    for (let i = 0; i < this.sheets.length; i++) {
      const path = `xl/worksheets/sheet${i + 1}.xml`;
      if (this._dirtySheets.has(i)) {
        entries.push({ name: path, data: strToBytes(sheetXmls.get(i) ?? '') });
      } else {
        // Preserve original sheet — inject any unknown parts back in
        const orig    = rr.sheets[i]?.originalXml ?? '';
        const unknown = rr.sheets[i]?.unknownParts.join('\n') ?? '';
        entries.push({ name: path, data: strToBytes(
          unknown ? orig.replace('</worksheet>', `${unknown}</worksheet>`) : orig
        )});
      }
    }

    // ── Unknown parts — verbatim ───────────────────────────────────────────
    for (const [path, data] of rr.unknownParts) {
      entries.push({ name: path, data });
    }

    // ── Rels ──────────────────────────────────────────────────────────────
    entries.push({ name: '_rels/.rels',                data: strToBytes(this._buildRootRels(customProps != null && customProps.length > 0)) });
    entries.push({ name: 'xl/_rels/workbook.xml.rels', data: strToBytes(this._buildWorkbookRels(rr)) });

    for (const [relPath, relMap] of rr.allRels) {
      if (relPath === 'xl/_rels/workbook.xml.rels') continue;
      entries.push({ name: relPath, data: strToBytes(this._relsToXml(relMap)) });
    }

    // ── Content types ──────────────────────────────────────────────────────
    entries.push({
      name: '[Content_Types].xml',
      data: strToBytes(this._patchContentTypes(rr.contentTypesXml, customProps != null && customProps.length > 0)),
    });

    return buildZip(entries, { level: this.compressionLevel });
  }

  // ── Fresh build ───────────────────────────────────────────────────────────

  private async _buildFresh(): Promise<Uint8Array> {
    const styles = new StyleRegistry();
    const shared = new SharedStrings();
    const entries: ZipEntry[] = [];

    let globalRId = 1;
    for (const ws of this.sheets) ws.rId = `rId${globalRId++}`;

    const allImages:  Array<{ ws: Worksheet; img: Image; ext: string; idx: number }> = [];
    const allCharts:  Array<{ ws: Worksheet; chartIdx: number; globalIdx: number }> = [];
    const allTables:  Array<{ ws: Worksheet; tableIdx: number; globalTableId: number }> = [];
    const sheetImageRIds  = new Map<Worksheet, string[]>();
    const sheetChartRIds  = new Map<Worksheet, string[]>();
    const sheetTableRIds  = new Map<Worksheet, string[]>();
    let imgCtr = 1, chartCtr = 1, tableCtr = 1;

    for (const ws of this.sheets) {
      const imgs = ws.getImages() as Image[];
      const charts = ws.getCharts();
      const tables = ws.getTables();
      const imgRIds: string[] = [], chartRIds: string[] = [], tblRIds: string[] = [];

      if (imgs.length || charts.length) ws.drawingRId = `rId${globalRId++}`;

      for (const img of imgs)    { const r = `rId${globalRId++}`; imgRIds.push(r);   allImages.push({ ws, img, ext: img.format === 'jpeg' ? 'jpg' : img.format, idx: imgCtr++ }); }
      for (let i=0;i<charts.length;i++) { const r = `rId${globalRId++}`; chartRIds.push(r); allCharts.push({ ws, chartIdx: i, globalIdx: chartCtr++ }); }
      for (let i=0;i<tables.length;i++) { const r = `rId${globalRId++}`; tblRIds.push(r);   allTables.push({ ws, tableIdx: i, globalTableId: tableCtr++ }); }

      sheetImageRIds.set(ws, imgRIds);
      sheetChartRIds.set(ws, chartRIds);
      sheetTableRIds.set(ws, tblRIds);
      ws.tableRIds = tblRIds;
    }

    const hasCustom = this.customProperties.length > 0;

    // Content types
    const imgCTs = new Set<string>();
    for (const { ext } of allImages) {
      const ct = ext === 'jpg' ? 'image/jpeg' : ext === 'png' ? 'image/png' : `image/${ext}`;
      imgCTs.add(`<Default Extension="${ext}" ContentType="${ct}"/>`);
    }
    entries.push({ name: '[Content_Types].xml', data: strToBytes(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
${[...imgCTs].join('')}
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
${this.sheets.map((_,i) => `<Override PartName="/xl/worksheets/sheet${i+1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('')}
${this.sheets.filter(ws=>ws.drawingRId).map((_,i) => `<Override PartName="/xl/drawings/drawing${i+1}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>`).join('')}
${allCharts.map(({globalIdx}) => `<Override PartName="/xl/charts/chart${globalIdx}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>`).join('')}
${allTables.map(({globalTableId}) => `<Override PartName="/xl/tables/table${globalTableId}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>`).join('')}
<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
${hasCustom ? '<Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/>' : ''}
</Types>`) });

    entries.push({ name: '_rels/.rels', data: strToBytes(this._buildRootRels(hasCustom)) });

    entries.push({ name: 'xl/_rels/workbook.xml.rels', data: strToBytes(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
${this.sheets.map((ws,i) => `<Relationship Id="${ws.rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i+1}.xml"/>`).join('')}
<Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rIdShared" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>`) });

    const date1904 = this.properties.date1904 ? `<workbookPr date1904="1"/>` : '<workbookPr/>';
    const namedRangesXml = this.namedRanges.length
      ? `<definedNames>${this.namedRanges.map(nr => `<definedName name="${escapeXml(nr.name)}"${nr.scope ? ` localSheetId="${this.sheets.findIndex(s=>s.name===nr.scope)}"` : ''}>${escapeXml(nr.ref)}</definedName>`).join('')}</definedNames>` : '';

    entries.push({ name: 'xl/workbook.xml', data: strToBytes(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
${date1904}
<bookViews><workbookView xWindow="0" yWindow="0" windowWidth="14400" windowHeight="8260"/></bookViews>
<sheets>${this.sheets.map((ws,i) => `<sheet name="${escapeXml(ws.name)}" sheetId="${i+1}" r:id="${ws.rId}"${ws.options?.state && ws.options.state !== 'visible' ? ` state="${ws.options.state}"` : ''}/>`).join('')}</sheets>
${namedRangesXml}
<calcPr calcId="191028"/>
</workbook>`) });

    // Per-sheet
    for (let i = 0; i < this.sheets.length; i++) {
      const ws       = this.sheets[i];
      const imgRIds  = sheetImageRIds.get(ws) ?? [];
      const cRIds    = sheetChartRIds.get(ws) ?? [];
      const tblEntries = allTables.filter(t => t.ws === ws);
      const tblRIds_ = sheetTableRIds.get(ws) ?? [];

      entries.push({ name: `xl/worksheets/sheet${i+1}.xml`, data: strToBytes(ws.toXml(styles, shared)) });

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
      if (wsRels.length) {
        entries.push({ name: `xl/worksheets/_rels/sheet${i+1}.xml.rels`, data: strToBytes(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
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
    for (const { ws, chartIdx, globalIdx } of allCharts) {
      entries.push({ name: `xl/charts/chart${globalIdx}.xml`, data: strToBytes(buildChartXml(ws.getCharts()[chartIdx])) });
    }
    for (const { ws, tableIdx, globalTableId } of allTables) {
      entries.push({ name: `xl/tables/table${globalTableId}.xml`, data: strToBytes(buildTableXml(ws.getTables()[tableIdx], globalTableId)) });
    }

    entries.push({ name: 'xl/styles.xml',        data: strToBytes(styles.toXml()) });
    entries.push({ name: 'xl/sharedStrings.xml', data: strToBytes(shared.toXml()) });

    const cp = { ...this.coreProperties, created: this.coreProperties.created ?? new Date(), modified: new Date() };
    if (!cp.creator && this.properties.author) cp.creator = this.properties.author;
    entries.push({ name: 'docProps/core.xml', data: strToBytes(buildCoreXml(cp)) });

    entries.push({ name: 'docProps/app.xml', data: strToBytes(buildAppXml({
      ...this.extendedProperties,
      application: this.extendedProperties.application ?? 'ExcelForge',
      company:     this.extendedProperties.company ?? this.properties.company,
      titlesOfParts: this.sheets.map(s => s.name),
      headingPairs:  [{ name: 'Worksheets', count: this.sheets.length }],
    })) });

    if (hasCustom) entries.push({ name: 'docProps/custom.xml', data: strToBytes(buildCustomXml(this.customProperties)) });

    return buildZip(entries, { level: this.compressionLevel });
  }

  // ─── Internal helpers ──────────────────────────────────────────────────────

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

  private _patchWorkbookXml(originalXml: string): string {
    let xml = originalXml;
    for (let i = 0; i < this.sheets.length; i++) {
      xml = xml.replace(
        new RegExp(`(<sheet[^>]+sheetId="${i+1}"[^>]+)name="[^"]*"`),
        `$1name="${escapeXml(this.sheets[i].name)}"`
      );
    }
    return xml;
  }

  private _buildWorkbookRels(rr: ReadResult): string {
    const rels = [...rr.workbookRels.entries()].map(([id, rel]) =>
      `<Relationship Id="${id}" Type="${rel.type}" Target="${rel.target}"/>`
    );
    if (![...rr.workbookRels.values()].some(r => r.type.includes('/styles')))
      rels.push(`<Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>`);
    if (![...rr.workbookRels.values()].some(r => r.type.includes('/sharedStrings')))
      rels.push(`<Relationship Id="rIdShared" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>`);
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
${rels.join('\n')}
</Relationships>`;
  }

  private _relsToXml(relMap: Map<string, { type: string; target: string }>): string {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
${[...relMap.entries()].map(([id,r]) => `<Relationship Id="${id}" Type="${r.type}" Target="${r.target}"/>`).join('\n')}
</Relationships>`;
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

  private _patchContentTypes(originalXml: string, addCustom: boolean): string {
    let xml = originalXml;
    if (!xml.includes('sharedStrings'))
      xml = xml.replace('</Types>', `<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>\n</Types>`);
    if (addCustom && !xml.includes('custom-properties'))
      xml = xml.replace('</Types>', `<Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/>\n</Types>`);
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

  async download(filename = 'workbook.xlsx'): Promise<void> {
    const bytes = await this.build();
    const blob = new Blob([bytes.buffer as ArrayBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url  = URL.createObjectURL(blob);
    const a    = document.createElement('a');
    a.href = url; a.download = filename; a.click();
    URL.revokeObjectURL(url);
  }
}
