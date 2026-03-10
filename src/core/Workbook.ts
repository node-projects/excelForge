import type {
  WorkbookProperties, NamedRange, WorksheetOptions, Image
} from '../core/types.js';
import { Worksheet } from './Worksheet.js';
import { StyleRegistry } from '../styles/StyleRegistry.js';
import { SharedStrings } from './SharedStrings.js';
import { buildChartXml } from '../features/ChartBuilder.js';
import { buildTableXml } from '../features/TableBuilder.js';
import { buildZip, type ZipEntry } from '../utils/zip.js';
import { strToBytes, base64ToBytes, escapeXml } from '../utils/helpers.js';

export class Workbook {
  private sheets: Worksheet[] = [];
  private namedRanges: NamedRange[] = [];
  properties: WorkbookProperties = {};

  // ─── Sheet management ──────────────────────────────────────────────────────

  addSheet(name: string, options: WorksheetOptions = {}): Worksheet {
    const ws = new Worksheet(name, options);
    ws.sheetIndex = this.sheets.length + 1;
    this.sheets.push(ws);
    return ws;
  }

  getSheet(name: string): Worksheet | undefined {
    return this.sheets.find(s => s.name === name);
  }

  getSheetByIndex(idx: number): Worksheet | undefined {
    return this.sheets[idx];
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

  // ─── Build XLSX ────────────────────────────────────────────────────────────

  async build(): Promise<Uint8Array> {
    const styles = new StyleRegistry();
    const shared = new SharedStrings();

    // Pre-register all cell styles
    for (const ws of this.sheets) {
      // (styles are registered on-demand in toXml)
    }

    // Assign relationship IDs
    let globalRId = 1;
    const sheetRIds: string[] = [];
    for (const ws of this.sheets) {
      ws.rId = `rId${globalRId++}`;
      sheetRIds.push(ws.rId);
    }

    const entries: ZipEntry[] = [];

    // Per-sheet resources
    const sheetDrawingRIds: Map<Worksheet, string>    = new Map();
    const sheetImageRIds:   Map<Worksheet, string[]>  = new Map();
    const sheetChartRIds:   Map<Worksheet, string[]>  = new Map();
    const sheetTableRIds:   Map<Worksheet, string[]>  = new Map();
    const allImages:  Array<{ ws: Worksheet; img: Image; ext: string; idx: number }> = [];
    const allCharts:  Array<{ ws: Worksheet; chartIdx: number; globalIdx: number }> = [];
    const allTables:  Array<{ ws: Worksheet; tableIdx: number; globalTableId: number }> = [];

    let imageCounter   = 1;
    let chartCounter   = 1;
    let tableCounter   = 1;
    let drawingCounter = 1;

    for (const ws of this.sheets) {
      const imgs   = ws.getImages() as Image[];
      const charts = ws.getCharts();
      const tables = ws.getTables();

      const imgRIds:   string[] = [];
      const chartRIds: string[] = [];
      const tableRIds_: string[] = [];

      let drawingRId = '';

      if (imgs.length || charts.length) {
        const localRId = globalRId++;
        drawingRId = `rId${localRId}`;
        ws.drawingRId = drawingRId;
        sheetDrawingRIds.set(ws, drawingRId);
      }

      for (const img of imgs) {
        const rId = `rId${globalRId++}`;
        imgRIds.push(rId);
        allImages.push({ ws, img, ext: img.format === 'jpeg' ? 'jpg' : img.format, idx: imageCounter++ });
      }

      for (let i = 0; i < charts.length; i++) {
        const rId = `rId${globalRId++}`;
        chartRIds.push(rId);
        allCharts.push({ ws, chartIdx: i, globalIdx: chartCounter++ });
      }

      for (let i = 0; i < tables.length; i++) {
        const rId = `rId${globalRId++}`;
        tableRIds_.push(rId);
        allTables.push({ ws, tableIdx: i, globalTableId: tableCounter++ });
      }

      sheetImageRIds.set(ws, imgRIds);
      sheetChartRIds.set(ws, chartRIds);
      sheetTableRIds.set(ws, tableRIds_);
      ws.tableRIds = tableRIds_;
    }

    // ── [Content_Types].xml ────────────────────────────────────────────────
    const imageContentTypes = new Set<string>();
    for (const { ext } of allImages) {
      const ct = ext === 'jpg' ? 'image/jpeg' : ext === 'png' ? 'image/png' : ext === 'gif' ? 'image/gif' : 'image/'+ext;
      imageContentTypes.add(`<Default Extension="${ext}" ContentType="${ct}"/>`);
    }

    const sheetCTs   = this.sheets.map((_, i) =>
      `<Override PartName="/xl/worksheets/sheet${i+1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`
    ).join('');
    const drawingCTs = [...new Set(this.sheets.filter(ws => ws.drawingRId).map((_, i) => i+1))].map(i =>
      `<Override PartName="/xl/drawings/drawing${i}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>`
    ).join('');
    const chartCTs = allCharts.map(({ globalIdx }) =>
      `<Override PartName="/xl/charts/chart${globalIdx}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>`
    ).join('');
    const tableCTs = allTables.map(({ globalTableId }) =>
      `<Override PartName="/xl/tables/table${globalTableId}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>`
    ).join('');

    entries.push({
      name: '[Content_Types].xml',
      data: strToBytes(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
${[...imageContentTypes].join('')}
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
${sheetCTs}${drawingCTs}${chartCTs}${tableCTs}
<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>`),
    });

    // ── _rels/.rels ────────────────────────────────────────────────────────
    entries.push({
      name: '_rels/.rels',
      data: strToBytes(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`),
    });

    // ── xl/_rels/workbook.xml.rels ─────────────────────────────────────────
    const wbRels = this.sheets.map((ws, i) =>
      `<Relationship Id="${ws.rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i+1}.xml"/>`
    ).join('');

    entries.push({
      name: 'xl/_rels/workbook.xml.rels',
      data: strToBytes(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
${wbRels}
<Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rIdShared" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>`),
    });

    // ── xl/workbook.xml ────────────────────────────────────────────────────
    const sheetsXml = this.sheets.map((ws, i) => {
      const state = ws.options?.state;
      return `<sheet name="${escapeXml(ws.name)}" sheetId="${i+1}" r:id="${ws.rId}"${state && state !== 'visible' ? ` state="${state}"` : ''}/>`;
    }).join('');

    const namedRangesXml = this.namedRanges.length
      ? `<definedNames>${this.namedRanges.map(nr =>
          `<definedName name="${escapeXml(nr.name)}"${nr.scope ? ` localSheetId="${this.sheets.findIndex(s=>s.name===nr.scope)}"` : ''}>${escapeXml(nr.ref)}</definedName>`
        ).join('')}</definedNames>` : '';

    const date1904 = this.properties.date1904 ? `<workbookPr date1904="1"/>` : '<workbookPr/>';

    entries.push({
      name: 'xl/workbook.xml',
      data: strToBytes(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
${date1904}
<bookViews><workbookView xWindow="0" yWindow="0" windowWidth="14400" windowHeight="8260"/></bookViews>
<sheets>${sheetsXml}</sheets>
${namedRangesXml}
<calcPr calcId="191028"/>
</workbook>`),
    });

    // ── Per-sheet files ────────────────────────────────────────────────────
    for (let i = 0; i < this.sheets.length; i++) {
      const ws    = this.sheets[i];
      const imgRIds   = sheetImageRIds.get(ws) ?? [];
      const chartRIds = sheetChartRIds.get(ws) ?? [];
      const tblRIds   = sheetTableRIds.get(ws) ?? [];

      // Sheet XML
      entries.push({
        name: `xl/worksheets/sheet${i+1}.xml`,
        data: strToBytes(ws.toXml(styles, shared)),
      });

      // Sheet rels
      const hasDrawing = ws.drawingRId !== '';
      const wsRels: string[] = [];

      if (hasDrawing) {
        const dIdx = this.sheets.filter((s, j) => j <= i && s.drawingRId).length;
        wsRels.push(`<Relationship Id="${ws.drawingRId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing${dIdx}.xml"/>`);
      }

      // Per-sheet image rels
      const wsImages   = ws.getImages() as Image[];
      const wsCharts   = ws.getCharts();
      const wsTables   = ws.getTables();

      for (let j = 0; j < wsImages.length; j++) {
        const img    = wsImages[j];
        const global = allImages.find(x => x.ws === ws && wsImages.indexOf(img) === wsImages.indexOf(x.img as any));
        if (!global) continue;
        wsRels.push(`<Relationship Id="${imgRIds[j]}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image${global.idx}.${global.ext}"/>`);
      }

      // Per-sheet chart rels (pointing from drawing to chart)
      // handled below in drawing rels

      // Table rels
      const wsTableEntries = allTables.filter(t => t.ws === ws);
      for (let j = 0; j < wsTableEntries.length; j++) {
        wsRels.push(`<Relationship Id="${tblRIds[j]}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table${wsTableEntries[j].globalTableId}.xml"/>`);
      }

      if (wsRels.length) {
        entries.push({
          name: `xl/worksheets/_rels/sheet${i+1}.xml.rels`,
          data: strToBytes(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
${wsRels.join('\n')}
</Relationships>`),
        });
      }

      // Drawing XML
      if (hasDrawing) {
        const wsChartRIds = chartRIds;
        const drawingXml  = ws.toDrawingXml(imgRIds, wsChartRIds);
        const dIdx = this.sheets.filter((s, j) => j <= i && s.drawingRId).length;
        entries.push({
          name: `xl/drawings/drawing${dIdx}.xml`,
          data: strToBytes(drawingXml),
        });

        // Drawing rels (image + chart relationships)
        const drawingRels: string[] = [];
        for (let j = 0; j < wsImages.length; j++) {
          const global = allImages.filter(x => x.ws === ws)[j];
          if (global) {
            drawingRels.push(`<Relationship Id="${imgRIds[j]}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image${global.idx}.${global.ext}"/>`);
          }
        }
        for (let j = 0; j < wsCharts.length; j++) {
          const global = allCharts.filter(x => x.ws === ws)[j];
          if (global) {
            drawingRels.push(`<Relationship Id="${wsChartRIds[j]}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart${global.globalIdx}.xml"/>`);
          }
        }
        if (drawingRels.length) {
          entries.push({
            name: `xl/drawings/_rels/drawing${dIdx}.xml.rels`,
            data: strToBytes(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
${drawingRels.join('\n')}
</Relationships>`),
          });
        }
      }
    }

    // ── Images ────────────────────────────────────────────────────────────
    for (const { img, ext, idx } of allImages) {
      let data: Uint8Array;
      if (typeof img.data === 'string') {
        data = base64ToBytes(img.data);
      } else {
        data = img.data;
      }
      entries.push({ name: `xl/media/image${idx}.${ext}`, data });
    }

    // ── Charts ────────────────────────────────────────────────────────────
    for (const { ws, chartIdx, globalIdx } of allCharts) {
      const chart = ws.getCharts()[chartIdx];
      entries.push({
        name: `xl/charts/chart${globalIdx}.xml`,
        data: strToBytes(buildChartXml(chart)),
      });
    }

    // ── Tables ────────────────────────────────────────────────────────────
    for (const { ws, tableIdx, globalTableId } of allTables) {
      const table = ws.getTables()[tableIdx];
      entries.push({
        name: `xl/tables/table${globalTableId}.xml`,
        data: strToBytes(buildTableXml(table, globalTableId)),
      });
    }

    // ── Styles ────────────────────────────────────────────────────────────
    // Re-trigger style registration by building sheet XMLs once more? No —
    // styles.register is called inside ws.toXml; but we already called it above.
    // The StyleRegistry accumulates all xfs on first call, so it's fine.
    entries.push({
      name: 'xl/styles.xml',
      data: strToBytes(styles.toXml()),
    });

    // ── Shared Strings ────────────────────────────────────────────────────
    entries.push({
      name: 'xl/sharedStrings.xml',
      data: strToBytes(shared.toXml()),
    });

    // ── docProps ──────────────────────────────────────────────────────────
    const p = this.properties;
    const now = (p.modified ?? new Date()).toISOString();
    const created = (p.created ?? new Date()).toISOString();

    entries.push({
      name: 'docProps/core.xml',
      data: strToBytes(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
  xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
${p.title       ? `<dc:title>${escapeXml(p.title)}</dc:title>` : ''}
${p.subject     ? `<dc:subject>${escapeXml(p.subject)}</dc:subject>` : ''}
${p.author      ? `<dc:creator>${escapeXml(p.author)}</dc:creator>` : ''}
${p.keywords    ? `<cp:keywords>${escapeXml(p.keywords)}</cp:keywords>` : ''}
${p.description ? `<dc:description>${escapeXml(p.description)}</dc:description>` : ''}
${p.lastModifiedBy ? `<cp:lastModifiedBy>${escapeXml(p.lastModifiedBy)}</cp:lastModifiedBy>` : ''}
${p.category    ? `<cp:category>${escapeXml(p.category)}</cp:category>` : ''}
${p.status      ? `<cp:contentStatus>${escapeXml(p.status)}</cp:contentStatus>` : ''}
<dcterms:created xsi:type="dcterms:W3CDTF">${created}</dcterms:created>
<dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified>
</cp:coreProperties>`),
    });

    entries.push({
      name: 'docProps/app.xml',
      data: strToBytes(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
<Application>ExcelForge</Application>
${p.company ? `<Company>${escapeXml(p.company)}</Company>` : ''}
<DocSecurity>0</DocSecurity>
<ScaleCrop>false</ScaleCrop>
<LinksUpToDate>false</LinksUpToDate>
<SharedDoc>false</SharedDoc>
<HyperlinksChanged>false</HyperlinksChanged>
<AppVersion>1.0</AppVersion>
</Properties>`),
    });

    return buildZip(entries);
  }

  /** Convenience: build and return as base64 string */
  async buildBase64(): Promise<string> {
    const bytes = await this.build();
    let bin = '';
    for (let i = 0; i < bytes.length; i++) bin += String.fromCharCode(bytes[i]);
    return btoa(bin);
  }

  /** Node.js: write to file */
  async writeFile(path: string): Promise<void> {
    const bytes = await this.build();
    // @ts-ignore
    const fs = await import('fs/promises');
    await fs.writeFile(path, bytes);
  }

  /** Browser: trigger download */
  async download(filename = 'workbook.xlsx'): Promise<void> {
    const bytes = await this.build();
    const blob  = new Blob([bytes.buffer as ArrayBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url   = URL.createObjectURL(blob);
    const a     = document.createElement('a');
    a.href = url; a.download = filename; a.click();
    URL.revokeObjectURL(url);
  }
}
