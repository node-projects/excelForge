# ExcelForge 📊

[![npm version](https://badge.fury.io/js/%40node-projects%2Fexcelforge.svg)](https://badge.fury.io/js/%40node-projects%2Fexcelforge)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

A **complete TypeScript library** for reading and writing Excel `.xlsx` and `.xlsm` (macro-enabled) files with **zero external dependencies**. Works in browsers, Node.js, Deno, Bun, and edge runtimes.

ExcelForge gives you the full power of the OOXML spec — including real DEFLATE compression, round-trip editing of existing files, and rich property support.

---

## Features

| Category | Features |
|---|---|
| **Read existing files** | Load `.xlsx` from file, `Uint8Array`, `base64`, or `Blob` |
| **Patch-only writes** | Re-serialise only changed sheets; preserve pivot tables, VBA, charts, unknown parts verbatim |
| **Compression** | Full LZ77 + Huffman DEFLATE (levels 0–9). Typical XML compresses 80–85% |
| **Cell Values** | Strings, numbers, booleans, dates, formulas, array formulas, rich text |
| **Styles** | Fonts, solid/pattern/gradient fills, all border styles, alignment, 30+ number format presets |
| **Layout** | Merge cells, freeze/split panes, column widths, row heights, hide rows/cols, outline grouping |
| **Charts** | Bar, column (stacked/100%), line, area, pie, doughnut, scatter, radar, bubble |
| **Images** | PNG, JPEG, GIF — two-cell or one-cell anchors |
| **Tables** | Styled Excel tables with totals row, filter buttons, column definitions |
| **Conditional Formatting** | Cell rules, color scales, data bars, icon sets, top/bottom N, above/below average |
| **Data Validation** | Dropdowns, whole number, decimal, date, time, text length, custom formula |
| **Sparklines** | Line, bar, stacked — with high/low/first/last/negative colors |
| **Page Setup** | Paper size, orientation, margins, headers/footers (odd/even/first), print options, page breaks |
| **Protection** | Sheet protection with password, cell locking/hiding |
| **Named Ranges** | Workbook and sheet-scoped |
| **Connections** | OLEDB, ODBC, text/CSV, web — create, read, round-trip |
| **Power Query** | Read M formulas from DataMashup; full round-trip preservation |
| **Pivot Tables** | Row/column/data fields, aggregation functions (sum, count, avg, max, min…), styles |
| **VBA Macros** | Create/read `.xlsm` with standard modules, class modules, document modules; full round-trip |
| **Auto Filter** | Dropdown filters on column headers |
| **Hyperlinks** | External URLs, mailto, internal navigation |
| **Form Controls** | Button, checkbox, combobox, listbox, radio, groupbox, label, scrollbar, spinner — with macro assignment |
| **Comments** | Cell comments with author |
| **Multiple Sheets** | Any number, hidden/veryHidden, tab colors |
| **Core Properties** | Title, author, subject, keywords, description, language, revision, category… |
| **Extended Properties** | Company, manager, application, appVersion, hyperlinkBase, word/line/page counts… |
| **Custom Properties** | Typed key-value store: string, int, decimal, bool, date, r8, i8 |

---

## Installation

```bash
# Copy the src/ directory into your project, or compile to dist/ first:
tsc --outDir dist --target ES2020 --module NodeNext --moduleResolution NodeNext \
    --declaration --strict --skipLibCheck src/index.ts [all src files]
```

No `npm install` required — zero runtime dependencies.

---

## Quick Start — Create a workbook

```typescript
import { Workbook, style, Colors, NumFmt } from './src/index.js';

const wb = new Workbook();
wb.coreProperties = { title: 'Q4 Report', creator: 'Alice', language: 'en-US' };
wb.extendedProperties = { company: 'Acme Corp', appVersion: '1.0' };

const ws = wb.addSheet('Sales Data');

// Header row
ws.writeRow(1, 1, ['Product', 'Q1', 'Q2', 'Q3', 'Q4', 'Total']);
for (let c = 1; c <= 6; c++) {
  ws.setStyle(1, c, style().bold().bg(Colors.ExcelBlue).fontColor(Colors.White).center().build());
}

// Data rows
ws.writeArray(2, 1, [
  ['Widget A', 1200, 1350, 1100, 1500],
  ['Widget B',  800,  950,  870, 1020],
  ['Gadget X', 2100, 1980, 2250, 2400],
]);

// SUM formulas
for (let r = 2; r <= 4; r++) {
  ws.setFormula(r, 6, `SUM(B${r}:E${r})`);
  ws.setStyle(r, 6, style().bold().build());
}

ws.freeze(1, 0); // freeze first row

// Output — compression level 6 by default (80–85% smaller than STORE)
await wb.writeFile('./report.xlsx');          // Node.js
await wb.download('report.xlsx');             // Browser
const bytes = await wb.build();              // Uint8Array (any runtime)
const b64   = await wb.buildBase64();        // base64 string
```

---

## Reading & modifying existing files

ExcelForge can load existing `.xlsx` files and either read their contents or patch them. Only the sheets you mark as dirty are re-serialised on write; everything else — pivot tables, VBA, drawings, slicers, macros — is preserved verbatim from the original ZIP.

### Loading

```typescript
// Node.js / Deno / Bun
const wb = await Workbook.fromFile('./existing.xlsx');

// Universal (Uint8Array)
const wb = await Workbook.fromBytes(uint8Array);

// Browser (File / Blob input element)
const wb = await Workbook.fromBlob(fileInputElement.files[0]);

// base64 string (e.g. from an API or email attachment)
const wb = await Workbook.fromBase64(base64String);
```

### Reading data

```typescript
console.log(wb.getSheetNames());         // ['Sheet1', 'Summary', 'Config']

const ws = wb.getSheet('Summary');
const cell = ws.getCell(3, 2);           // row 3, col 2
console.log(cell.value);                 // 'Q4 Revenue'
console.log(cell.formula);              // 'SUM(B10:B20)'
console.log(cell.style?.font?.bold);    // true
```

### Modifying and saving

```typescript
const wb = await Workbook.fromFile('./report.xlsx');
const ws = wb.getSheet('Sales');

// Make changes
ws.setValue(5, 3, 99000);
ws.setStyle(5, 3, style().bg(Colors.LightGreen).build());
ws.writeRow(20, 1, ['TOTAL', '', '=SUM(C2:C19)']);

// Mark the sheet dirty — it will be re-serialised on write.
// Sheets NOT marked dirty are written back byte-for-byte from the original.
wb.markDirty('Sales');

// Patch properties without re-serialising any sheets
wb.coreProperties.title = 'Updated Report';
wb.setCustomProperty('Status', { type: 'string', value: 'Approved' });

await wb.writeFile('./report_updated.xlsx');
```

> **Tip:** If you forget to call `markDirty()`, your cell changes won't appear in the output because the original sheet XML will be used. Always call it after modifying a loaded sheet.

---

## Compression

ExcelForge includes a full pure-TypeScript DEFLATE implementation (LZ77 lazy matching + dynamic/fixed Huffman coding) with no external dependencies. XML content — the bulk of any `.xlsx` — typically compresses to 80–85% of its original size.

### Setting the compression level

```typescript
const wb = new Workbook();
wb.compressionLevel = 6;  // 0–9, default 6
```

| Level | Description | Typical size vs STORE |
|---|---|---|
| `0` | STORE — no compression, fastest | baseline |
| `1` | FAST — fixed Huffman, minimal LZ77 | ~75% smaller |
| `6` | **DEFAULT** — dynamic Huffman + lazy LZ77 | ~82% smaller |
| `9` | BEST — maximum LZ77 effort | ~83% smaller (marginal gain over 6) |

Level 6 is the default and the recommended choice — it achieves most of the compression benefit of level 9 at a fraction of the CPU cost.

### Per-entry level override

The `buildZip` function used internally also supports per-entry overrides, useful if you want images (already compressed) stored uncompressed while XML entries are compressed:

```typescript
import { buildZip } from './src/utils/zip.js';

const zip = buildZip([
  { name: 'xl/worksheets/sheet1.xml', data: xmlBytes },         // uses global level
  { name: 'xl/media/image1.png',      data: pngBytes, level: 0 }, // forced STORE
  { name: 'xl/styles.xml',            data: stylesBytes, level: 9 }, // max compression
], { level: 6 });
```

By default, `buildZip` automatically stores image file types (`png`, `jpg`, `gif`, `tiff`, `emf`, `wmf`) uncompressed since they're already compressed formats.

---

## Document Properties

ExcelForge reads and writes all three OOXML property namespaces.

### Core properties (`docProps/core.xml`)

```typescript
wb.coreProperties = {
  title:          'Annual Report 2024',
  subject:        'Financial Summary',
  creator:        'Finance Team',
  keywords:       'excel quarterly finance',
  description:    'Auto-generated from ERP export',
  lastModifiedBy: 'Alice',
  revision:       '3',
  language:       'en-US',
  category:       'Finance',
  contentStatus:  'Final',
  created:        new Date('2024-01-01'),
  // modified is always set to current time on write
};
```

### Extended properties (`docProps/app.xml`)

```typescript
wb.extendedProperties = {
  application:       'ExcelForge',
  appVersion:        '1.0.0',
  company:           'Acme Corp',
  manager:           'Bob Smith',
  hyperlinkBase:     'https://intranet.acme.com/',
  docSecurity:       0,
  linksUpToDate:     true,
  // These are computed automatically on write:
  // titlesOfParts, headingPairs
};
```

### Custom properties (`docProps/custom.xml`)

Custom properties support typed values — they appear in Excel under **File → Properties → Custom**.

```typescript
// Set custom properties at workbook level
wb.customProperties = [
  { name: 'ProjectCode',  value: { type: 'string',  value: 'PRJ-2024-007' } },
  { name: 'Revision',     value: { type: 'int',     value: 5             } },
  { name: 'Budget',       value: { type: 'decimal', value: 125000.00     } },
  { name: 'IsApproved',   value: { type: 'bool',    value: true          } },
  { name: 'ReviewDate',   value: { type: 'date',    value: new Date()    } },
];

// Or use the helper methods
wb.setCustomProperty('Status', { type: 'string', value: 'In Review' });
wb.setCustomProperty('Score',  { type: 'decimal', value: 9.7 });
wb.removeCustomProperty('OldField');

// Read back
const proj = wb.getCustomProperty('ProjectCode');
console.log(proj?.value.value);  // 'PRJ-2024-007'

// Full list
for (const p of wb.customProperties) {
  console.log(p.name, p.value.type, p.value.value);
}
```

Available value types: `string`, `int`, `decimal`, `bool`, `date`, `r8` (8-byte float), `i8` (BigInt).

---

## Cell API reference

### Writing values

```typescript
ws.setValue(row, col, value);          // string | number | boolean | Date
ws.setFormula(row, col, 'SUM(A1:A5)');
ws.setArrayFormula(row, col, 'row*col formula', 'A1:C3');
ws.setStyle(row, col, cellStyle);
ws.setCell(row, col, { value, formula, style, comment, hyperlink });

// Bulk writes
ws.writeRow(row, startCol, [v1, v2, v3]);
ws.writeArray(startRow, startCol, [[...], [...], ...]);
```

### Reading values

```typescript
const cell = ws.getCell(row, col);
cell.value     // the stored value (string | number | boolean | undefined)
cell.formula   // formula string if present
cell.style     // CellStyle object
```

### Styles

```typescript
import { style, Colors, NumFmt, Styles } from './src/index.js';

// Fluent builder
const headerStyle = style()
  .bold()
  .italic()
  .fontSize(13)
  .fontColor(Colors.White)
  .bg(Colors.ExcelBlue)
  .border('thin')
  .center()
  .wrapText()
  .numFmt(NumFmt.Currency)
  .build();

// Built-in presets
ws.setStyle(1, 1, Styles.bold);
ws.setStyle(1, 2, Styles.headerBlue);
ws.setStyle(2, 3, Styles.currency);
ws.setStyle(3, 4, Styles.percent);
```

### Number formats

```typescript
NumFmt.General      // General
NumFmt.Integer      // 0
NumFmt.Decimal2     // #,##0.00
NumFmt.Currency     // $#,##0.00
NumFmt.Percent      // 0%
NumFmt.Percent2     // 0.00%
NumFmt.Scientific   // 0.00E+00
NumFmt.ShortDate    // mm-dd-yy
NumFmt.LongDate     // d-mmm-yy
NumFmt.Time         // h:mm:ss AM/PM
NumFmt.DateTime     // m/d/yy h:mm
NumFmt.Accounting   // _($* #,##0.00_)
NumFmt.Text         // @
```

### Layout

```typescript
ws.merge(r1, c1, r2, c2);              // merge a range
ws.mergeByRef('A1:D1');
ws.freeze(rows, cols);                 // freeze panes
ws.setColumn(colIndex, { width: 20, hidden: false, style });
ws.setRow(rowIndex, { height: 30, hidden: false });
ws.autoFilter = { ref: 'A1:E1' };
```

### Conditional formatting

```typescript
ws.addConditionalFormat({
  sqref: 'C2:C100',
  type: 'colorScale',
  colorScale: {
    min: { type: 'min', color: 'FFF8696B' },
    max: { type: 'max', color: 'FF63BE7B' },
  },
  priority: 1,
});

ws.addConditionalFormat({
  sqref: 'D2:D100',
  type: 'dataBar',
  dataBar: { color: 'FF638EC6' },
  priority: 2,
});
```

### Data validation

```typescript
ws.addDataValidation({
  sqref: 'B2:B100',
  type: 'list',
  formula1: '"North,South,East,West"',
  showDropDown: false,
  errorTitle: 'Invalid Region',
  error: 'Please select a valid region.',
});
```

### Charts

```typescript
ws.addChart({
  type: 'bar',
  title: 'Sales by Region',
  series: [{ name: 'Q1 Sales', dataRange: 'Sheet1!B2:B6', catRange: 'Sheet1!A2:A6' }],
  position: { from: { row: 1, col: 8 }, to: { row: 20, col: 16 } },
  legend: { position: 'bottom' },
});
```

Supported chart types: `bar`, `col`, `colStacked`, `col100`, `barStacked`, `bar100`, `line`, `lineStacked`, `area`, `pie`, `doughnut`, `scatter`, `radar`, `bubble`.

### Images

```typescript
import { readFileSync } from 'fs';
const imgData = readFileSync('./logo.png');

ws.addImage({
  data:   imgData,          // Buffer, Uint8Array, or base64 string
  format: 'png',
  from:   { row: 1, col: 1 },
  to:     { row: 8, col: 4 },
});
```

### Pivot tables

```typescript
const wb = new Workbook();

// Source data sheet
const wsData = wb.addSheet('Data');
wsData.writeRow(1, 1, ['Region', 'Product', 'Sales', 'Units']);
wsData.writeArray(2, 1, [
  ['North', 'Widget', 12000, 150],
  ['South', 'Widget', 9500,  120],
  ['North', 'Gadget', 8700,  90],
  ['South', 'Gadget', 11200, 140],
]);

// Pivot table on a separate sheet
const wsPivot = wb.addSheet('Summary');
wsPivot.addPivotTable({
  name:        'SalesBreakdown',
  sourceSheet: 'Data',
  sourceRef:   'A1:D5',
  targetCell:  'A1',
  rowFields:   ['Region'],
  colFields:   ['Product'],
  dataFields:  [{ field: 'Sales', name: 'Sum of Sales', func: 'sum' }],
  style:       'PivotStyleMedium9',
  rowGrandTotals: true,
  colGrandTotals: true,
});

await wb.writeFile('./pivot_report.xlsx');
```

Available aggregation functions: `sum`, `count`, `average`, `max`, `min`, `product`, `countNums`, `stdDev`, `stdDevp`, `var`, `varp`.

### VBA macros

ExcelForge can create, read, and round-trip `.xlsm` files with VBA macros. All module types are supported: standard modules, class modules, and document modules (auto-created for `ThisWorkbook` and each worksheet).

```typescript
import { Workbook, VbaProject } from './src/index.js';

const wb = new Workbook();
const ws = wb.addSheet('Sheet1');
ws.setValue(1, 1, 'Hello');

const vba = new VbaProject();

// Standard module
vba.addModule({
  name: 'Module1',
  type: 'standard',
  code: 'Sub HelloWorld()\r\n    MsgBox "Hello from VBA!"\r\nEnd Sub\r\n',
});

// Class module
vba.addModule({
  name: 'MyClass',
  type: 'class',
  code: [
    'Private pValue As String',
    'Public Property Get Value() As String',
    '    Value = pValue',
    'End Property',
    'Public Property Let Value(v As String)',
    '    pValue = v',
    'End Property',
  ].join('\r\n') + '\r\n',
});

wb.vbaProject = vba;
await wb.writeFile('./macros.xlsm');  // must use .xlsm extension
```

Reading VBA from existing files:

```typescript
const wb = await Workbook.fromFile('./macros.xlsm');
if (wb.vbaProject) {
  for (const mod of wb.vbaProject.modules) {
    console.log(`${mod.name} (${mod.type}): ${mod.code.length} chars`);
  }
}

// Modify and re-save — existing modules are preserved
wb.vbaProject.addModule({ name: 'Module2', type: 'standard', code: '...' });
wb.vbaProject.removeModule('OldModule');
await wb.writeFile('./macros_updated.xlsm');
```

> **Note:** Document modules for `ThisWorkbook` and each worksheet are automatically created if not explicitly provided. VBA code uses `\r\n` line endings.

### Page setup

```typescript
ws.pageSetup = {
  paperSize:   9,             // A4
  orientation: 'landscape',
  scale:       90,
  fitToPage:   true,
  fitToWidth:  1,
  fitToHeight: 0,
};

ws.pageMargins = {
  left: 0.5, right: 0.5, top: 0.75, bottom: 0.75,
  header: 0.3, footer: 0.3,
};

ws.headerFooter = {
  oddHeader: '&C&BQ4 Report&B',
  oddFooter: '&LExcelForge&RPage &P of &N',
};
```

### Page breaks

```typescript
// Add manual page breaks for printing
ws.addRowBreak(20);    // page break after row 20
ws.addRowBreak(40);    // page break after row 40
ws.addColBreak(5);     // page break after column E

// Read page breaks from an existing file
const wb = await Workbook.fromBytes(data);
const ws = wb.getSheet('Sheet1')!;
for (const brk of ws.getRowBreaks()) {
  console.log(`Row break at ${brk.id}, manual: ${brk.manual}`);
}
for (const brk of ws.getColBreaks()) {
  console.log(`Col break at ${brk.id}, manual: ${brk.manual}`);
}
```

Page breaks are fully preserved during round-trip editing, even when sheets are modified.

### Named ranges

```typescript
// Define workbook-scoped named ranges
wb.addNamedRange({ name: 'SalesData', ref: 'Data!$A$1:$A$5' });
wb.addNamedRange({ name: 'Products', ref: 'Data!$B$1:$B$5', comment: 'Product list' });

// Define sheet-scoped named range
wb.addNamedRange({ name: 'LocalTotal', ref: 'Data!$A$6', scope: 'Data' });

// Use in formulas
ws.setFormula(1, 1, 'SUM(SalesData)');

// Read named ranges from an existing file
const wb2 = await Workbook.fromBytes(data);
const ranges = wb2.getNamedRanges();         // all named ranges
const sales = wb2.getNamedRange('SalesData'); // find by name
console.log(sales?.ref);                      // "Data!$A$1:$A$5"

// Remove a named range
wb2.removeNamedRange('SalesData');
```

Named ranges (including scope and comments) are fully preserved during round-trip editing.

### Connections & Power Query

```typescript
// Add a data connection (OLEDB, ODBC, text/CSV, web, etc.)
wb.addConnection({
  id: 1,
  name: 'SalesDB',
  type: 'oledb',  // 'odbc' | 'dao' | 'file' | 'web' | 'oledb' | 'text' | 'dsp'
  connectionString: 'Provider=SQLOLEDB;Data Source=server;Initial Catalog=Sales;',
  command: 'SELECT * FROM Orders',
  commandType: 'sql',  // 'sql' | 'table' | 'default' | 'web' | 'oledb'
  description: 'Sales database connection',
  background: true,
  saveData: true,
});

// Read connections from an existing file
const wb2 = await Workbook.fromBytes(data);
const conns = wb2.getConnections();           // all connections
const sales = wb2.getConnection('SalesDB');   // find by name
wb2.removeConnection('SalesDB');              // remove by name

// Read Power Query M formulas (extracted from DataMashup)
const queries = wb2.getPowerQueries();        // all queries
const q = wb2.getPowerQuery('MyQuery');       // find by name
console.log(q?.formula);                       // Power Query M code
```

Connections are fully preserved during round-trip editing. Power Query formulas (M code) stored in DataMashup binary blobs are automatically extracted for read access. Power Query/Power Pivot data models created in Excel are preserved verbatim during round-trip — you can safely open, modify cells, and save without losing any Power Query or Power Pivot features.

### Form Controls

```typescript
// Add a button with a macro
ws.addFormControl({
  type: 'button',
  from: { col: 1, row: 2 },
  to:   { col: 3, row: 4 },
  text: 'Run Report',
  macro: 'Sheet1.RunReport',
});

// CheckBox linked to a cell
ws.addFormControl({
  type: 'checkBox',
  from: { col: 1, row: 5 },
  to:   { col: 3, row: 6 },
  text: 'Enable Feature',
  linkedCell: '$B$10',
  checked: 'checked',   // 'checked' | 'unchecked' | 'mixed'
});

// ComboBox (dropdown) with input range
ws.addFormControl({
  type: 'comboBox',
  from: { col: 1, row: 7 },
  to:   { col: 3, row: 8 },
  linkedCell: '$B$11',
  inputRange: '$D$1:$D$5',
  dropLines: 5,
});

// ListBox, OptionButton, GroupBox, Label, ScrollBar, Spinner
ws.addFormControl({
  type: 'scrollBar',
  from: { col: 4, row: 6 },
  to:   { col: 6, row: 7 },
  linkedCell: '$B$14',
  min: 0, max: 100, inc: 1, page: 10, val: 50,
});

// Read form controls from an existing file
const wb2 = await Workbook.fromBytes(data);
const controls = ws.getFormControls();
for (const ctrl of controls) {
  console.log(ctrl.type, ctrl.linkedCell, ctrl.macro);
}
```

Supported control types: `button`, `checkBox`, `comboBox`, `listBox`, `optionButton`, `groupBox`, `label`, `scrollBar`, `spinner`. All control types support `macro` assignment and are fully preserved during round-trip editing.

### Sheet protection

```typescript
ws.protect('mypassword', {
  formatCells:   false,   // allow formatting
  insertRows:    false,   // allow inserting rows
  deleteRows:    false,
  sort:          false,
  autoFilter:    false,
});

// Lock individual cells (requires sheet protection to take effect)
ws.setCell(1, 1, { value: 'Locked', style: { locked: true } });
ws.setCell(2, 1, { value: 'Editable', style: { locked: false } });
```

---

## Output methods

```typescript
// Node.js: write to file
await wb.writeFile('./output.xlsx');

// Browser: trigger download
await wb.download('report.xlsx');

// Any runtime: get bytes
const bytes: Uint8Array = await wb.build();
const b64: string       = await wb.buildBase64();
```

---

## ZIP / Compression API

The `buildZip` and `deflateRaw` utilities are exported for direct use:

```typescript
import { buildZip, deflateRaw, type ZipEntry, type ZipOptions } from './src/utils/zip.js';

// deflateRaw: compress bytes with raw DEFLATE (no zlib header)
const compressed = deflateRaw(data, 6);  // level 0–9

// buildZip: assemble a ZIP archive
const zip = buildZip(entries, { level: 6 });

// ZipEntry shape
interface ZipEntry {
  name:   string;
  data:   Uint8Array;
  level?: number;  // per-entry override
}

// ZipOptions shape
interface ZipOptions {
  level?:       number;    // global default (0–9)
  noCompress?:  string[];  // extensions to always STORE
}
```

---

## Architecture overview

```
ExcelForge
├── core/
│   ├── Workbook.ts         — orchestrates build/read/patch, holds properties
│   ├── Worksheet.ts        — cells, formulas, styles, drawings, page setup
│   ├── WorkbookReader.ts   — parse existing XLSX (ZIP → XML → object model)
│   ├── SharedStrings.ts    — string deduplication table
│   ├── properties.ts       — core / extended / custom property read+write
│   └── types.ts            — all 80+ TypeScript interfaces
├── styles/
│   ├── StyleRegistry.ts    — interns fonts/fills/borders/xfs, emits styles.xml
│   └── builders.ts         — fluent style() builder, Colors/NumFmt/Styles presets
├── features/
│   ├── ChartBuilder.ts     — DrawingML chart XML for 15+ chart types
│   ├── TableBuilder.ts     — Excel table XML
│   └── PivotTableBuilder.ts — pivot table + cache XML
├── vba/
│   ├── VbaProject.ts       — VBA project build/parse, module management
│   ├── cfb.ts              — Compound Binary File (OLE2) reader & writer
│   └── ovba.ts             — MS-OVBA compression/decompression
└── utils/
    ├── zip.ts              — ZIP writer with full LZ77+Huffman DEFLATE
    ├── zipReader.ts        — ZIP reader (STORE + DEFLATE via DecompressionStream)
    ├── xmlParser.ts        — roundtrip-safe XML parser (preserves unknown nodes)
    └── helpers.ts          — cell ref math, XML escaping, date serials, EMU conversion
```

### Round-trip / patch strategy

When you load an existing `.xlsx` and call `wb.build()`:

1. The original ZIP is read and every entry is retained as raw bytes.
2. Sheets **not** marked dirty via `wb.markDirty(name)` are written back verbatim — their original bytes are preserved unchanged.
3. Sheets that **are** marked dirty are re-serialised with any changes applied.
4. Core/extended/custom properties are always rewritten (they're cheap and typically user-modified).
5. Styles and shared strings are always rewritten (dirty sheets need fresh indices).
6. All other parts — drawings, charts, images, pivot tables, VBA modules, custom XML, connections, theme — are preserved verbatim.

This means you can safely open a complex Excel file produced by another tool, change a few cells, and save without losing any features ExcelForge doesn't understand.

---

## Browser usage

ExcelForge is fully tree-shakeable and has zero runtime dependencies. In the browser, use `CompressionStream` / `DecompressionStream` (available in all modern browsers since 2022) for decompression when reading files.

```html
<input type="file" id="file" accept=".xlsx">
<script type="module">
import { Workbook } from './dist/index.js';

document.getElementById('file').addEventListener('change', async (e) => {
  const file = e.target.files[0];
  const wb   = await Workbook.fromBlob(file);

  console.log('Sheets:', wb.getSheetNames());
  console.log('Title:', wb.coreProperties.title);

  const ws = wb.getSheet(wb.getSheetNames()[0]);
  console.log('A1:', ws.getCell(1, 1).value);

  // Modify and re-download
  ws.setValue(1, 1, 'Modified!');
  wb.markDirty(wb.getSheetNames()[0]);
  await wb.download('modified.xlsx');
});
</script>
```

---

## Changelog

### v2.4 — Pivot Tables & VBA Macros

- **Pivot tables** — create pivot tables with row/column/data fields, 11 aggregation functions, customisable styles
- **VBA macros** — create, read, and round-trip `.xlsm` files with standard, class, and document modules
- **CFB (OLE2) support** — MS-CFB reader/writer for vbaProject.bin, with MS-OVBA compression
- **Automatic sheet modules** — document modules for ThisWorkbook and each worksheet are auto-generated

### v2.0 — Read, Modify, Compress

- **Read existing XLSX files** — `Workbook.fromFile()`, `fromBytes()`, `fromBase64()`, `fromBlob()`
- **Patch-only writes** — preserve unknown parts verbatim, only re-serialise dirty sheets
- **Full DEFLATE compression** — pure-TypeScript LZ77 + dynamic Huffman (levels 0–9), 80–85% smaller output
- **Extended & custom properties** — full read/write of `core.xml`, `app.xml`, `custom.xml`
- **New utilities** — `zipReader.ts`, `xmlParser.ts`, `properties.ts`

### v1.0 — Initial release

- Full XLSX write support: cells, formulas, styles, charts, images, tables, conditional formatting, data validation, sparklines, page setup, protection, named ranges, auto filter, hyperlinks, comments
