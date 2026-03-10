# ExcelForge 📊

A **complete TypeScript library** for creating Excel `.xlsx` files with **zero external dependencies**. Works in browsers, Node.js, Deno, Bun, and edge runtimes.

Inspired by EPPlus (C#), ExcelForge gives you the full power of the OOXML spec from TypeScript.

---

## Features

| Category | Features |
|---|---|
| **Cell Values** | Strings, numbers, booleans, dates, formulas, array formulas, rich text |
| **Styles** | Fonts (size, bold, italic, color, underline, strike, family), fills (solid, pattern, gradient), borders (all sides, styles, colors), alignment (horizontal, vertical, wrap, rotate, indent, shrink), number formats |
| **Layout** | Merge cells, freeze/split panes, column widths, row heights, hide rows/cols, outline/grouping |
| **Charts** | Bar, column (stacked/100%), line, area, pie, doughnut, scatter, radar, bubble — with titles, axes, legends |
| **Images** | PNG, JPEG, GIF — positioned with two-cell or one-cell anchors |
| **Tables** | Styled Excel tables with totals row, filter buttons, column definitions |
| **Conditional Formatting** | Cell rules, color scales, data bars, icon sets, top/bottom N, above/below average |
| **Data Validation** | Dropdown lists, whole number, decimal, date, time, text length, custom formula |
| **Sparklines** | Line, bar, stacked — with colors for high/low/first/last/negative |
| **Page Setup** | Paper size, orientation, margins, headers/footers (odd/even/first), print options |
| **Protection** | Sheet protection with password (XOR hash), cell locking/hiding |
| **Named Ranges** | Workbook and sheet-scoped named ranges |
| **Auto Filter** | Dropdown filters on column headers |
| **Hyperlinks** | External URLs, mailto, internal sheet navigation |
| **Comments** | Cell comments with author |
| **Multiple Sheets** | Any number of sheets, hidden/veryHidden state, tab colors |
| **Workbook Properties** | Title, author, company, subject, keywords, dates, etc. |
| **1904 date system** | Optional 1904 date epoch |

---

## Quick Start

```typescript
import { Workbook, style, Colors, NumFmt } from './src/index.js';

const wb = new Workbook();
wb.properties.title = 'My Report';
wb.properties.author = 'ExcelForge';

const ws = wb.addSheet('Sales Data');

// Write a header row
ws.writeRow(1, 1, ['Product', 'Q1', 'Q2', 'Q3', 'Q4', 'Total']);
for (let c = 1; c <= 6; c++) {
  ws.setStyle(1, c, style().bold().bg(Colors.ExcelBlue).fontColor(Colors.White).center().build());
}

// Write data rows
ws.writeArray(2, 1, [
  ['Widget A', 1200, 1350, 1100, 1500],
  ['Widget B',  800,  950,  870, 1020],
  ['Gadget X', 2100, 1980, 2250, 2400],
]);

// Add SUM formulas in column F
for (let r = 2; r <= 4; r++) {
  ws.setFormula(r, 6, `SUM(B${r}:E${r})`);
  ws.setStyle(r, 6, style().bold().numFmt(NumFmt.Integer).build());
}

// Set column widths
ws.setColumnWidth(1, 20);
for (let c = 2; c <= 6; c++) ws.setColumnWidth(c, 12);

// Freeze header row
ws.freeze(1, 0);

// Browser: trigger download
await wb.download('sales_report.xlsx');

// Node.js: write to file
await wb.writeFile('./sales_report.xlsx');
```

---

## API Reference

### Workbook

```typescript
const wb = new Workbook();

wb.properties = {
  title, author, company, subject, description,
  keywords, category, status, created, modified,
  lastModifiedBy, date1904
};

const ws   = wb.addSheet('Name', options?);
const ws   = wb.getSheet('Name');
const ws   = wb.getSheetByIndex(0);
wb.removeSheet('Name');
wb.addNamedRange({ name, ref, scope?, comment? });

// Output
const bytes  = await wb.build();           // Uint8Array
const b64    = await wb.buildBase64();     // base64 string
await wb.writeFile('./report.xlsx');        // Node.js
await wb.download('report.xlsx');           // Browser
```

### Worksheet

```typescript
// Cell access
ws.setValue(row, col, value)
ws.setFormula(row, col, 'SUM(A1:A5)')
ws.setStyle(row, col, cellStyle)
ws.setCell(row, col, { value, formula, richText, style, comment, hyperlink })
ws.getCell(row, col)
ws.getCellByRef('A1')

// Batch writing
ws.writeRow(row, startCol, values[])
ws.writeColumn(startRow, col, values[])
ws.writeArray(startRow, startCol, values[][])

// Layout
ws.setColumnWidth(col, width)
ws.setColumn(col, { width, hidden, outlineLevel, collapsed, style })
ws.setRowHeight(row, height)
ws.setRow(row, { height, hidden, outlineLevel, collapsed, style })

// Merging
ws.merge(startRow, startCol, endRow, endCol)
ws.mergeByRef('A1:C3')

// Freeze / split panes
ws.freeze(rows?, cols?)    // e.g. ws.freeze(1) to freeze top row

// Features
ws.addImage(image)
ws.addChart(chart)
ws.addTable(table)
ws.addConditionalFormat(cf)
ws.addDataValidation(sqref, dv)
ws.addSparkline(sparkline)

// Sheet properties
ws.autoFilter = { ref: 'A1:D1' }
ws.view = { showGridLines, zoomScale, rightToLeft, view }
ws.pageSetup = { paperSize, orientation, fitToPage, ... }
ws.pageMargins = { left, right, top, bottom, header, footer }
ws.headerFooter = { oddHeader, oddFooter, ... }
ws.printOptions = { gridLines, headings, centerHorizontal }
ws.protection = { sheet, password, ... }
ws.options = { tabColor, state: 'hidden' | 'veryHidden' }
```

### Style Builder (Fluent API)

```typescript
import { style, Styles, Colors, NumFmt } from './src/index.js';

// Build from scratch
const myStyle = style()
  .bold()
  .italic()
  .fontSize(14)
  .fontName('Arial')
  .fontColor('#FF0000')         // or Colors.Red
  .bg('#4472C4')                // solid fill
  .fill('gray125')              // pattern fill
  .border('thin')               // all sides
  .borderTop('thick', Colors.Red)
  .align('center', 'middle')
  .wrap()
  .rotate(45)
  .indent(2)
  .numFmt(NumFmt.Currency)
  .locked(false)
  .build();

// Pre-built styles
Styles.headerBlue      // Bold, blue bg, white text, centered
Styles.headerGray      // Bold, dark gray bg, white text
Styles.tableHeader     // Bold, border, blue bg, centered
Styles.currency        // $#,##0.00
Styles.percent         // 0.00%
Styles.date            // Short date
Styles.bordered        // All thin borders
Styles.highlight       // Yellow background
Styles.bold            // Just bold
Styles.centered        // Just centered
Styles.redText         // Red font
// ... and more

// Color constants
Colors.White, Colors.Black, Colors.Red, Colors.Green
Colors.Blue, Colors.Yellow, Colors.Orange, Colors.Gray
Colors.ExcelBlue, Colors.ExcelOrange, Colors.ExcelGreen
Colors.DarkBlue, Colors.LightGray, Colors.Transparent

// Number format strings
NumFmt.General, NumFmt.Integer, NumFmt.Decimal2
NumFmt.Currency, NumFmt.CurrencyNeg, NumFmt.Accounting
NumFmt.Percent, NumFmt.Percent2
NumFmt.ShortDate, NumFmt.LongDate, NumFmt.DateTime
NumFmt.Scientific, NumFmt.Text, NumFmt.Multiple, NumFmt.ZeroDash
```

### Charts

```typescript
ws.addChart({
  type: 'column',        // bar, column, line, area, pie, doughnut,
                         // scatter, radar, bubble, barStacked, columnStacked100, ...
  title: 'My Chart',
  series: [
    {
      name: 'Series 1',
      values:     'Sheet1!$B$2:$B$10',
      categories: 'Sheet1!$A$2:$A$10',
      color: Colors.ExcelBlue,
    }
  ],
  from:   { col: 5, row: 1 },
  to:     { col: 14, row: 18 },
  xAxis:  { title: 'Month', gridLines: false },
  yAxis:  { title: 'Revenue ($)', min: 0, max: 5000 },
  legend: 'b',           // top | bottom | left | right | false
  style:  2,             // built-in chart style 1–48
  varyColors: true,
});
```

### Images

```typescript
import { base64ToBytes } from './src/index.js';

ws.addImage({
  data:   pngUint8Array,   // or base64 string
  format: 'png',           // png | jpeg | gif
  from:   { col: 1, row: 2, colOff: 0, rowOff: 0 },
  to:     { col: 5, row: 10 },  // optional, use width/height instead
  width:  300,             // pixels (used when 'to' is omitted)
  height: 200,
  altText: 'My image',
});
```

### Tables

```typescript
ws.addTable({
  name:        'SalesTable',
  ref:         'A1:F10',
  style:       'TableStyleMedium2',
  showRowStripes: true,
  totalsRow:   true,
  columns: [
    { name: 'Product', totalsRowLabel: 'Total' },
    { name: 'Q1', totalsRowFunction: 'sum', numFmt: '#,##0' },
    { name: 'Q2', totalsRowFunction: 'sum' },
  ]
});
```

### Conditional Formatting

```typescript
// Color scale
ws.addConditionalFormat({
  sqref: 'A1:A20', type: 'colorScale',
  colorScale: {
    type: 'colorScale',
    cfvo: [{ type: 'min' }, { type: 'max' }],
    color: ['FFFF0000', 'FF00B050'],
  }
});

// Data bars
ws.addConditionalFormat({
  sqref: 'B1:B20', type: 'dataBar',
  dataBar: { type: 'dataBar', minColor: 'FF638EC6', maxColor: 'FF638EC6' }
});

// Icon sets
ws.addConditionalFormat({
  sqref: 'C1:C20', type: 'iconSet',
  iconSet: {
    type: 'iconSet', iconSet: '3TrafficLights1',
    cfvo: [{ type: 'num', val: '0' }, { type: 'num', val: '33' }, { type: 'num', val: '67' }]
  }
});

// Cell value rule
ws.addConditionalFormat({
  sqref: 'D1:D20', type: 'cellIs',
  operator: 'greaterThan', formula: '100',
  style: style().bg(Colors.Green).fontColor(Colors.White).build(),
  priority: 1,
});
```

### Data Validation

```typescript
// Dropdown list
ws.addDataValidation('B2:B100', {
  type: 'list',
  list: ['Active', 'Inactive', 'Pending'],
  showDropDown: true,
  showErrorAlert: true,
  errorTitle: 'Invalid',
  error: 'Please select from the dropdown.',
});

// Number range
ws.addDataValidation('C2:C100', {
  type: 'whole',
  operator: 'between',
  formula1: '1',
  formula2: '100',
  showErrorAlert: true,
});
```

### Sparklines

```typescript
ws.addSparkline({
  type:       'line',        // line | bar | stacked
  dataRange:  'B2:G2',
  location:   'H2',
  color:      Colors.ExcelBlue,
  highColor:  Colors.Green,
  lowColor:   Colors.Red,
  showMarkers: true,
  showHigh:   true,
  showLow:    true,
});
```

### Page Setup

```typescript
ws.pageSetup = {
  paperSize:   9,           // 9 = A4, 1 = Letter
  orientation: 'landscape',
  fitToPage:   true,
  fitToWidth:  1,
  fitToHeight: 0,
  scale:       90,
};

ws.pageMargins = { left: 0.5, right: 0.5, top: 0.75, bottom: 0.75 };

// Headers/footers use Excel codes:
// &L = left  &C = center  &R = right
// &P = page number  &N = total pages  &D = date  &T = time
// &[Tab] = sheet name  &"Font,Style" = font  &nn = font size
ws.headerFooter = {
  oddHeader: '&L&"Calibri,Bold"My Company&C&[Tab]&RConfidential',
  oddFooter: '&LPrinted: &D &T&CPage &P of &N',
};
```

### Sheet Protection

```typescript
ws.protection = {
  sheet:    true,
  password: 'mypassword',
  selectLockedCells:   true,
  selectUnlockedCells: true,
  formatCells:   true,    // false = prevent formatting
  insertRows:    true,    // false = prevent insert
  deleteRows:    true,
  sort:          true,
  autoFilter:    true,
};

// Make specific cells editable (unlocked)
ws.setStyle(2, 1, { locked: false });
```

---

## Architecture

```
src/
├── index.ts              # Public API surface (all exports)
├── core/
│   ├── types.ts          # All TypeScript types/interfaces
│   ├── Workbook.ts       # Orchestrates everything, builds ZIP
│   ├── Worksheet.ts      # Sheet data, XML serialization
│   └── SharedStrings.ts  # String deduplication table
├── styles/
│   ├── StyleRegistry.ts  # Converts CellStyle → xf/font/fill/border indices
│   └── builders.ts       # StyleBuilder fluent API, Styles, Colors, NumFmt
├── features/
│   ├── ChartBuilder.ts   # Produces chart XML (c:chartSpace)
│   └── TableBuilder.ts   # Produces table XML
└── utils/
    ├── helpers.ts         # Cell ref math, XML escaping, date serial, EMU
    └── zip.ts            # Zero-dependency ZIP writer (STORE blocks)
```

The output `.xlsx` file is a ZIP archive containing:
- `xl/workbook.xml` — sheet list, named ranges
- `xl/worksheets/sheet{N}.xml` — cell data, formulas, merges, CF, tables
- `xl/styles.xml` — fonts, fills, borders, number formats
- `xl/sharedStrings.xml` — deduplicated strings
- `xl/drawings/drawing{N}.xml` — images and charts positions
- `xl/charts/chart{N}.xml` — chart definitions
- `xl/media/image{N}.{ext}` — embedded image data
- `xl/tables/table{N}.xml` — table definitions
- `docProps/core.xml` & `docProps/app.xml` — metadata
- `[Content_Types].xml` & `_rels/.rels` — OOXML manifest

---

## Browser Usage

```html
<script type="module">
  import { Workbook, style } from './src/index.js';

  document.getElementById('download').addEventListener('click', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.setValue(1, 1, 'Hello, Excel!');
    ws.setStyle(1, 1, style().bold().fontSize(20).bg('#4472C4').fontColor('#FFFFFF').center().build());
    await wb.download('hello.xlsx');
  });
</script>
```

---

## No External Dependencies

ExcelForge uses only native Web APIs:
- `TextEncoder` — UTF-8 string encoding
- `atob` / `btoa` — base64 encode/decode (for images)
- `URL.createObjectURL` — browser downloads only
- `fs/promises` — Node.js file writes only

The ZIP implementation uses STORE (uncompressed) blocks for maximum compatibility.

---

## License

MIT
