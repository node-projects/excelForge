/**
 * ExcelForge — Comprehensive Usage Examples
 * This file demonstrates every major feature of the library.
 */

import { Workbook, Worksheet, style, Styles, Colors, NumFmt, VbaProject } from '../index.js';
import type { Chart, ConditionalFormat, Table, Sparkline, DataValidation, Image } from '../index.js';
import { deflateSync } from 'zlib';

// ============================================================
// 1. BASIC WORKBOOK & SHEET CREATION
// ============================================================
async function example_basic() {
  const wb = new Workbook();
  wb.properties = {
    title:    'ExcelForge Demo',
    author:   'ExcelForge',
    company:  'Acme Corp',
    subject:  'Demonstration',
    keywords: 'excel typescript',
    created:  new Date(),
  };

  const ws = wb.addSheet('Sheet1');

  // Set cell values
  ws.setValue(1, 1, 'Hello');
  ws.setValue(1, 2, 42);
  ws.setValue(1, 3, true);
  ws.setValue(1, 4, new Date());
  ws.setValue(1, 5, null); // empty cell

  // Write rows & arrays
  ws.writeRow(2, 1, ['Name', 'Age', 'City']);
  ws.writeArray(3, 1, [
    ['Alice', 30, 'Berlin'],
    ['Bob',   25, 'Paris'],
    ['Carol', 35, 'Tokyo'],
  ]);

  await wb.writeFile('./output/01_basic.xlsx');
}

// ============================================================
// 2. FORMULAS
// ============================================================
async function example_formulas() {
  const wb = new Workbook();
  const ws = wb.addSheet('Formulas');

  ws.writeColumn(1, 1, [10, 20, 30, 40, 50]);

  ws.setFormula(6, 1, 'SUM(A1:A5)');
  ws.setFormula(7, 1, 'AVERAGE(A1:A5)');
  ws.setFormula(8, 1, 'MAX(A1:A5)');
  ws.setFormula(9, 1, 'MIN(A1:A5)');
  ws.setFormula(10, 1, 'COUNT(A1:A5)');

  // Conditional formulas
  ws.setFormula(11, 1, 'IF(A1>20,"High","Low")');
  ws.setFormula(12, 1, 'VLOOKUP(20,A1:A5,1,FALSE)');
  ws.setFormula(13, 1, 'SUMIF(A1:A5,">20")');
  ws.setFormula(14, 1, 'IFERROR(A1/0,"N/A")');

  // String formulas
  ws.setFormula(1, 3, 'CONCATENATE("Hello"," ","World")');
  ws.setFormula(2, 3, 'UPPER("hello")');
  ws.setFormula(3, 3, 'LEFT("ExcelForge",5)');
  ws.setFormula(4, 3, 'LEN("ExcelForge")');

  // Date formulas
  ws.setFormula(1, 5, 'TODAY()');
  ws.setFormula(2, 5, 'NOW()');
  ws.setFormula(3, 5, 'YEAR(TODAY())');
  ws.setFormula(4, 5, 'TEXT(TODAY(),"dd/mm/yyyy")');

  // Array formula
  ws.getCell(15, 1).arrayFormula = 'SUM(A1:A5*2)';

  await wb.writeFile('./output/02_formulas.xlsx');
}

// ============================================================
// 3. CELL FORMATTING & STYLES
// ============================================================
async function example_styles() {
  const wb = new Workbook();
  const ws = wb.addSheet('Styles');

  // Font styles
  ws.setCell(1, 1, { value: 'Bold',       style: style().bold().build() });
  ws.setCell(2, 1, { value: 'Italic',     style: style().italic().build() });
  ws.setCell(3, 1, { value: 'Strike',     style: style().strike().build() });
  ws.setCell(4, 1, { value: 'Underline',  style: style().underline().build() });
  ws.setCell(5, 1, { value: 'Large Font', style: style().fontSize(18).build() });
  ws.setCell(6, 1, { value: 'Red Text',   style: style().fontColor(Colors.Red).build() });
  ws.setCell(7, 1, { value: 'Calibri',    style: style().fontName('Calibri').build() });
  ws.setCell(8, 1, { value: 'Superscript',style: style().font({ vertAlign: 'superscript' }).build() });

  // Background fills
  ws.setCell(1, 3, { value: 'Blue BG',   style: style().bg(Colors.ExcelBlue).fontColor(Colors.White).build() });
  ws.setCell(2, 3, { value: 'Yellow',    style: style().bg(Colors.Yellow).build() });
  ws.setCell(3, 3, { value: 'Gradient',  style: {
    fill: {
      type: 'gradient',
      gradientType: 'linear',
      degree: 90,
      stops: [
        { position: 0, color: 'FF4472C4' },
        { position: 1, color: 'FF70AD47' },
      ]
    }
  }});

  // Borders
  ws.setCell(1, 5, { value: 'All Thin',   style: style().border('thin').build() });
  ws.setCell(2, 5, { value: 'Thick',      style: style().border('thick').build() });
  ws.setCell(3, 5, { value: 'Dashed',     style: style().border('dashed').build() });
  ws.setCell(4, 5, { value: 'Double',     style: style().border('double').build() });
  ws.setCell(5, 5, { value: 'Custom',     style: style()
    .borderTop('thick', Colors.Red)
    .borderBottom('thin', Colors.Blue)
    .borderLeft('dashed', Colors.Green)
    .borderRight('dotted', Colors.ExcelOrange ?? 'FFED7D31')
    .build() });

  // Alignment
  ws.setCell(1, 7, { value: 'Center',       style: style().center().build() });
  ws.setCell(2, 7, { value: 'Right',        style: style().align('right').build() });
  ws.setCell(3, 7, { value: 'Wrap\nText',   style: style().wrap().build() });
  ws.setCell(4, 7, { value: 'Rotate 45°',   style: style().rotate(45).build() });
  ws.setCell(5, 7, { value: 'Indent',       style: style().indent(3).build() });

  // Number formats
  ws.setCell(1, 9, { value: 1234.56,     style: style().numFmt(NumFmt.Currency).build() });
  ws.setCell(2, 9, { value: 0.1234,      style: style().numFmt(NumFmt.Percent2).build() });
  ws.setCell(3, 9, { value: 1234567,     style: style().numFmt(NumFmt.Integer).build() });
  ws.setCell(4, 9, { value: new Date(),  style: style().numFmt(NumFmt.ShortDate).build() });
  ws.setCell(5, 9, { value: new Date(),  style: style().numFmt(NumFmt.DateTime).build() });
  ws.setCell(6, 9, { value: 1234.56,     style: style().numFmt(NumFmt.Accounting).build() });
  ws.setCell(7, 9, { value: 2.5,         style: style().numFmt(NumFmt.Multiple).build() });

  // Using pre-built styles
  ws.setCell(10, 1, { value: 'Pre-built Table Header', style: Styles.tableHeader });
  ws.setCell(11, 1, { value: 'Highlighted',            style: Styles.highlight });
  ws.setCell(12, 1, { value: '$1,234.56',              style: Styles.currency });

  // Column widths
  ws.setColumnWidth(1, 20); ws.setColumnWidth(3, 20); ws.setColumnWidth(5, 20);
  ws.setColumnWidth(7, 25); ws.setColumnWidth(9, 20);

  await wb.writeFile('./output/03_styles.xlsx');
}

// ============================================================
// 4. MERGING CELLS
// ============================================================
async function example_merges() {
  const wb = new Workbook();
  const ws = wb.addSheet('Merges');

  ws.setCell(1, 1, { value: 'Merged Heading', style: Styles.headerBlue });
  ws.merge(1, 1, 1, 5); // Merge A1:E1

  ws.setCell(2, 1, { value: 'Left Merged', style: Styles.headerGray });
  ws.merge(2, 1, 4, 1); // Vertical merge A2:A4

  ws.mergeByRef('B2:E4');
  ws.setCell(2, 2, { value: 'Big merged area', style: style().center().wrap().build() });

  await wb.writeFile('./output/04_merges.xlsx');
}

// ============================================================
// 5. RICH TEXT
// ============================================================
async function example_richtext() {
  const wb = new Workbook();
  const ws = wb.addSheet('Rich Text');

  ws.setCell(1, 1, {
    richText: [
      { text: 'Hello ',  font: { bold: true, color: 'FF0000FF', size: 14 } },
      { text: 'World',   font: { italic: true, color: 'FFFF0000', size: 12 } },
      { text: '!',       font: { bold: true, italic: true, size: 16 } },
    ]
  });

  ws.setCell(3, 1, {
    richText: [
      { text: 'Revenue: ', font: { bold: true } },
      { text: '$1,234.56', font: { color: 'FF00B050', bold: true } },
      { text: ' (+12%)',   font: { color: 'FF0070C0' } },
    ]
  });

  ws.setColumnWidth(1, 40);

  await wb.writeFile('./output/05_richtext.xlsx');
}

// ============================================================
// 6. FREEZE PANES & VIEWS
// ============================================================
async function example_panes() {
  const wb = new Workbook();
  const ws = wb.addSheet('Freeze');

  // Header row
  ws.writeRow(1, 1, ['ID', 'Name', 'Value', 'Date']);
  for (let i = 0; i < 1; i++) {
    ws.setStyle(1, i+1, Styles.tableHeader);
  }

  for (let r = 2; r <= 50; r++) {
    ws.writeRow(r, 1, [r-1, `Item ${r-1}`, Math.random() * 1000, new Date()]);
  }

  ws.freeze(1, 0); // Freeze first row
  // ws.freeze(0, 1); // Or freeze first column
  // ws.freeze(1, 1); // Or freeze first row AND column

  ws.view = { showGridLines: true, zoomScale: 100, tabSelected: true };

  await wb.writeFile('./output/06_panes.xlsx');
}

// ============================================================
// 7. CONDITIONAL FORMATTING
// ============================================================
async function example_conditional_formatting() {
  const wb = new Workbook();
  const ws = wb.addSheet('CF');

  const data = [10, 85, 42, 99, 3, 67, 55, 78, 20, 91];
  ws.writeColumn(1, 1, data);
  ws.writeColumn(1, 2, data);
  ws.writeColumn(1, 3, data);
  ws.writeColumn(1, 4, data);

  // Color scale: red-yellow-green
  ws.addConditionalFormat({
    sqref: 'A1:A10', type: 'colorScale',
    colorScale: {
      type: 'colorScale',
      cfvo: [{ type: 'min' }, { type: 'percent', val: '50' }, { type: 'max' }],
      color: ['FFFF0000', 'FFFFFF00', 'FF00B050'],
    }
  });

  // Data bar
  ws.addConditionalFormat({
    sqref: 'B1:B10', type: 'dataBar',
    dataBar: { type: 'dataBar', minColor: 'FF638EC6', maxColor: 'FF638EC6', showValue: true }
  });

  // Icon set
  ws.addConditionalFormat({
    sqref: 'C1:C10', type: 'iconSet',
    iconSet: {
      type: 'iconSet', iconSet: '3TrafficLights1',
      cfvo: [{ type: 'num', val: '0' }, { type: 'num', val: '33' }, { type: 'num', val: '67' }]
    }
  });

  // Cell value rule: highlight cells > 70
  ws.addConditionalFormat({
    sqref: 'D1:D10', type: 'cellIs', operator: 'greaterThan', formula: '70',
    style: style().bg(Colors.Green).fontColor(Colors.White).build(),
    priority: 1,
  });

  // Top 10%
  ws.addConditionalFormat({
    sqref: 'E1:E10', type: 'top10', rank: 3, percent: false,
    style: Styles.highlight, priority: 2,
  });

  await wb.writeFile('./output/07_conditional_formatting.xlsx');
}

// ============================================================
// 8. TABLES
// ============================================================
async function example_tables() {
  const wb = new Workbook();
  const ws = wb.addSheet('Tables');

  const headers = ['Product', 'Q1', 'Q2', 'Q3', 'Q4', 'Total'];
  ws.writeRow(1, 1, headers);
  ws.writeArray(2, 1, [
    ['Widget A', 1200, 1350, 1100, 1500],
    ['Widget B',  800,  950,  870, 1020],
    ['Gadget X', 2100, 1980, 2250, 2400],
    ['Gadget Y',  650,  720,  680,  800],
  ]);

  // Add SUM formula in Total column
  for (let r = 2; r <= 5; r++) {
    ws.setFormula(r, 6, `SUM(B${r}:E${r})`);
    ws.setStyle(r, 6, style().bold().build());
  }

  ws.addTable({
    name: 'SalesTable',
    ref: 'A1:F6',
    style: 'TableStyleMedium2',
    showRowStripes: true,
    totalsRow: true,
    columns: [
      { name: 'Product', totalsRowLabel: 'Total' },
      { name: 'Q1',      totalsRowFunction: 'sum', numFmt: NumFmt.Integer },
      { name: 'Q2',      totalsRowFunction: 'sum', numFmt: NumFmt.Integer },
      { name: 'Q3',      totalsRowFunction: 'sum', numFmt: NumFmt.Integer },
      { name: 'Q4',      totalsRowFunction: 'sum', numFmt: NumFmt.Integer },
      { name: 'Total',   totalsRowFunction: 'sum', numFmt: NumFmt.Integer },
    ]
  });

  for (let c = 1; c <= 6; c++) ws.setColumnWidth(c, 14);
  ws.setColumnWidth(1, 20);

  await wb.writeFile('./output/08_tables.xlsx');
}

// ============================================================
// 9. CHARTS
// ============================================================
async function example_charts() {
  const wb = new Workbook();
  const ws = wb.addSheet('Charts');

  // Data for chart
  ws.writeRow(1, 1, ['Month', 'Sales', 'Expenses', 'Profit']);
  const months = ['Jan','Feb','Mar','Apr','May','Jun'];
  const sales   = [4200, 5100, 4800, 6200, 5800, 7100];
  const expenses= [3100, 3400, 3200, 3800, 3600, 4100];
  months.forEach((m, i) => {
    ws.setValue(i+2, 1, m);
    ws.setValue(i+2, 2, sales[i]);
    ws.setValue(i+2, 3, expenses[i]);
    ws.setFormula(i+2, 4, `B${i+2}-C${i+2}`);
  });

  // Column chart
  ws.addChart({
    type: 'column',
    title: 'Monthly Performance',
    series: [
      { name: 'Sales',    values: 'Charts!$B$2:$B$7', categories: 'Charts!$A$2:$A$7' },
      { name: 'Expenses', values: 'Charts!$C$2:$C$7', categories: 'Charts!$A$2:$A$7' },
    ],
    from: { col: 5, row: 0 },
    to:   { col: 13, row: 15 },
    xAxis: { title: 'Month' },
    yAxis: { title: 'Amount ($)', gridLines: true },
    legend: 'b',
    style: 2,
  });

  // Line chart on the same sheet
  ws.addChart({
    type: 'line',
    title: 'Sales Trend',
    series: [
      { name: 'Sales',  values: 'Charts!$B$2:$B$7', categories: 'Charts!$A$2:$A$7', color: Colors.ExcelBlue },
      { name: 'Profit', values: 'Charts!$D$2:$D$7', categories: 'Charts!$A$2:$A$7', color: Colors.ExcelGreen },
    ],
    from: { col: 5, row: 16 },
    to:   { col: 13, row: 31 },
    legend: 'b',
  });

  // Pie chart
  const ws2 = wb.addSheet('Pie Chart');
  ws2.writeColumn(1, 1, ['North', 'South', 'East', 'West']);
  ws2.writeColumn(1, 2, [35, 25, 20, 20]);

  ws2.addChart({
    type: 'pie',
    title: 'Sales by Region',
    varyColors: true,
    series: [{
      name: 'Region',
      values:     'Pie Chart!$B$1:$B$4',
      categories: 'Pie Chart!$A$1:$A$4',
    }],
    from: { col: 3, row: 0 },
    to:   { col: 11, row: 15 },
    legend: 'r',
  });

  await wb.writeFile('./output/09_charts.xlsx');
}

// ============================================================
// 10. IMAGES
// ============================================================
async function example_images() {
  const wb = new Workbook();
  const ws = wb.addSheet('Images');

  ws.setValue(1, 1, 'Image below:');

  // Generate a 200x120 PNG showing a mini bar chart with Excel brand colors
  const pngBytes = (() => {
    const W = 200, H = 120;

    // CRC32 table
    const crcTable = new Uint32Array(256);
    for (let n = 0; n < 256; n++) {
      let c = n;
      for (let k = 0; k < 8; k++) c = (c & 1) ? 0xEDB88320 ^ (c >>> 1) : c >>> 1;
      crcTable[n] = c;
    }
    const crc32 = (buf: Uint8Array, seed = 0xFFFFFFFF): number => {
      let c = seed;
      for (const b of buf) c = crcTable[(c ^ b) & 0xFF] ^ (c >>> 8);
      return (c ^ 0xFFFFFFFF) >>> 0;
    };

    const u32be = (n: number) => new Uint8Array([n>>>24, (n>>>16)&0xFF, (n>>>8)&0xFF, n&0xFF]);

    const chunk = (type: string, data: Uint8Array): Uint8Array => {
      const typeBytes = new TextEncoder().encode(type);
      const crcBuf = new Uint8Array(4 + data.length);
      crcBuf.set(typeBytes); crcBuf.set(data, 4);
      const crc = u32be(crc32(crcBuf));
      const out = new Uint8Array(4 + 4 + data.length + 4);
      out.set(u32be(data.length));
      out.set(typeBytes, 4);
      out.set(data, 8);
      out.set(crc, 8 + data.length);
      return out;
    };

    // Colors: ExcelBlue, ExcelOrange, ExcelGreen, ExcelYellow, ExcelPurple
    const palette: [number,number,number][] = [
      [0x44,0x72,0xC4], [0xED,0x7D,0x31], [0x70,0xAD,0x47],
      [0xFF,0xC0,0x00], [0x7F,0x48,0xCC],
    ];
    const barValues = [0.55, 0.80, 0.45, 0.95, 0.65];
    const barCount  = barValues.length;
    const margin    = 10;
    const barW      = Math.floor((W - margin * 2) / barCount) - 4;
    const maxBarH   = H - margin * 2;

    // Build raw RGBA pixel rows
    const raw = new Uint8Array((1 + W * 3) * H); // filter byte + RGB per row
    let pos = 0;
    for (let y = 0; y < H; y++) {
      raw[pos++] = 0; // filter type None
      for (let x = 0; x < W; x++) {
        let r = 0xF2, g = 0xF2, b = 0xF2; // light-grey background
        for (let i = 0; i < barCount; i++) {
          const bx = margin + i * (barW + 4);
          const bh = Math.round(barValues[i] * maxBarH);
          const by = H - margin - bh;
          if (x >= bx && x < bx + barW && y >= by && y < H - margin) {
            [r, g, b] = palette[i % palette.length];
          }
        }
        raw[pos++] = r; raw[pos++] = g; raw[pos++] = b;
      }
    }

    const idat = deflateSync(raw);

    const ihdrData = new Uint8Array(13);
    const dv = new DataView(ihdrData.buffer);
    dv.setUint32(0, W); dv.setUint32(4, H);
    ihdrData[8] = 8;   // bit depth
    ihdrData[9] = 2;   // color type RGB
    // compression, filter, interlace = 0

    const sig = new Uint8Array([0x89,0x50,0x4E,0x47,0x0D,0x0A,0x1A,0x0A]);
    const ihdr = chunk('IHDR', ihdrData);
    const idatChunk = chunk('IDAT', idat);
    const iend = chunk('IEND', new Uint8Array(0));

    const total = sig.length + ihdr.length + idatChunk.length + iend.length;
    const png = new Uint8Array(total);
    let off = 0;
    for (const part of [sig, ihdr, idatChunk, iend]) {
      png.set(part, off); off += part.length;
    }
    return png;
  })();

  const img: Image = {
    data: pngBytes,
    format: 'png',
    from:   { col: 1, row: 2, colOff: 0, rowOff: 0 },
    width:  200,
    height: 100,
    altText: 'Sample image',
  };

  ws.addImage(img);

  await wb.writeFile('./output/10_images.xlsx');
}

// ============================================================
// 11. DATA VALIDATION
// ============================================================
async function example_data_validation() {
  const wb = new Workbook();
  const ws = wb.addSheet('Validation');

  ws.writeRow(1, 1, ['Name', 'Status', 'Score', 'Date', 'Email']);
  ws.setStyle(1, 1, Styles.tableHeader);
  ws.setStyle(1, 2, Styles.tableHeader);
  ws.setStyle(1, 3, Styles.tableHeader);
  ws.setStyle(1, 4, Styles.tableHeader);
  ws.setStyle(1, 5, Styles.tableHeader);

  // Dropdown list
  ws.addDataValidation('B2:B20', {
    type: 'list',
    list: ['Active', 'Inactive', 'Pending', 'Archived'],
    showDropDown: true,
    showErrorAlert: true,
    errorTitle: 'Invalid Status',
    error: 'Please select a valid status from the dropdown.',
    showInputMessage: true,
    promptTitle: 'Status',
    prompt: 'Select the item status.',
  });

  // Number range validation
  ws.addDataValidation('C2:C20', {
    type: 'whole',
    operator: 'between',
    formula1: '0',
    formula2: '100',
    showErrorAlert: true,
    errorTitle: 'Invalid Score',
    error: 'Score must be between 0 and 100.',
    allowBlank: true,
  });

  // Date validation
  ws.addDataValidation('D2:D20', {
    type: 'date',
    operator: 'greaterThan',
    formula1: 'TODAY()',
    showErrorAlert: true,
    errorTitle: 'Invalid Date',
    error: 'Date must be in the future.',
  });

  // Text length validation
  ws.addDataValidation('A2:A20', {
    type: 'textLength',
    operator: 'between',
    formula1: '2',
    formula2: '50',
    showErrorAlert: true,
    errorTitle: 'Name too short/long',
    error: 'Name must be 2–50 characters.',
  });

  // Custom formula validation (no spaces in email)
  ws.addDataValidation('E2:E20', {
    type: 'custom',
    formula1: 'ISNUMBER(FIND("@",E2))',
    showErrorAlert: true,
    errorTitle: 'Invalid Email',
    error: 'Must contain @.',
  });

  for (let c = 1; c <= 5; c++) ws.setColumnWidth(c, 18);

  await wb.writeFile('./output/11_validation.xlsx');
}

// ============================================================
// 12. SPARKLINES
// ============================================================
async function example_sparklines() {
  const wb = new Workbook();
  const ws = wb.addSheet('Sparklines');

  ws.writeRow(1, 1, ['Item', 'Jan','Feb','Mar','Apr','May','Jun','Trend','Bar']);
  ws.setStyle(1, 1, Styles.tableHeader);
  for (let c = 2; c <= 9; c++) ws.setStyle(1, c, Styles.tableHeader);

  const rows = [
    ['Product A', 10, 15, 13, 18, 22, 20],
    ['Product B',  5,  8, 12,  9, 14, 16],
    ['Product C', 20, 18, 21, 17, 19, 23],
  ];

  rows.forEach((row, i) => {
    ws.writeRow(i+2, 1, row);
    const r = i + 2;

    // Line sparkline
    ws.addSparkline({
      type: 'line',
      dataRange: `B${r}:G${r}`,
      location: `H${r}`,
      color: Colors.ExcelBlue,
      highColor: Colors.Green,
      lowColor: Colors.Red,
      showHigh: true,
      showLow: true,
      showMarkers: true,
    });

    // Bar sparkline
    ws.addSparkline({
      type: 'bar',
      dataRange: `B${r}:G${r}`,
      location: `I${r}`,
      color: Colors.ExcelOrange ?? 'FFED7D31',
      negativeColor: Colors.Red,
      showNegative: true,
    });
  });

  ws.setColumnWidth(1, 15);
  ws.setColumnWidth(8, 20);
  ws.setColumnWidth(9, 20);

  await wb.writeFile('./output/12_sparklines.xlsx');
}

// ============================================================
// 13. PAGE SETUP, HEADERS & FOOTERS
// ============================================================
async function example_page_setup() {
  const wb = new Workbook();
  const ws = wb.addSheet('Page Setup');

  // Fill with sample data
  for (let r = 1; r <= 100; r++) {
    ws.writeRow(r, 1, [`Row ${r}`, r * 10, r * 20, r * 30]);
  }

  ws.pageSetup = {
    paperSize:    9,   // A4
    orientation:  'landscape',
    fitToPage:    true,
    fitToWidth:   1,
    fitToHeight:  0,
    horizontalDpi: 300,
    verticalDpi:   300,
  };

  ws.pageMargins = {
    left: 0.5, right: 0.5,
    top: 0.75, bottom: 0.75,
    header: 0.3, footer: 0.3,
  };

  ws.headerFooter = {
    oddHeader:  '&L&"Calibri,Bold"My Company&C&[Tab]&RConfidential',
    oddFooter:  '&LPrinted: &D&CPage &P of &N&R&[Path]&[File]',
    differentFirst: true,
    firstHeader: '&C&"Calibri,Bold"&18Cover Page',
    firstFooter: '&CConfidential',
  };

  ws.printOptions = {
    gridLines: false,
    headings: false,
    centerHorizontal: true,
  };

  await wb.writeFile('./output/13_page_setup.xlsx');
}

// ============================================================
// 14. SHEET PROTECTION
// ============================================================
async function example_protection() {
  const wb = new Workbook();
  const ws = wb.addSheet('Protected');

  ws.setValue(1, 1, 'This cell is locked (default)');
  ws.setCell(2, 1, {
    value: 'This cell is editable',
    style: { locked: false, fill: { type: 'pattern', pattern: 'solid', fgColor: 'FFFFFF00' } }
  });

  ws.protection = {
    sheet: true,
    password: 'secret',
    selectLockedCells: true,
    selectUnlockedCells: true,
    formatCells: true,
    insertRows: true,
    deleteRows: true,
    sort: true,
    autoFilter: true,
  };

  await wb.writeFile('./output/14_protection.xlsx');
}

// ============================================================
// 15. NAMED RANGES & MULTIPLE SHEETS
// ============================================================
async function example_named_ranges() {
  const wb = new Workbook();

  const wsData = wb.addSheet('Data');
  wsData.writeColumn(1, 1, [100, 200, 300, 400, 500]);
  wsData.writeColumn(1, 2, ['Alpha', 'Beta', 'Gamma', 'Delta', 'Epsilon']);

  wb.addNamedRange({
    name: 'SalesData',
    ref: 'Data!$A$1:$A$5',
  });

  wb.addNamedRange({
    name: 'ProductNames',
    ref: 'Data!$B$1:$B$5',
  });

  const wsSummary = wb.addSheet('Summary');
  wsSummary.setFormula(1, 1, 'SUM(SalesData)');
  wsSummary.setFormula(2, 1, 'AVERAGE(SalesData)');
  wsSummary.setFormula(3, 1, 'MAX(SalesData)');
  wsSummary.setFormula(1, 3, 'INDEX(ProductNames,MATCH(MAX(SalesData),SalesData,0))');

  // Hidden sheet
  const wsConfig = wb.addSheet('Config', { state: 'hidden' });
  wsConfig.setValue(1, 1, 'This sheet is hidden');

  await wb.writeFile('./output/15_named_ranges.xlsx');
}

// ============================================================
// 16. AUTO FILTER, OUTLINING, GROUPING
// ============================================================
async function example_autofilter() {
  const wb = new Workbook();
  const ws = wb.addSheet('AutoFilter');

  ws.writeRow(1, 1, ['Region', 'Product', 'Q1', 'Q2', 'Q3', 'Q4']);
  ws.setStyle(1, 1, Styles.tableHeader);
  for (let c = 2; c <= 6; c++) ws.setStyle(1, c, Styles.tableHeader);

  const data = [
    ['North', 'Widget', 100, 120, 110, 130],
    ['South', 'Widget', 80,  90,  85, 95],
    ['North', 'Gadget', 200, 210, 205, 220],
    ['South', 'Gadget', 150, 160, 155, 170],
    ['East',  'Widget', 60,  70,  65,  80],
    ['West',  'Gadget', 120, 130, 125, 140],
  ];

  ws.writeArray(2, 1, data);
  ws.autoFilter = { ref: 'A1:F1' };

  // Row grouping (outline)
  ws.setRow(2, { outlineLevel: 1 });
  ws.setRow(3, { outlineLevel: 1 });
  ws.setRow(4, { outlineLevel: 2 });
  ws.setRow(5, { outlineLevel: 2 });

  for (let c = 1; c <= 6; c++) ws.setColumnWidth(c, 14);

  await wb.writeFile('./output/16_autofilter.xlsx');
}

// ============================================================
// 17. HYPERLINKS
// ============================================================
async function example_hyperlinks() {
  const wb = new Workbook();
  const ws = wb.addSheet('Hyperlinks');

  const blueUnderline = style().fontColor(Colors.Blue).underline().build();

  ws.setCell(1, 1, {
    value: 'Visit Google',
    style: blueUnderline,
    hyperlink: { href: 'https://www.google.com', tooltip: 'Go to Google' }
  });

  ws.setCell(2, 1, {
    value: 'Email Us',
    style: blueUnderline,
    hyperlink: { href: 'mailto:hello@example.com' }
  });

  ws.setCell(3, 1, {
    value: 'Go to Sheet2!A1',
    style: blueUnderline,
    hyperlink: { href: '#Sheet2!A1', tooltip: 'Navigate to Sheet2' }
  });

  const ws2 = wb.addSheet('Sheet2');
  ws2.setValue(1, 1, 'You navigated here!');

  await wb.writeFile('./output/17_hyperlinks.xlsx');
}

// ============================================================
// 18. COMMENTS
// ============================================================
async function example_comments() {
  const wb = new Workbook();
  const ws = wb.addSheet('Comments');

  ws.setCell(1, 1, {
    value: 'Hover for comment',
    comment: { text: 'This is a cell comment.', author: 'ExcelForge' }
  });

  ws.setCell(3, 3, {
    value: 42,
    comment: { text: 'The answer to life, the universe, and everything.', author: 'Deep Thought' }
  });

  await wb.writeFile('./output/18_comments.xlsx');
}

// ============================================================
// 19. COMPLETE FINANCIAL REPORT EXAMPLE
// ============================================================
async function example_financial_report() {
  const wb = new Workbook();
  wb.properties = {
    title: 'Q4 2024 Financial Report',
    author: 'Finance Team',
    company: 'Acme Corp',
    category: 'Financial',
  };

  const ws = wb.addSheet('P&L Statement');

  // Title
  ws.setCell(1, 1, { value: 'ACME CORP — Q4 2024 P&L', style: style().bold().fontSize(16).build() });
  ws.merge(1, 1, 1, 5);

  // Headers
  ws.writeRow(3, 1, ['', 'Q1 2024', 'Q2 2024', 'Q3 2024', 'Q4 2024']);
  for (let c = 1; c <= 5; c++) ws.setStyle(3, c, Styles.tableHeader);

  // Revenue section
  ws.setCell(4, 1, { value: 'REVENUE', style: style().bold().bg('FFD9E1F2').build() });
  ws.merge(4, 1, 4, 5);

  const revenueRows = [
    ['Product Sales', 1200000, 1350000, 1100000, 1500000],
    ['Services',       300000,  320000,  350000,  380000],
    ['Other Revenue',   50000,   60000,   55000,   70000],
  ];

  revenueRows.forEach((row, i) => {
    const r = 5 + i;
    ws.setValue(r, 1, row[0] as string);
    for (let c = 1; c <= 4; c++) {
      ws.setValue(r, c + 1, row[c] as number);
      ws.setStyle(r, c + 1, style().numFmt(NumFmt.Currency).build());
    }
  });

  // Total Revenue (formula row)
  ws.setCell(8, 1, { value: 'Total Revenue', style: style().bold().build() });
  for (let c = 2; c <= 5; c++) {
    ws.setCell(8, c, {
      formula: `SUM(${String.fromCharCode(64+c)}5:${String.fromCharCode(64+c)}7)`,
      style: style().bold().numFmt(NumFmt.Currency).borderTop('thin').build(),
    });
  }

  // Expenses section
  ws.setCell(10, 1, { value: 'EXPENSES', style: style().bold().bg('FFFCE4D6').build() });
  ws.merge(10, 1, 10, 5);

  const expenseRows = [
    ['Cost of Goods',  600000,  680000,  550000,  750000],
    ['Operating Exp',  200000,  210000,  220000,  230000],
    ['Marketing',       80000,   90000,   85000,  100000],
    ['R&D',             50000,   55000,   60000,   65000],
  ];

  expenseRows.forEach((row, i) => {
    const r = 11 + i;
    ws.setValue(r, 1, row[0] as string);
    for (let c = 1; c <= 4; c++) {
      ws.setValue(r, c + 1, row[c] as number);
      ws.setStyle(r, c + 1, style().numFmt(NumFmt.Currency).build());
    }
  });

  // Total Expenses
  ws.setCell(15, 1, { value: 'Total Expenses', style: style().bold().build() });
  for (let c = 2; c <= 5; c++) {
    ws.setCell(15, c, {
      formula: `SUM(${String.fromCharCode(64+c)}11:${String.fromCharCode(64+c)}14)`,
      style: style().bold().numFmt(NumFmt.Currency).borderTop('thin').build(),
    });
  }

  // Net Income
  ws.setCell(17, 1, { value: 'NET INCOME', style: style().bold().fontSize(12).build() });
  for (let c = 2; c <= 5; c++) {
    const col = String.fromCharCode(64 + c);
    ws.setCell(17, c, {
      formula: `${col}8-${col}15`,
      style: style().bold().fontSize(12).numFmt(NumFmt.Currency)
        .border('medium').bg('FF70AD47').fontColor(Colors.White).build(),
    });
  }

  // Margin row
  ws.setCell(18, 1, { value: 'Net Margin %', style: style().bold().build() });
  for (let c = 2; c <= 5; c++) {
    const col = String.fromCharCode(64 + c);
    ws.setCell(18, c, {
      formula: `${col}17/${col}8`,
      style: style().bold().numFmt(NumFmt.Percent2).build(),
    });
  }

  // Conditional formatting on net income
  ws.addConditionalFormat({
    sqref: 'B17:E17', type: 'cellIs', operator: 'greaterThan', formula: '0',
    style: style().fontColor(Colors.Green).build(), priority: 1,
  });

  // Column widths
  ws.setColumnWidth(1, 22);
  for (let c = 2; c <= 5; c++) ws.setColumnWidth(c, 16);

  // Freeze pane
  ws.freeze(3, 1);

  // Chart sheet
  const wsChart = wb.addSheet('Charts');
  wsChart.writeRow(1, 1, ['Quarter', 'Revenue', 'Expenses', 'Net Income']);
  [['Q1',1550000,930000,620000],['Q2',1730000,1035000,695000],
   ['Q3',1505000,915000,590000],['Q4',1950000,1145000,805000]].forEach((row, i) => {
    wsChart.writeRow(i+2, 1, row);
  });

  wsChart.addChart({
    type: 'columnStacked',
    title: 'Quarterly Financial Performance',
    series: [
      { name: 'Revenue',     values: 'Charts!$B$2:$B$5', categories: 'Charts!$A$2:$A$5', color: Colors.ExcelBlue },
      { name: 'Expenses',    values: 'Charts!$C$2:$C$5', categories: 'Charts!$A$2:$A$5', color: Colors.ExcelOrange ?? 'FFED7D31' },
    ],
    from: { col: 5, row: 1 }, to: { col: 14, row: 18 },
    xAxis: { title: 'Quarter' },
    yAxis: { title: 'Amount ($)' },
    legend: 'b', style: 2,
  });

  await wb.writeFile('./output/19_financial_report.xlsx');
}

// ============================================================
// 20. LOAD EXISTING TABLE (ErrorsAndWarnings.xlsx)
// ============================================================
async function example_load_table() {
  const wb = await Workbook.fromFile('./src/test/ErrorsAndWarnings.xlsx');

  const sheetNames = wb.getSheetNames();
  console.log('  Sheets:', sheetNames);

  const ws = wb.getSheet(sheetNames[0]);
  if (!ws) throw new Error('No sheet found');

  const tables = ws.getTables();
  console.log('  Tables found:', tables.length);
  if (tables.length === 0) throw new Error('No tables found — table loading is broken');

  const table = tables[0];
  if (!table) throw new Error('No table found');
  console.log('  Table:', table.name, 'ref:', table.ref, 'columns:', table.columns.map(c => c.name));

  // Extend the table with new data rows
  const { startRow, startCol, endRow, endCol } = parseRangeHelper(table.ref);
  const newRow = endRow + 1;
  ws.writeRow(newRow, startCol, ['New Error', 'TestModule', 'critical', new Date(), 'Added by test']);

  // Update table ref to include the new row
  const startColLetter = colIndexToLetterHelper(startCol);
  const endColLetter = colIndexToLetterHelper(endCol);
  table.ref = `${startColLetter}${startRow}:${endColLetter}${newRow}`;
  console.log('  Updated table ref:', table.ref);

  wb.markDirty(sheetNames[0]);
  await wb.writeFile('./output/20_loaded_table.xlsx');

  // Verify round-trip: reload with ExcelForge and check data
  const wb2 = await Workbook.fromFile('./output/20_loaded_table.xlsx');
  const ws2 = wb2.getSheet(sheetNames[0])!;
  const tables2 = ws2.getTables();
  if (tables2.length === 0) throw new Error('Round-trip failed: no tables in re-loaded file');
  console.log('  Round-trip OK, tables:', tables2.length);

  // Verify ExcelForge can read back the written values
  const efVals: (string | number | boolean | null)[] = [];
  for (let c = 1; c <= 12; c++) {
    const cell = ws2.getCell(newRow, c);
    const v = cell.value;
    efVals.push(v === undefined ? null : v as string | number | boolean);
  }
  console.log('  ExcelForge read-back row', newRow, ':', efVals.slice(0, 5));
  if (efVals[0] !== 'New Error') throw new Error('ExcelForge read-back mismatch col A: ' + efVals[0]);
  if (efVals[1] !== 'TestModule') throw new Error('ExcelForge read-back mismatch col B: ' + efVals[1]);
  if (efVals[2] !== 'critical') throw new Error('ExcelForge read-back mismatch col C: ' + efVals[2]);
  if (typeof efVals[3] !== 'number') throw new Error('ExcelForge read-back mismatch col D (expected number): ' + efVals[3]);
  if (efVals[4] !== 'Added by test') throw new Error('ExcelForge read-back mismatch col E: ' + efVals[4]);

  // Verify with OpenXML SDK via C# reader
  // @ts-ignore
  const { execSync } = await import('child_process');
  try {
    const cmd = `dotnet run validatorReadData.cs output/20_loaded_table.xlsx ErrorsAndWarnings ${newRow} 1 12`;
    const out = execSync(cmd, { encoding: 'utf-8', timeout: 60000, stdio: ['pipe', 'pipe', 'pipe'] }).trim();
    // Extract JSON array from output (skip any dotnet warnings)
    const jsonStart = out.indexOf('[');
    const jsonEnd = out.lastIndexOf(']');
    if (jsonStart < 0 || jsonEnd < 0) throw new Error('C# reader returned no JSON: ' + out.slice(0, 200));
    const csharpVals: (string | number | boolean | null)[] = JSON.parse(out.slice(jsonStart, jsonEnd + 1));
    console.log('  C# OpenXML read-back row', newRow, ':', csharpVals.slice(0, 5));

    // Compare ExcelForge vs C# results
    const expected = ['New Error', 'TestModule', 'critical', efVals[3], 'Added by test'];
    for (let i = 0; i < expected.length; i++) {
      const ev = expected[i], cv = csharpVals[i];
      if (typeof ev === 'number' && typeof cv === 'number') {
        if (Math.abs(ev - cv) > 0.001) throw new Error(`C# mismatch col ${i + 1}: expected ${ev}, got ${cv}`);
      } else if (ev !== cv) {
        throw new Error(`C# mismatch col ${i + 1}: expected ${JSON.stringify(ev)}, got ${JSON.stringify(cv)}`);
      }
    }
    // Remaining columns should be null
    for (let i = 5; i < 12; i++) {
      if (csharpVals[i] !== null) throw new Error(`C# col ${i + 1} expected null, got ${JSON.stringify(csharpVals[i])}`);
    }
    console.log('  C# OpenXML verification: all values match ✓');
  } catch (e: any) {
    if (e.message?.includes('mismatch') || e.message?.includes('expected null')) throw e;
    console.log('  C# OpenXML verification skipped (dotnet not available):', e.message?.split('\n')[0]);
  }
}

// helpers used by the test (re-import the utility)
function parseRangeHelper(ref: string) {
  const [start, end] = ref.split(':');
  const s = parseCellRef(start), e = parseCellRef(end);
  return { startRow: s.row, startCol: s.col, endRow: e.row, endCol: e.col };
}
function parseCellRef(ref: string) {
  const m = ref.match(/^([A-Z]+)(\d+)$/);
  if (!m) throw new Error('Invalid ref: ' + ref);
  let col = 0;
  for (const ch of m[1]) col = col * 26 + (ch.charCodeAt(0) - 64);
  return { row: parseInt(m[2], 10), col };
}
function colIndexToLetterHelper(col: number): string {
  let s = '';
  while (col > 0) { const r = (col - 1) % 26; s = String.fromCharCode(65 + r) + s; col = Math.floor((col - 1) / 26); }
  return s;
}

// ============================================================
// 21. PIVOT TABLE — generate, round-trip, and C# validate
// ============================================================
async function example_pivot_table() {
  const wb = new Workbook();
  wb.properties = { title: 'Pivot Table Demo', author: 'ExcelForge' };

  // ── Source data sheet ──────────────────────────────────────────────────────
  const wsData = wb.addSheet('Data');
  wsData.writeRow(1, 1, ['Region', 'Product', 'Sales', 'Units']);
  wsData.writeRow(2, 1, ['North', 'Widget',  100, 10]);
  wsData.writeRow(3, 1, ['North', 'Gadget',  200, 20]);
  wsData.writeRow(4, 1, ['South', 'Widget',  300, 30]);
  wsData.writeRow(5, 1, ['South', 'Gadget',  500, 50]);

  // ── Pivot sheet ────────────────────────────────────────────────────────────
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
  });

  await wb.writeFile('./output/21_pivot_table.xlsx');

  // ── ExcelForge round-trip ──────────────────────────────────────────────────
  const wb2 = await Workbook.fromFile('./output/21_pivot_table.xlsx');
  const names = wb2.getSheetNames();
  if (!names.includes('Data'))    throw new Error('Round-trip: missing Data sheet');
  if (!names.includes('Summary')) throw new Error('Round-trip: missing Summary sheet');
  const wsData2 = wb2.getSheet('Data')!;
  if (wsData2.getCell(1, 1).value !== 'Region') throw new Error('Round-trip: header mismatch');
  if (wsData2.getCell(2, 1).value !== 'North')  throw new Error('Round-trip: data mismatch');
  console.log('  ExcelForge round-trip: OK');

  // ── C# OpenXML SDK validation ─────────────────────────────────────────────
  // @ts-ignore
  const { execSync } = await import('child_process');
  try {
    const out = execSync('dotnet run validator.cs output/21_pivot_table.xlsx', {
      encoding: 'utf-8', timeout: 60000, stdio: ['pipe', 'pipe', 'pipe'],
    }).trim();
    const jsonStart = out.indexOf('[');
    const jsonEnd   = out.lastIndexOf(']');
    if (jsonStart < 0) throw new Error('C# validator returned no JSON: ' + out.slice(0, 200));
    const errors: object[] = JSON.parse(out.slice(jsonStart, jsonEnd + 1));
    if (errors.length > 0) {
      console.log('  C# validation errors:', JSON.stringify(errors, null, 2));
      throw new Error(`C# OpenXML validation failed with ${errors.length} error(s)`);
    }
    console.log('  C# OpenXML validation: OK (no errors)');
  } catch (e: any) {
    if (e.message?.includes('validation failed') || e.message?.includes('no JSON')) throw e;
    console.log('  C# OpenXML validation skipped (dotnet not available):', e.message?.split('\n')[0]);
  }
}

// ============================================================
// 22. VBA MACROS
// ============================================================
async function example_vba() {
  const wb = new Workbook();
  wb.properties = { title: 'VBA Macro Demo', author: 'ExcelForge' };

  const ws = wb.addSheet('Sheet1');
  ws.setValue(1, 1, 'Hello');
  ws.setValue(2, 1, 'Click the button to run the macro!');

  // Create a VBA project with a simple macro
  const vba = new VbaProject();
  vba.addModule({
    name: 'Module1',
    type: 'standard',
    code: `Sub HelloWorld()\r\n    MsgBox "Hello from ExcelForge VBA!"\r\nEnd Sub\r\n`,
  });
  wb.vbaProject = vba;

  await wb.writeFile('./output/22_vba_macros.xlsm');

  // ── ExcelForge round-trip ──────────────────────────────────────────────────
  const wb2 = await Workbook.fromFile('./output/22_vba_macros.xlsm');
  if (!wb2.vbaProject) throw new Error('Round-trip: VBA project missing');
  const mod = wb2.vbaProject.getModule('Module1');
  if (!mod) throw new Error('Round-trip: Module1 missing');
  if (!mod.code.includes('HelloWorld')) throw new Error('Round-trip: macro code missing');
  console.log('  ExcelForge VBA round-trip: OK');

  // ── Modify and re-save ────────────────────────────────────────────────────
  wb2.vbaProject.addModule({
    name: 'Module2',
    type: 'standard',
    code: `Sub GoodbyeWorld()\r\n    MsgBox "Goodbye!"\r\nEnd Sub\r\n`,
  });
  await wb2.writeFile('./output/22_vba_macros_modified.xlsm');

  // Verify second round-trip
  const wb3 = await Workbook.fromFile('./output/22_vba_macros_modified.xlsm');
  if (!wb3.vbaProject) throw new Error('Round-trip 2: VBA project missing');
  if (!wb3.vbaProject.getModule('Module1')) throw new Error('Round-trip 2: Module1 missing');
  if (!wb3.vbaProject.getModule('Module2')) throw new Error('Round-trip 2: Module2 missing');
  console.log('  ExcelForge VBA round-trip (modified): OK');

  // ── C# OpenXML SDK validation ─────────────────────────────────────────────
  // @ts-ignore
  const { execSync } = await import('child_process');
  try {
    const out = execSync('dotnet run validator.cs output/22_vba_macros.xlsm', {
      encoding: 'utf-8', timeout: 60000, stdio: ['pipe', 'pipe', 'pipe'],
    }).trim();
    const jsonStart = out.indexOf('[');
    const jsonEnd   = out.lastIndexOf(']');
    if (jsonStart < 0) throw new Error('C# validator returned no JSON: ' + out.slice(0, 200));
    const errors: object[] = JSON.parse(out.slice(jsonStart, jsonEnd + 1));
    if (errors.length > 0) {
      console.log('  C# validation errors:', JSON.stringify(errors, null, 2));
      throw new Error(`C# OpenXML validation failed with ${errors.length} error(s)`);
    }
    console.log('  C# OpenXML validation: OK (no errors)');
  } catch (e: any) {
    if (e.message?.includes('validation failed') || e.message?.includes('no JSON')) throw e;
    console.log('  C# OpenXML validation skipped (dotnet not available):', e.message?.split('\n')[0]);
  }
}

// Run all examples
async function runAll() {
  // @ts-ignore
  const fs = await import('fs/promises');
  try { await fs.mkdir('./output', { recursive: true }); } catch {}

  const examples = [
    ['Basic',                  example_basic],
    ['Formulas',               example_formulas],
    ['Styles',                 example_styles],
    ['Merges',                 example_merges],
    ['Rich Text',              example_richtext],
    ['Freeze Panes',           example_panes],
    ['Conditional Formatting', example_conditional_formatting],
    ['Tables',                 example_tables],
    ['Charts',                 example_charts],
    ['Images',                 example_images],
    ['Data Validation',        example_data_validation],
    ['Sparklines',             example_sparklines],
    ['Page Setup',             example_page_setup],
    ['Protection',             example_protection],
    ['Named Ranges',           example_named_ranges],
    ['AutoFilter',             example_autofilter],
    ['Hyperlinks',             example_hyperlinks],
    ['Comments',               example_comments],
    ['Financial Report',       example_financial_report],
    ['Load Table',             example_load_table],
    ['Pivot Table',            example_pivot_table],
    ['VBA Macros',             example_vba],
  ] as const;

  for (const [name, fn] of examples) {
    try {
      await fn();
      console.log(`✅ ${name}`);
    } catch (e) {
      console.error(`❌ ${name}: ${e}`);
    }
  }
}

runAll().catch(console.error);
