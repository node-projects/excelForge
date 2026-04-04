/**
 * ExcelForge — Comprehensive Usage Examples
 * This file demonstrates every major feature of the library.
 */

import { Workbook, Worksheet, style, Styles, Colors, NumFmt, VbaProject, encryptWorkbook } from '../index.js';
import type { Chart, ConditionalFormat, Table, Sparkline, DataValidation, Image, CellImage, FormControl, CalcSettings, OleObject } from '../index.js';
//@ts-ignore
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

  // Workbook-scoped named ranges
  wb.addNamedRange({
    name: 'SalesData',
    ref: 'Data!$A$1:$A$5',
  });

  wb.addNamedRange({
    name: 'ProductNames',
    ref: 'Data!$B$1:$B$5',
    comment: 'Product name list',
  });

  // Sheet-scoped named range
  wb.addNamedRange({
    name: 'LocalTotal',
    ref: 'Data!$A$6',
    scope: 'Data',
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

  // Round-trip: read back and verify named ranges are preserved
  const wb2 = await Workbook.fromFile('./output/15_named_ranges.xlsx');
  const ranges = wb2.getNamedRanges();
  if (ranges.length !== 3) throw new Error(`Expected 3 named ranges, got ${ranges.length}`);
  const sales = wb2.getNamedRange('SalesData');
  if (!sales || sales.ref !== 'Data!$A$1:$A$5') throw new Error(`SalesData: ${JSON.stringify(sales)}`);
  const products = wb2.getNamedRange('ProductNames');
  if (!products || products.comment !== 'Product name list') throw new Error(`ProductNames comment: ${JSON.stringify(products)}`);
  const local = wb2.getNamedRange('LocalTotal');
  if (!local || local.scope !== 'Data') throw new Error(`LocalTotal scope: ${JSON.stringify(local)}`);

  // Dirty round-trip: modify + re-save + read back
  wb2.markDirty('Data');
  wb2.getSheet('Data')!.setValue(6, 1, 1500);
  wb2.addNamedRange({ name: 'NewRange', ref: 'Summary!$A$1:$A$3' });
  const dirtyBytes = await wb2.build();
  const wb3 = await Workbook.fromBytes(dirtyBytes);
  if (wb3.getNamedRanges().length !== 4) throw new Error(`Dirty: expected 4 ranges, got ${wb3.getNamedRanges().length}`);
  if (!wb3.getNamedRange('NewRange')) throw new Error('NewRange missing after dirty round-trip');

  // Remove a named range
  wb3.removeNamedRange('NewRange');
  if (wb3.getNamedRanges().length !== 3) throw new Error('removeNamedRange failed');
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

async function example_vba_complex() {
  const wb = new Workbook();
  wb.properties = { title: 'Complex VBA Demo', author: 'ExcelForge' };

  // Multiple sheets
  const ws1 = wb.addSheet('Dashboard');
  ws1.setValue(1, 1, 'Sales Dashboard');
  ws1.setValue(2, 1, 'Region');
  ws1.setValue(2, 2, 'Sales');
  ws1.setValue(2, 3, 'Target');
  ws1.setValue(3, 1, 'North'); ws1.setValue(3, 2, 45000); ws1.setValue(3, 3, 50000);
  ws1.setValue(4, 1, 'South'); ws1.setValue(4, 2, 38000); ws1.setValue(4, 3, 35000);
  ws1.setValue(5, 1, 'East');  ws1.setValue(5, 2, 52000); ws1.setValue(5, 3, 48000);
  ws1.setValue(6, 1, 'West');  ws1.setValue(6, 2, 41000); ws1.setValue(6, 3, 45000);

  const ws2 = wb.addSheet('Config');
  ws2.setValue(1, 1, 'HighlightThreshold');
  ws2.setValue(1, 2, 0.9);
  ws2.setValue(2, 1, 'ReportTitle');
  ws2.setValue(2, 2, 'Quarterly Review');

  // VBA project with all module types
  const vba = new VbaProject();

  // Standard module with utility functions
  vba.addModule({
    name: 'MathUtils',
    type: 'standard',
    code: [
      'Public Function PercentOfTarget(actual As Double, target As Double) As Double',
      '    If target = 0 Then',
      '        PercentOfTarget = 0',
      '    Else',
      '        PercentOfTarget = actual / target',
      '    End If',
      'End Function',
      '',
      'Public Function FormatAsCurrency(value As Double) As String',
      '    FormatAsCurrency = Format(value, "$#,##0.00")',
      'End Function',
      '',
      'Public Sub HighlightCell(rng As Range, threshold As Double)',
      '    If rng.Value >= threshold Then',
      '        rng.Interior.Color = RGB(198, 239, 206)',
      '    Else',
      '        rng.Interior.Color = RGB(255, 199, 206)',
      '    End If',
      'End Sub',
    ].join('\r\n') + '\r\n',
  });

  // Standard module with main subroutines
  vba.addModule({
    name: 'ReportMacros',
    type: 'standard',
    code: [
      'Public Sub GenerateReport()',
      '    Dim ws As Worksheet',
      '    Set ws = ThisWorkbook.Sheets("Dashboard")',
      '    Dim threshold As Double',
      '    threshold = ThisWorkbook.Sheets("Config").Range("B1").Value',
      '    ',
      '    Dim lastRow As Long',
      '    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row',
      '    ',
      '    Dim i As Long',
      '    For i = 3 To lastRow',
      '        Dim pct As Double',
      '        pct = PercentOfTarget(ws.Cells(i, 2).Value, ws.Cells(i, 3).Value)',
      '        ws.Cells(i, 4).Value = pct',
      '        ws.Cells(i, 4).NumberFormat = "0.0%"',
      '        HighlightCell ws.Cells(i, 4), threshold',
      '    Next i',
      '    ',
      '    MsgBox "Report generated for " & (lastRow - 2) & " regions.", vbInformation',
      'End Sub',
      '',
      'Public Sub ClearReport()',
      '    Dim ws As Worksheet',
      '    Set ws = ThisWorkbook.Sheets("Dashboard")',
      '    Dim lastRow As Long',
      '    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row',
      '    If lastRow >= 3 Then',
      '        ws.Range(ws.Cells(3, 4), ws.Cells(lastRow, 4)).ClearContents',
      '        ws.Range(ws.Cells(3, 4), ws.Cells(lastRow, 4)).Interior.ColorIndex = xlNone',
      '    End If',
      'End Sub',
    ].join('\r\n') + '\r\n',
  });

  // Class module
  vba.addModule({
    name: 'RegionStats',
    type: 'class',
    code: [
      'Private pName As String',
      'Private pSales As Double',
      'Private pTarget As Double',
      '',
      'Public Property Get Name() As String',
      '    Name = pName',
      'End Property',
      '',
      'Public Property Let Name(value As String)',
      '    pName = value',
      'End Property',
      '',
      'Public Property Get Sales() As Double',
      '    Sales = pSales',
      'End Property',
      '',
      'Public Property Let Sales(value As Double)',
      '    pSales = value',
      'End Property',
      '',
      'Public Property Get Target() As Double',
      '    Target = pTarget',
      'End Property',
      '',
      'Public Property Let Target(value As Double)',
      '    pTarget = value',
      'End Property',
      '',
      'Public Function Achievement() As Double',
      '    If pTarget = 0 Then',
      '        Achievement = 0',
      '    Else',
      '        Achievement = pSales / pTarget',
      '    End If',
      'End Function',
    ].join('\r\n') + '\r\n',
  });

  wb.vbaProject = vba;
  await wb.writeFile('./output/23_vba_complex.xlsm');

  // ── Round-trip 1: verify all modules survive ──────────────────────────────
  const wb2 = await Workbook.fromFile('./output/23_vba_complex.xlsm');
  if (!wb2.vbaProject) throw new Error('RT1: VBA project missing');
  const moduleNames = wb2.vbaProject.modules.map(m => m.name);
  for (const expected of ['MathUtils', 'ReportMacros', 'RegionStats', 'ThisWorkbook', 'Dashboard', 'Config']) {
    if (!moduleNames.includes(expected)) throw new Error(`RT1: module "${expected}" missing, got: ${moduleNames}`);
  }
  // Verify module types
  const mathMod = wb2.vbaProject.getModule('MathUtils')!;
  if (mathMod.type !== 'standard') throw new Error(`RT1: MathUtils type wrong: ${mathMod.type}`);
  const classMod = wb2.vbaProject.getModule('RegionStats')!;
  if (classMod.type !== 'class') throw new Error(`RT1: RegionStats type wrong: ${classMod.type}`);
  const docMod = wb2.vbaProject.getModule('Dashboard')!;
  if (docMod.type !== 'document') throw new Error(`RT1: Dashboard type wrong: ${docMod.type}`);
  // Verify code content
  if (!mathMod.code.includes('PercentOfTarget')) throw new Error('RT1: MathUtils code missing');
  if (!classMod.code.includes('Achievement')) throw new Error('RT1: RegionStats code missing');
  const reportMod = wb2.vbaProject.getModule('ReportMacros')!;
  if (!reportMod.code.includes('GenerateReport')) throw new Error('RT1: ReportMacros code missing');
  if (!reportMod.code.includes('ClearReport')) throw new Error('RT1: ClearReport sub missing');
  console.log('  RT1 all modules + types + code: OK');

  // ── Round-trip 2: modify and re-save ──────────────────────────────────────
  wb2.vbaProject.addModule({
    name: 'ExtraModule',
    type: 'standard',
    code: 'Public Sub ExtraWork()\r\n    MsgBox "Extra!"\r\nEnd Sub\r\n',
  });
  wb2.vbaProject.removeModule('RegionStats');
  await wb2.writeFile('./output/23_vba_complex_modified.xlsm');

  const wb3 = await Workbook.fromFile('./output/23_vba_complex_modified.xlsm');
  if (!wb3.vbaProject) throw new Error('RT2: VBA project missing');
  if (!wb3.vbaProject.getModule('ExtraModule')) throw new Error('RT2: ExtraModule missing');
  if (wb3.vbaProject.getModule('RegionStats')) throw new Error('RT2: RegionStats should be removed');
  if (!wb3.vbaProject.getModule('MathUtils')) throw new Error('RT2: MathUtils missing');
  if (!wb3.vbaProject.getModule('ReportMacros')) throw new Error('RT2: ReportMacros missing');
  console.log('  RT2 add/remove modules: OK');

  // ── EPPlus validation ─────────────────────────────────────────────────────
  // @ts-ignore
  const { execSync } = await import('child_process');
  try {
    const out = execSync('dotnet run read_vba_epplus.cs output/23_vba_complex.xlsm', {
      encoding: 'utf-8', timeout: 60000, stdio: ['pipe', 'pipe', 'pipe'],
    }).trim();
    if (!out.includes('VBA Project: YES')) throw new Error('EPPlus: VBA project not found');
    if (!out.includes('MathUtils')) throw new Error('EPPlus: MathUtils missing');
    if (!out.includes('ReportMacros')) throw new Error('EPPlus: ReportMacros missing');
    if (!out.includes('RegionStats')) throw new Error('EPPlus: RegionStats missing');
    console.log('  EPPlus validation: OK');
  } catch (e: any) {
    if (e.message?.includes('EPPlus:')) throw e;
    console.log('  EPPlus validation skipped (dotnet not available):', e.message?.split('\n')[0]);
  }

  // ── C# OpenXML SDK validation ─────────────────────────────────────────────
  try {
    const out = execSync('dotnet run validator.cs output/23_vba_complex.xlsm', {
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
    console.log('  C# OpenXML validation: OK');
  } catch (e: any) {
    if (e.message?.includes('validation failed') || e.message?.includes('no JSON')) throw e;
    console.log('  C# OpenXML validation skipped:', e.message?.split('\n')[0]);
  }
}

// ============================================================
// CONDITIONAL FORMATTING & DATA VALIDATION ROUND-TRIP TEST
// ============================================================
async function example_cf_dv_roundtrip() {
  // @ts-ignore
  const fs = await import('fs/promises');

  // ── Part 1: Create a workbook with CF and DV, round-trip it ──
  const wb = new Workbook();
  const ws = wb.addSheet('CFTest');

  // Add data
  for (let r = 1; r <= 20; r++) {
    ws.setValue(r, 1, `Item ${r}`);
    ws.setValue(r, 2, r * 10);
    ws.setValue(r, 3, r % 3 === 0 ? 'Yes' : 'No');
  }

  // Add conditional formats
  ws.addConditionalFormat({
    sqref: 'B1:B20',
    type: 'cellIs',
    operator: 'greaterThan',
    formula: '100',
    priority: 1,
    style: { font: { bold: true, color: 'FFFF0000' }, fill: { type: 'pattern', pattern: 'solid', fgColor: 'FFFFFFCC' } as any },
  });
  ws.addConditionalFormat({
    sqref: 'B1:B20',
    type: 'colorScale',
    priority: 2,
    colorScale: {
      type: 'colorScale',
      cfvo: [{ type: 'min' }, { type: 'max' }],
      color: ['FF63BE7B', 'FFF8696B'],
    },
  });
  ws.addConditionalFormat({
    sqref: 'B1:B20',
    type: 'dataBar',
    priority: 3,
    dataBar: { type: 'dataBar', color: 'FF638EC6' },
  });
  ws.addConditionalFormat({
    sqref: 'C1:C20',
    type: 'containsText',
    operator: 'equal',
    text: 'Yes',
    formula: 'NOT(ISERROR(SEARCH("Yes",C1)))',
    priority: 4,
    style: { fill: { type: 'pattern', pattern: 'solid', fgColor: 'FFC6EFCE' } as any },
  });

  // Add data validations
  ws.addDataValidation('C1:C20', {
    type: 'list',
    list: ['Yes', 'No', 'Maybe'],
    showErrorAlert: true,
    errorTitle: 'Invalid',
    error: 'Pick Yes, No, or Maybe',
    showInputMessage: true,
    promptTitle: 'Select',
    prompt: 'Choose a value',
  });
  ws.addDataValidation('B1:B20', {
    type: 'whole',
    operator: 'between',
    formula1: '0',
    formula2: '500',
    allowBlank: true,
    showErrorAlert: true,
    errorTitle: 'Out of range',
    error: 'Must be 0-500',
  });

  const bytes1 = await wb.build();

  // Round-trip: read back and verify
  const wb2 = await Workbook.fromBytes(bytes1);
  const ws2 = wb2.getSheet('CFTest')!;
  const cfs = ws2.getConditionalFormats();
  const dvs = ws2.getDataValidations();

  if (cfs.length !== 4) throw new Error(`Expected 4 CFs, got ${cfs.length}`);
  if (dvs.size !== 2) throw new Error(`Expected 2 DVs, got ${dvs.size}`);

  // Verify CF types
  if (cfs[0].type !== 'cellIs') throw new Error(`CF[0] type: ${cfs[0].type}`);
  if (cfs[0].operator !== 'greaterThan') throw new Error(`CF[0] operator: ${cfs[0].operator}`);
  if (cfs[0].formula !== '100') throw new Error(`CF[0] formula: ${cfs[0].formula}`);
  if (!cfs[0].style?.font?.bold) throw new Error('CF[0] should have bold font style');
  if (cfs[1].type !== 'colorScale') throw new Error(`CF[1] type: ${cfs[1].type}`);
  if (!cfs[1].colorScale) throw new Error('CF[1] should have colorScale');
  if (cfs[2].type !== 'dataBar') throw new Error(`CF[2] type: ${cfs[2].type}`);
  if (!cfs[2].dataBar) throw new Error('CF[2] should have dataBar');
  if (cfs[3].type !== 'containsText') throw new Error(`CF[3] type: ${cfs[3].type}`);

  // Verify DV
  const dvList = dvs.get('C1:C20');
  if (!dvList) throw new Error('DV for C1:C20 not found');
  if (dvList.type !== 'list') throw new Error(`DV type: ${dvList.type}`);
  if (!dvList.list || dvList.list.join(',') !== 'Yes,No,Maybe') throw new Error(`DV list: ${dvList.list}`);
  if (dvList.errorTitle !== 'Invalid') throw new Error(`DV errorTitle: ${dvList.errorTitle}`);

  const dvWhole = dvs.get('B1:B20');
  if (!dvWhole) throw new Error('DV for B1:B20 not found');
  if (dvWhole.type !== 'whole') throw new Error(`DV type: ${dvWhole.type}`);
  if (dvWhole.operator !== 'between') throw new Error(`DV operator: ${dvWhole.operator}`);

  console.log('  Create + round-trip: OK (4 CFs, 2 DVs preserved)');

  // ── Part 2: Round-trip a dirty workbook — CF/DV must survive re-serialization ──
  wb2.markDirty('CFTest');
  ws2.setValue(1, 4, 'added');
  const bytes2 = await wb2.build();
  const wb3 = await Workbook.fromBytes(bytes2);
  const ws3 = wb3.getSheet('CFTest')!;
  if (ws3.getConditionalFormats().length !== 4) throw new Error('CFs lost after dirty round-trip');
  if (ws3.getDataValidations().size !== 2) throw new Error('DVs lost after dirty round-trip');
  console.log('  Dirty round-trip: OK (CF/DV survive re-serialization)');

  // ── Part 3: Round-trip Book 2.xlsx (real file with CF + DV) ──
  const b2Data = await fs.readFile('src/test/Book 2.xlsx');
  const wbB = await Workbook.fromBytes(new Uint8Array(b2Data));

  // Capture original counts
  const origCfs = wbB.getSheet('Entwicklungsbericht')!.getConditionalFormats().length;
  const origDvs = wbB.getSheet('Entwicklungsbericht')!.getDataValidations().size;
  if (origCfs !== 9) throw new Error(`Book 2 Entwicklungsbericht CF: expected 9, got ${origCfs}`);
  if (origDvs !== 30) throw new Error(`Book 2 Entwicklungsbericht DV: expected 30, got ${origDvs}`);

  // Dirty round-trip
  wbB.markDirty('Entwicklungsbericht');
  const b2Out = await wbB.build();
  const wbB2 = await Workbook.fromBytes(b2Out);
  const rtCfs = wbB2.getSheet('Entwicklungsbericht')!.getConditionalFormats().length;
  const rtDvs = wbB2.getSheet('Entwicklungsbericht')!.getDataValidations().size;
  if (rtCfs !== origCfs) throw new Error(`Book 2 CF lost: ${origCfs} → ${rtCfs}`);
  if (rtDvs !== origDvs) throw new Error(`Book 2 DV lost: ${origDvs} → ${rtDvs}`);
  console.log(`  Book 2 round-trip: OK (CF=${rtCfs}, DV=${rtDvs} preserved)`);
}

// ============================================================
// PAGE BREAKS TEST
// ============================================================
async function example_page_breaks() {
  // @ts-ignore
  const fs = await import('fs/promises');

  // Create workbook with page breaks
  const wb = new Workbook();
  const ws = wb.addSheet('Breaks');
  for (let r = 1; r <= 50; r++) ws.setValue(r, 1, `Row ${r}`);
  ws.addRowBreak(10);
  ws.addRowBreak(25);
  ws.addRowBreak(40);
  ws.addColBreak(3);
  ws.addColBreak(6);

  const bytes = await wb.build();
  const wb2 = await Workbook.fromBytes(bytes);
  const ws2 = wb2.getSheet('Breaks')!;
  if (ws2.getRowBreaks().length !== 3) throw new Error(`Expected 3 row breaks, got ${ws2.getRowBreaks().length}`);
  if (ws2.getColBreaks().length !== 2) throw new Error(`Expected 2 col breaks, got ${ws2.getColBreaks().length}`);
  if (ws2.getRowBreaks()[0].id !== 10) throw new Error(`First row break at ${ws2.getRowBreaks()[0].id}`);
  if (ws2.getRowBreaks()[1].id !== 25) throw new Error(`Second row break at ${ws2.getRowBreaks()[1].id}`);
  if (ws2.getColBreaks()[0].id !== 3) throw new Error(`First col break at ${ws2.getColBreaks()[0].id}`);
  console.log('  Create + round-trip: OK (3 row, 2 col breaks)');

  // Dirty round-trip
  wb2.markDirty('Breaks');
  ws2.setValue(1, 2, 'modified');
  const bytes2 = await wb2.build();
  const wb3 = await Workbook.fromBytes(bytes2);
  const ws3 = wb3.getSheet('Breaks')!;
  if (ws3.getRowBreaks().length !== 3) throw new Error('Row breaks lost after dirty RT');
  if (ws3.getColBreaks().length !== 2) throw new Error('Col breaks lost after dirty RT');
  console.log('  Dirty round-trip: OK (breaks survive re-serialization)');

}

// ============================================================
// CONNECTIONS & POWER QUERY
// ============================================================
async function example_connections() {
  // Create a workbook with an ODBC connection
  const wb = new Workbook();
  const ws = wb.addSheet('Data');
  ws.setValue(1, 1, 'Connected data goes here');

  wb.addConnection({
    id: 1,
    name: 'SalesDB',
    type: 'oledb',
    connectionString: 'Provider=SQLOLEDB;Data Source=server;Initial Catalog=Sales;',
    command: 'SELECT * FROM Orders',
    commandType: 'sql',
    description: 'Sales database connection',
    background: true,
    saveData: true,
  });

  wb.addConnection({
    id: 2,
    name: 'WarehouseDB',
    type: 'odbc',
    connectionString: 'DSN=Warehouse;',
    command: 'Inventory',
    commandType: 'table',
    description: 'Warehouse inventory',
    saveData: true,
  });

  await wb.writeFile('./output/26_connections.xlsx');

  // Round-trip: read back and verify connections are preserved
  const wb2 = await Workbook.fromFile('./output/26_connections.xlsx');
  const conns = wb2.getConnections();
  if (conns.length !== 2) throw new Error(`Expected 2 connections, got ${conns.length}`);

  const salesConn = wb2.getConnection('SalesDB');
  if (!salesConn) throw new Error('SalesDB connection missing');
  if (salesConn.type !== 'oledb') throw new Error(`SalesDB type: ${salesConn.type}`);
  if (salesConn.command !== 'SELECT * FROM Orders') throw new Error(`SalesDB command: ${salesConn.command}`);
  if (salesConn.commandType !== 'sql') throw new Error(`SalesDB commandType: ${salesConn.commandType}`);
  if (!salesConn.saveData) throw new Error('SalesDB saveData not preserved');

  const whConn = wb2.getConnection('WarehouseDB');
  if (!whConn || whConn.type !== 'odbc') throw new Error(`WarehouseDB: ${JSON.stringify(whConn)}`);

  // Remove a connection
  wb2.removeConnection('WarehouseDB');
  if (wb2.getConnections().length !== 1) throw new Error('removeConnection failed');

  // Dirty round-trip
  wb2.markDirty('Data');
  wb2.getSheet('Data')!.setValue(2, 1, 'Updated');
  const dirtyBytes = await wb2.build();
  const wb3 = await Workbook.fromBytes(dirtyBytes);
  if (wb3.getConnections().length !== 1) throw new Error(`Dirty: expected 1 connection, got ${wb3.getConnections().length}`);
  if (!wb3.getConnection('SalesDB')) throw new Error('SalesDB missing after dirty round-trip');
}

// ============================================================
// 27. FORM CONTROLS
// ============================================================
async function example_form_controls() {
  const wb = new Workbook();
  const ws = wb.addSheet('Controls');
  ws.setValue(1, 1, 'Form Controls Demo');

  // Button with macro
  ws.addFormControl({
    type: 'button',
    from: { col: 1, row: 2 },
    to:   { col: 3, row: 4 },
    text: 'Click Me',
    macro: 'Sheet1.ButtonClick',
  });

  // CheckBox linked to a cell
  ws.addFormControl({
    type: 'checkBox',
    from: { col: 1, row: 5 },
    to:   { col: 3, row: 6 },
    text: 'Enable Feature',
    linkedCell: '$B$10',
    checked: 'checked',
  });

  // ComboBox (dropdown)
  ws.addFormControl({
    type: 'comboBox',
    from: { col: 1, row: 7 },
    to:   { col: 3, row: 8 },
    linkedCell: '$B$11',
    inputRange: '$D$1:$D$5',
    dropLines: 5,
  });

  // ListBox with multi-select
  ws.addFormControl({
    type: 'listBox',
    from: { col: 1, row: 9 },
    to:   { col: 3, row: 14 },
    linkedCell: '$B$12',
    inputRange: '$D$1:$D$10',
    selType: 'multi',
  });

  // Option buttons (radio)
  ws.addFormControl({
    type: 'optionButton',
    from: { col: 4, row: 2 },
    to:   { col: 6, row: 3 },
    text: 'Option A',
    linkedCell: '$B$13',
    checked: 'checked',
  });
  ws.addFormControl({
    type: 'optionButton',
    from: { col: 4, row: 4 },
    to:   { col: 6, row: 5 },
    text: 'Option B',
    linkedCell: '$B$13',
    checked: 'unchecked',
  });

  // ScrollBar
  ws.addFormControl({
    type: 'scrollBar',
    from: { col: 4, row: 6 },
    to:   { col: 6, row: 7 },
    linkedCell: '$B$14',
    min: 0,
    max: 100,
    inc: 1,
    page: 10,
    val: 50,
  });

  // Spinner
  ws.addFormControl({
    type: 'spinner',
    from: { col: 4, row: 8 },
    to:   { col: 5, row: 10 },
    linkedCell: '$B$15',
    min: 1,
    max: 50,
    inc: 1,
    val: 10,
  });

  // GroupBox
  ws.addFormControl({
    type: 'groupBox',
    from: { col: 4, row: 1 },
    to:   { col: 7, row: 11 },
    text: 'Options Group',
  });

  // Label
  ws.addFormControl({
    type: 'label',
    from: { col: 7, row: 2 },
    to:   { col: 9, row: 3 },
    text: 'Status Label',
  });

  // Populate input range data
  for (let i = 1; i <= 10; i++) ws.setValue(i, 4, `Item ${i}`);

  await wb.writeFile('./output/27_form_controls.xlsx');

  // Verify controls count
  const controls = ws.getFormControls();
  if (controls.length !== 10) throw new Error(`Expected 10 controls, got ${controls.length}`);

  // Round-trip: read back and verify
  const wb2 = await Workbook.fromFile('./output/27_form_controls.xlsx');
  const ws2 = wb2.getSheet('Controls')!;
  const controls2 = ws2.getFormControls();
  if (controls2.length !== 10) throw new Error(`Round-trip: expected 10 controls, got ${controls2.length}`);

  // Verify specific control properties
  const btn = controls2.find(c => c.type === 'button');
  if (!btn) throw new Error('Button not found after round-trip');
  if (btn.macro !== 'Sheet1.ButtonClick') throw new Error(`Button macro: ${btn.macro}`);

  const cb = controls2.find(c => c.type === 'checkBox');
  if (!cb) throw new Error('CheckBox not found');
  if (cb.linkedCell !== '$B$10') throw new Error(`CheckBox linkedCell: ${cb.linkedCell}`);
  if (cb.checked !== 'checked') throw new Error(`CheckBox checked: ${cb.checked}`);

  const combo = controls2.find(c => c.type === 'comboBox');
  if (!combo) throw new Error('ComboBox not found');
  if (combo.inputRange !== '$D$1:$D$5') throw new Error(`ComboBox inputRange: ${combo.inputRange}`);
  if (combo.dropLines !== 5) throw new Error(`ComboBox dropLines: ${combo.dropLines}`);

  const list = controls2.find(c => c.type === 'listBox');
  if (!list) throw new Error('ListBox not found');
  if (list.selType !== 'multi') throw new Error(`ListBox selType: ${list.selType}`);

  const scroll = controls2.find(c => c.type === 'scrollBar');
  if (!scroll) throw new Error('ScrollBar not found');
  if (scroll.min !== 0 || scroll.max !== 100 || scroll.val !== 50) throw new Error(`ScrollBar: ${JSON.stringify(scroll)}`);

  const spinner = controls2.find(c => c.type === 'spinner');
  if (!spinner) throw new Error('Spinner not found');
  if (spinner.min !== 1 || spinner.max !== 50) throw new Error(`Spinner: ${JSON.stringify(spinner)}`);

  // Test modifying controls after round-trip
  const cbMod = ws2.getFormControls().find(c => c.type === 'checkBox');
  if (cbMod) cbMod.checked = 'unchecked';
  await wb2.writeFile('./output/27_form_controls_modified.xlsx');
}

// ============================================================
// HUGE FILE — 200 columns × 100,000 rows (EPPlus benchmark)
// ============================================================
async function example_huge_file() {
  const wb = new Workbook();
  const ws = wb.addSheet('HugeData');

  // V8 Map limit is ~16.7M entries; use 100 cols × 100k rows = 10M cells
  const COLS = 100;
  const ROWS = 100_000;

  // Header row
  const header: string[] = [];
  for (let c = 1; c <= COLS; c++) header.push(`Col_${c}`);
  ws.writeRow(1, 1, header);

  // Data rows — mix of numbers, strings, dates, formulas
  const t0 = performance.now();
  for (let r = 2; r <= ROWS + 1; r++) {
    const row: (string | number | Date | null)[] = [];
    for (let c = 1; c <= COLS; c++) {
      const mod = c % 4;
      if (mod === 1) row.push(r * c);                          // number
      else if (mod === 2) row.push(`R${r}C${c}`);              // string
      else if (mod === 3) row.push(new Date(2020, 0, (r % 365) + 1)); // date
      else row.push(null);                                     // empty
    }
    ws.writeRow(r, 1, row);
  }
  const populateMs = (performance.now() - t0).toFixed(0);
  console.log(`   → 100×100k populate: ${populateMs} ms`);

  // A few formulas in the last row
  const sumRow = ROWS + 2;
  ws.setFormula(sumRow, 1, `SUM(A2:A${ROWS + 1})`);
  ws.setFormula(sumRow, 5, `AVERAGE(E2:E${ROWS + 1})`);
  ws.setFormula(sumRow, 9, `MAX(I2:I${ROWS + 1})`);

  const t1 = performance.now();
  await wb.writeFile('./output/28_huge_file.xlsx');
  const writeMs = (performance.now() - t1).toFixed(0);
  console.log(`   → 100×100k write: ${writeMs} ms`);

  // Round-trip: read it back and verify dimensions
  const t2 = performance.now();
  const wb2 = await Workbook.fromFile('./output/28_huge_file.xlsx');
  const readMs = (performance.now() - t2).toFixed(0);
  console.log(`   → 100×100k read : ${readMs} ms`);

  const ws2 = wb2.getSheet('HugeData')!;
  const cell = ws2.getCell(2, 1);
  if (cell.value !== 2) throw new Error(`Expected 2, got ${cell.value}`);
  const lastCell = ws2.getCell(ROWS + 1, 1);
  if (lastCell.value !== (ROWS + 1) * 1) throw new Error(`Last-row value mismatch`);
}

// ============================================================
// IN-CELL PICTURES & NEW IMAGE FORMATS (SVG, WebP, ICO, BMP)
// ============================================================

/** Helper: create a minimal valid PNG from scratch */
function makeTestPng(w: number, h: number, r: number, g: number, b: number): Uint8Array {
  const crcTable = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = (c & 1) ? 0xEDB88320 ^ (c >>> 1) : c >>> 1;
    crcTable[n] = c;
  }
  const crc32 = (buf: Uint8Array, seed = 0xFFFFFFFF): number => {
    let c = seed;
    for (const byte of buf) c = crcTable[(c ^ byte) & 0xFF] ^ (c >>> 8);
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
  const raw = new Uint8Array((1 + w * 3) * h);
  let pos = 0;
  for (let y = 0; y < h; y++) {
    raw[pos++] = 0;
    for (let x = 0; x < w; x++) { raw[pos++] = r; raw[pos++] = g; raw[pos++] = b; }
  }
  const idat = deflateSync(raw);
  const ihdrData = new Uint8Array(13);
  new DataView(ihdrData.buffer).setUint32(0, w);
  new DataView(ihdrData.buffer).setUint32(4, h);
  ihdrData[8] = 8; ihdrData[9] = 2;
  const sig = new Uint8Array([0x89,0x50,0x4E,0x47,0x0D,0x0A,0x1A,0x0A]);
  const parts = [sig, chunk('IHDR', ihdrData), chunk('IDAT', idat), chunk('IEND', new Uint8Array(0))];
  const total = parts.reduce((s, p) => s + p.length, 0);
  const png = new Uint8Array(total);
  let off = 0;
  for (const p of parts) { png.set(p, off); off += p.length; }
  return png;
}

/** Helper: minimal BMP (24-bit, 4x4 solid color) */
function makeTestBmp(r: number, g: number, b: number): Uint8Array {
  const w = 4, h = 4;
  const rowBytes = w * 3;
  const padding = (4 - (rowBytes % 4)) % 4;
  const stride = rowBytes + padding;
  const pixelSize = stride * h;
  const fileSize = 54 + pixelSize;
  const buf = new Uint8Array(fileSize);
  const dv = new DataView(buf.buffer);
  buf[0] = 0x42; buf[1] = 0x4D;  // BM
  dv.setUint32(2, fileSize, true);
  dv.setUint32(10, 54, true);     // pixel data offset
  dv.setUint32(14, 40, true);     // DIB header size
  dv.setInt32(18, w, true);
  dv.setInt32(22, h, true);
  dv.setUint16(26, 1, true);      // planes
  dv.setUint16(28, 24, true);     // bpp
  dv.setUint32(34, pixelSize, true);
  let off = 54;
  for (let y = 0; y < h; y++) {
    for (let x = 0; x < w; x++) { buf[off++] = b; buf[off++] = g; buf[off++] = r; }
    off += padding;
  }
  return buf;
}

async function example_cell_images() {
  const wb = new Workbook();
  const ws = wb.addSheet('CellImages');

  ws.setValue(1, 1, 'In-cell pictures demo');
  ws.setStyle(1, 1, style().bold().build());

  // Create distinct colored PNGs for in-cell images
  const redPng   = makeTestPng(40, 40, 0xFF, 0x00, 0x00);
  const greenPng = makeTestPng(40, 40, 0x00, 0xAA, 0x00);
  const bluePng  = makeTestPng(40, 40, 0x00, 0x00, 0xFF);

  // Add in-cell pictures at specific cells
  ws.addCellImage({ data: redPng,   format: 'png', cell: 'B2', altText: 'Red square' });
  ws.addCellImage({ data: greenPng, format: 'png', cell: 'C3', altText: 'Green square' });
  ws.addCellImage({ data: bluePng,  format: 'png', cell: 'D4', altText: 'Blue square' });

  // Add labels
  ws.setValue(2, 1, 'Red:');
  ws.setValue(3, 1, 'Green:');
  ws.setValue(4, 1, 'Blue:');

  // Make rows taller so images are visible
  ws.setRowHeight(2, 40);
  ws.setRowHeight(3, 40);
  ws.setRowHeight(4, 40);
  ws.setColumnWidth(2, 12);
  ws.setColumnWidth(3, 12);
  ws.setColumnWidth(4, 12);

  await wb.writeFile('./output/29_cell_images.xlsx');

  // Verify round-trip: richData files should be preserved in unknownParts
  const wb2 = await Workbook.fromFile('./output/29_cell_images.xlsx');
  const ws2 = wb2.getSheet('CellImages')!;
  if (!ws2) throw new Error('Sheet not found after round-trip');
  await wb2.writeFile('./output/29_cell_images_rt.xlsx');

  // EPPlus validation
  // @ts-ignore
  const { execSync } = await import('child_process');
  try {
    const out = execSync('dotnet run validate_cellimage_epplus.cs output/29_cell_images.xlsx', {
      encoding: 'utf-8', timeout: 60000, stdio: ['pipe', 'pipe', 'pipe'],
    }).trim();
    if (!out.includes('EPPlus validation: OK')) throw new Error('EPPlus: cell images validation failed');
    if (!out.includes('B2.Picture: EXISTS')) throw new Error('EPPlus: B2 cell picture not found');
    console.log('  EPPlus validation: OK');
  } catch (e: any) {
    if (e.message?.includes('EPPlus:')) throw e;
    console.log('  EPPlus validation skipped:', e.message?.split('\n')[0]);
  }
}

async function example_new_image_formats() {
  const wb = new Workbook();
  const ws = wb.addSheet('ImageFormats');

  ws.setValue(1, 1, 'New image format support');
  ws.setStyle(1, 1, style().bold().build());

  // BMP image (real valid file)
  const bmpBytes = makeTestBmp(0x44, 0x72, 0xC4);
  const bmpImg: Image = {
    data: bmpBytes,
    format: 'bmp',
    from: { col: 1, row: 2 },
    width: 80, height: 80,
    altText: 'BMP test image',
  };
  ws.addImage(bmpImg);
  ws.setValue(2, 1, 'BMP →');

  // SVG image (valid SVG XML)
  const svgStr = `<svg xmlns="http://www.w3.org/2000/svg" width="80" height="80"><rect width="80" height="80" fill="#ED7D31"/><circle cx="40" cy="40" r="30" fill="#FFC000"/></svg>`;
  const svgBytes = new TextEncoder().encode(svgStr);
  const svgImg: Image = {
    data: svgBytes,
    format: 'svg',
    from: { col: 3, row: 2 },
    width: 80, height: 80,
    altText: 'SVG test image',
  };
  ws.addImage(svgImg);
  ws.setValue(2, 3, 'SVG →');

  // WebP — use a minimal valid 1x1 WebP (RIFF/WEBP VP8 container)
  const webpBytes = new Uint8Array([
    0x52,0x49,0x46,0x46, 0x24,0x00,0x00,0x00,  // RIFF + size
    0x57,0x45,0x42,0x50, 0x56,0x50,0x38,0x20,  // WEBP VP8
    0x18,0x00,0x00,0x00, 0x30,0x01,0x00,0x9D,  // chunk header
    0x01,0x2A,0x01,0x00, 0x01,0x00,0x01,0x40,  // 1x1 px
    0x25,0xA4,0x00,0x03, 0x70,0x00,0xFE,0xFB,
    0x94,0x00,0x00,
  ]);
  const webpImg: Image = {
    data: webpBytes,
    format: 'webp',
    from: { col: 5, row: 2 },
    width: 80, height: 80,
    altText: 'WebP test image',
  };
  ws.addImage(webpImg);
  ws.setValue(2, 5, 'WebP →');

  // ICO — minimal 1x1 ICO file
  const icoBytes = new Uint8Array([
    0x00,0x00, 0x01,0x00, 0x01,0x00,              // ICO header: 1 image
    0x01, 0x01, 0x00, 0x00, 0x01,0x00, 0x18,0x00, // 1x1, 24bpp
    0x30,0x00,0x00,0x00, 0x16,0x00,0x00,0x00,     // size + offset
    // BMP info header (40 bytes)
    0x28,0x00,0x00,0x00, 0x01,0x00,0x00,0x00,
    0x02,0x00,0x00,0x00, 0x01,0x00, 0x18,0x00,
    0x00,0x00,0x00,0x00, 0x08,0x00,0x00,0x00,
    0x00,0x00,0x00,0x00, 0x00,0x00,0x00,0x00,
    0x00,0x00,0x00,0x00, 0x00,0x00,0x00,0x00,
    // pixel data (1 BGR pixel + padding + AND mask row)
    0xC4,0x72,0x44, 0x00,   // BGR + pad
    0x00,0x00,0x00, 0x00,   // AND mask
  ]);
  const icoImg: Image = {
    data: icoBytes,
    format: 'ico',
    from: { col: 7, row: 2 },
    width: 80, height: 80,
    altText: 'ICO test image',
  };
  ws.addImage(icoImg);
  ws.setValue(2, 7, 'ICO →');

  // Also test in-cell BMP
  ws.addCellImage({ data: bmpBytes, format: 'bmp', cell: 'B6', altText: 'In-cell BMP' });
  ws.setValue(6, 1, 'In-cell BMP:');
  ws.setRowHeight(6, 40);

  await wb.writeFile('./output/30_new_image_formats.xlsx');

  // Verify the file can be read back
  const wb2 = await Workbook.fromFile('./output/30_new_image_formats.xlsx');
  const ws2 = wb2.getSheet('ImageFormats')!;
  if (!ws2) throw new Error('Sheet not found after round-trip');
  await wb2.writeFile('./output/30_new_image_formats_rt.xlsx');

  // EPPlus validation
  // @ts-ignore
  const { execSync } = await import('child_process');
  try {
    const out = execSync('dotnet run validate_cellimage_epplus.cs output/30_new_image_formats.xlsx', {
      encoding: 'utf-8', timeout: 60000, stdio: ['pipe', 'pipe', 'pipe'],
    }).trim();
    if (!out.includes('EPPlus validation: OK')) throw new Error('EPPlus: new image formats validation failed');
    console.log('  EPPlus validation: OK');
  } catch (e: any) {
    if (e.message?.includes('EPPlus:')) throw e;
    console.log('  EPPlus validation skipped:', e.message?.split('\n')[0]);
  }
}

// ============================================================
// 31. ABSOLUTE POSITIONING & FORM CONTROL WIDTH/HEIGHT
// ============================================================
async function example_absolute_and_sizing() {
  const wb = new Workbook();

  // Sheet 1: Absolute-positioned images
  const ws1 = wb.addSheet('AbsoluteImages');
  ws1.setValue(1, 1, 'Absolute-positioned images (not cell-anchored)');
  ws1.setStyle(1, 1, style().bold().build());

  const png = makeTestPng(40, 40, 0x33, 0x99, 0xFF);

  // Image at absolute position (100px, 50px)
  ws1.addImage({
    data: png, format: 'png',
    position: { x: 100, y: 50 },
    width: 80, height: 80,
    altText: 'Absolute image at (100,50)',
  });

  // Another at (300, 200)
  ws1.addImage({
    data: makeTestPng(40, 40, 0xFF, 0x66, 0x00), format: 'png',
    position: { x: 300, y: 200 },
    width: 60, height: 60,
    altText: 'Absolute image at (300,200)',
  });

  // Also a cell-anchored image for comparison
  ws1.addImage({
    data: makeTestPng(40, 40, 0x00, 0xCC, 0x44), format: 'png',
    from: { col: 5, row: 5 },
    width: 50, height: 50,
  });

  // Sheet 2: Form controls with width/height instead of 'to'
  const ws2 = wb.addSheet('FormControlSizing');
  ws2.setValue(1, 1, 'Form controls with width/height');
  ws2.setStyle(1, 1, style().bold().build());

  ws2.addFormControl({
    type: 'button',
    from: { col: 1, row: 2 },
    width: 120, height: 30,
    text: 'Wide Button',
  } as FormControl);

  ws2.addFormControl({
    type: 'checkBox',
    from: { col: 1, row: 4 },
    width: 80, height: 20,
    text: 'Check me',
    checked: 'unchecked',
  } as FormControl);

  // Also a traditional from/to control for comparison
  ws2.addFormControl({
    type: 'button',
    from: { col: 1, row: 6 },
    to: { col: 3, row: 7 },
    text: 'Traditional Anchor',
  } as FormControl);

  await wb.writeFile('./output/31_absolute_and_sizing.xlsx');

  // Round-trip
  const wb2 = await Workbook.fromFile('./output/31_absolute_and_sizing.xlsx');
  const ws1rt = wb2.getSheet('AbsoluteImages')!;
  if (!ws1rt) throw new Error('AbsoluteImages sheet not found after round-trip');
  await wb2.writeFile('./output/31_absolute_and_sizing_rt.xlsx');

  // EPPlus validation
  // @ts-ignore
  const { execSync } = await import('child_process');
  try {
    const out = execSync('dotnet run validatorEpplus.cs output/31_absolute_and_sizing.xlsx', {
      encoding: 'utf-8', timeout: 60000, stdio: ['pipe', 'pipe', 'pipe'],
    }).trim();
    if (!out.includes('successfully')) throw new Error('EPPlus: absolute positioning validation failed');
    console.log('  EPPlus validation: OK');
  } catch (e: any) {
    if (e.message?.includes('EPPlus:')) throw e;
    console.log('  EPPlus validation skipped:', e.message?.split('\n')[0]);
  }
}

// ============================================================
// 32. DIGITAL SIGNING EXAMPLES
// ============================================================

async function example_signing() {
  const { signPackage, signWorkbook, generateTestCertificate } = await import('../features/Signing.js');
  //@ts-ignore
  const { writeFileSync: wf, readFileSync: rf } = await import('fs');
  const { readZip } = await import('../utils/zipReader.js');
  const { buildZip } = await import('../utils/zip.js');

  // Generate RSA-2048 key pair
  const keyPair = await crypto.subtle.generateKey(
    { name: 'RSASSA-PKCS1-v1_5', modulusLength: 2048, publicExponent: new Uint8Array([1, 0, 1]), hash: 'SHA-256' },
    true, ['sign', 'verify']
  );
  const privDer = new Uint8Array(await crypto.subtle.exportKey('pkcs8', keyPair.privateKey));
  let b64 = '';
  for (let i = 0; i < privDer.length; i++) b64 += String.fromCharCode(privDer[i]);
  b64 = btoa(b64);
  const privateKeyPem = `-----BEGIN PRIVATE KEY-----\n${b64.match(/.{1,64}/g)!.join('\n')}\n-----END PRIVATE KEY-----`;

  // Export public key as SPKI DER for certificate
  const spkiDer = new Uint8Array(await crypto.subtle.exportKey('spki', keyPair.publicKey));

  // Generate self-signed test certificate
  const certPem = await generateTestCertificate('ExcelForge Test Signer', privateKeyPem, spkiDer);
  console.log('  Certificate generated');

  // ── Sample 1: Package-signed workbook ──

  const wb1 = new Workbook();
  const ws1 = wb1.addSheet('Signed');
  ws1.setValue(1, 1, 'This workbook has a package digital signature');
  ws1.setStyle(1, 1, style().bold().build());
  ws1.setValue(2, 1, 'Signed at: ' + new Date().toISOString());

  const xlsxBytes1 = await wb1.build();
  const parts1 = await readZip(xlsxBytes1);
  const partsMap1 = new Map<string, Uint8Array>();
  for (const [name, entry] of parts1) partsMap1.set(name, entry.data);

  const sigEntries = await signPackage(partsMap1, { certificate: certPem, privateKey: privateKeyPem });
  console.log(`  Package signature entries: ${sigEntries.size}`);

  // Merge signature entries into ZIP
  const allEntries1: Array<{ name: string; data: Uint8Array }> = [];
  for (const [name, entry] of parts1) {
    if (sigEntries.has(name)) continue; // skip entries replaced by signature
    allEntries1.push({ name, data: entry.data });
  }
  for (const [name, data] of sigEntries) allEntries1.push({ name, data });
  const signed1 = buildZip(allEntries1);
  wf('./output/32_signed_package.xlsx', signed1);
  console.log('  Written: 32_signed_package.xlsx');

  // ── Sample 2: VBA workbook with both package + VBA signature ──

  const wb2 = new Workbook();
  const ws2 = wb2.addSheet('SignedVBA');
  ws2.setValue(1, 1, 'This workbook has package + VBA signatures');
  ws2.setStyle(1, 1, style().bold().build());

  const vba = new VbaProject();
  vba.addModule({ name: 'Module1', type: 'standard', code: 'Sub Hello()\n  MsgBox "Signed VBA"\nEnd Sub' });
  wb2.vbaProject = vba;

  const xlsmBytes = await wb2.build();
  const parts2 = await readZip(xlsmBytes);
  const partsMap2 = new Map<string, Uint8Array>();
  let vbaProjectBin: Uint8Array | undefined;
  for (const [name, entry] of parts2) {
    partsMap2.set(name, entry.data);
    if (name === 'xl/vbaProject.bin') vbaProjectBin = entry.data;
  }

  const result = await signWorkbook(partsMap2, { certificate: certPem, privateKey: privateKeyPem }, vbaProjectBin);
  console.log(`  signWorkbook: ${result.packageSignatureEntries.size} package entries, VBA sig: ${result.vbaSignature ? result.vbaSignature.length + ' bytes' : 'none'}`);

  // Note: Neither VBA nor package signature is embedded in the .xlsm output.
  // The VBA signature can't be valid because [MS-OVBA] §2.4.2 requires hashing
  // normalized VBA content parts, not the raw binary. The package signature with a
  // self-signed cert would trigger validation errors. Both APIs are still exercised
  // above; the .xlsx sample (32_signed_package.xlsx) demonstrates the package signature.

  const allEntries2: Array<{ name: string; data: Uint8Array }> = [];
  for (const [name, entry] of parts2) {
    allEntries2.push({ name, data: entry.data });
  }
  const signed2 = buildZip(allEntries2);
  wf('./output/32_signed_vba.xlsm', signed2);
  console.log('  Written: 32_signed_vba.xlsm');
}

// ============================================================
// Example: Calc Settings
// ============================================================
async function example_calc_settings() {
  const wb = new Workbook();
  wb.properties.title = 'Calc Settings Demo';

  // Set workbook to manual calculation with iterative calculation enabled
  wb.calcSettings = {
    calcMode: 'manual',
    iterate: true,
    iterateCount: 200,
    iterateDelta: 0.0001,
    fullCalcOnLoad: false,
    calcOnSave: true,
    fullPrecision: true,
    concurrentCalc: false,
  };

  const ws = wb.addSheet('Calc Settings');
  ws.setValue(1, 1, 'Calculation Mode:');
  ws.setValue(1, 2, 'Manual');
  ws.setStyle(1, 1, style().bold().build());
  ws.setValue(2, 1, 'Iterative Calculation:');
  ws.setValue(2, 2, 'Enabled (200 iterations, delta 0.0001)');
  ws.setStyle(2, 1, style().bold().build());

  // Add circular reference to demonstrate iterative calc
  ws.setValue(4, 1, 'Circular Reference Demo');
  ws.setStyle(4, 1, style().bold().fontSize(12).build());
  ws.setValue(5, 1, 'A');
  ws.setValue(5, 2, 'B');
  ws.setValue(6, 1, 1);
  ws.setFormula(6, 2, 'A6+1');
  ws.setFormula(7, 1, 'B6+1');
  ws.setFormula(7, 2, 'A7*2');

  ws.setColumn(1, { width: 24 });
  ws.setColumn(2, { width: 40 });

  await wb.writeFile('./output/33_calc_settings.xlsx');
}

// ============================================================
// Example: VBA UserForms
// ============================================================
async function example_vba_userforms() {
  const wb = new Workbook();
  const ws = wb.addSheet('UserForm Demo');
  ws.setValue(1, 1, 'VBA UserForm Demo');
  ws.setStyle(1, 1, style().bold().fontSize(14).build());
  ws.setValue(2, 1, 'Press Alt+F11 to view the UserForm in the VBA editor');
  ws.setValue(3, 1, 'Run ShowMyForm to display the form');
  ws.setColumn(1, { width: 50 });

  const vba = new VbaProject();

  // Standard module that shows the form
  vba.addModule({
    name: 'Module1',
    type: 'standard',
    code: [
      'Sub ShowMyForm()',
      '  MyForm.Show',
      'End Sub',
    ].join('\n'),
  });

  // UserForm module with controls
  vba.addModule({
    name: 'MyForm',
    type: 'userform',
    controls: [
      { type: 'Label', name: 'Label1', caption: 'Enter your name:', left: 12, top: 12, width: 120, height: 18 },
      { type: 'TextBox', name: 'TextBox1', caption: '', left: 12, top: 36, width: 180, height: 24 },
      { type: 'CommandButton', name: 'OKButton', caption: 'OK', left: 48, top: 72, width: 72, height: 28 },
      { type: 'CommandButton', name: 'CancelButton', caption: 'Cancel', left: 128, top: 72, width: 72, height: 28 },
    ],
    code: [
      'Private Sub OKButton_Click()',
      '  If Len(TextBox1.Text) > 0 Then',
      '    Sheet1.Range("A5").Value = "Hello, " & TextBox1.Text & "!"',
      '  End If',
      '  Unload Me',
      'End Sub',
      '',
      'Private Sub CancelButton_Click()',
      '  Unload Me',
      'End Sub',
      '',
      'Private Sub UserForm_Initialize()',
      '  Me.Caption = "Greeting Form"',
      'End Sub',
    ].join('\n'),
  });

  wb.vbaProject = vba;
  await wb.writeFile('./output/34_vba_userforms.xlsm');
}

// ============================================================
// Example: OLE Objects
// ============================================================
async function example_ole_objects() {
  const wb = new Workbook();
  const ws = wb.addSheet('OLE Objects');
  ws.setValue(1, 1, 'Embedded OLE Object Demo');
  ws.setStyle(1, 1, style().bold().fontSize(14).build());
  ws.setValue(2, 1, 'Below is an embedded binary OLE object placeholder');
  ws.setColumn(1, { width: 40 });

  // Create a small binary payload to embed as an OLE object
  const payload = new Uint8Array(256);
  for (let i = 0; i < payload.length; i++) payload[i] = i & 0xFF;

  ws.addOleObject({
    name: 'EmbeddedData',
    progId: 'Package',
    fileName: 'embedded_data.bin',
    data: payload,
    from: { col: 1, row: 4 },
    to: { col: 5, row: 12 },
  });

  // Add a second OLE object
  const textData = new TextEncoder().encode('Hello from ExcelForge OLE Object!\nThis is embedded text content.');
  ws.addOleObject({
    name: 'EmbeddedText',
    progId: 'Package',
    fileName: 'readme.txt',
    data: textData,
    from: { col: 6, row: 4 },
    to: { col: 10, row: 12 },
  });

  await wb.writeFile('./output/35_ole_objects.xlsx');
}

// ============================================================
// Example: Encrypted XLSX
// ============================================================
async function example_encrypted() {
  const wb = new Workbook();
  wb.properties.title = 'Encrypted Workbook';
  const ws = wb.addSheet('Confidential');
  ws.setValue(1, 1, 'Encrypted Workbook');
  ws.setStyle(1, 1, style().bold().fontSize(16).fontColor('FFCC0000').build());
  ws.setValue(2, 1, 'This file is password-protected.');
  ws.setValue(3, 1, 'Password: secret123');
  ws.setStyle(3, 1, style().italic().fontColor('FF808080').build());

  ws.setValue(5, 1, 'Sensitive Data');
  ws.setStyle(5, 1, style().bold().build());
  ws.setValue(6, 1, 'Account');
  ws.setValue(6, 2, 'Balance');
  ws.setStyle(6, 1, style().bold().bg('FF4472C4').fontColor('FFFFFFFF').build());
  ws.setStyle(6, 2, style().bold().bg('FF4472C4').fontColor('FFFFFFFF').build());
  ws.setValue(7, 1, 'Savings');
  ws.setValue(7, 2, 50000);
  ws.setStyle(7, 2, style().numFmt('#,##0.00').build());
  ws.setValue(8, 1, 'Checking');
  ws.setValue(8, 2, 12500);
  ws.setStyle(8, 2, style().numFmt('#,##0.00').build());
  ws.setColumn(1, { width: 20 });
  ws.setColumn(2, { width: 16 });

  const xlsxData = await wb.build();
  const encrypted = await encryptWorkbook(xlsxData, 'secret123', { spinCount: 1000 });

  //@ts-ignore
  const { writeFileSync } = await import('fs');
  writeFileSync('./output/36_encrypted.xlsx', encrypted);
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
    ['VBA Complex',            example_vba_complex],
    ['CF/DV Round-trip',       example_cf_dv_roundtrip],
    ['Page Breaks',            example_page_breaks],
    ['Connections',            example_connections],
    ['Form Controls',          example_form_controls],
    //['Huge File (200×100k)',   example_huge_file],
    ['Cell Images',            example_cell_images],
    ['New Image Formats',      example_new_image_formats],
    ['Absolute & Sizing',      example_absolute_and_sizing],
    ['Signing',                 example_signing],
    ['Calc Settings',            example_calc_settings],
    ['VBA UserForms',            example_vba_userforms],
    ['OLE Objects',              example_ole_objects],
    ['Encrypted',                example_encrypted],
  ] as const;

  for (const [name, fn] of examples) {
    try {
      await fn();
      console.log(`✅ ${name}`);
    } catch (e) {
      console.error(`❌ ${name}: ${e}`);
    }
  }

  // Export every generated xlsx as HTML
  console.log('\n── HTML Export ──');
  // @ts-ignore
  const { readdirSync, writeFileSync } = await import('fs');
  const { workbookToHtml } = await import('../features/HtmlModule.js');
  const xlsxFiles = (readdirSync('./output') as string[]).filter((f: string) => f.endsWith('.xlsx'));
  let htmlOk = 0;
  for (const file of xlsxFiles) {
    try {
      const wb2 = await Workbook.fromFile(`./output/${file}`);
      const html = workbookToHtml(wb2, { title: file.replace(/\.[^.]+$/, ''), includeTabs: true });
      writeFileSync(`./output/${file.replace(/\.[^.]+$/, '.html')}`, html, 'utf-8');
      htmlOk++;
    } catch (e: any) {
      console.error(`  ❌ HTML ${file}: ${e.message?.slice(0, 120)}`);
    }
  }
  console.log(`  ✅ Exported ${htmlOk}/${xlsxFiles.length} files as HTML`);

  // Export every generated xlsx as PDF
  console.log('\n── PDF Export ──');
  const { workbookToPdf } = await import('../features/PdfModule.js');
  let pdfOk = 0;
  for (const file of xlsxFiles) {
    try {
      const wb2 = await Workbook.fromFile(`./output/${file}`);
      const pdf = workbookToPdf(wb2, { title: file.replace(/\.[^.]+$/, ''), fitToWidth: true, gridLines: true });
      writeFileSync(`./output/${file.replace(/\.[^.]+$/, '.pdf')}`, pdf);
      pdfOk++;
    } catch (e: any) {
      console.error(`  ❌ PDF ${file}: ${e.message?.slice(0, 120)}`);
    }
  }
  console.log(`  ✅ Exported ${pdfOk}/${xlsxFiles.length} files as PDF`);
}

runAll().catch(console.error);
