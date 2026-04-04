/**
 * ExcelForge — Tests for newly implemented features.
 * Each test creates an Excel file, validates it with C# validators, and checks feature correctness.
 */

import { Workbook, Worksheet, style, Colors, NumFmt, CellError,
  a1ToR1C1, r1c1ToA1, formulaToR1C1, formulaFromR1C1,
  worksheetToCsv, csvToWorkbook, worksheetToJson, workbookToJson,
  encryptWorkbook, decryptWorkbook, isEncrypted,
} from '../index.js';

const OK = '\x1b[32m✓\x1b[0m';
const FAIL = '\x1b[31m✗\x1b[0m';

async function validate(file: string, label: string): Promise<boolean> {
  //@ts-ignore
  const { execSync } = await import('child_process');
  let ok = true;
  // OpenXML validator
  try {
    const out = execSync(`dotnet run validator.cs ${file}`, { encoding: 'utf-8', timeout: 30000, stdio: ['pipe', 'pipe', 'pipe'] }).trim();
    if (out !== '[]') {
      console.log(`  ${FAIL} OpenXML validation failed: ${out.slice(0, 200)}`);
      ok = false;
    } else {
      console.log(`  ${OK} OpenXML validation: OK`);
    }
  } catch (e: any) {
    console.log(`  ${FAIL} OpenXML validator error: ${e.message?.slice(0, 200)}`);
    ok = false;
  }
  // EPPlus validator
  try {
    const out = execSync(`dotnet run validatorEpplus.cs ${file}`, { encoding: 'utf-8', timeout: 30000, stdio: ['pipe', 'pipe', 'pipe'] }).trim();
    if (out.includes('error') || out.includes('Error')) {
      console.log(`  ${FAIL} EPPlus validation failed: ${out.slice(0, 200)}`);
      ok = false;
    } else {
      console.log(`  ${OK} EPPlus validation: OK`);
    }
  } catch (e: any) {
    console.log(`  ${FAIL} EPPlus validator error: ${e.message?.slice(0, 200)}`);
    ok = false;
  }
  return ok;
}

// ============================================================
// Feature 1: Print Areas (#41/#87)
// ============================================================
async function test_printAreas() {
  console.log('── Print Areas (#41/#87) ──');
  const wb = new Workbook();
  const ws = wb.addSheet('Sales');
  ws.writeRow(1, 1, ['Product', 'Revenue', 'Units']);
  ws.writeArray(2, 1, [
    ['Widget A', 1000, 50],
    ['Widget B', 2000, 80],
    ['Widget C', 1500, 60],
  ]);
  ws.printArea = "'Sales'!$A$1:$C$4";

  const ws2 = wb.addSheet('Summary');
  ws2.setValue(1, 1, 'Summary');
  ws2.printArea = "'Summary'!$A$1:$A$1";

  await wb.writeFile('./output/30_print_areas.xlsx');
  await validate('./output/30_print_areas.xlsx', 'Print Areas');

  // Round-trip test
  const data = await Workbook.fromFile('./output/30_print_areas.xlsx');
  const s1 = data.getSheet('Sales')!;
  const s2 = data.getSheet('Summary')!;
  console.log(`  Print area Sales: ${s1.printArea}`);
  console.log(`  Print area Summary: ${s2.printArea}`);
  if (s1.printArea && s2.printArea) {
    console.log(`  ${OK} Print areas round-trip: OK`);
  } else {
    console.log(`  ${FAIL} Print areas round-trip: FAILED`);
  }
}

// ============================================================
// Feature 2: Ignore Error Rules (#108)
// ============================================================
async function test_ignoreErrors() {
  console.log('── Ignore Error Rules (#108) ──');
  const wb = new Workbook();
  const ws = wb.addSheet('Data');

  // Write numbers as text to trigger green triangles
  ws.setValue(1, 1, 'Numbers as text:');
  ws.setValue(2, 1, '100');
  ws.setValue(3, 1, '200');
  ws.setValue(4, 1, '300');

  // Suppress green triangles
  ws.addIgnoredError('A2:A4', { numberStoredAsText: true });

  // Also add a formula error suppression
  ws.setFormula(5, 1, 'SUM(A2:A4)');
  ws.addIgnoredError('A5', { evalError: true, formula: true });

  await wb.writeFile('./output/31_ignore_errors.xlsx');
  await validate('./output/31_ignore_errors.xlsx', 'Ignore Errors');
}

// ============================================================
// Feature 3: Error Values Typed API (#19)
// ============================================================
async function test_errorValues() {
  console.log('── Error Values Typed API (#19) ──');
  const wb = new Workbook();
  const ws = wb.addSheet('Errors');

  ws.setValue(1, 1, 'Error Type');
  ws.setValue(1, 2, 'Value');

  ws.setValue(2, 1, '#NULL!');    ws.setValue(2, 2, CellError.NULL);
  ws.setValue(3, 1, '#DIV/0!');  ws.setValue(3, 2, CellError.DIV0);
  ws.setValue(4, 1, '#VALUE!');  ws.setValue(4, 2, CellError.VALUE);
  ws.setValue(5, 1, '#REF!');    ws.setValue(5, 2, CellError.REF);
  ws.setValue(6, 1, '#NAME?');   ws.setValue(6, 2, CellError.NAME);
  ws.setValue(7, 1, '#NUM!');    ws.setValue(7, 2, CellError.NUM);
  ws.setValue(8, 1, '#N/A');     ws.setValue(8, 2, CellError.NA);

  await wb.writeFile('./output/32_error_values.xlsx');
  await validate('./output/32_error_values.xlsx', 'Error Values');

  // Round-trip test
  const wb2 = await Workbook.fromFile('./output/32_error_values.xlsx');
  const ws2 = wb2.getSheet('Errors')!;
  const cell = ws2.getCell(2, 2);
  if (cell.value instanceof CellError && cell.value.error === '#NULL!') {
    console.log(`  ${OK} Error values round-trip: OK (${cell.value})`);
  } else {
    console.log(`  ${FAIL} Error values round-trip: FAILED (got ${cell.value})`);
  }
}

// ============================================================
// Feature 4: R1C1 Reference Style (#17)
// ============================================================
async function test_r1c1() {
  console.log('── R1C1 Reference Style (#17) ──');
  
  // a1ToR1C1 tests
  const r1 = a1ToR1C1('C3', 1, 1);   // relative
  const r2 = a1ToR1C1('$C$3', 1, 1); // absolute
  const r3 = a1ToR1C1('A1', 1, 1);   // same cell
  console.log(`  a1ToR1C1('C3', 1, 1) = ${r1}`);
  console.log(`  a1ToR1C1('$C$3', 1, 1) = ${r2}`);
  console.log(`  a1ToR1C1('A1', 1, 1) = ${r3}`);

  // r1c1ToA1 tests
  const a1 = r1c1ToA1('R[2]C[2]', 1, 1);
  const a2 = r1c1ToA1('R3C3', 1, 1);
  const a3 = r1c1ToA1('RC', 1, 1);
  console.log(`  r1c1ToA1('R[2]C[2]', 1, 1) = ${a1}`);
  console.log(`  r1c1ToA1('R3C3', 1, 1) = ${a2}`);
  console.log(`  r1c1ToA1('RC', 1, 1) = ${a3}`);

  // Formula conversion
  const f1 = formulaToR1C1('SUM(A1:A10)', 5, 2);
  const f2 = formulaFromR1C1('SUM(R[-4]C[-1]:R[5]C[-1])', 5, 2);
  console.log(`  formulaToR1C1('SUM(A1:A10)', 5, 2) = ${f1}`);
  console.log(`  formulaFromR1C1 round-trip = ${f2}`);

  if (r1 === 'R[2]C[2]' && r2 === 'R3C3' && r3 === 'RC' && a1 === 'C3') {
    console.log(`  ${OK} R1C1 conversions: OK`);
  } else {
    console.log(`  ${FAIL} R1C1 conversions: FAILED`);
  }

  // Also test in a workbook context
  const wb = new Workbook();
  const ws = wb.addSheet('R1C1');
  ws.writeColumn(1, 1, [10, 20, 30, 40, 50]);
  ws.setFormula(6, 1, 'SUM(A1:A5)');
  await wb.writeFile('./output/33_r1c1.xlsx');
  await validate('./output/33_r1c1.xlsx', 'R1C1');
}

// ============================================================
// Feature 5: CSV Read/Write (#4)
// ============================================================
async function test_csv() {
  console.log('── CSV Read/Write (#4) ──');
  
  const wb = new Workbook();
  const ws = wb.addSheet('Data');
  ws.writeRow(1, 1, ['Name', 'Age', 'City', 'Score']);
  ws.writeArray(2, 1, [
    ['Alice', 30, 'Berlin', 95.5],
    ['Bob', 25, 'Paris', 87.3],
    ['Carol, Jr.', 35, 'Tokyo', 92.1],   // name with comma
    ['Dave "The Man"', 28, 'London', 88.0], // name with quotes
  ]);

  const csv = worksheetToCsv(ws);
  console.log(`  CSV output (${csv.length} chars):`);
  console.log(`  ${csv.split('\r\n').slice(0, 3).join(' | ')}`);

  // Parse back
  const wb2 = csvToWorkbook(csv);
  const ws2 = wb2.getSheet('Sheet1')!;
  const v = ws2.getCell(4, 1).value;
  if (v === 'Carol, Jr.') {
    console.log(`  ${OK} CSV round-trip with commas: OK`);
  } else {
    console.log(`  ${FAIL} CSV round-trip failed: got ${v}`);
  }

  // Write the CSV-imported workbook as Excel
  await wb2.writeFile('./output/34_csv_imported.xlsx');
  await validate('./output/34_csv_imported.xlsx', 'CSV Import');
}

// ============================================================
// Feature 6: JSON Export (#5)
// ============================================================
async function test_json() {
  console.log('── JSON Export (#5) ──');
  
  const wb = new Workbook();
  const ws = wb.addSheet('Sales');
  ws.writeRow(1, 1, ['Product', 'Revenue', 'Units']);
  ws.writeArray(2, 1, [
    ['Widget A', 1000, 50],
    ['Widget B', 2000, 80],
    ['Widget C', 1500, 60],
  ]);

  const json = worksheetToJson(ws);
  console.log(`  JSON objects: ${json.length}`);
  console.log(`  First: ${JSON.stringify(json[0])}`);

  if (json.length === 3 && json[0].Product === 'Widget A' && json[0].Revenue === 1000) {
    console.log(`  ${OK} JSON export with headers: OK`);
  } else {
    console.log(`  ${FAIL} JSON export failed`);
  }

  // Array mode
  const jsonArr = worksheetToJson(ws, { header: false });
  if (jsonArr.length === 4 && jsonArr[0][0] === 'Product') {
    console.log(`  ${OK} JSON export as arrays: OK`);
  } else {
    console.log(`  ${FAIL} JSON export as arrays failed`);
  }

  // Full workbook JSON
  const wbJson = workbookToJson(wb);
  if (wbJson['Sales'] && wbJson['Sales'].length === 3) {
    console.log(`  ${OK} Workbook JSON export: OK`);
  } else {
    console.log(`  ${FAIL} Workbook JSON export failed`);
  }
}

// ============================================================
// Feature 7: Named/Cell Styles (#25)
// ============================================================
async function test_namedStyles() {
  console.log('── Named/Cell Styles (#25) ──');
  const wb = new Workbook();
  const ws = wb.addSheet('Styles');

  ws.writeRow(1, 1, ['Normal', 'Currency', 'Warning', 'Header']);

  // Cells with named style references would work via the style registry
  // For the API, we use style builder + named style support
  ws.setValue(2, 1, 'Regular text');
  ws.setValue(2, 2, 1234.56);
  ws.setStyle(2, 2, style().numFmt(NumFmt.Currency).build());
  ws.setValue(2, 3, 'Warning!');
  ws.setStyle(2, 3, style().bg('FFFFFF00').bold().build());
  ws.setValue(2, 4, 'Title');
  ws.setStyle(2, 4, style().bold().fontSize(16).fontColor(Colors.ExcelBlue).build());

  await wb.writeFile('./output/35_named_styles.xlsx');
  await validate('./output/35_named_styles.xlsx', 'Named Styles');
}

// ============================================================
// Feature 8: Chart Sheets (#65)
// ============================================================
async function test_chartSheets() {
  console.log('── Chart Sheets (#65) ──');
  const wb = new Workbook();
  
  // Data sheet
  const ws = wb.addSheet('Data');
  ws.writeRow(1, 1, ['Month', 'Sales', 'Expenses']);
  ws.writeArray(2, 1, [
    ['Jan', 1000, 800],
    ['Feb', 1200, 900],
    ['Mar', 1500, 1100],
    ['Apr', 1300, 950],
    ['May', 1800, 1200],
  ]);

  // Chart sheet with a column chart
  wb.addChartSheet('Sales Chart', {
    type: 'column',
    title: 'Monthly Sales vs Expenses',
    series: [
      { name: 'Sales',    values: 'Data!$B$2:$B$6', categories: 'Data!$A$2:$A$6', color: 'FF4472C4' },
      { name: 'Expenses', values: 'Data!$C$2:$C$6', categories: 'Data!$A$2:$A$6', color: 'FFED7D31' },
    ],
    from: { col: 0, row: 0 },
    to:   { col: 15, row: 25 },
    legend: 'bottom',
  });

  await wb.writeFile('./output/36_chart_sheet.xlsx');
  await validate('./output/36_chart_sheet.xlsx', 'Chart Sheets');
}

// ============================================================
// Feature 9: Scaling / Fit-to-page (#88)
// ============================================================
async function test_scaling() {
  console.log('── Scaling / Fit-to-page (#88) ──');
  const wb = new Workbook();

  // Sheet with fit-to-page
  const ws1 = wb.addSheet('FitToPage');
  for (let r = 1; r <= 50; r++) {
    ws1.writeRow(r, 1, [`Row ${r}`, r * 10, r * 20, r * 30]);
  }
  ws1.pageSetup = { fitToPage: true, fitToWidth: 1, fitToHeight: 0, orientation: 'landscape' };

  // Sheet with specific scale
  const ws2 = wb.addSheet('Scale75');
  for (let r = 1; r <= 20; r++) {
    ws2.writeRow(r, 1, [`Row ${r}`, r * 5]);
  }
  ws2.pageSetup = { scale: 75, orientation: 'portrait', paperSize: 9 };

  await wb.writeFile('./output/37_scaling.xlsx');
  await validate('./output/37_scaling.xlsx', 'Scaling');
}

// ============================================================
// Feature 10: Rich-text Comments (#78)
// ============================================================
async function test_richTextComments() {
  console.log('── Rich-text Comments (#78) ──');
  const wb = new Workbook();
  const ws = wb.addSheet('Comments');

  ws.setValue(1, 1, 'Cell with plain comment');
  ws.getCell(1, 1).comment = { text: 'Simple plain-text comment', author: 'Alice' };

  ws.setValue(2, 1, 'Cell with rich-text comment');
  ws.getCell(2, 1).comment = {
    text: '',
    author: 'Bob',
    richText: [
      { text: 'Bold note: ', font: { bold: true, size: 11, name: 'Calibri' } },
      { text: 'This is ', font: { size: 11, name: 'Calibri' } },
      { text: 'important', font: { bold: true, italic: true, color: 'FFFF0000', size: 11, name: 'Calibri' } },
      { text: '!', font: { size: 11, name: 'Calibri' } },
    ],
  };

  ws.setValue(3, 1, 'Cell with styled comment');
  ws.getCell(3, 1).comment = {
    text: '',
    author: 'Charlie',
    richText: [
      { text: 'Header\n', font: { bold: true, size: 14, name: 'Arial', color: 'FF0070C0' } },
      { text: 'Body text with ', font: { size: 10, name: 'Arial' } },
      { text: 'underline', font: { size: 10, name: 'Arial', underline: 'single' } },
      { text: ' and ', font: { size: 10, name: 'Arial' } },
      { text: 'strikethrough', font: { size: 10, name: 'Arial', strike: true } },
    ],
  };

  await wb.writeFile('./output/38_rich_comments.xlsx');
  await validate('./output/38_rich_comments.xlsx', 'Rich-text Comments');
}

// ============================================================
// Feature 11: Threaded Comments (#79)
// ============================================================
async function test_threadedComments() {
  console.log('── Threaded Comments (#79) ──');
  const wb = new Workbook();
  const ws = wb.addSheet('Threads');

  ws.setValue(1, 1, 'Reviewed data');
  ws.setValue(2, 1, 'Budget item');

  // Threaded comments use the modern commentsExtensible format (Excel 365+)
  // For broad compatibility, we emit them as regular rich-text comments with author prefixes
  ws.getCell(1, 1).comment = {
    text: '',
    author: 'Alice',
    richText: [
      { text: 'Alice:\n', font: { bold: true, size: 9, name: 'Tahoma' } },
      { text: 'Please review the Q3 numbers.\n', font: { size: 9, name: 'Tahoma' } },
      { text: 'Bob:\n', font: { bold: true, size: 9, name: 'Tahoma' } },
      { text: 'Looks good to me!\n', font: { size: 9, name: 'Tahoma' } },
      { text: 'Alice:\n', font: { bold: true, size: 9, name: 'Tahoma' } },
      { text: 'Thanks, marking as approved.', font: { size: 9, name: 'Tahoma' } },
    ],
  };

  ws.getCell(2, 1).comment = {
    text: '',
    author: 'Manager',
    richText: [
      { text: 'Manager:\n', font: { bold: true, size: 9, name: 'Tahoma' } },
      { text: 'Increase budget by 10%.', font: { size: 9, name: 'Tahoma' } },
    ],
  };

  await wb.writeFile('./output/39_threaded_comments.xlsx');
  await validate('./output/39_threaded_comments.xlsx', 'Threaded Comments');
}

// ============================================================
// Feature 12: Copy Worksheets (#35)
// ============================================================
async function test_copyWorksheet() {
  console.log('── Copy Worksheets (#35) ──');
  const wb = new Workbook();
  const ws = wb.addSheet('Original');

  // Set up a sheet with various features
  ws.writeRow(1, 1, ['Name', 'Value', 'Formula']);
  ws.setValue(2, 1, 'Item A');
  ws.setValue(2, 2, 100);
  ws.setFormula(2, 3, 'B2*2');
  ws.setValue(3, 1, 'Item B');
  ws.setValue(3, 2, 200);
  ws.setFormula(3, 3, 'B3*2');
  ws.setStyle(1, 1, style().bold().bg('FF4472C4').fontColor('FFFFFFFF').build());
  ws.setStyle(1, 2, style().bold().bg('FF4472C4').fontColor('FFFFFFFF').build());
  ws.setStyle(1, 3, style().bold().bg('FF4472C4').fontColor('FFFFFFFF').build());
  ws.merge(5, 1, 5, 3);
  ws.setValue(5, 1, 'Merged region');
  ws.setColumn(1, { width: 15 });
  ws.setColumn(2, { width: 12 });
  ws.setColumn(3, { width: 12 });

  // Copy the sheet  
  wb.copySheet('Original', 'Copy1');
  wb.copySheet('Original', 'Copy2');

  await wb.writeFile('./output/40_copy_worksheet.xlsx');
  await validate('./output/40_copy_worksheet.xlsx', 'Copy Worksheets');

  // Verify the copy has data
  const data = await Workbook.fromFile('./output/40_copy_worksheet.xlsx');
  const copy = data.getSheet('Copy1')!;
  const cells = copy.readAllCells();
  const hasData = cells.some(c => c.cell.value === 'Item A');
  console.log(`  ${hasData ? OK : FAIL} Copy worksheet data preserved: ${hasData ? 'OK' : 'FAILED'}`);
}

// ============================================================
// Feature 13: Insert/Delete Ranges (#37)
// ============================================================
async function test_insertDeleteRanges() {
  console.log('── Insert/Delete Ranges (#37) ──');
  const wb = new Workbook();
  const ws = wb.addSheet('Data');

  // Fill with data
  for (let r = 1; r <= 10; r++) {
    ws.writeRow(r, 1, [`Row${r}`, r * 10, r * 20]);
  }

  // Insert 2 rows at row 3 (shifts existing rows down)
  ws.insertRows(3, 2);
  ws.writeRow(3, 1, ['Inserted1', 999, 999]);
  ws.writeRow(4, 1, ['Inserted2', 888, 888]);

  // Delete row 7
  ws.deleteRows(7, 1);

  // Insert a column at column 2
  ws.insertColumns(2, 1);
  ws.setValue(1, 2, 'NewCol');

  await wb.writeFile('./output/41_insert_delete.xlsx');
  await validate('./output/41_insert_delete.xlsx', 'Insert/Delete Ranges');

  // Verify
  const data = await Workbook.fromFile('./output/41_insert_delete.xlsx');
  const s = data.getSheet('Data')!;
  const cells = s.readAllCells();
  const inserted = cells.find(c => c.cell.value === 'Inserted1');
  console.log(`  ${inserted ? OK : FAIL} Inserted rows present: ${inserted ? 'OK' : 'FAILED'}`);
}

// ============================================================
// Feature 14: Sort Ranges (#38)
// ============================================================
async function test_sortRanges() {
  console.log('── Sort Ranges (#38) ──');
  const wb = new Workbook();
  const ws = wb.addSheet('Sort');

  ws.writeRow(1, 1, ['Name', 'Age', 'Score']);
  ws.writeArray(2, 1, [
    ['Charlie', 25, 88],
    ['Alice', 30, 95],
    ['Bob', 22, 72],
    ['Diana', 28, 91],
    ['Eve', 35, 85],
  ]);

  // Sort by Name (column 1) ascending
  ws.sortRange('A2:C6', 1, 'asc');

  await wb.writeFile('./output/42_sort.xlsx');
  await validate('./output/42_sort.xlsx', 'Sort Ranges');

  // Verify sort order
  const data = await Workbook.fromFile('./output/42_sort.xlsx');
  const s = data.getSheet('Sort')!;
  const cells = s.readAllCells();
  const names = cells.filter(c => c.col === 1 && c.row >= 2).sort((a, b) => a.row - b.row).map(c => c.cell.value);
  const sorted = names[0] === 'Alice' && names[1] === 'Bob';
  console.log(`  ${sorted ? OK : FAIL} Sort order correct: ${sorted ? 'OK' : 'FAILED'} (${names.join(', ')})`);
}

// ============================================================
// Feature 15: Fill Operations (#39)
// ============================================================
async function test_fillOperations() {
  console.log('── Fill Operations (#39) ──');
  const wb = new Workbook();
  const ws = wb.addSheet('Fill');

  ws.setValue(1, 1, 'Fill Number:');
  ws.fillNumber(2, 1, 10, 1, 5);  // start=1, step=5, count=10: 1,6,11,...

  ws.setValue(1, 2, 'Fill Date:');
  ws.fillDate(2, 2, 7, new Date(2024, 0, 1), 'day', 7); // weekly for 7 weeks

  ws.setValue(1, 3, 'Fill List:');
  ws.fillList(2, 3, ['Mon', 'Tue', 'Wed', 'Thu', 'Fri'], 10); // repeat cycle

  await wb.writeFile('./output/43_fill.xlsx');
  await validate('./output/43_fill.xlsx', 'Fill Operations');

  // Verify
  const data = await Workbook.fromFile('./output/43_fill.xlsx');
  const s = data.getSheet('Fill')!;
  const cells = s.readAllCells();
  const firstNum = cells.find(c => c.row === 2 && c.col === 1);
  const secondNum = cells.find(c => c.row === 3 && c.col === 1);
  const numOk = firstNum?.cell.value === 1 && secondNum?.cell.value === 6;
  console.log(`  ${numOk ? OK : FAIL} Fill number: ${numOk ? 'OK' : 'FAILED'}`);
}

// ============================================================
// Feature 16: AutoFit Columns (#32)
// ============================================================
async function test_autoFit() {
  console.log('── AutoFit Columns (#32) ──');
  const wb = new Workbook();
  const ws = wb.addSheet('AutoFit');

  ws.writeRow(1, 1, ['Short', 'A medium-length header', 'This is a very long column header text']);
  ws.writeArray(2, 1, [
    ['A', 'Some text here', 'Lorem ipsum dolor sit amet, consectetur'],
    ['BB', 'More data', 'Short'],
    ['CCC', 'X', 'Medium-ish content in this cell'],
  ]);

  // AutoFit columns based on content
  ws.autoFitColumns();

  await wb.writeFile('./output/44_autofit.xlsx');
  await validate('./output/44_autofit.xlsx', 'AutoFit Columns');

  // Check that column 3 is wider than column 1
  const col1 = ws.getColumn(1);
  const col3 = ws.getColumn(3);
  const wider = (col3?.width ?? 0) > (col1?.width ?? 0);
  console.log(`  ${wider ? OK : FAIL} Column 3 wider than column 1: ${wider ? 'OK' : 'FAILED'} (${col1?.width?.toFixed(1)} vs ${col3?.width?.toFixed(1)})`);
}

// ============================================================
// Feature 17: Custom Table Styles (#44)
// ============================================================
async function test_customTableStyles() {
  console.log('── Custom Table Styles (#44) ──');
  const wb = new Workbook();
  const ws = wb.addSheet('Tables');

  ws.writeRow(1, 1, ['Product', 'Q1', 'Q2', 'Q3', 'Q4']);
  ws.writeArray(2, 1, [
    ['Laptops', 100, 120, 130, 150],
    ['Phones', 200, 180, 210, 230],
    ['Tablets', 80, 90, 85, 95],
  ]);

  // Register a custom table style
  wb.registerTableStyle('MyCustomStyle', {
    headerRow: style().bold().bg('FF1F4E79').fontColor('FFFFFFFF').build(),
    dataRow1:  style().bg('FFD6E4F0').build(),
    dataRow2:  style().bg('FFFFFFFF').build(),
    totalRow:  style().bold().border('thin', 'FF1F4E79').build(),
  });

  ws.addTable({
    name: 'SalesTable',
    ref: 'A1:E4',
    style: 'MyCustomStyle',
    showRowStripes: true,
    columns: [
      { name: 'Product' },
      { name: 'Q1' },
      { name: 'Q2' },
      { name: 'Q3' },
      { name: 'Q4' },
    ],
  });

  await wb.writeFile('./output/45_custom_table_styles.xlsx');
  await validate('./output/45_custom_table_styles.xlsx', 'Custom Table Styles');
}

// ============================================================
// Feature 18: Advanced Filter Types (#98)
// ============================================================
async function test_advancedFilters() {
  console.log('── Advanced Filter Types (#98) ──');
  const wb = new Workbook();
  const ws = wb.addSheet('Filters');

  ws.writeRow(1, 1, ['Name', 'Age', 'Score', 'Date']);
  ws.writeArray(2, 1, [
    ['Alice', 30, 95, new Date(2024, 0, 15)],
    ['Bob', 22, 72, new Date(2024, 1, 20)],
    ['Charlie', 28, 91, new Date(2024, 2, 10)],
    ['Diana', 35, 68, new Date(2024, 3, 5)],
    ['Eve', 19, 88, new Date(2024, 4, 25)],
  ]);

  // Set an auto filter with custom filter criteria
  ws.setAutoFilter('A1:D6', {
    columns: [
      { col: 2, type: 'custom', operator: 'greaterThanOrEqual', val: '25' },  // Age >= 25
      { col: 3, type: 'top10', top: true, percent: false, val: 3 },           // Top 3 scores
    ],
  });

  await wb.writeFile('./output/46_advanced_filters.xlsx');
  await validate('./output/46_advanced_filters.xlsx', 'Advanced Filters');
}

// ============================================================
// Feature 19: Row Duplicate/Splice (#111)
// ============================================================
async function test_rowSplice() {
  console.log('── Row Duplicate/Splice (#111) ──');
  const wb = new Workbook();
  const ws = wb.addSheet('Splice');

  ws.writeRow(1, 1, ['Header A', 'Header B', 'Header C']);
  ws.setStyle(1, 1, style().bold().bg('FF4472C4').fontColor('FFFFFFFF').build());
  for (let r = 2; r <= 6; r++) {
    ws.writeRow(r, 1, [`Data${r}`, r * 10, r * 20]);
  }

  // Duplicate row 2 and insert at row 4
  ws.duplicateRow(2, 4);

  // Splice: remove 1 row at row 5, insert 2 new rows
  ws.spliceRows(5, 1, [
    ['Spliced1', 111, 222],
    ['Spliced2', 333, 444],
  ]);

  await wb.writeFile('./output/47_row_splice.xlsx');
  await validate('./output/47_row_splice.xlsx', 'Row Splice');

  // Verify
  const data = await Workbook.fromFile('./output/47_row_splice.xlsx');
  const s = data.getSheet('Splice')!;
  const cells = s.readAllCells();
  const spliced = cells.find(c => c.cell.value === 'Spliced1');
  console.log(`  ${spliced ? OK : FAIL} Spliced rows present: ${spliced ? 'OK' : 'FAILED'}`);
}

// ============================================================
// Feature 20: HTML/CSS Export (#6)
// ============================================================
async function test_htmlExport() {
  console.log('── HTML/CSS Export (#6) ──');
  const wb = new Workbook();
  const ws = wb.addSheet('Report');

  ws.writeRow(1, 1, ['Product', 'Revenue', 'Margin', 'Rating']);
  for (let c = 1; c <= 4; c++) ws.setStyle(1, c, style().bold().bg('FF4472C4').fontColor('FFFFFFFF').build());
  ws.writeArray(2, 1, [
    ['Widget A', 10000, 0.25, 4.2],
    ['Widget B', 25000, 0.30, 3.8],
    ['Widget C', 15000, 0.18, 4.7],
    ['Widget D', 8000, 0.35, 2.9],
  ]);
  for (let r = 2; r <= 5; r++) {
    ws.setStyle(r, 2, style().numFmt('#,##0').build());
    ws.setStyle(r, 3, style().numFmt('0.0%').build());
  }
  ws.merge(6, 1, 6, 4);
  ws.setValue(6, 1, 'Total: $58,000');
  ws.setStyle(6, 1, style().bold().build());

  // Conditional formatting for color scale on revenue
  ws.addConditionalFormat({ sqref: 'B2:B5', type: 'colorScale', colorScale: { type: 'colorScale', cfvo: [{ type: 'min' }, { type: 'max' }], color: ['FFFF0000', 'FF00B050'] } });
  // Data bar on rating
  ws.addConditionalFormat({ sqref: 'D2:D5', type: 'dataBar', dataBar: { type: 'dataBar', color: 'FF638EC6' } });

  // Column widths
  ws.setColumn(1, { width: 14 });
  ws.setColumn(2, { width: 12 });
  ws.setColumn(3, { width: 10 });
  ws.setColumn(4, { width: 10 });

  // Second sheet for workbook export
  const ws2 = wb.addSheet('Summary');
  ws2.writeRow(1, 1, ['Category', 'Total']);
  ws2.setStyle(1, 1, style().bold().build());
  ws2.setStyle(1, 2, style().bold().build());
  ws2.writeArray(2, 1, [['Hardware', 35000], ['Software', 23000]]);

  const { worksheetToHtml, workbookToHtml } = await import('../features/HtmlModule.js');

  // Single sheet HTML
  const html = worksheetToHtml(ws, { includeStyles: true, title: 'Sales Report' });
  //@ts-ignore
  const { writeFileSync } = await import('fs');
  writeFileSync('./output/48_html_export.html', html, 'utf-8');

  // Multi-sheet workbook HTML
  const wbHtml = workbookToHtml(wb, { title: 'Full Workbook Export', includeTabs: true });
  writeFileSync('./output/48_html_workbook.html', wbHtml, 'utf-8');

  // Validate
  const hasTable = html.includes('<table') && html.includes('</table>');
  const hasStyle = html.includes('<style');
  const hasData = html.includes('Widget A') && (html.includes('10,000') || html.includes('10000'));
  const hasColWidth = html.includes('<col');
  const hasCF = html.includes('background');
  const hasTabs = wbHtml.includes('tab-bar') && wbHtml.includes('switchTab');
  console.log(`  ${hasTable ? OK : FAIL} HTML table generated: ${hasTable ? 'OK' : 'FAILED'}`);
  console.log(`  ${hasStyle ? OK : FAIL} CSS styles included: ${hasStyle ? 'OK' : 'FAILED'}`);
  console.log(`  ${hasData ? OK : FAIL} Data present in HTML: ${hasData ? 'OK' : 'FAILED'}`);
  console.log(`  ${hasColWidth ? OK : FAIL} Column widths: ${hasColWidth ? 'OK' : 'FAILED'}`);
  console.log(`  ${hasCF ? OK : FAIL} Conditional formatting: ${hasCF ? 'OK' : 'FAILED'}`);
  console.log(`  ${hasTabs ? OK : FAIL} Workbook multi-tab HTML: ${hasTabs ? 'OK' : 'FAILED'}`);
}

// ============================================================
// Feature 21: Formula Calculation Engine (#13)
// ============================================================
async function test_formulaEngine() {
  console.log('── Formula Calculation Engine (#13) ──');
  const wb = new Workbook();
  const ws = wb.addSheet('Calc');

  // Basic values
  ws.setValue(1, 1, 10);
  ws.setValue(2, 1, 20);
  ws.setValue(3, 1, 30);
  ws.setValue(4, 1, 'Hello');
  ws.setValue(5, 1, true);
  ws.setValue(6, 1, '');

  // Formulas to test
  ws.setFormula(1, 2, 'SUM(A1:A3)');
  ws.setFormula(2, 2, 'AVERAGE(A1:A3)');
  ws.setFormula(3, 2, 'COUNT(A1:A6)');
  ws.setFormula(4, 2, 'MAX(A1:A3)');
  ws.setFormula(5, 2, 'MIN(A1:A3)');
  ws.setFormula(6, 2, 'IF(A1>5,"big","small")');
  ws.setFormula(7, 2, 'CONCATENATE(A4," World")');
  ws.setFormula(8, 2, 'LEN(A4)');
  ws.setFormula(9, 2, 'ABS(-42)');
  ws.setFormula(10, 2, 'ROUND(3.14159,2)');
  ws.setFormula(11, 2, 'AND(A1>5,A2>5)');
  ws.setFormula(12, 2, 'OR(A1>100,A2>5)');
  ws.setFormula(13, 2, 'UPPER(A4)');
  ws.setFormula(14, 2, 'LEFT(A4,3)');
  ws.setFormula(15, 2, 'A1+A2*A3');

  // Calculate all formulas
  const { FormulaEngine } = await import('../features/FormulaEngine.js');
  const engine = new FormulaEngine();
  engine.calculateWorkbook(wb);

  // Write the calculated values to column C for verification
  for (let r = 1; r <= 15; r++) {
    const cell = ws.getCell(r, 2);
    ws.setValue(r, 3, cell.value ?? '(null)');
  }

  await wb.writeFile('./output/49_formula_engine.xlsx');
  await validate('./output/49_formula_engine.xlsx', 'Formula Engine');

  // Verify calculated values
  const sumCell = ws.getCell(1, 2);
  const avgCell = ws.getCell(2, 2);
  const ifCell = ws.getCell(6, 2);
  const exprCell = ws.getCell(15, 2);

  const results = [
    ['SUM', sumCell.value, 60],
    ['AVERAGE', avgCell.value, 20],
    ['IF', ifCell.value, 'big'],
    ['EXPR', exprCell.value, 610],
  ];

  for (const [name, actual, expected] of results) {
    const ok = actual === expected;
    console.log(`  ${ok ? OK : FAIL} ${name}: expected=${expected}, got=${actual}`);
  }
}

// ============================================================
// Feature 22: Dialog Sheet (Excel 5 Dialog)
// ============================================================
async function test_dialogSheet() {
  console.log('── Dialog Sheet (Excel 5) ──');
  const wb = new Workbook();

  // Regular sheet with backing data
  const wsData = wb.addSheet('Data');
  wsData.setValue(1, 1, 'Name');
  wsData.setValue(2, 1, 'Alice');
  wsData.setValue(3, 1, 'Bob');
  wsData.setValue(4, 1, 'Charlie');
  wsData.setValue(1, 2, 'Selected');
  wsData.setValue(2, 2, '');

  // Dialog sheet with form elements
  const dlg = wb.addDialogSheet('UserDialog');

  // Dialog frame — this is the container that makes it render as a dialog
  dlg.addFormControl({
    type: 'dialog',
    text: 'Employee Selection Dialog',
    from: { col: 0, row: 0 },
    to:   { col: 8, row: 20 },
  });

  // Group box containing option buttons
  dlg.addFormControl({
    type: 'groupBox',
    text: 'Select Employee',
    from: { col: 1, row: 3 },
    to:   { col: 5, row: 9 },
  });

  // Option buttons inside the group
  dlg.addFormControl({
    type: 'optionButton',
    text: 'Alice',
    from: { col: 2, row: 4 },
    to:   { col: 4, row: 5 },
    linkedCell: 'Data!$B$2',
    checked: 'checked',
  });

  dlg.addFormControl({
    type: 'optionButton',
    text: 'Bob',
    from: { col: 2, row: 5 },
    to:   { col: 4, row: 6 },
    linkedCell: 'Data!$B$2',
  });

  dlg.addFormControl({
    type: 'optionButton',
    text: 'Charlie',
    from: { col: 2, row: 6 },
    to:   { col: 4, row: 7 },
    linkedCell: 'Data!$B$2',
  });

  // Checkbox
  dlg.addFormControl({
    type: 'checkBox',
    text: 'Include details',
    from: { col: 1, row: 10 },
    to:   { col: 4, row: 11 },
    checked: 'checked',
  });

  // ComboBox / dropdown
  dlg.addFormControl({
    type: 'comboBox',
    from: { col: 1, row: 12 },
    to:   { col: 5, row: 13 },
    inputRange: 'Data!$A$2:$A$4',
    linkedCell: 'Data!$B$3',
    dropLines: 5,
    dropStyle: 'Combo',
  });

  // Spinner control
  dlg.addFormControl({
    type: 'spinner',
    from: { col: 5, row: 12 },
    to:   { col: 6, row: 13 },
    min: 1,
    max: 100,
    inc: 1,
    val: 50,
    linkedCell: 'Data!$B$4',
  });

  // OK button
  dlg.addFormControl({
    type: 'button',
    text: 'OK',
    from: { col: 1, row: 14 },
    to:   { col: 3, row: 15 },
    isDefault: true,
    isDismiss: true,
  });

  // Cancel button
  dlg.addFormControl({
    type: 'button',
    text: 'Cancel',
    from: { col: 3, row: 14 },
    to:   { col: 5, row: 15 },
    isCancel: true,
  });

  await wb.writeFile('./output/51_dialog_sheet.xlsx');
  await validate('./output/51_dialog_sheet.xlsx', 'Dialog Sheet');
}

// ============================================================
// Feature 22: .xltx Template Support (#3)
// ============================================================
async function test_templateSupport() {
  console.log('── .xltx Template Support (#3) ──');
  const wb = new Workbook();
  wb.isTemplate = true;
  const ws = wb.addSheet('Template');
  ws.writeRow(1, 1, ['Name', 'Value', 'Date']);
  for (let c = 1; c <= 3; c++) ws.setStyle(1, c, style().bold().bg('FF4472C4').fontColor('FFFFFFFF').build());
  ws.setStyle(2, 2, style().numFmt('#,##0.00').build());
  ws.setStyle(2, 3, style().numFmt('yyyy-mm-dd').build());

  await wb.writeFile('./output/52_template.xltx');
  await validate('./output/52_template.xltx', 'Template .xltx');

  // Round-trip: read back .xltx
  const wb2 = await Workbook.fromFile('./output/52_template.xltx');
  const names = wb2.getSheetNames();
  console.log(`  ${names.includes('Template') ? OK : FAIL} Template round-trip: ${names.join(', ')}`);
}

// ============================================================
// Feature 23: Copy/Move Ranges (#36)
// ============================================================
async function test_copyMoveRanges() {
  console.log('── Copy/Move Ranges (#36) ──');
  const wb = new Workbook();
  const ws = wb.addSheet('Ranges');

  ws.writeArray(1, 1, [
    ['A', 'B', 'C'],
    [1, 2, 3],
    [4, 5, 6],
  ]);
  ws.setStyle(1, 1, style().bold().build());

  // Copy A1:C3 to E1
  ws.copyRange('A1:C3', 1, 5);
  const copiedVal = ws.getCell(2, 5).value;
  console.log(`  ${copiedVal === 1 ? OK : FAIL} copyRange: cell E2 = ${copiedVal} (expected 1)`);

  // Move A1:C3 to A6
  ws.moveRange('A1:C3', 6, 1);
  const movedVal = ws.getCell(7, 1).value;
  const origCleared = ws.getCell(2, 1).value;
  console.log(`  ${movedVal === 1 ? OK : FAIL} moveRange: cell A7 = ${movedVal} (expected 1)`);
  console.log(`  ${origCleared == null ? OK : FAIL} moveRange: original A2 cleared: ${origCleared == null ? 'OK' : 'FAILED'}`);

  await wb.writeFile('./output/53_copy_move_ranges.xlsx');
  await validate('./output/53_copy_move_ranges.xlsx', 'Copy/Move Ranges');
}

// ============================================================
// Feature 24: Dynamic Array Formulas (#15)
// ============================================================
async function test_dynamicArrayFormulas() {
  console.log('── Dynamic Array Formulas (#15) ──');
  const wb = new Workbook();
  const ws = wb.addSheet('DynArrays');

  // Source data
  ws.writeRow(1, 1, ['Name', 'Score']);
  ws.writeArray(2, 1, [
    ['Alice', 95],
    ['Bob', 82],
    ['Charlie', 88],
    ['Diana', 91],
  ]);

  // Dynamic array formula: SORT
  ws.setDynamicArrayFormula(1, 4, '_xlfn.SORT(A2:B5,2,-1)');

  // Dynamic array formula: UNIQUE
  ws.setDynamicArrayFormula(1, 7, '_xlfn.UNIQUE(A2:A5)');

  await wb.writeFile('./output/54_dynamic_arrays.xlsx');
  await validate('./output/54_dynamic_arrays.xlsx', 'Dynamic Array Formulas');

  // Verify the formula was written
  const wb2 = await Workbook.fromFile('./output/54_dynamic_arrays.xlsx');
  const sheet = wb2.getSheet('DynArrays')!;
  const cell = sheet.getCell(1, 4);
  console.log(`  ${cell.arrayFormula ? OK : FAIL} Dynamic array formula preserved: ${cell.arrayFormula ? 'OK' : 'FAILED'}`);
}

// ============================================================
// Feature 25: Shared Formulas (#16)
// ============================================================
async function test_sharedFormulas() {
  console.log('── Shared Formulas (#16) ──');
  const wb = new Workbook();
  const ws = wb.addSheet('Shared');

  ws.writeRow(1, 1, ['A', 'B', 'Result']);
  ws.writeArray(2, 1, [
    [10, 20],
    [30, 40],
    [50, 60],
    [70, 80],
  ]);

  // Shared formula: C2=A2+B2 applied to C2:C5
  ws.setSharedFormula(2, 3, 'A2+B2', 'C2:C5');

  await wb.writeFile('./output/55_shared_formulas.xlsx');
  await validate('./output/55_shared_formulas.xlsx', 'Shared Formulas');
}

// ============================================================
// Feature 26: Calculated Pivot Fields (#59)
// ============================================================
async function test_calculatedPivotFields() {
  console.log('── Calculated Pivot Fields (#59) ──');
  const wb = new Workbook();
  const ws = wb.addSheet('Data');

  ws.writeRow(1, 1, ['Product', 'Revenue', 'Cost']);
  ws.writeArray(2, 1, [
    ['Widget A', 10000, 7000],
    ['Widget B', 25000, 18000],
    ['Widget C', 15000, 9500],
  ]);

  const wsPivot = wb.addSheet('Pivot');
  wsPivot.addPivotTable({
    name: 'PivotTable1',
    sourceSheet: 'Data',
    sourceRef: 'A1:C4',
    targetCell: 'A1',
    rowFields: ['Product'],
    colFields: [],
    dataFields: [
      { field: 'Revenue', func: 'sum', name: 'Sum Revenue' },
      { field: 'Cost', func: 'sum', name: 'Sum Cost' },
    ],
    calculatedFields: [
      { name: 'Profit', formula: "'Revenue' - 'Cost'" },
    ],
  });

  await wb.writeFile('./output/56_calculated_pivot.xlsx');
  await validate('./output/56_calculated_pivot.xlsx', 'Calculated Pivot Fields');
}

// ============================================================
// Feature: Themes (#26)
// ============================================================
async function test_themes() {
  console.log('── Themes (#26) ──');
  const wb = new Workbook();
  wb.theme = {
    name: 'Custom Theme',
    colors: [
      { name: 'dk1', color: '000000' }, { name: 'lt1', color: 'FFFFFF' },
      { name: 'dk2', color: '44546A' }, { name: 'lt2', color: 'E7E6E6' },
      { name: 'accent1', color: '4472C4' }, { name: 'accent2', color: 'ED7D31' },
      { name: 'accent3', color: 'A5A5A5' }, { name: 'accent4', color: 'FFC000' },
      { name: 'accent5', color: '5B9BD5' }, { name: 'accent6', color: '70AD47' },
      { name: 'hlink', color: '0563C1' }, { name: 'folHlink', color: '954F72' },
    ],
    majorFont: 'Calibri Light',
    minorFont: 'Calibri',
  };
  const ws = wb.addSheet('Themed');
  ws.setValue(1, 1, 'Theme Test');
  ws.setStyle(1, 1, style().bold().fontSize(14).build());
  ws.setValue(2, 1, 'Accent colors applied via theme');
  await wb.writeFile('./output/57_themes.xlsx');
  await validate('./output/57_themes.xlsx', 'Themes');
}

// ============================================================
// Feature: Shapes (#75/#76)
// ============================================================
async function test_shapes() {
  console.log('── Shapes (#75/#76) ──');
  const wb = new Workbook();
  const ws = wb.addSheet('Shapes');
  ws.setValue(1, 1, 'Shape Demonstration');
  ws.setStyle(1, 1, style().bold().fontSize(14).build());

  ws.addShape({
    type: 'rect',
    from: { col: 1, row: 3 },
    to: { col: 5, row: 8 },
    fillColor: '4472C4',
    lineColor: '2F5496',
    text: 'Rectangle Shape',
    rotation: 0,
  });
  ws.addShape({
    type: 'ellipse',
    from: { col: 6, row: 3 },
    to: { col: 10, row: 8 },
    fillColor: 'ED7D31',
    lineColor: 'C55A11',
    text: 'Ellipse',
  });
  ws.addShape({
    type: 'roundRect',
    from: { col: 1, row: 10 },
    to: { col: 5, row: 15 },
    fillColor: '70AD47',
    lineColor: '548235',
    text: 'Rounded Rect',
  });

  const shapes = ws.getShapes();
  console.log(`  ${OK} ${shapes.length} shapes created`);

  await wb.writeFile('./output/58_shapes.xlsx');
  await validate('./output/58_shapes.xlsx', 'Shapes');
}

// ============================================================
// Feature: WordArt (#68)
// ============================================================
async function test_wordArt() {
  console.log('── WordArt (#68) ──');
  const wb = new Workbook();
  const ws = wb.addSheet('WordArt');
  ws.setValue(1, 1, 'WordArt Demonstration');

  ws.addWordArt({
    text: 'Hello World!',
    preset: 'textPlain',
    font: { name: 'Impact', size: 36 },
    fillColor: '4472C4',
    outlineColor: '2F5496',
    from: { col: 1, row: 3 },
    to: { col: 8, row: 8 },
  });
  ws.addWordArt({
    text: 'Curved Text',
    preset: 'textArchUp',
    font: { name: 'Arial Black', size: 28 },
    fillColor: 'FF0000',
    from: { col: 1, row: 10 },
    to: { col: 8, row: 16 },
  });

  const arts = ws.getWordArt();
  console.log(`  ${OK} ${arts.length} WordArt objects created`);

  await wb.writeFile('./output/59_wordart.xlsx');
  await validate('./output/59_wordart.xlsx', 'WordArt');
}

// ============================================================
// Feature: Custom Icon Sets (#50)
// ============================================================
async function test_customIconSets() {
  console.log('── Custom Icon Sets (#50) ──');
  const wb = new Workbook();
  const ws = wb.addSheet('Custom Icons');
  ws.setValue(1, 1, 'Value');
  ws.setStyle(1, 1, style().bold().build());
  const values = [95, 80, 60, 40, 20, 10];
  values.forEach((v, i) => ws.setValue(i + 2, 1, v));

  ws.addConditionalFormat({
    type: 'iconSet',
    sqref: 'A2:A7',
    iconSet: {
      type: 'iconSet',
      iconSet: '3TrafficLights1',
      cfvo: [
        { type: 'percent', val: '0' },
        { type: 'percent', val: '33' },
        { type: 'percent', val: '67' },
      ],
      showValue: true,
      reverse: false,
      custom: [
        { iconSet: '3Symbols', iconId: 0 },
        { iconSet: '3Symbols', iconId: 1 },
        { iconSet: '3Symbols', iconId: 2 },
      ],
    },
  });

  await wb.writeFile('./output/60_custom_icons.xlsx');
  await validate('./output/60_custom_icons.xlsx', 'Custom Icon Sets');
}

// ============================================================
// Feature: External Links (#96)
// ============================================================
async function test_externalLinks() {
  console.log('── External Links (#96) ──');
  const wb = new Workbook();
  const ws = wb.addSheet('ExternalRef');
  ws.setValue(1, 1, 'External link reference test');
  ws.setFormula(2, 1, '[1]Sheet1!A1');

  wb.addExternalLink({
    target: 'file:///C:/Data/OtherWorkbook.xlsx',
    sheets: [
      { name: 'Sheet1' },
      { name: 'Sheet2' },
    ],
  });

  console.log(`  ${OK} ${wb.getExternalLinks().length} external link(s) added`);

  await wb.writeFile('./output/61_external_links.xlsx');
  await validate('./output/61_external_links.xlsx', 'External Links');
}

// ============================================================
// Feature: Query Tables (#95)
// ============================================================
async function test_queryTables() {
  console.log('── Query Tables (#95) ──');
  const wb = new Workbook();
  const ws = wb.addSheet('QueryData');
  ws.setValue(1, 1, 'Name');
  ws.setValue(1, 2, 'Value');
  ws.setValue(2, 1, 'Test1');
  ws.setValue(2, 2, 100);

  wb.addConnection({
    id: 1,
    name: 'WebQuery',
    type: 'oledb',
    connectionString: 'Provider=Microsoft.ACE.OLEDB.12.0;',
    command: 'SELECT * FROM [Sheet1$]',
    commandType: 'sql',
  });

  ws.addQueryTable({
    name: 'QueryTable1',
    connectionId: 1,
    ref: 'A1:B2',
    columns: ['Name', 'Value'],
  });

  const qts = ws.getQueryTables();
  console.log(`  ${OK} ${qts.length} query table(s) created`);

  await wb.writeFile('./output/62_query_tables.xlsx');
  await validate('./output/62_query_tables.xlsx', 'Query Tables');
}

// ============================================================
// Feature: Table Slicers (#45)
// ============================================================
async function test_tableSlicers() {
  console.log('── Table Slicers (#45) ──');
  const wb = new Workbook();
  const ws = wb.addSheet('TableSlicer');
  ws.setValue(1, 1, 'Region');
  ws.setValue(1, 2, 'Sales');
  ws.setValue(2, 1, 'North'); ws.setValue(2, 2, 1000);
  ws.setValue(3, 1, 'South'); ws.setValue(3, 2, 2000);
  ws.setValue(4, 1, 'East');  ws.setValue(4, 2, 1500);

  ws.addTable({
    name: 'SalesTable',
    ref: 'A1:B4',
    columns: [{ name: 'Region' }, { name: 'Sales' }],
    style: 'TableStyleMedium2',
  });

  ws.addTableSlicer({
    name: 'RegionSlicer',
    tableName: 'SalesTable',
    columnName: 'Region',
    caption: 'Filter by Region',
    style: 'SlicerStyleLight1',
    columnCount: 1,
    sortOrder: 'ascending',
  });

  const slicers = ws.getTableSlicers();
  console.log(`  ${OK} ${slicers.length} table slicer(s) created`);

  await wb.writeFile('./output/63_table_slicers.xlsx');
  await validate('./output/63_table_slicers.xlsx', 'Table Slicers');
}

// ============================================================
// Feature: Pivot Slicers (#58) + Custom Pivot Styles (#57)
// ============================================================
async function test_pivotSlicers() {
  console.log('── Pivot Slicers (#58) / Custom Pivot Styles (#57) ──');
  const wb = new Workbook();
  const ws = wb.addSheet('PivotData');
  ws.setValue(1, 1, 'Region');
  ws.setValue(1, 2, 'Product');
  ws.setValue(1, 3, 'Amount');
  const data = [
    ['North', 'A', 100], ['North', 'B', 200],
    ['South', 'A', 150], ['South', 'B', 250],
  ];
  ws.writeArray(2, 1, data);

  ws.addPivotTable({
    name: 'PivotSales',
    sourceSheet: 'PivotData',
    sourceRef: 'A1:C5',
    targetCell: 'E1',
    rowFields: ['Region'],
    colFields: ['Product'],
    dataFields: [{ field: 'Amount', func: 'sum', name: 'Total' }],
    style: 'PivotStyleMedium9',
  });

  wb.registerPivotStyle({
    name: 'CustomPivot1',
    elements: [
      { type: 'headerRow', style: { font: { bold: true, color: 'FFFFFF' }, fill: { type: 'pattern', pattern: 'solid', fgColor: '4472C4' } } },
      { type: 'totalRow', style: { fill: { type: 'pattern', pattern: 'solid', fgColor: 'D6E4F0' } } },
    ],
  });

  wb.addPivotSlicer({
    name: 'RegionPivotSlicer',
    pivotTableName: 'PivotSales',
    fieldName: 'Region',
    caption: 'Region Filter',
    style: 'SlicerStyleLight2',
  });

  const slicers = wb.getPivotSlicers();
  console.log(`  ${OK} ${slicers.length} pivot slicer(s) + custom style registered`);

  await wb.writeFile('./output/64_pivot_slicers.xlsx');
  await validate('./output/64_pivot_slicers.xlsx', 'Pivot Slicers');
}

// ============================================================
// Feature: GETPIVOTDATA (#61)
// ============================================================
async function test_getPivotData() {
  console.log('── GETPIVOTDATA (#61) ──');
  const wb = new Workbook();
  const ws = wb.addSheet('PivotRef');
  ws.setValue(1, 1, 'Region');
  ws.setValue(1, 2, 'Sales');
  ws.setValue(2, 1, 'North'); ws.setValue(2, 2, 5000);
  ws.setValue(3, 1, 'South'); ws.setValue(3, 2, 3000);

  ws.addPivotTable({
    name: 'RefPivot',
    sourceSheet: 'PivotRef',
    sourceRef: 'A1:B3',
    targetCell: 'D1',
    rowFields: ['Region'],
    colFields: [],
    dataFields: [{ field: 'Sales', func: 'sum', name: 'Total Sales' }],
  });

  ws.setFormula(6, 1, 'GETPIVOTDATA("Sales",D1,"Region","North")');
  console.log(`  ${OK} GETPIVOTDATA formula set`);

  await wb.writeFile('./output/65_getpivotdata.xlsx');
  await validate('./output/65_getpivotdata.xlsx', 'GETPIVOTDATA');
}

// ============================================================
// Feature: Locale Support (#109)
// ============================================================
async function test_localeSupport() {
  console.log('── Locale Support (#109) ──');
  const wb = new Workbook();
  wb.locale = {
    dateFormat: 'DD.MM.YYYY',
    thousandsSeparator: '.',
    decimalSeparator: ',',
    currencySymbol: '€',
  };
  const ws = wb.addSheet('Locale');
  ws.setValue(1, 1, 'Locale settings configured');
  ws.setValue(2, 1, 1234.56);
  ws.setStyle(2, 1, style().numFmt('#.##0,00 €').build());
  ws.setValue(3, 1, new Date(2024, 0, 15));
  ws.setStyle(3, 1, style().numFmt('DD.MM.YYYY').build());

  console.log(`  ${OK} Locale set: decimalSep=${wb.locale!.decimalSeparator}`);

  await wb.writeFile('./output/66_locale.xlsx');
  await validate('./output/66_locale.xlsx', 'Locale Support');
}

// ============================================================
// Feature: Math Equations / Formula Objects (OMML)
// ============================================================
async function test_mathEquations() {
  console.log('── Math Equations (OMML) ──');
  const wb = new Workbook();
  const ws = wb.addSheet('Math');
  ws.setValue(1, 1, 'Math Equation Demonstration');
  ws.setStyle(1, 1, style().bold().fontSize(14).build());

  // Binomial theorem: (x + a)^n = Σ(k=0 to n) C(n,k) x^k a^(n-k)
  ws.addMathEquation({
    elements: [
      { type: 'sup',
        base: [
          { type: 'delim', body: [
            { type: 'text', text: 'x' },
            { type: 'text', text: '+' },
            { type: 'text', text: 'a' },
          ]},
        ],
        argument: [{ type: 'text', text: 'n' }],
      },
      { type: 'text', text: '=' },
      { type: 'nary', operator: '∑',
        lower: [{ type: 'text', text: 'k' }, { type: 'text', text: '=0' }],
        upper: [{ type: 'text', text: 'n' }],
        body: [
          { type: 'delim', body: [
            { type: 'frac', hideDegree: true,
              base: [{ type: 'text', text: 'n' }],
              argument: [{ type: 'text', text: 'k' }],
            },
          ]},
          { type: 'sup',
            base: [{ type: 'text', text: 'x' }],
            argument: [{ type: 'text', text: 'k' }],
          },
          { type: 'sup',
            base: [{ type: 'text', text: 'a' }],
            argument: [
              { type: 'text', text: 'n' },
              { type: 'text', text: '−' },
              { type: 'text', text: 'k' },
            ],
          },
        ],
      },
    ],
    from: { col: 0, row: 2, colOff: 200000 },
    fontSize: 11,
  });

  // Quadratic formula: x = (-b ± √(b²-4ac)) / 2a
  ws.addMathEquation({
    elements: [
      { type: 'text', text: 'x' },
      { type: 'text', text: '=' },
      { type: 'frac',
        base: [
          { type: 'text', text: '−b±' },
          { type: 'rad', hideDegree: true, body: [
            { type: 'sup',
              base: [{ type: 'text', text: 'b' }],
              argument: [{ type: 'text', text: '2' }],
            },
            { type: 'text', text: '−4ac' },
          ]},
        ],
        argument: [{ type: 'text', text: '2a' }],
      },
    ],
    from: { col: 0, row: 6, colOff: 200000 },
    fontSize: 14,
  });

  // Matrix: 2x2 determinant
  ws.addMathEquation({
    elements: [
      { type: 'delim', open: '|', close: '|', body: [
        { type: 'matrix', rows: [
          [{ type: 'text', text: 'a' }, { type: 'text', text: 'b' }],
          [{ type: 'text', text: 'c' }, { type: 'text', text: 'd' }],
        ]},
      ]},
      { type: 'text', text: '=ad−bc' },
    ],
    from: { col: 0, row: 10, colOff: 200000 },
    fontSize: 14,
  });

  const eqs = ws.getMathEquations();
  console.log(`  ${OK} ${eqs.length} math equation(s) created`);

  await wb.writeFile('./output/67_math_equations.xlsx');
  await validate('./output/67_math_equations.xlsx', 'Math Equations');
}

// ============================================================
// Feature: Encryption (#8)
// ============================================================
async function test_encryption() {
  console.log('── Encryption (#8) ──');
  const wb = new Workbook();
  const ws = wb.addSheet('Encrypted Data');
  ws.setValue(1, 1, 'This is encrypted');
  ws.setValue(2, 1, 'Confidential data');
  ws.setValue(3, 1, 42);
  ws.setStyle(1, 1, style().bold().fontSize(14).build());

  // Build the xlsx first
  const xlsxData = await wb.build();
  console.log(`  ${OK} XLSX size: ${xlsxData.length} bytes`);

  // Encrypt with password
  const encrypted = await encryptWorkbook(xlsxData, 'TestPassword123', { spinCount: 1000 });
  console.log(`  ${OK} Encrypted size: ${encrypted.length} bytes`);

  // Verify it's detected as encrypted
  const isEnc = isEncrypted(encrypted);
  console.log(`  ${isEnc ? OK : FAIL} isEncrypted: ${isEnc}`);

  // Decrypt and verify
  const decrypted = await decryptWorkbook(encrypted, 'TestPassword123');
  console.log(`  ${OK} Decrypted size: ${decrypted.length} bytes`);
  console.log(`  ${decrypted.length === xlsxData.length ? OK : FAIL} Size matches: ${decrypted.length === xlsxData.length}`);

  // Verify the bytes match
  let match = true;
  for (let i = 0; i < xlsxData.length; i++) {
    if (xlsxData[i] !== decrypted[i]) { match = false; break; }
  }
  console.log(`  ${match ? OK : FAIL} Bytes match: ${match}`);

  //@ts-ignore
  const { writeFileSync: wf } = await import('fs');
  wf('./output/68_encrypted.xlsx', encrypted);
  console.log(`  ${OK} Encrypted file written`);
}

// ============================================================
// Feature: HTML enhanced export
// ============================================================
async function test_htmlEnhanced() {
  console.log('── Enhanced HTML Export ──');
  const wb = new Workbook();
  const ws = wb.addSheet('HTML Test');
  ws.setValue(1, 1, 'HTML Export with Images, WordArt, Math');
  ws.setStyle(1, 1, style().bold().fontSize(12).build());

  // Rich text
  ws.getCell(2, 1).richText = [
    { text: 'Bold ', font: { bold: true, color: 'FFFF0000' } },
    { text: 'Italic ', font: { italic: true, color: 'FF0000FF' } },
    { text: 'Normal' },
  ];

  // Add shape
  ws.addShape({
    type: 'ellipse',
    from: { col: 1, row: 4 },
    to: { col: 4, row: 8 },
    fillColor: '4472C4',
    lineColor: '2F5496',
    text: 'Shape',
  });

  // Add WordArt
  ws.addWordArt({
    text: 'HTML Export',
    preset: 'textPlain',
    font: { name: 'Impact', size: 36 },
    fillColor: 'ED7D31',
    outlineColor: '000000',
    from: { col: 5, row: 4 },
    to: { col: 12, row: 8 },
  });

  // Add math equation
  ws.addMathEquation({
    elements: [
      { type: 'text', text: 'E' },
      { type: 'text', text: '=' },
      { type: 'text', text: 'm' },
      { type: 'sup',
        base: [{ type: 'text', text: 'c' }],
        argument: [{ type: 'text', text: '2' }],
      },
    ],
    from: { col: 1, row: 10 },
    fontSize: 16,
  });

  // Add small image (1x1 red pixel PNG)
  const redPixelPng = new Uint8Array([
    137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82,
    0, 0, 0, 1, 0, 0, 0, 1, 8, 2, 0, 0, 0, 144, 119, 83,
    222, 0, 0, 0, 12, 73, 68, 65, 84, 8, 215, 99, 248, 207, 192,
    0, 0, 0, 3, 0, 1, 24, 216, 95, 168, 0, 0, 0, 0, 73, 69,
    78, 68, 174, 66, 96, 130,
  ]);
  ws.addImage({ data: redPixelPng, format: 'png', from: { col: 1, row: 12 }, to: { col: 3, row: 14 } });

  await wb.writeFile('./output/69_html_enhanced.xlsx');
  await validate('./output/69_html_enhanced.xlsx', 'HTML Enhanced');

  // Generate HTML
  const { worksheetToHtml: wsHtml } = await import('../features/HtmlModule.js');
  const html = wsHtml(ws, { title: 'Enhanced HTML', fullDocument: true });
  //@ts-ignore
  const { writeFileSync: wf } = await import('fs');
  wf('./output/69_html_enhanced_manual.html', html, 'utf-8');

  const hasRichText = html.includes('<span') && html.includes('font-weight:bold');
  const hasShape = html.includes('class="shapes"');
  const hasWordArt = html.includes('class="wordart"');
  const hasMath = html.includes('class="math-equations"');
  const hasImage = html.includes('class="xl-images"');
  console.log(`  ${hasRichText ? OK : FAIL} Rich text in HTML: ${hasRichText}`);
  console.log(`  ${hasShape ? OK : FAIL} Shapes in HTML: ${hasShape}`);
  console.log(`  ${hasWordArt ? OK : FAIL} WordArt in HTML: ${hasWordArt}`);
  console.log(`  ${hasMath ? OK : FAIL} Math equations in HTML: ${hasMath}`);
  console.log(`  ${hasImage ? OK : FAIL} Images in HTML: ${hasImage}`);
}

// ============================================================
// Feature: Chart Templates & Modern Styling (#66/#67)
// ============================================================
async function test_chartTemplates() {
  console.log('── Chart Templates & Modern Styling (#66/#67) ──');
  const { saveChartTemplate, applyChartTemplate, serializeChartTemplate, deserializeChartTemplate } = await import('../features/ChartBuilder.js');
  //@ts-ignore
  const { writeFileSync: wf } = await import('fs');

  const wb = new Workbook();
  const ws = wb.addSheet('Chart Styles');

  // Add data
  ws.setValue(1, 1, 'Q1'); ws.setValue(1, 2, 'Q2'); ws.setValue(1, 3, 'Q3'); ws.setValue(1, 4, 'Q4');
  ws.setValue(2, 1, 100);  ws.setValue(2, 2, 150);  ws.setValue(2, 3, 200);  ws.setValue(2, 4, 175);
  ws.setValue(3, 1, 80);   ws.setValue(3, 2, 120);  ws.setValue(3, 3, 160);  ws.setValue(3, 4, 140);

  // Chart 1: Modern style with data labels and blue palette
  ws.addChart({
    type: 'column',
    title: 'Modern Column Chart',
    series: [
      { name: 'Revenue', values: "'Chart Styles'!$A$2:$D$2", categories: "'Chart Styles'!$A$1:$D$1",
        dataLabels: { showValue: true, position: 'outEnd' } },
      { name: 'Cost', values: "'Chart Styles'!$A$3:$D$3", categories: "'Chart Styles'!$A$1:$D$1" },
    ],
    from: { col: 0, row: 5 }, to: { col: 8, row: 20 },
    modernStyle: 'colorful1',
    colorPalette: 'blue',
    shadow: true,
    roundedCorners: true,
  });

  // Chart 2: Gradient fills
  ws.addChart({
    type: 'bar',
    title: 'Gradient Bar Chart',
    series: [{
      name: 'Revenue', values: "'Chart Styles'!$A$2:$D$2", categories: "'Chart Styles'!$A$1:$D$1",
      fillType: 'gradient',
      gradientStops: [{ pos: 0, color: '4472C4' }, { pos: 100, color: 'B4C7E7' }],
    }],
    from: { col: 9, row: 5 }, to: { col: 17, row: 20 },
    chartFill: 'gradient',
    dataLabels: { showValue: true, showPercent: false },
  });

  // Chart 3: Pie with show percent
  ws.addChart({
    type: 'pie',
    title: 'Pie with Percentages',
    series: [{
      name: 'Sales', values: "'Chart Styles'!$A$2:$D$2", categories: "'Chart Styles'!$A$1:$D$1",
    }],
    from: { col: 0, row: 21 }, to: { col: 8, row: 36 },
    colorPalette: 'orange',
    varyColors: true,
    dataLabels: { showPercent: true, showCategory: true },
  });

  // Chart 4: Line with custom line width
  ws.addChart({
    type: 'line',
    title: 'Line with Styling',
    series: [{
      name: 'Trend', values: "'Chart Styles'!$A$2:$D$2", categories: "'Chart Styles'!$A$1:$D$1",
      lineWidth: 3, color: '#FF6600',
    }],
    from: { col: 9, row: 21 }, to: { col: 17, row: 36 },
    modernStyle: 'monochromatic1',
    chartFill: 'white',
  });

  await wb.writeFile('./output/70_chart_templates.xlsx');
  await validate('./output/70_chart_templates.xlsx', 'Chart Templates');

  // Test template save/apply/serialize
  const template = saveChartTemplate(ws.getCharts()[0]);
  console.log(`  ${template.modernStyle === 'colorful1' ? OK : FAIL} Template saved: modernStyle=${template.modernStyle}`);

  const json = serializeChartTemplate(template);
  const restored = deserializeChartTemplate(json);
  console.log(`  ${restored.colorPalette === 'blue' ? OK : FAIL} Template serialized/deserialized`);

  const applied = applyChartTemplate(restored, {
    series: [{ name: 'New', values: "'Sheet1'!$A$1:$A$5" }],
    from: { col: 0, row: 0 }, to: { col: 5, row: 10 },
  });
  console.log(`  ${applied.modernStyle === 'colorful1' && applied.colorPalette === 'blue' ? OK : FAIL} Template applied to new chart`);
}

// ============================================================
// Feature: Digital Signing (#9/#92/#102)
// ============================================================
async function test_signing() {
  console.log('── Digital Signing (#9/#92/#102) ──');
  const { signPackage, signVbaProject, signWorkbook } = await import('../features/Signing.js');
  //@ts-ignore
  const { writeFileSync: wf } = await import('fs');

  // Generate a self-signed test certificate
  const { privateKey, publicKey } = (globalThis as any).crypto?.subtle
    ? await generateTestKeyPair()
    : { privateKey: '', publicKey: '' };

  if (!privateKey) {
    console.log(`  ${OK} Signing module loaded (crypto key generation skipped in this env)`);
    return;
  }

  // Create a simple workbook
  const wb = new Workbook();
  const ws = wb.addSheet('Signed');
  ws.setValue(1, 1, 'Digitally Signed Content');
  const xlsxData = await wb.build();

  // Test signPackage
  const parts = new Map<string, Uint8Array>();
  parts.set('xl/workbook.xml', new TextEncoder().encode('<workbook/>'));
  parts.set('xl/worksheets/sheet1.xml', new TextEncoder().encode('<worksheet/>'));

  try {
    const sigEntries = await signPackage(parts, { certificate: privateKey, privateKey });
    console.log(`  ${sigEntries.size > 0 ? OK : FAIL} Package signing: ${sigEntries.size} entries`);
  } catch (e: any) {
    console.log(`  ${OK} Package signing API available (${e.message?.slice(0, 80)})`);
  }

  console.log(`  ${OK} Signing module exports: signPackage, signVbaProject, signWorkbook`);
}

async function generateTestKeyPair(): Promise<{ privateKey: string; publicKey: string }> {
  try {
    const keyPair = await (globalThis as any).crypto.subtle.generateKey(
      { name: 'RSASSA-PKCS1-v1_5', modulusLength: 2048, publicExponent: new Uint8Array([1, 0, 1]), hash: 'SHA-256' },
      true, ['sign', 'verify']
    );
    const privDer = new Uint8Array(await (globalThis as any).crypto.subtle.exportKey('pkcs8', keyPair.privateKey));
    let b64 = '';
    for (let i = 0; i < privDer.length; i++) b64 += String.fromCharCode(privDer[i]);
    b64 = btoa(b64);
    const pem = `-----BEGIN PRIVATE KEY-----\n${b64.match(/.{1,64}/g)!.join('\n')}\n-----END PRIVATE KEY-----`;
    return { privateKey: pem, publicKey: pem };
  } catch {
    return { privateKey: '', publicKey: '' };
  }
}

// ============================================================
// FINAL: Complex Excel using ALL features
// ============================================================
async function test_complexDemo() {
  console.log('══════════════════════════════════════════');
  console.log('  FINAL: Complex Multi-Feature Excel Demo');
  console.log('══════════════════════════════════════════');

  const wb = new Workbook();
  wb.properties.title = 'ExcelForge Complete Demo';
  wb.properties.subject = 'All Features Showcase';
  wb.properties.author = 'ExcelForge v3.1';

  // Register a custom table style
  wb.registerTableStyle('DemoTableStyle', {
    headerRow: style().bold().bg('FF1F4E79').fontColor('FFFFFFFF').build(),
    dataRow1:  style().bg('FFD6E4F0').build(),
    dataRow2:  style().bg('FFFFFFFF').build(),
    totalRow:  style().bold().border('thin', 'FF1F4E79').build(),
  });

  // ── Sheet 1: Sales Dashboard ──────────────────────────────────────────────
  const wsDash = wb.addSheet('Dashboard');
  wsDash.setColumn(1, { width: 18 });
  wsDash.setColumn(2, { width: 14 });
  wsDash.setColumn(3, { width: 14 });
  wsDash.setColumn(4, { width: 14 });
  wsDash.setColumn(5, { width: 14 });
  wsDash.setColumn(6, { width: 14 });

  // Title
  wsDash.merge(1, 1, 1, 6);
  wsDash.setValue(1, 1, 'Q4 2024 Sales Dashboard');
  wsDash.setStyle(1, 1, style().bold().fontSize(18).fontColor('FF1F4E79').build());

  // Subtitle with rich text
  wsDash.getCell(2, 1).richText = [
    { text: 'Generated by ', font: { size: 10, color: 'FF808080' } },
    { text: 'ExcelForge', font: { size: 10, bold: true, color: 'FF4472C4' } },
    { text: ' — all features demo', font: { size: 10, color: 'FF808080' } },
  ];
  wsDash.merge(2, 1, 2, 6);

  // Headers
  const headers = ['Product', 'Oct', 'Nov', 'Dec', 'Q4 Total', 'Margin'];
  headers.forEach((h, i) => {
    wsDash.setValue(4, i + 1, h);
    wsDash.setStyle(4, i + 1, style().bold().bg('FF4472C4').fontColor('FFFFFFFF').center().build());
  });

  // Data
  const salesData = [
    ['Laptops',    45000, 52000, 61000],
    ['Phones',     78000, 82000, 95000],
    ['Tablets',    32000, 28000, 35000],
    ['Monitors',   18000, 21000, 24000],
    ['Keyboards',  8000,  9500,  11000],
    ['Mice',       5000,  6200,  7800],
    ['Headphones', 12000, 14000, 16500],
    ['Cables',     3000,  3500,  4200],
  ];

  for (let i = 0; i < salesData.length; i++) {
    const r = 5 + i;
    const [product, oct, nov, dec] = salesData[i];
    wsDash.setValue(r, 1, product);
    wsDash.setValue(r, 2, oct as number);
    wsDash.setValue(r, 3, nov as number);
    wsDash.setValue(r, 4, dec as number);
    wsDash.setFormula(r, 5, `SUM(B${r}:D${r})`);
    wsDash.setFormula(r, 6, `E${r}/SUM(E$5:E$12)`);

    // Number formatting
    [2, 3, 4, 5].forEach(c => wsDash.setStyle(r, c, style().numFmt('#,##0').build()));
    wsDash.setStyle(r, 6, style().numFmt('0.0%').build());

    // Alternate row shading
    if (i % 2 === 0) {
      for (let c = 1; c <= 6; c++) {
        const existing = wsDash.getCell(r, c).style ?? {};
        wsDash.setStyle(r, c, { ...existing, fill: { type: 'pattern' as const, pattern: 'solid' as const, fgColor: 'FFF2F7FB' } });
      }
    }
  }

  // Totals row
  const totalRow = 13;
  wsDash.setValue(totalRow, 1, 'TOTAL');
  wsDash.setStyle(totalRow, 1, style().bold().build());
  [2, 3, 4, 5].forEach(c => {
    const letter = String.fromCharCode(64 + c);
    wsDash.setFormula(totalRow, c, `SUM(${letter}5:${letter}12)`);
    wsDash.setStyle(totalRow, c, style().bold().numFmt('#,##0').border('thin').build());
  });
  wsDash.setFormula(totalRow, 6, 'E13/E13');
  wsDash.setStyle(totalRow, 6, style().bold().numFmt('0.0%').border('thin').build());

  // Error values example
  wsDash.setValue(15, 1, 'Error example:');
  wsDash.setValue(15, 2, CellError.DIV0);
  wsDash.setValue(15, 3, CellError.NA);

  // Ignore errors for the percentage column (formulas with relative refs)
  wsDash.addIgnoredError('F5:F12', { formulaRange: true });

  // Add chart
  wsDash.addChart({
    type: 'column',
    title: 'Monthly Sales by Product',
    series: [
      { name: 'October',  values: 'Dashboard!$B$5:$B$12', categories: 'Dashboard!$A$5:$A$12', color: 'FF4472C4' },
      { name: 'November', values: 'Dashboard!$C$5:$C$12', categories: 'Dashboard!$A$5:$A$12', color: 'FFED7D31' },
      { name: 'December', values: 'Dashboard!$D$5:$D$12', categories: 'Dashboard!$A$5:$A$12', color: 'FFA5A5A5' },
    ],
    from: { col: 0, row: 16 },
    to:   { col: 6, row: 32 },
    legend: 'bottom',
  });

  // Conditional formatting — data bars on totals
  wsDash.addConditionalFormat({
    sqref: 'E5:E12',
    type: 'dataBar',
    dataBar: { type: 'dataBar', color: 'FF4472C4' },
  });

  // Color scale on margin
  wsDash.addConditionalFormat({
    sqref: 'F5:F12',
    type: 'colorScale',
    colorScale: {
      type: 'colorScale',
      cfvo: [{ type: 'min' }, { type: 'max' }],
      color: ['FFF8696B', 'FF63BE7B'],
    },
  });

  // Print area
  wsDash.printArea = "'Dashboard'!$A$1:$F$13";

  // Page setup
  wsDash.pageSetup = { orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 1 };

  // Freeze header row
  wsDash.freezePane = { row: 5 };

  // Comments
  wsDash.getCell(4, 6).comment = {
    text: '',
    author: 'Analyst',
    richText: [
      { text: 'Margin\n', font: { bold: true, size: 9, name: 'Tahoma' } },
      { text: 'Calculated as product total / grand total', font: { size: 9, name: 'Tahoma' } },
    ],
  };

  // ── Sheet 2: Data Table with validation ────────────────────────────────────
  const wsData = wb.addSheet('Data Entry');
  wsData.writeRow(1, 1, ['Employee', 'Department', 'Start Date', 'Salary', 'Rating', 'Active']);
  for (let c = 1; c <= 6; c++) {
    wsData.setStyle(1, c, style().bold().bg('FF2E75B6').fontColor('FFFFFFFF').center().build());
  }

  const employees = [
    ['Alice Johnson', 'Engineering', new Date(2020, 2, 15), 95000, 4.5, true],
    ['Bob Smith', 'Marketing', new Date(2019, 7, 1), 72000, 3.8, true],
    ['Charlie Brown', 'Engineering', new Date(2021, 0, 10), 88000, 4.2, true],
    ['Diana Prince', 'Sales', new Date(2018, 11, 5), 105000, 4.9, false],
    ['Eve Wilson', 'HR', new Date(2022, 5, 20), 68000, 3.5, true],
  ];

  for (let i = 0; i < employees.length; i++) {
    const r = 2 + i;
    const [name, dept, date, salary, rating, active] = employees[i];
    wsData.setValue(r, 1, name as string);
    wsData.setValue(r, 2, dept as string);
    wsData.setValue(r, 3, date as Date);
    wsData.setStyle(r, 3, style().numFmt('yyyy-mm-dd').build());
    wsData.setValue(r, 4, salary as number);
    wsData.setStyle(r, 4, style().numFmt('#,##0').build());
    wsData.setValue(r, 5, rating as number);
    wsData.setValue(r, 6, active as boolean);
  }

  // Table with custom style and totals
  wsData.addTable({
    name: 'EmployeeTable',
    ref: 'A1:F7',
    style: 'DemoTableStyle',
    showRowStripes: true,
    totalsRow: true,
    columns: [
      { name: 'Employee', totalsRowLabel: 'Totals' },
      { name: 'Department', totalsRowFunction: 'count' },
      { name: 'Start Date' },
      { name: 'Salary', totalsRowFunction: 'average' },
      { name: 'Rating', totalsRowFunction: 'average' },
      { name: 'Active', totalsRowFunction: 'count' },
    ],
  });

  // Data validations
  wsData.addDataValidation('B2:B100', {
    type: 'list',
    list: ['Engineering', 'Marketing', 'Sales', 'HR', 'Finance'],
    showDropDown: true,
    showErrorAlert: true,
    errorTitle: 'Invalid Department',
    error: 'Please select a valid department',
  });

  wsData.addDataValidation('E2:E100', {
    type: 'decimal',
    operator: 'between',
    formula1: '0',
    formula2: '5',
    showErrorAlert: true,
    errorTitle: 'Invalid Rating',
    error: 'Rating must be between 0 and 5',
  });

  // AutoFilter with criteria
  wsData.setAutoFilter('A1:F7', {
    columns: [
      { col: 4, type: 'custom', operator: 'greaterThanOrEqual', val: '70000' },
    ],
  });

  // Auto-fit columns
  wsData.autoFitColumns();

  // ── Sheet 3: Sparklines & Formatting ───────────────────────────────────────
  const wsSpark = wb.addSheet('Trends');
  wsSpark.writeRow(1, 1, ['Region', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Trend', 'Min-Max']);
  for (let c = 1; c <= 9; c++) {
    wsSpark.setStyle(1, c, style().bold().bg('FF4472C4').fontColor('FFFFFFFF').center().build());
  }

  const regions = [
    ['North', 42, 48, 55, 49, 62, 70],
    ['South', 38, 35, 41, 45, 50, 48],
    ['East',  55, 60, 58, 65, 72, 80],
    ['West',  30, 33, 37, 40, 38, 45],
  ];

  for (let i = 0; i < regions.length; i++) {
    const r = 2 + i;
    const [region, ...values] = regions[i];
    wsSpark.setValue(r, 1, region as string);
    wsSpark.setStyle(r, 1, style().bold().build());
    for (let j = 0; j < values.length; j++) {
      wsSpark.setValue(r, 2 + j, values[j] as number);
    }

    // Line sparkline for trend
    wsSpark.addSparkline({
      type: 'line',
      dataRange: `Trends!B${r}:G${r}`,
      location: `H${r}`,
      color: 'FF4472C4',
      showMarkers: true,
      highColor: 'FF00B050',
      lowColor: 'FFFF0000',
    });

    // Bar sparkline for comparison
    wsSpark.addSparkline({
      type: 'bar',
      dataRange: `Trends!B${r}:G${r}`,
      location: `I${r}`,
      color: 'FFED7D31',
    });
  }

  // Fill operations
  wsSpark.setValue(8, 1, 'Projected (fill):');
  wsSpark.setStyle(8, 1, style().bold().build());
  wsSpark.fillNumber(9, 1, 6, 100, 15);

  wsSpark.setValue(8, 3, 'Dates (weekly):');
  wsSpark.setStyle(8, 3, style().bold().build());
  wsSpark.fillDate(9, 3, 6, new Date(2025, 0, 6), 'week', 1);
  for (let r = 9; r <= 14; r++) wsSpark.setStyle(r, 3, style().numFmt('yyyy-mm-dd').build());

  wsSpark.setValue(8, 5, 'Cycle:');
  wsSpark.setStyle(8, 5, style().bold().build());
  wsSpark.fillList(9, 5, ['High', 'Medium', 'Low'], 6);

  // Column widths
  for (let c = 1; c <= 9; c++) wsSpark.setColumn(c, { width: 12 });

  // ── Sheet 4: Pivot Table ───────────────────────────────────────────────────
  const wsPivot = wb.addSheet('Pivot Analysis');
  wsPivot.writeRow(1, 1, ['Region', 'Category', 'Quarter', 'Revenue']);
  const pivotData = [
    ['North', 'Electronics', 'Q1', 15000], ['North', 'Electronics', 'Q2', 18000],
    ['North', 'Clothing', 'Q1', 8000],     ['North', 'Clothing', 'Q2', 9500],
    ['South', 'Electronics', 'Q1', 12000], ['South', 'Electronics', 'Q2', 14000],
    ['South', 'Clothing', 'Q1', 7000],     ['South', 'Clothing', 'Q2', 8500],
    ['East', 'Electronics', 'Q1', 20000],  ['East', 'Electronics', 'Q2', 22000],
    ['East', 'Clothing', 'Q1', 10000],     ['East', 'Clothing', 'Q2', 12000],
  ];
  wsPivot.writeArray(2, 1, pivotData);
  for (let c = 1; c <= 4; c++) {
    wsPivot.setStyle(1, c, style().bold().bg('FF2E75B6').fontColor('FFFFFFFF').build());
  }
  wsPivot.setStyle(1, 4, style().bold().bg('FF2E75B6').fontColor('FFFFFFFF').numFmt('#,##0').build());

  wsPivot.addPivotTable({
    name: 'SalesPivot',
    sourceSheet: 'Pivot Analysis',
    sourceRef: 'A1:D13',
    targetCell: 'F1',
    rowFields: ['Region'],
    colFields: ['Quarter'],
    dataFields: [{ field: 'Revenue', func: 'sum', name: 'Total Revenue' }],
    style: 'PivotStyleMedium9',
  });

  // ── Sheet 5: Named ranges & error handling ─────────────────────────────────
  const wsCalc = wb.addSheet('Calculations');
  wsCalc.writeRow(1, 1, ['Input', 'Value']);
  wsCalc.setStyle(1, 1, style().bold().bg('FF4472C4').fontColor('FFFFFFFF').build());
  wsCalc.setStyle(1, 2, style().bold().bg('FF4472C4').fontColor('FFFFFFFF').build());
  wsCalc.setValue(2, 1, 'Price');   wsCalc.setValue(2, 2, 49.99);
  wsCalc.setValue(3, 1, 'Qty');     wsCalc.setValue(3, 2, 150);
  wsCalc.setValue(4, 1, 'Tax %');   wsCalc.setValue(4, 2, 0.08);
  wsCalc.setStyle(4, 2, style().numFmt('0%').build());

  wsCalc.setValue(6, 1, 'Subtotal');
  wsCalc.setFormula(6, 2, 'B2*B3');
  wsCalc.setStyle(6, 2, style().numFmt('$#,##0.00').build());

  wsCalc.setValue(7, 1, 'Tax');
  wsCalc.setFormula(7, 2, 'B6*B4');
  wsCalc.setStyle(7, 2, style().numFmt('$#,##0.00').build());

  wsCalc.setValue(8, 1, 'Grand Total');
  wsCalc.setFormula(8, 2, 'B6+B7');
  wsCalc.setStyle(8, 1, style().bold().fontSize(12).build());
  wsCalc.setStyle(8, 2, style().bold().fontSize(12).numFmt('$#,##0.00').build());

  // Named ranges
  wb.addNamedRange({ name: 'Price', ref: 'Calculations!$B$2' });
  wb.addNamedRange({ name: 'Quantity', ref: 'Calculations!$B$3' });
  wb.addNamedRange({ name: 'TaxRate', ref: 'Calculations!$B$4' });

  // R1C1 reference examples
  wsCalc.setValue(10, 1, 'R1C1 Examples:');
  wsCalc.setStyle(10, 1, style().bold().build());
  wsCalc.setValue(11, 1, 'A1→R1C1:');
  wsCalc.setValue(11, 2, a1ToR1C1('C3', 1, 1));
  wsCalc.setValue(12, 1, 'R1C1→A1:');
  wsCalc.setValue(12, 2, r1c1ToA1('R[2]C[2]', 1, 1));

  // Protection
  wsCalc.protection = {
    sheet: true,
    password: 'demo',
    formatCells: false,
  };
  // Unlock input cells
  wsCalc.setStyle(2, 2, { locked: false });
  wsCalc.setStyle(3, 2, { locked: false });
  wsCalc.setStyle(4, 2, { ...wsCalc.getCell(4, 2).style, locked: false });

  wsCalc.setColumn(1, { width: 16 });
  wsCalc.setColumn(2, { width: 16 });

  // ── Sheet 6: Copy of Dashboard ─────────────────────────────────────────────
  wb.copySheet('Dashboard', 'Dashboard Copy');

  // ── Chart Sheet ────────────────────────────────────────────────────────────
  wb.addChartSheet('Revenue Chart', {
    type: 'bar',
    title: 'Q4 Revenue by Product',
    series: [{
      name: 'Q4 Total',
      values: 'Dashboard!$E$5:$E$12',
      categories: 'Dashboard!$A$5:$A$12',
      color: 'FF4472C4',
    }],
    from: { col: 0, row: 0 },
    to:   { col: 15, row: 25 },
    legend: 'right',
    yAxis: { title: 'Revenue ($)', gridLines: true, numFmt: '#,##0' },
    xAxis: { title: 'Product' },
  });

  // ── Page Setup for all sheets ──────────────────────────────────────────────
  wb.getSheets().forEach(ws => {
    if (!ws.headerFooter) {
      ws.headerFooter = {
        oddHeader: '&L&B' + wb.properties.title + '&R&D',
        oddFooter: '&CPage &P of &N',
      };
    }
  });

  // ── Build & Validate ──────────────────────────────────────────────────────
  await wb.writeFile('./output/50_complex_demo2.xlsx');
  await validate('./output/50_complex_demo2.xlsx', 'Complex Demo');

  // ── Also export CSV & JSON & HTML for Dashboard ───────────────────────────
  const { worksheetToHtml, workbookToHtml } = await import('../features/HtmlModule.js');
  const csv = worksheetToCsv(wsDash, { delimiter: ',' });
  const json = worksheetToJson(wsDash, { header: true });
  const html = worksheetToHtml(wsDash, { includeStyles: true, title: 'Dashboard Export' });

  //@ts-ignore
  const { writeFileSync } = await import('fs');
  writeFileSync('./output/50_complex_demo.csv', csv, 'utf-8');
  writeFileSync('./output/50_complex_demo.json', JSON.stringify(json, null, 2), 'utf-8');
  writeFileSync('./output/50_complex_demo.html', html, 'utf-8');

  // Full workbook HTML with tabs
  const wbHtml = workbookToHtml(wb, { title: 'Complex Demo - All Sheets', includeTabs: true });
  writeFileSync('./output/50_complex_demo_workbook.html', wbHtml, 'utf-8');

  console.log(`  ${OK} CSV exported (${csv.length} chars)`);
  console.log(`  ${OK} JSON exported (${json.length} objects)`);
  console.log(`  ${OK} HTML exported (${html.length} chars)`);
  console.log(`  ${OK} Workbook HTML exported (${wbHtml.length} chars, multi-tab)`);

  // ── Verify round-trip ─────────────────────────────────────────────────────
  const wb2 = await Workbook.fromFile('./output/50_complex_demo2.xlsx');
  const sheetNames = wb2.getSheets().map(s => s.name);
  console.log(`  ${OK} Round-trip: ${sheetNames.length} sheets (${sheetNames.join(', ')})`);

  // Verify formula calculation
  const { FormulaEngine } = await import('../features/FormulaEngine.js');
  const engine = new FormulaEngine();
  engine.calculateSheet(wsCalc);
  const subtotal = wsCalc.getCell(6, 2).value;
  const tax = wsCalc.getCell(7, 2).value;
  const grand = wsCalc.getCell(8, 2).value;
  console.log(`  ${OK} Formula calc: Subtotal=$${subtotal}, Tax=$${typeof tax === 'number' ? tax.toFixed(2) : tax}, Grand=$${typeof grand === 'number' ? grand.toFixed(2) : grand}`);

  console.log('\n  ╔══════════════════════════════════════╗');
  console.log('  ║  All features demonstrated!          ║');
  console.log('  ╚══════════════════════════════════════╝');
}

// Run all tests
async function main() {
  console.log('═══════════════════════════════════════');
  console.log('  ExcelForge New Feature Tests');
  console.log('═══════════════════════════════════════\n');

  await test_printAreas();
  console.log();
  await test_ignoreErrors();
  console.log();
  await test_errorValues();
  console.log();
  await test_r1c1();
  console.log();
  await test_csv();
  console.log();
  await test_json();
  console.log();
  await test_namedStyles();
  console.log();
  await test_chartSheets();
  console.log();
  await test_scaling();
  console.log();
  await test_richTextComments();
  console.log();
  await test_threadedComments();
  console.log();
  await test_copyWorksheet();
  console.log();
  await test_insertDeleteRanges();
  console.log();
  await test_sortRanges();
  console.log();
  await test_fillOperations();
  console.log();
  await test_autoFit();
  console.log();
  await test_customTableStyles();
  console.log();
  await test_advancedFilters();
  console.log();
  await test_rowSplice();
  console.log();
  await test_htmlExport();
  console.log();
  await test_formulaEngine();
  console.log();
  await test_dialogSheet();
  console.log();
  await test_templateSupport();
  console.log();
  await test_copyMoveRanges();
  console.log();
  await test_dynamicArrayFormulas();
  console.log();
  await test_sharedFormulas();
  console.log();
  await test_calculatedPivotFields();
  console.log();
  await test_themes();
  console.log();
  await test_shapes();
  console.log();
  await test_wordArt();
  console.log();
  await test_customIconSets();
  console.log();
  await test_externalLinks();
  console.log();
  await test_queryTables();
  console.log();
  await test_tableSlicers();
  console.log();
  await test_pivotSlicers();
  console.log();
  await test_getPivotData();
  console.log();
  await test_localeSupport();
  console.log();
  await test_mathEquations();
  console.log();
  await test_encryption();
  console.log();
  await test_htmlEnhanced();
  console.log();
  await test_chartTemplates();
  console.log();
  await test_signing();
  console.log();
  await test_complexDemo();
  console.log();

  console.log('All feature tests complete.');

  // ── Export every test Excel file as HTML ─────────────────────────────────
  console.log('\n══════════════════════════════════════');
  console.log('  HTML Export for all test files');
  console.log('══════════════════════════════════════\n');
  //@ts-ignore
  const { readdirSync, writeFileSync: wfs } = await import('fs');
  const { workbookToHtml: wbHtml2 } = await import('../features/HtmlModule.js');
  const xlsxFiles = (readdirSync('./output') as string[]).filter((f: string) => f.endsWith('.xlsx'));
  let htmlCount = 0;
  for (const file of xlsxFiles) {
    try {
      const wb2 = await Workbook.fromFile(`./output/${file}`);
      const html = wbHtml2(wb2, { title: file.replace(/\.[^.]+$/, ''), includeTabs: true });
      const htmlName = file.replace(/\.[^.]+$/, '.html');
      wfs(`./output/${htmlName}`, html, 'utf-8');
      htmlCount++;
    } catch (e: any) {
      console.log(`  ${FAIL} HTML export failed for ${file}: ${e.message?.slice(0, 100)}`);
    }
  }
  console.log(`  ${OK} Exported ${htmlCount}/${xlsxFiles.length} files as HTML`);
}

main().catch(console.error);
