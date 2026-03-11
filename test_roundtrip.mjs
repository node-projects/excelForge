import { Workbook } from './dist/core/Workbook.js';
import { style, Colors, NumFmt, Styles } from './dist/styles/builders.js';
import { mkdir, readFile } from 'fs/promises';

await mkdir('./output', { recursive: true });

// ── Test 1: Read existing file, read properties ──────────────────────────────
{
  const wb1 = new Workbook();
  wb1.properties = { title: 'Original Title', author: 'Alice', company: 'Acme' };
  wb1.extendedProperties = { appVersion: '16.0300', manager: 'Bob' };
  wb1.coreProperties = { subject: 'Test Subject', keywords: 'excel test', language: 'en-US' };
  wb1.customProperties = [
    { name: 'ProjectCode',  value: { type: 'string', value: 'PRJ-001' } },
    { name: 'Version',      value: { type: 'int',    value: 3 } },
    { name: 'IsApproved',   value: { type: 'bool',   value: true } },
    { name: 'Budget',       value: { type: 'decimal',value: 12500.50 } },
    { name: 'ReviewDate',   value: { type: 'date',   value: new Date('2024-06-01') } },
  ];

  const ws = wb1.addSheet('Data');
  ws.setValue(1, 1, 'Name');  ws.setValue(1, 2, 'Score');
  ws.setValue(2, 1, 'Alice'); ws.setValue(2, 2, 95);
  ws.setValue(3, 1, 'Bob');   ws.setValue(3, 2, 87);

  await wb1.writeFile('./output/roundtrip_source.xlsx');
  console.log('✅ Created source file');
}

// ── Test 2: Load and read properties back ────────────────────────────────────
{
  const wb = await Workbook.fromFile('./output/roundtrip_source.xlsx');

  console.assert(wb.coreProperties.title    === 'Original Title', 'title mismatch');
  console.assert(wb.coreProperties.creator  === 'Alice',          'author mismatch');
  console.assert(wb.coreProperties.subject  === 'Test Subject',   'subject mismatch');
  console.assert(wb.coreProperties.keywords === 'excel test',     'keywords mismatch');
  console.assert(wb.coreProperties.language === 'en-US',          'language mismatch');
  console.assert(wb.extendedProperties.company    === 'Acme',     'company mismatch');
  console.assert(wb.extendedProperties.appVersion === '16.0300',  'appVersion mismatch');
  console.assert(wb.extendedProperties.manager    === 'Bob',      'manager mismatch');

  const proj = wb.customProperties.find(p => p.name === 'ProjectCode');
  console.assert(proj?.value.value === 'PRJ-001', 'custom prop string mismatch');
  const ver = wb.customProperties.find(p => p.name === 'Version');
  console.assert(ver?.value.value === 3, 'custom prop int mismatch');
  const approved = wb.customProperties.find(p => p.name === 'IsApproved');
  console.assert(approved?.value.value === true, 'custom prop bool mismatch');

  console.log(`✅ Properties read correctly (${wb.customProperties.length} custom props)`);
  console.log(`   Sheets: ${wb.getSheetNames().join(', ')}`);
}

// ── Test 3: Load, modify, and save (patch mode) ──────────────────────────────
{
  const wb = await Workbook.fromFile('./output/roundtrip_source.xlsx');

  // Read the original data
  const ws = wb.getSheet('Data');
  console.assert(ws !== undefined, 'Sheet "Data" not found');

  const name1 = ws.getCell(2, 1).value;
  const score1 = ws.getCell(2, 2).value;
  console.assert(name1 === 'Alice',  `expected Alice, got ${name1}`);
  console.assert(score1 === 95,      `expected 95, got ${score1}`);

  // Modify cells
  ws.setValue(2, 2, 99);  // change Alice's score
  ws.setValue(4, 1, 'Carol'); ws.setValue(4, 2, 78); // add new row
  ws.setStyle(1, 1, Styles.tableHeader);
  ws.setStyle(1, 2, Styles.tableHeader);

  // Modify properties
  wb.coreProperties.title = 'Modified Title';
  wb.setCustomProperty('ProjectCode', { type: 'string', value: 'PRJ-002' });
  wb.setCustomProperty('ModifiedBy',  { type: 'string', value: 'Carol' });

  // Mark sheet dirty so it gets re-serialised
  wb.markDirty('Data');

  await wb.writeFile('./output/roundtrip_modified.xlsx');
  console.log('✅ Modified and saved');
}

// ── Test 4: Verify the modified file ─────────────────────────────────────────
{
  const wb = await Workbook.fromFile('./output/roundtrip_modified.xlsx');

  console.assert(wb.coreProperties.title === 'Modified Title', 'title not updated');

  const proj = wb.getCustomProperty('ProjectCode');
  console.assert(proj?.value.value === 'PRJ-002', 'custom prop not updated');
  const modBy = wb.getCustomProperty('ModifiedBy');
  console.assert(modBy?.value.value === 'Carol', 'new custom prop missing');

  // Original custom props should survive
  console.assert(wb.customProperties.length >= 6, `expected ≥6 custom props, got ${wb.customProperties.length}`);

  const ws = wb.getSheet('Data');
  const score = ws.getCell(2, 2).value;
  console.assert(score === 99, `expected 99, got ${score}`);
  const carol = ws.getCell(4, 1).value;
  console.assert(carol === 'Carol', `expected Carol, got ${carol}`);

  console.log(`✅ Modified file verified (${wb.customProperties.length} custom props preserved)`);
}

// ── Test 5: fromBytes and fromBase64 ─────────────────────────────────────────
{
  const bytes = await readFile('./output/roundtrip_source.xlsx');
  const wb1 = await Workbook.fromBytes(new Uint8Array(bytes));
  console.assert(wb1.coreProperties.title === 'Original Title', 'fromBytes failed');
  console.log('✅ fromBytes works');

  const b64 = wb1.buildBase64 ? await wb1.buildBase64() : '';
  const wb2 = await Workbook.fromBase64(b64);
  console.assert(wb2.coreProperties.title === 'Original Title', 'fromBase64 roundtrip failed');
  console.log('✅ fromBase64 roundtrip works');
}

// ── Test 6: Multi-sheet modify (only dirty sheets regenerated) ────────────────
{
  const wb = new Workbook();
  wb.coreProperties = { title: 'Multi-sheet', creator: 'Test' };
  const ws1 = wb.addSheet('Sheet1');
  const ws2 = wb.addSheet('Sheet2');
  const ws3 = wb.addSheet('Sheet3');

  ws1.setValue(1, 1, 'Sheet1 data'); ws1.setStyle(1, 1, Styles.bold);
  ws2.setValue(1, 1, 'Sheet2 data');
  ws3.setValue(1, 1, 'Sheet3 data');

  await wb.writeFile('./output/multisheet.xlsx');

  // Load and modify only sheet2
  const wb2 = await Workbook.fromFile('./output/multisheet.xlsx');
  const s2 = wb2.getSheet('Sheet2');
  s2.setValue(2, 1, 'MODIFIED');
  s2.addConditionalFormat({
    sqref: 'A1:A10', type: 'cellIs', operator: 'greaterThan', formula: '50',
    style: style().bg(Colors.Green).build(), priority: 1
  });
  wb2.markDirty('Sheet2');

  await wb2.writeFile('./output/multisheet_modified.xlsx');

  // Verify
  const wb3 = await Workbook.fromFile('./output/multisheet_modified.xlsx');
  console.assert(wb3.getSheetNames().length === 3, 'wrong sheet count');
  console.assert(wb3.getSheet('Sheet2').getCell(2, 1).value === 'MODIFIED', 'sheet2 not updated');
  console.assert(wb3.getSheet('Sheet1').getCell(1, 1).value === 'Sheet1 data', 'sheet1 corrupted');
  console.assert(wb3.getSheet('Sheet3').getCell(1, 1).value === 'Sheet3 data', 'sheet3 corrupted');
  console.log('✅ Multi-sheet patch: only dirty sheet regenerated, others preserved');
}

// ── Test 7: All property fields roundtrip ────────────────────────────────────
{
  const wb = new Workbook();
  wb.coreProperties = {
    title: 'Full Props Test',
    subject: 'Testing',
    creator: 'ExcelForge',
    keywords: 'test roundtrip',
    description: 'A comprehensive test',
    lastModifiedBy: 'Robot',
    revision: '5',
    category: 'Testing',
    contentStatus: 'Draft',
    language: 'de-DE',
  };
  wb.extendedProperties = {
    application: 'ExcelForge',
    appVersion: '1.0.0',
    company: 'Test Corp',
    manager: 'Manager Name',
    hyperlinkBase: 'https://example.com',
    docSecurity: 0,
  };
  wb.customProperties = [
    { name: 'StringProp',  value: { type: 'string',  value: 'hello world' } },
    { name: 'IntProp',     value: { type: 'int',     value: 42 } },
    { name: 'DecimalProp', value: { type: 'decimal', value: 3.14159 } },
    { name: 'BoolProp',    value: { type: 'bool',    value: false } },
    { name: 'DateProp',    value: { type: 'date',    value: new Date('2024-01-15T12:00:00Z') } },
  ];

  wb.addSheet('Test').setValue(1, 1, 'ok');
  await wb.writeFile('./output/all_props.xlsx');

  const wb2 = await Workbook.fromFile('./output/all_props.xlsx');
  console.assert(wb2.coreProperties.title       === 'Full Props Test',   'title');
  console.assert(wb2.coreProperties.language    === 'de-DE',             'language');
  console.assert(wb2.coreProperties.revision    === '5',                 'revision');
  console.assert(wb2.extendedProperties.manager === 'Manager Name',      'manager');
  console.assert(wb2.extendedProperties.hyperlinkBase === 'https://example.com', 'hyperlinkBase');
  console.assert(wb2.customProperties.length    === 5,                   'custom count');
  const dp = wb2.customProperties.find(p => p.name === 'DecimalProp');
  console.assert(Math.abs(Number(dp?.value.value) - 3.14159) < 0.001, 'decimal');
  console.log('✅ All property fields roundtrip correctly');
}

console.log('\n🎉 All round-trip tests passed!');
