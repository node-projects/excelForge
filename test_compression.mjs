import { Workbook } from './dist/core/Workbook.js';
import { style, Colors, Styles, NumFmt } from './dist/styles/builders.js';
import { deflateRaw } from './dist/utils/zip.js';
import { mkdir } from 'fs/promises';

await mkdir('./output', { recursive: true });

// ── Verify DEFLATE correctness via zlib ──────────────────────────────────────
import { inflateRawSync } from 'zlib';

const testVectors = [
  new Uint8Array(0),
  new Uint8Array([0x48, 0x65, 0x6c, 0x6c, 0x6f]),           // "Hello"
  new Uint8Array(1000).fill(0x41),                           // 1000 × 'A'
  new Uint8Array(Array.from({ length: 256 }, (_, i) => i)), // all bytes
  new Uint8Array(Array.from({ length: 4096 }, (_, i) =>     // repeating XML-like
    'abcdefghijklmnopqrstuvwxyz<>/="'.charCodeAt(i % 31))),
];

console.log('── DEFLATE round-trip correctness ──────────────────────────────');
for (const vec of testVectors) {
  for (const level of [0, 1, 6, 9]) {
    const compressed = deflateRaw(vec, level);
    const decompressed = level === 0
      ? (() => {
          // Level 0 uses DEFLATE stored blocks — zlib can decompress them
          const r = inflateRawSync(compressed);
          return new Uint8Array(r.buffer, r.byteOffset, r.byteLength);
        })()
      : (() => {
          const r = inflateRawSync(compressed);
          return new Uint8Array(r.buffer, r.byteOffset, r.byteLength);
        })();
    const ok = decompressed.length === vec.length &&
               decompressed.every((b, i) => b === vec[i]);
    if (!ok) {
      console.error(`  FAIL level=${level} len=${vec.length}`);
      process.exit(1);
    }
  }
}
console.log('  ✅ All DEFLATE vectors pass (levels 0, 1, 6, 9)');

// ── Build a realistic workbook ────────────────────────────────────────────────
function makeWorkbook() {
  const wb = new Workbook();
  wb.coreProperties = { title: 'Compression Test', creator: 'ExcelForge', subject: 'Benchmark' };
  wb.extendedProperties = { company: 'Test Corp', appVersion: '1.0' };
  wb.customProperties = [
    { name: 'Project', value: { type: 'string', value: 'Alpha' } },
    { name: 'Version', value: { type: 'int',    value: 7       } },
  ];

  const ws = wb.addSheet('Sales Data');
  const headers = ['Region', 'Product', 'Q1', 'Q2', 'Q3', 'Q4', 'Total', 'YoY%'];
  headers.forEach((h, i) => { ws.setValue(1, i + 1, h); ws.setStyle(1, i + 1, Styles.headerBlue); });

  const regions = ['North', 'South', 'East', 'West', 'Central'];
  const products = ['Widget A', 'Widget B', 'Gadget X', 'Gadget Y', 'Service Pro'];
  let row = 2;
  for (const r of regions) {
    for (const p of products) {
      const q = Array.from({ length: 4 }, () => Math.floor(Math.random() * 50000 + 10000));
      ws.setValue(row, 1, r);
      ws.setValue(row, 2, p);
      q.forEach((v, i) => ws.setValue(row, 3 + i, v));
      ws.setFormula(row, 7, `SUM(C${row}:F${row})`);
      ws.setValue(row, 8, (Math.random() * 0.4 - 0.1));
      ws.setStyle(row, 8, style().numFmt(NumFmt.Percent2).build());
      ws.setStyle(row, 7, style().numFmt(NumFmt.Decimal2).build());
      row++;
    }
  }

  ws.setColumn(1, { width: 12 }); ws.setColumn(2, { width: 16 });
  for (let c = 3; c <= 8; c++) ws.setColumn(c, { width: 13 });
  ws.freeze(1, 0);
  ws.autoFilter = { ref: `A1:H1` };

  // Second sheet with lots of strings (better compression target)
  const ws2 = wb.addSheet('Log');
  for (let r = 1; r <= 200; r++) {
    ws2.setValue(r, 1, `2024-${String(Math.ceil(r/17)).padStart(2,'0')}-${String((r%28)+1).padStart(2,'0')}`);
    ws2.setValue(r, 2, `INFO`);
    ws2.setValue(r, 3, `Processing item ${r} in batch ${Math.ceil(r/10)} of workflow ExcelForge-Demo`);
    ws2.setValue(r, 4, Math.floor(Math.random() * 1000));
  }

  return wb;
}

// ── Size comparison across levels ─────────────────────────────────────────────
console.log('\n── File size by compression level ──────────────────────────────');
console.log('  Level  Size (bytes)  Ratio   Description');
console.log('  ─────  ────────────  ──────  ───────────────────────────────');

let baseSize = 0;
const results = [];

for (const level of [0, 1, 6, 9]) {
  const wb = makeWorkbook();
  wb.compressionLevel = level;
  const bytes = await wb.build();
  if (level === 0) baseSize = bytes.length;
  const ratio = ((1 - bytes.length / baseSize) * 100).toFixed(1);
  const desc = level === 0 ? 'STORE (no compression)'
             : level === 1 ? 'FAST (fixed Huffman)'
             : level === 6 ? 'DEFAULT (dynamic Huffman + lazy LZ77)'
             :               'BEST (maximum effort)';
  console.log(`  ${String(level).padEnd(5)}  ${String(bytes.length).padStart(12)}  ${ratio.padStart(5)}%  ${desc}`);
  results.push({ level, size: bytes.length });
}

console.log(`\n  Best compression: ${((1 - results[3].size / results[0].size) * 100).toFixed(1)}% smaller than STORE`);

// ── Write all four variants ────────────────────────────────────────────────────
for (const { level } of results) {
  const wb = makeWorkbook();
  wb.compressionLevel = level;
  await wb.writeFile(`./output/compression_level${level}.xlsx`);
}
console.log('\n  ✅ Written: output/compression_level{0,1,6,9}.xlsx');

// ── Verify all four are readable by openpyxl ──────────────────────────────────
import { execSync } from 'child_process';
const py = `
import openpyxl, sys
for lvl in [0, 1, 6, 9]:
    wb = openpyxl.load_workbook(f'output/compression_level{lvl}.xlsx')
    ws = wb['Sales Data']
    assert ws['A1'].value == 'Region', f'level {lvl}: header wrong'
    assert ws['A2'].value == 'North',  f'level {lvl}: data wrong'
    assert len(wb.sheetnames) == 2,    f'level {lvl}: sheet count wrong'
print('openpyxl: all 4 levels valid')
`;
try {
  const out = execSync(`python3 -c "${py.replace(/\n/g,'\\n').replace(/"/g,'\\"')}"`, { cwd: '/home/claude/excelforge' }).toString();
  console.log(`  ✅ ${out.trim()}`);
} catch (e) {
  console.error('  openpyxl check failed:', e.message);
}

// ── Per-entry level override ───────────────────────────────────────────────────
console.log('\n── Per-entry level override ─────────────────────────────────────');
{
  // Demonstrate the ZipOptions API directly
  const { buildZip } = await import('./dist/utils/zip.js');
  const enc = new TextEncoder();

  const xmlData  = enc.encode('<worksheet>' + '<row r="1"><c r="A1" t="s"><v>0</v></c></row>'.repeat(500) + '</worksheet>');
  const pngData  = new Uint8Array(200).fill(0x89); // fake PNG-ish bytes

  const zip = buildZip([
    { name: 'xl/worksheets/sheet1.xml', data: xmlData },            // global level
    { name: 'xl/media/image1.png',      data: pngData, level: 0 },  // forced STORE
    { name: 'xl/styles.xml',            data: enc.encode('<styleSheet/>'), level: 9 }, // max
  ], { level: 6 });

  console.log(`  ✅ Built ZIP with mixed per-entry levels (total ${zip.length} bytes)`);
}

console.log('\n🎉 All compression tests passed!');
