/**
 * ExcelForge — Round-trip tests for real-world Excel files.
 *
 * For each test file:
 *   1. Read the file with ExcelForge
 *   2. Inspect parsed features (sheets, CF, DV, tables, breaks, etc.)
 *   3. Clean round-trip (no modifications) — validate with OpenXML + EPPlus
 *   4. Dirty round-trip (modify first sheet) — validate with OpenXML + EPPlus
 *   5. Verify feature counts survive the dirty round-trip
 */

import { Workbook } from '../index.js';

// @ts-ignore
const { readFile, writeFile, mkdir, rm } = await import('fs/promises');
// @ts-ignore
const { execSync } = await import('child_process');
// @ts-ignore
const { join } = await import('path');

// ── Helpers ──────────────────────────────────────────────────────────────────

const TMP_DIR = './output';

function countOpenXmlErrors(file: string): number {
  try {
    const json = execSync(`dotnet run validator.cs -- "${file}"`, {
      encoding: 'utf8',
      timeout: 60000,
    });
    return JSON.parse(json).length;
  } catch (e: any) {
    // validator exits with code 1 when there are errors — parse stdout
    const stdout = e.stdout ?? '';
    try { return JSON.parse(stdout).length; } catch { return -1; }
  }
}

function epplusOk(file: string): boolean {
  try {
    const out = execSync(`dotnet run validatorEpplus.cs -- "${file}"`, {
      encoding: 'utf8',
      timeout: 30000,
    }).trim();
    return out.includes('success');
  } catch { return false; }
}

interface SheetFeatures {
  cells: number;
  cfs: number;
  dvs: number;
  rowBreaks: number;
  colBreaks: number;
  tables: number;
}

function inspectWorkbook(wb: InstanceType<typeof Workbook>): Map<string, SheetFeatures> {
  const map = new Map<string, SheetFeatures>();
  for (const name of wb.getSheetNames()) {
    const ws = wb.getSheet(name)!;
    // Count cells in a sample area
    let cells = 0;
    for (let r = 1; r <= 100; r++) {
      for (let c = 1; c <= 30; c++) {
        if (ws.getCell(r, c).value !== undefined) cells++;
      }
    }
    map.set(name, {
      cells,
      cfs: ws.getConditionalFormats().length,
      dvs: ws.getDataValidations().size,
      rowBreaks: ws.getRowBreaks().length,
      colBreaks: ws.getColBreaks().length,
      tables: ws.getTables().length,
    });
  }
  return map;
}

// ── Test runner ──────────────────────────────────────────────────────────────

interface TestFile {
  path: string;
  name: string;
  /** Expected OpenXML errors in the ORIGINAL file (pre-existing) */
  expectedOriginalErrors: number;
}

const TEST_FILES: TestFile[] = [
  { path: 'src/test/Book 1.xlsx',                                    name: 'Book 1',           expectedOriginalErrors: 0 },
  { path: 'src/test/Book 2.xlsx',                                    name: 'Book 2',           expectedOriginalErrors: 0 },
  { path: 'src/test/Book 3.xlsx',                                    name: 'Book 3',           expectedOriginalErrors: 0 },
  { path: 'src/test/IC-Gantt-Chart-Template.xlsx',                   name: 'Gantt Chart',      expectedOriginalErrors: 0 },
  { path: 'src/test/IC-Invoice-Tracking-Template.xlsx',              name: 'Invoice Tracking', expectedOriginalErrors: 0 },
  { path: 'src/test/IC-School-Assignments-Tracking-Template.xlsx',   name: 'School Tracking',  expectedOriginalErrors: 6 },
];

async function runRoundtripTest(tf: TestFile): Promise<void> {
  const data = await readFile(tf.path);
  const wb = await Workbook.fromBytes(new Uint8Array(data));
  const sheetNames = wb.getSheetNames();
  const features = inspectWorkbook(wb);

  // Print what we found
  const featureSummary: string[] = [];
  for (const [name, f] of features) {
    const parts: string[] = [];
    if (f.cfs)       parts.push(`CF=${f.cfs}`);
    if (f.dvs)       parts.push(`DV=${f.dvs}`);
    if (f.tables)    parts.push(`Tables=${f.tables}`);
    if (f.rowBreaks) parts.push(`RowBrk=${f.rowBreaks}`);
    if (f.colBreaks) parts.push(`ColBrk=${f.colBreaks}`);
    if (parts.length) featureSummary.push(`${name}(${parts.join(',')})`);
  }
  console.log(`  Sheets: ${sheetNames.length}, Features: ${featureSummary.join('; ') || 'none'}`);

  // ── Clean round-trip ──────────────────────────────────────────────────────
  const cleanBytes = await wb.build();
  const cleanPath = join(TMP_DIR, `rt_clean_${tf.name}.xlsx`);
  await writeFile(cleanPath, cleanBytes);

  const cleanErrors = countOpenXmlErrors(cleanPath);
  if (cleanErrors > tf.expectedOriginalErrors) {
    throw new Error(`Clean round-trip introduced errors: ${cleanErrors} (original had ${tf.expectedOriginalErrors})`);
  }
  if (!epplusOk(cleanPath)) {
    throw new Error('Clean round-trip: EPPlus failed to open');
  }

  // ── Dirty round-trip (modify first sheet → forces all sheets to re-serialise) ─
  wb.markDirty(sheetNames[0]);
  wb.getSheet(sheetNames[0])!.setValue(999, 1, '__roundtrip_test__');

  const dirtyBytes = await wb.build();
  const dirtyPath = join(TMP_DIR, `rt_dirty_${tf.name}.xlsx`);
  await writeFile(dirtyPath, dirtyBytes);

  const dirtyErrors = countOpenXmlErrors(dirtyPath);
  if (dirtyErrors > tf.expectedOriginalErrors) {
    throw new Error(`Dirty round-trip introduced errors: ${dirtyErrors} (original had ${tf.expectedOriginalErrors})`);
  }
  if (!epplusOk(dirtyPath)) {
    throw new Error('Dirty round-trip: EPPlus failed to open');
  }

  // ── Verify features survived the dirty round-trip ─────────────────────────
  const wb2 = await Workbook.fromBytes(dirtyBytes);
  const features2 = inspectWorkbook(wb2);

  for (const [name, orig] of features) {
    const rt = features2.get(name);
    if (!rt) throw new Error(`Sheet "${name}" missing after round-trip`);
    if (rt.cfs !== orig.cfs) throw new Error(`"${name}" CF: ${orig.cfs} → ${rt.cfs}`);
    if (rt.dvs !== orig.dvs) throw new Error(`"${name}" DV: ${orig.dvs} → ${rt.dvs}`);
    if (rt.tables !== orig.tables) throw new Error(`"${name}" Tables: ${orig.tables} → ${rt.tables}`);
    if (rt.rowBreaks !== orig.rowBreaks) throw new Error(`"${name}" RowBreaks: ${orig.rowBreaks} → ${rt.rowBreaks}`);
    if (rt.colBreaks !== orig.colBreaks) throw new Error(`"${name}" ColBreaks: ${orig.colBreaks} → ${rt.colBreaks}`);
  }

  // Verify the test value survived
  const testCell = wb2.getSheet(sheetNames[0])!.getCell(999, 1);
  if (testCell.value !== '__roundtrip_test__') {
    throw new Error(`Test cell value lost: ${testCell.value}`);
  }

  console.log(`  Clean: ${cleanErrors} errors, Dirty: ${dirtyErrors} errors, EPPlus: OK, Features: preserved`);
}

// ── Main ─────────────────────────────────────────────────────────────────────

async function main() {
  try { await mkdir(TMP_DIR, { recursive: true }); } catch {}

  let pass = 0;
  let fail = 0;
  for (const tf of TEST_FILES) {
    try {
      await runRoundtripTest(tf);
      console.log(`✅ ${tf.name}`);
      pass++;
    } catch (e: any) {
      console.error(`❌ ${tf.name}: ${e.message}`);
      fail++;
    }
  }


  console.log(`\n${pass} passed, ${fail} failed out of ${TEST_FILES.length} files`);
  // @ts-ignore
  if (fail > 0) process.exit(1);
}

main();
