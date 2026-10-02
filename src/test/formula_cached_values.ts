import { Workbook, CellError } from '../index.js';
import { buildZip } from '../utils/zip.js';

function check(condition: boolean, message: string): void {
  if (!condition) throw new Error(`Formula cached values: ${message}`);
}

// Use worksheet XML directly so the fixture does not depend on formula evaluation
// or on the writer's handling of cached results.
const parts: Record<string, string> = {
  'xl/workbook.xml': '<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Cached" sheetId="1" r:id="rId1"/></sheets></workbook>',
  'xl/_rels/workbook.xml.rels': '<Relationships><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>',
  'xl/worksheets/sheet1.xml': `<worksheet><sheetData><row r="2">
    <c r="E2"><f ca="1">RANDBETWEEN(250, 500)</f><v>422</v></c>
    <c r="F2"><f>1-1</f><v>0</v></c>
    <c r="G2" t="str"><f>"hello"</f><v>hello</v></c>
    <c r="H2" t="b"><f>1=1</f><v>1</v></c>
    <c r="I2" t="b"><f>1=2</f><v>0</v></c>
    <c r="J2" t="e"><f>1/0</f><v>#DIV/0!</v></c>
    <c r="K2"><f t="array" ref="K2:K3">ROW(K2:K3)</f><v>2</v></c>
    <c r="L2"><f>SUM(A1:A2)</f></c>
    <c r="M2" t="str"><f>""</f><v/></c>
    <c r="N2"><v>123</v></c>
  </row></sheetData></worksheet>`,
};
const encoder = new TextEncoder();
const bytes = buildZip(Object.entries(parts).map(([name, xml]) => ({
  name, data: encoder.encode(xml),
})));
const wb = await Workbook.fromBytes(bytes);
const ws = wb.getSheet('Cached')!;
const expected: Array<[number, string, unknown]> = [
  [5, 'RANDBETWEEN(250, 500)', 422],
  [6, '1-1', 0],
  [7, '"hello"', 'hello'],
  [8, '1=1', true],
  [9, '1=2', false],
  [12, 'SUM(A1:A2)', undefined],
  [13, '""', ''],
];
for (const [col, formula, value] of expected) {
  const cell = ws.getCell(2, col);
  check(cell.formula === formula, `column ${col} lost its formula`);
  check(cell.value === value, `column ${col}: expected ${String(value)}, got ${String(cell.value)}`);
}
const errorCell = ws.getCell(2, 10);
check(errorCell.formula === '1/0', 'error cell lost its formula');
check(errorCell.value instanceof CellError && errorCell.value.error === '#DIV/0!', 'cached error was not parsed');
const arrayCell = ws.getCell(2, 11);
check(arrayCell.arrayFormula === 'ROW(K2:K3)' && arrayCell.value === 2, 'array formula lost its cached value');
check(ws.getCell(2, 14).value === 123, 'ordinary numeric cell changed');
console.log('Formula cached value regression tests passed.');
