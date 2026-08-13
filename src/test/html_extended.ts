import { Workbook, style, workbookToHtml } from '../index.js';

function check(condition: boolean, message: string): void {
  if (!condition) throw new Error(`HTML extended export: ${message}`);
}

const wb = new Workbook();
const ws = wb.addSheet('Warnings');
ws.writeRow(1, 1, ['Time', 'Type', 'Message']);
ws.writeArray(2, 1, [
  ['00:07:01', 'Warning', 'First'],
  ['00:08:02', 'Error', 'Second'],
  ['00:09:03', 'Info', 'Third'],
]);
ws.setStyle(2, 1, style().numFmt('[$-F400]h:mm:ss\\ AM/PM').build());
ws.setStyle(1000, 1, style().bg('FFFFFFFF').build());
ws.setStyle(1, 50, style().bg('FFFFFFFF').build());
ws.addTable({
  name: 'WarningsTable',
  ref: 'A1:C1000',
  style: 'TableStyleMedium2',
  showRowStripes: true,
  columns: [{ name: 'Time' }, { name: 'Type' }, { name: 'Message' }],
});
ws.addConditionalFormat({
  sqref: 'A2:C1000',
  type: 'expression',
  formula: '$B2="Warning"',
  priority: 1,
  style: style().bg('FFFFD966').build(),
});

const html = workbookToHtml(wb, {
  title: 'Interactive HTML test',
  includeTabs: true,
  mode: 'interactive',
});

check(html.includes('class="xl-filter-button"'), 'missing filter buttons');
check(html.includes('.xl-filter-menu'), 'missing filter menu styles');
check(html.includes('Filter values'), 'missing interactive filter script');
check(html.includes('tr.xl-sticky-header { position: sticky'), 'sticky positioning is not applied to the header row');
check(!html.includes('tr.xl-sticky-header td { position: sticky'), 'sticky positioning still displaces individual header cells');
check(!html.includes('--xl-sticky-top: 40px'), 'sticky header is offset over the first data row');
check(html.includes('white-space: nowrap !important'), 'filter headers can wrap');
check(html.includes('.xl-filter-button { position: absolute'), 'filter buttons are not anchored within header cells');
check(!html.includes('.xl-filter-button { float: right'), 'filter buttons can still fall onto a second line');
check(html.includes('background-color:#4472C4'), 'missing Excel table header style');
check(html.includes('background-color:#FFD966'), 'missing expression conditional formatting');
check(html.includes('>00:07:01</td>'), 'time text was reformatted incorrectly');
check((html.match(/data-xl-row=/g) ?? []).length === 4, 'trailing formatted rows were not trimmed');
check(!html.includes('data-xl-row="1000"'), 'pre-sized empty table tail was exported');
check(!html.includes('data-xl-col="50"'), 'trailing formatted columns were not trimmed');

const rotationWb = new Workbook();
const rotationWs = rotationWb.addSheet('Rotation');
rotationWs.setCell(1, 1, {
  value: 'Vertical label',
  style: style().bg('FF95AEC2').align('center', 'center').rotate(90).build(),
});
rotationWs.setCell(2, 1, { value: 'Sized row' });
rotationWs.setRow(2, { height: 30 });

const rotationHtml = workbookToHtml(rotationWb, {
  includeTabs: false,
  includeStyles: true,
});
const rotatedCellTag = rotationHtml.match(/<td[^>]*data-cell="A1"[^>]*>/)?.[0] ?? '';
check(rotatedCellTag.includes('background-color:#95AEC2'), 'rotated cell lost its background fill');
check(!rotatedCellTag.includes('transform:'), 'rotation is still applied to the table cell');
check(rotationHtml.includes('class="xl-cell-content xl-text-rotation-90"'), 'rotated text content wrapper is missing');
check(rotationHtml.includes('writing-mode:vertical-rl'), '90-degree text does not participate in row auto-height');
check(rotationHtml.includes('data-xl-row="2" style="height:40px"'), 'Excel row height points were not converted to CSS pixels');
const basicRotationHtml = workbookToHtml(rotationWb, { includeTabs: false, mode: 'basic' });
check(!basicRotationHtml.includes('xl-text-rotation-90'), 'basic mode unexpectedly rendered cell rotation styling');

const formulaWb = new Workbook();
const formulaSource = formulaWb.addSheet('Inputs - Outputs');
formulaSource.setValue(5, 1, 'Input');
formulaSource.setValue(5, 2, 'Description');
const formulaTarget = formulaWb.addSheet('Outputs - Inputs');
formulaTarget.setCell(1, 5, {
  formula: "='Inputs - Outputs'!A5",
  style: style().bg('FFEEC18F').align('center', 'center').rotate(90).build(),
});
formulaTarget.setCell(2, 5, {
  formula: "='Inputs - Outputs'!B5",
  style: style().bg('FFEEC18F').align('center', 'center').rotate(90).build(),
});
const evaluatedHtml = workbookToHtml(formulaWb, {
  includeTabs: true,
  includeStyles: true,
  evaluateFormulas: true,
});
check(formulaTarget.getCell(1, 5).value === 'Input', 'leading-equals cross-sheet formula was not evaluated');
check(formulaTarget.getCell(2, 5).value === 'Description', 'quoted sheet reference was not evaluated');
check(evaluatedHtml.includes('>Input</span>'), 'evaluated formula text is missing from HTML');
check(evaluatedHtml.includes('>Description</span>'), 'second evaluated formula text is missing from HTML');

console.log('Extended HTML export checks passed.');
