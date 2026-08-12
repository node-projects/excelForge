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

console.log('Extended HTML export checks passed.');
