import type { Table } from '../core/types.js';
import { escapeXml, parseRange, colIndexToLetter } from '../utils/helpers.js';

export function buildTableXml(table: Table, tableId: number): string {
  const { startRow, startCol, endRow, endCol } = parseRange(table.ref);
  const displayName = table.displayName ?? table.name;

  const colsXml = table.columns.map((col, i) => {
    const colIdx = startCol + i;
    const id = i + 1;
    const totFn = col.totalsRowFunction ?? 'none';
    const totAttrs = table.totalsRow && totFn !== 'none'
      ? ` totalsRowFunction="${totFn}"${col.totalsRowFormula ? ` totalsRowFormula="${escapeXml(col.totalsRowFormula)}"` : ''}`
      : '';
    const totLabel = table.totalsRow && col.totalsRowLabel
      ? ` totalsRowLabel="${escapeXml(col.totalsRowLabel)}"` : '';
    return `<tableColumn id="${id}" name="${escapeXml(col.name)}"${totAttrs}${totLabel}/>`;
  }).join('');

  const styleAttrs = [
    `name="${escapeXml(table.style ?? 'TableStyleMedium2')}"`,
    table.showFirstColumn   ? 'showFirstColumn="1"'   : '',
    table.showLastColumn    ? 'showLastColumn="1"'    : '',
    table.showRowStripes !== false ? 'showRowStripes="1"' : '',
    table.showColumnStripes ? 'showColumnStripes="1"' : '',
  ].filter(Boolean).join(' ');

  const totalsRow = table.totalsRow ? 1 : 0;

  // autoFilter must exclude the totals row
  const startColLetter = colIndexToLetter(startCol);
  const endColLetter   = colIndexToLetter(endCol);
  const autoFilterRef  = `${startColLetter}${startRow}:${endColLetter}${endRow - totalsRow}`;

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  id="${tableId}" name="${escapeXml(table.name)}" displayName="${escapeXml(displayName)}"
  ref="${table.ref}" totalsRowCount="${totalsRow}">
  <autoFilter ref="${autoFilterRef}"/>
  <tableColumns count="${table.columns.length}">${colsXml}</tableColumns>
  <tableStyleInfo ${styleAttrs}/>
</table>`;
}
