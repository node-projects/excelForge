import type { PivotTable, CellValue } from '../core/types.js';
import { escapeXml, cellRefToIndices, indicesToCellRef } from '../utils/helpers.js';

const FUNC_MAP: Record<string, string> = {
  sum: 'sum', count: 'count', average: 'average', max: 'max', min: 'min',
  product: 'product', countNums: 'countNums', stdDev: 'stdDev',
  stdDevp: 'stdDevp', var: 'var', varp: 'varp',
};

export interface PivotBuildResult {
  pivotTableXml:   string;
  cacheDefXml:     string;
  cacheRecordsXml: string;
}

/**
 * Build all three XML documents for a single pivot table.
 *
 * @param pt          - Pivot table configuration
 * @param sourceData  - 2-D array; row 0 = headers, rows 1+ = data records
 * @param pivotId     - 1-based file index  (pivotTable1.xml, pivotCacheDefinition1.xml, …)
 * @param cacheId     - workbook-scoped numeric cache ID for workbook.xml <pivotCache cacheId="…">
 */
export function buildPivotTableFiles(
  pt: PivotTable,
  sourceData: CellValue[][],
  pivotId: number,
  cacheId: number,
): PivotBuildResult {
  const headers   = (sourceData[0] ?? []).map(v => String(v ?? ''));
  const dataRows  = sourceData.slice(1);
  const numFields = headers.length;

  const rowGT = pt.rowGrandTotals !== false;
  const colGT = pt.colGrandTotals !== false;

  // Map header names → 0-based field indices
  const fieldIdx = new Map<string, number>(headers.map((h, i) => [h, i] as [string, number]));

  const rowFldIdxs  = pt.rowFields.map(f  => fieldIdx.get(f)    ?? 0);
  const colFldIdxs  = pt.colFields.map(f  => fieldIdx.get(f)    ?? 0);
  const dataFldIdxs = pt.dataFields.map(df => fieldIdx.get(df.field) ?? 0);

  // ── Collect unique values per field (in order of first appearance) ────────
  const uniqueVals: string[][]           = Array.from({ length: numFields }, () => []);
  const uniqueMap:  Map<string, number>[] = Array.from({ length: numFields }, () => new Map());
  const isNumeric:  boolean[]            = new Array(numFields).fill(true);

  for (const row of dataRows) {
    for (let fi = 0; fi < numFields; fi++) {
      const v  = row[fi];
      const vs = v === null || v === undefined ? '' : String(v);
      if (typeof v !== 'number') isNumeric[fi] = false;
      if (!uniqueMap[fi].has(vs)) {
        uniqueMap[fi].set(vs, uniqueVals[fi].length);
        uniqueVals[fi].push(vs);
      }
    }
  }

  // ── Compute pivot table bounding box ──────────────────────────────────────
  const { row: tRow, col: tCol } = cellRefToIndices(pt.targetCell);

  const hasColFields = colFldIdxs.length > 0;
  const numColCombos = hasColFields
    ? colFldIdxs.reduce((n, fi) => n * Math.max(uniqueVals[fi].length, 1), 1)
    : 1;
  const numDataRowsPT = rowFldIdxs.length
    ? rowFldIdxs.reduce((n, fi) => n * Math.max(uniqueVals[fi].length, 1), 1)
    : 1;

  const totalRows = 1 /* header */ + numDataRowsPT + (rowGT ? 1 : 0);
  const totalCols = rowFldIdxs.length
    + numColCombos * pt.dataFields.length
    + (colGT ? pt.dataFields.length : 0);

  const locationRef  = `${indicesToCellRef(tRow, tCol)}:${indicesToCellRef(tRow + totalRows - 1, tCol + totalCols - 1)}`;
  const firstDataCol = rowFldIdxs.length + 1; // 1-based column within bounding box where data starts

  // ── Axis field set ────────────────────────────────────────────────────────
  const isAxisField = new Set([...rowFldIdxs, ...colFldIdxs]);

  // ── <cacheFields> ────────────────────────────────────────────────────────
  const cacheFieldsXml = headers.map((name, fi) => {
    if (isAxisField.has(fi)) {
      const items = uniqueVals[fi].map(v => `<s v="${escapeXml(v)}"/>`).join('');
      return `<cacheField name="${escapeXml(name)}" numFmtId="0"><sharedItems count="${uniqueVals[fi].length}">${items}</sharedItems></cacheField>`;
    }
    if (isNumeric[fi]) {
      const nums = dataRows.map(r => Number(r[fi])).filter(n => !isNaN(n));
      const min  = nums.length ? Math.min(...nums) : 0;
      const max  = nums.length ? Math.max(...nums) : 0;
      return `<cacheField name="${escapeXml(name)}" numFmtId="0"><sharedItems containsSemiMixedTypes="0" containsString="0" containsNumber="1" minValue="${min}" maxValue="${max}"/></cacheField>`;
    }
    return `<cacheField name="${escapeXml(name)}" numFmtId="0"><sharedItems/></cacheField>`;
  }).join('');

  // ── <pivotCacheRecords> ───────────────────────────────────────────────────
  const recordsXml = dataRows.map(row => {
    const cells = headers.map((_, fi) => {
      const v = row[fi];
      if (isAxisField.has(fi)) {
        const vs  = v === null || v === undefined ? '' : String(v);
        return `<x v="${uniqueMap[fi].get(vs) ?? 0}"/>`;
      }
      if (typeof v === 'number') return `<n v="${v}"/>`;
      if (typeof v === 'boolean') return `<b v="${v ? 1 : 0}"/>`;
      return `<s v="${escapeXml(String(v ?? ''))}"/>`;
    });
    return `<r>${cells.join('')}</r>`;
  }).join('');

  // ── <pivotFields> ────────────────────────────────────────────────────────
  const pivotFieldsXml = headers.map((_, fi) => {
    const isRow  = rowFldIdxs.includes(fi);
    const isCol  = colFldIdxs.includes(fi);
    const isData = dataFldIdxs.includes(fi);

    if (isRow) {
      const items = uniqueVals[fi].map((_, vi) => `<item x="${vi}"/>`).join('') + '<item t="default"/>';
      return `<pivotField axis="axisRow" showAll="0"><items count="${uniqueVals[fi].length + 1}">${items}</items></pivotField>`;
    }
    if (isCol) {
      const items = uniqueVals[fi].map((_, vi) => `<item x="${vi}"/>`).join('') + '<item t="default"/>';
      return `<pivotField axis="axisCol" showAll="0"><items count="${uniqueVals[fi].length + 1}">${items}</items></pivotField>`;
    }
    if (isData) return `<pivotField dataField="1" showAll="0"/>`;
    return `<pivotField showAll="0"/>`;
  }).join('');

  // ── <rowFields> / <rowItems> ─────────────────────────────────────────────
  let rowFieldsXml = '';
  let rowItemsXml  = '';
  if (rowFldIdxs.length) {
    rowFieldsXml = `<rowFields count="${rowFldIdxs.length}">${rowFldIdxs.map(fi => `<field x="${fi}"/>`).join('')}</rowFields>`;
    const fi    = rowFldIdxs[0];
    const items = uniqueVals[fi].map((_, vi) => `<i><x v="${vi}"/></i>`).join('');
    const grand = rowGT ? '<i t="grand"><x/></i>' : '';
    rowItemsXml = `<rowItems count="${uniqueVals[fi].length + (rowGT ? 1 : 0)}">${items}${grand}</rowItems>`;
  }

  // ── <colFields> / <colItems> ─────────────────────────────────────────────
  let colFieldsXml = '';
  let colItemsXml  = '';
  if (hasColFields) {
    colFieldsXml = `<colFields count="${colFldIdxs.length}">${colFldIdxs.map(fi => `<field x="${fi}"/>`).join('')}</colFields>`;
    const fi    = colFldIdxs[0];
    const items = uniqueVals[fi].map((_, vi) => `<i><x v="${vi}"/></i>`).join('');
    const grand = colGT ? '<i t="grand"><x/></i>' : '';
    colItemsXml = `<colItems count="${uniqueVals[fi].length + (colGT ? 1 : 0)}">${items}${grand}</colItems>`;
  }

  // ── <dataFields> ─────────────────────────────────────────────────────────
  const dataFieldsXml = `<dataFields count="${pt.dataFields.length}">${
    pt.dataFields.map((df, i) => {
      const fi   = dataFldIdxs[i];
      const func = FUNC_MAP[df.func ?? 'sum'] ?? 'sum';
      const name = df.name ?? `Sum of ${df.field}`;
      return `<dataField name="${escapeXml(name)}" fld="${fi}" subtotal="${func}" showDataAs="normal" baseField="0" baseItem="0"/>`;
    }).join('')
  }</dataFields>`;

  // ── Grand-total attrs ────────────────────────────────────────────────────
  const gtParts = [rowGT ? '' : 'rowGrandTotals="0"', colGT ? '' : 'colGrandTotals="0"'].filter(Boolean);
  const gtAttrStr = gtParts.length ? ' ' + gtParts.join(' ') : '';

  const style = escapeXml(pt.style ?? 'PivotStyleMedium9');

  // ── Assemble XMLs ─────────────────────────────────────────────────────────
  const pivotTableXml =
`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" name="${escapeXml(pt.name)}" cacheId="${cacheId}" applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0" applyPatternFormats="0" applyAlignmentFormats="0" applyWidthHeightFormats="1" dataCaption="Values" updatedVersion="6" minRefreshableVersion="3" useAutoFormatting="1" itemPrintTitles="1" createdVersion="6" indent="2" outline="1" outlineData="1" multipleFieldFilters="0"${gtAttrStr}>
<location ref="${locationRef}" firstHeaderRow="1" firstDataRow="2" firstDataCol="${firstDataCol}"/>
<pivotFields count="${numFields}">${pivotFieldsXml}</pivotFields>
${rowFieldsXml}${rowItemsXml}${colFieldsXml}${colItemsXml}${dataFieldsXml}
<pivotTableStyleInfo name="${style}" showRowHeaders="1" showColHeaders="1" showRowStripes="0" showColStripes="0" showLastColumn="1"/>
</pivotTableDefinition>`;

  const cacheDefXml =
`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1" refreshedBy="ExcelForge" refreshedDate="45000" createdVersion="6" refreshedVersion="6" minRefreshableVersion="3" recordCount="${dataRows.length}" saveData="1">
<cacheSource type="worksheet"><worksheetSource ref="${pt.sourceRef}" sheet="${escapeXml(pt.sourceSheet)}"/></cacheSource>
<cacheFields count="${numFields}">${cacheFieldsXml}</cacheFields>
</pivotCacheDefinition>`;

  const cacheRecordsXml =
`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" count="${dataRows.length}">${recordsXml}</pivotCacheRecords>`;

  return { pivotTableXml, cacheDefXml, cacheRecordsXml };
}
