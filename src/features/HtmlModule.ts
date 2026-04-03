/**
 * ExcelForge — HTML/CSS Export Module (tree-shakeable).
 * Converts worksheets to HTML tables with inline CSS styling.
 */

import type { Worksheet } from '../core/Worksheet.js';
import type { CellStyle, Font, Fill, PatternFill, Border, BorderSide, Alignment } from '../core/types.js';
import { escapeXml } from '../utils/helpers.js';

export interface HtmlExportOptions {
  /** Include <style> block with CSS */
  includeStyles?: boolean;
  /** Full HTML document or just the <table> */
  fullDocument?: boolean;
  /** Document/page title */
  title?: string;
  /** CSS class prefix */
  classPrefix?: string;
}

function colorToCSS(c: string | undefined): string {
  if (!c) return '';
  if (c.startsWith('#')) return c;
  if (c.startsWith('theme:')) return '#000'; // theme colors need a theme, fallback to black
  // AARRGGBB → #RRGGBB
  if (c.length === 8) return '#' + c.slice(2);
  return '#' + c;
}

function fontToCSS(f: Font): string {
  const parts: string[] = [];
  if (f.bold) parts.push('font-weight:bold');
  if (f.italic) parts.push('font-style:italic');
  if (f.underline && f.underline !== 'none') parts.push('text-decoration:underline');
  if (f.strike) parts.push('text-decoration:line-through');
  if (f.size) parts.push(`font-size:${f.size}pt`);
  if (f.color) parts.push(`color:${colorToCSS(f.color)}`);
  if (f.name) parts.push(`font-family:'${f.name}',sans-serif`);
  return parts.join(';');
}

function fillToCSS(fill: Fill): string {
  if (fill.type === 'pattern') {
    const pf = fill as PatternFill;
    if (pf.pattern === 'solid' && pf.fgColor) {
      return `background-color:${colorToCSS(pf.fgColor)}`;
    }
  }
  return '';
}

function borderSideCSS(side: BorderSide | undefined): string {
  if (!side || !side.style) return '';
  const widthMap: Record<string, string> = {
    thin: '1px', medium: '2px', thick: '3px', dashed: '1px', dotted: '1px',
    double: '3px', hair: '1px',
  };
  const styleMap: Record<string, string> = {
    thin: 'solid', medium: 'solid', thick: 'solid', dashed: 'dashed', dotted: 'dotted',
    double: 'double', hair: 'solid',
  };
  const w = widthMap[side.style] ?? '1px';
  const s = styleMap[side.style] ?? 'solid';
  const c = side.color ? colorToCSS(side.color) : '#000';
  return `${w} ${s} ${c}`;
}

function alignmentCSS(a: Alignment): string {
  const parts: string[] = [];
  if (a.horizontal) parts.push(`text-align:${a.horizontal}`);
  if (a.vertical) {
    const vMap: Record<string, string> = { top: 'top', center: 'middle', bottom: 'bottom' };
    parts.push(`vertical-align:${vMap[a.vertical] ?? 'bottom'}`);
  }
  if (a.wrapText) parts.push('white-space:normal;word-wrap:break-word');
  return parts.join(';');
}

function styleToCSS(s: CellStyle): string {
  const parts: string[] = [];
  if (s.font) parts.push(fontToCSS(s.font));
  if (s.fill) parts.push(fillToCSS(s.fill));
  if (s.alignment) parts.push(alignmentCSS(s.alignment));
  if (s.border) {
    if (s.border.top) parts.push(`border-top:${borderSideCSS(s.border.top)}`);
    if (s.border.bottom) parts.push(`border-bottom:${borderSideCSS(s.border.bottom)}`);
    if (s.border.left) parts.push(`border-left:${borderSideCSS(s.border.left)}`);
    if (s.border.right) parts.push(`border-right:${borderSideCSS(s.border.right)}`);
  }
  return parts.filter(Boolean).join(';');
}

/**
 * Convert a worksheet to an HTML table string.
 */
export function worksheetToHtml(ws: Worksheet, options: HtmlExportOptions = {}): string {
  const range = ws.getUsedRange();
  if (!range) return options.fullDocument !== false ? `<!DOCTYPE html><html><head><title>${escapeXml(options.title ?? '')}</title></head><body><p>Empty worksheet</p></body></html>` : '<table></table>';

  const { startRow, startCol, endRow, endCol } = range;
  const merges = ws.getMerges();

  // Build merge map: "row,col" → { rowSpan, colSpan } for top-left, "skip" for covered cells
  const mergeMap = new Map<string, { rowSpan: number; colSpan: number } | 'skip'>();
  for (const m of merges) {
    const rs = m.endRow - m.startRow + 1;
    const cs = m.endCol - m.startCol + 1;
    mergeMap.set(`${m.startRow},${m.startCol}`, { rowSpan: rs, colSpan: cs });
    for (let r = m.startRow; r <= m.endRow; r++) {
      for (let c = m.startCol; c <= m.endCol; c++) {
        if (r !== m.startRow || c !== m.startCol) mergeMap.set(`${r},${c}`, 'skip');
      }
    }
  }

  const rows: string[] = [];
  for (let r = startRow; r <= endRow; r++) {
    const cells: string[] = [];
    for (let c = startCol; c <= endCol; c++) {
      const key = `${r},${c}`;
      const merge = mergeMap.get(key);
      if (merge === 'skip') continue;

      const cell = ws.getCell(r, c);
      let val = '';
      if (cell.richText) {
        val = cell.richText.map(run => {
          const s = run.font ? fontToCSS(run.font) : '';
          return s ? `<span style="${s}">${escapeXml(run.text)}</span>` : escapeXml(run.text);
        }).join('');
      } else if (cell.value != null) {
        val = escapeXml(String(cell.value));
      }

      const attrs: string[] = [];
      if (merge && typeof merge !== 'string') {
        if (merge.rowSpan > 1) attrs.push(`rowspan="${merge.rowSpan}"`);
        if (merge.colSpan > 1) attrs.push(`colspan="${merge.colSpan}"`);
      }
      if (options.includeStyles && cell.style) {
        const css = styleToCSS(cell.style);
        if (css) attrs.push(`style="${css}"`);
      }

      const attrStr = attrs.length ? ' ' + attrs.join(' ') : '';
      const tag = r === startRow ? 'th' : 'td';
      cells.push(`<${tag}${attrStr}>${val}</${tag}>`);
    }
    rows.push(`<tr>${cells.join('')}</tr>`);
  }

  const tableHtml = `<table border="1" cellpadding="4" cellspacing="0">\n${rows.join('\n')}\n</table>`;

  if (options.fullDocument === false) return tableHtml;

  const title = escapeXml(options.title ?? 'Export');
  const styleBlock = options.includeStyles ? `<style>
  table { border-collapse: collapse; font-family: Calibri, sans-serif; font-size: 11pt; }
  th, td { padding: 4px 8px; border: 1px solid #ccc; vertical-align: bottom; }
  th { background-color: #4472C4; color: white; font-weight: bold; }
</style>` : '';

  return `<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<title>${title}</title>
${styleBlock}
</head>
<body>
${tableHtml}
</body>
</html>`;
}
