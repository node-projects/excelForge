import type {
  CellStyle, Font, Fill, Border, Alignment, NumberFormat,
  PatternFill, GradientFill, BorderSide, Color
} from '../core/types.js';
import { escapeXml } from '../utils/helpers.js';

// ─── Built-in number format IDs ───────────────────────────────────────────────
const BUILTIN_NUMFMT: Record<number, string> = {
  0: 'General', 1: '0', 2: '0.00', 3: '#,##0', 4: '#,##0.00',
  9: '0%', 10: '0.00%', 11: '0.00E+00', 12: '# ?/?', 13: '# ??/??',
  14: 'mm-dd-yy', 15: 'd-mmm-yy', 16: 'd-mmm', 17: 'mmm-yy',
  18: 'h:mm AM/PM', 19: 'h:mm:ss AM/PM', 20: 'h:mm', 21: 'h:mm:ss',
  22: 'm/d/yy h:mm', 37: '#,##0 ;(#,##0)', 38: '#,##0 ;[Red](#,##0)',
  39: '#,##0.00;(#,##0.00)', 40: '#,##0.00;[Red](#,##0.00)',
  45: 'mm:ss', 46: '[h]:mm:ss', 47: 'mmss.0', 48: '##0.0E+0', 49: '@',
};

function colorXml(color: Color | undefined, prefix = 'rgb'): string {
  if (!color) return '';
  if (color.startsWith('theme:')) return `<color theme="${color.slice(6)}"/>`;
  if (color.startsWith('#')) color = 'FF' + color.slice(1);
  return `<color ${prefix}="${color}"/>`;
}

/** Emit a color attribute for inline use (font, fill, border) */
function colorAttrXml(tag: string, color: Color | undefined): string {
  if (!color) return '';
  if (color.startsWith('theme:')) return `<${tag} theme="${color.slice(6)}"/>`;
  const rgb = color.startsWith('#') ? 'FF' + color.slice(1) : color;
  return `<${tag} rgb="${rgb}"/>`;
}

function fontXml(f: Font): string {
  const parts: string[] = [];
  if (f.bold)   parts.push('<b/>');
  if (f.italic) parts.push('<i/>');
  if (f.strike) parts.push('<strike/>');
  if (f.underline && f.underline !== 'none')
    parts.push(`<u val="${f.underline}"/>`);
  if (f.vertAlign)
    parts.push(`<vertAlign val="${f.vertAlign}"/>`);
  if (f.size)   parts.push(`<sz val="${f.size}"/>`);
  if (f.color)  parts.push(colorAttrXml('color', f.color));
  if (f.name)   parts.push(`<name val="${escapeXml(f.name)}"/>`);
  if (f.family) parts.push(`<family val="${f.family}"/>`);
  if (f.scheme) parts.push(`<scheme val="${f.scheme}"/>`);
  if (f.charset) parts.push(`<charset val="${f.charset}"/>`);
  return parts.join('');
}

function fillXml(fill: Fill): string {
  if (fill.type === 'pattern') {
    const f = fill as PatternFill;
    const fg = colorAttrXml('fgColor', f.fgColor);
    const bg = colorAttrXml('bgColor', f.bgColor);
    return `<patternFill patternType="${f.pattern}">${fg}${bg}</patternFill>`;
  }
  // gradient
  const f = fill as GradientFill;
  const stops = f.stops.map(s =>
    `<stop position="${s.position}">${colorAttrXml('color', s.color)}</stop>`
  ).join('');
  const attrs = [
    f.gradientType ? `type="${f.gradientType}"` : '',
    f.degree !== undefined ? `degree="${f.degree}"` : '',
    f.left !== undefined ? `left="${f.left}"` : '',
    f.right !== undefined ? `right="${f.right}"` : '',
    f.top !== undefined ? `top="${f.top}"` : '',
    f.bottom !== undefined ? `bottom="${f.bottom}"` : '',
  ].filter(Boolean).join(' ');
  return `<gradientFill ${attrs}>${stops}</gradientFill>`;
}

function borderSideXml(tag: string, s: BorderSide | undefined): string {
  if (!s) return `<${tag}/>`;
  const color = colorAttrXml('color', s.color);
  return s.style
    ? `<${tag} style="${s.style}">${color}</${tag}>`
    : `<${tag}/>`;
}

function borderXml(b: Border): string {
  const diag = [
    b.diagonalUp   ? 'diagonalUp="1"'   : '',
    b.diagonalDown ? 'diagonalDown="1"' : '',
  ].filter(Boolean).join(' ');
  return `<border${diag ? ' '+diag : ''}>` +
    borderSideXml('left',     b.left) +
    borderSideXml('right',    b.right) +
    borderSideXml('top',      b.top) +
    borderSideXml('bottom',   b.bottom) +
    borderSideXml('diagonal', b.diagonal) +
    '</border>';
}

function alignmentXml(a: Alignment): string {
  const attrs = [
    a.horizontal    ? `horizontal="${a.horizontal}"` : '',
    a.vertical      ? `vertical="${a.vertical}"` : '',
    a.wrapText      ? `wrapText="1"` : '',
    a.shrinkToFit   ? `shrinkToFit="1"` : '',
    a.textRotation !== undefined ? `textRotation="${a.textRotation}"` : '',
    a.indent        ? `indent="${a.indent}"` : '',
    a.readingOrder !== undefined ? `readingOrder="${a.readingOrder}"` : '',
  ].filter(Boolean).join(' ');
  return `<alignment ${attrs}/>`;
}

// ─── StyleRegistry ────────────────────────────────────────────────────────────

export class StyleRegistry {
  private fonts:   string[] = [];
  private fills:   string[] = [];
  private borders: string[] = [];
  private fontIdx:   Map<string, number> = new Map();
  private fillIdx:   Map<string, number> = new Map();
  private borderIdx: Map<string, number> = new Map();
  private numFmts: Map<string, number> = new Map();
  private xfs:     string[] = [];
  private styleKey: Map<string, number> = new Map();
  private dxfs: string[] = [];  // differential formats for conditional formatting
  private nextNumFmtId = 164;  // custom formats start at 164

  // Named/cell styles
  private cellStyleXfs: string[] = [];
  private cellStyleNames: Array<{ name: string; xfId: number; builtinId?: number }> = [];

  constructor() {
    // Default required entries
    const defFont = fontXml({ name: 'Calibri', size: 11, scheme: 'minor' });
    this.fonts.push(defFont); this.fontIdx.set(defFont, 0);
    const fill0 = fillXml({ type: 'pattern', pattern: 'none' });
    const fill1 = fillXml({ type: 'pattern', pattern: 'gray125' });
    this.fills.push(fill0, fill1); this.fillIdx.set(fill0, 0); this.fillIdx.set(fill1, 1);
    const defBorder = borderXml({});
    this.borders.push(defBorder); this.borderIdx.set(defBorder, 0);
    // Default cellStyleXf (Normal style)
    this.cellStyleXfs.push(`<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>`);
    this.cellStyleNames.push({ name: 'Normal', xfId: 0, builtinId: 0 });
    // Default xf (style 0)
    this.xfs.push(`<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>`);
  }

  /**
   * Register a named cell style (appears in Excel's Cell Styles gallery).
   * Returns the xfId that can be referenced via CellStyle.namedStyleId.
   */
  registerNamedStyle(name: string, style: CellStyle, builtinId?: number): number {
    const fontId   = this.internFont(style.font);
    const fillId   = this.internFill(style.fill);
    const borderId = this.internBorder(style.border);
    const numFmtId = this.internNumFmt(style.numberFormat, style.numFmtId);
    const applyFont      = style.font      ? ' applyFont="1"'      : '';
    const applyFill      = style.fill      ? ' applyFill="1"'      : '';
    const applyBorder    = style.border    ? ' applyBorder="1"'    : '';
    const applyAlignment = style.alignment ? ' applyAlignment="1"' : '';
    const applyNumFmt    = (style.numberFormat || style.numFmtId !== undefined) ? ' applyNumberFormat="1"' : '';
    const align = style.alignment ? alignmentXml(style.alignment) : '';
    const xml = `<xf numFmtId="${numFmtId}" fontId="${fontId}" fillId="${fillId}" borderId="${borderId}"${applyFont}${applyFill}${applyBorder}${applyAlignment}${applyNumFmt}>${align}</xf>`;

    this.cellStyleXfs.push(xml);
    const xfId = this.cellStyleXfs.length - 1;
    this.cellStyleNames.push({ name, xfId, builtinId });
    return xfId;
  }

  private internFont(f: Font | undefined): number {
    if (!f) return 0;
    const xml = fontXml(f);
    const existing = this.fontIdx.get(xml);
    if (existing !== undefined) return existing;
    const idx = this.fonts.length;
    this.fonts.push(xml); this.fontIdx.set(xml, idx);
    return idx;
  }

  private internFill(f: Fill | undefined): number {
    if (!f) return 0;
    const xml = fillXml(f);
    const existing = this.fillIdx.get(xml);
    if (existing !== undefined) return existing;
    const idx = this.fills.length;
    this.fills.push(xml); this.fillIdx.set(xml, idx);
    return idx;
  }

  private internBorder(b: Border | undefined): number {
    if (!b) return 0;
    const xml = borderXml(b);
    const existing = this.borderIdx.get(xml);
    if (existing !== undefined) return existing;
    const idx = this.borders.length;
    this.borders.push(xml); this.borderIdx.set(xml, idx);
    return idx;
  }

  private internNumFmt(fmt: NumberFormat | undefined, builtinId?: number): number {
    if (builtinId !== undefined) return builtinId;
    if (!fmt) return 0;
    if (this.numFmts.has(fmt.formatCode)) return this.numFmts.get(fmt.formatCode)!;
    const id = this.nextNumFmtId++;
    this.numFmts.set(fmt.formatCode, id);
    return id;
  }

  /** Register a CellStyle and return its xf index */
  register(style: CellStyle | undefined): number {
    if (!style) return 0;
    const key = JSON.stringify(style);
    if (this.styleKey.has(key)) return this.styleKey.get(key)!;

    const fontId   = this.internFont(style.font);
    const fillId   = this.internFill(style.fill);
    const borderId = this.internBorder(style.border);
    const numFmtId = this.internNumFmt(style.numberFormat, style.numFmtId);
    const applyFont      = style.font      ? ' applyFont="1"'      : '';
    const applyFill      = style.fill      ? ' applyFill="1"'      : '';
    const applyBorder    = style.border    ? ' applyBorder="1"'    : '';
    const applyAlignment = style.alignment ? ' applyAlignment="1"' : '';
    const applyNumFmt    = (style.numberFormat || style.numFmtId !== undefined) ? ' applyNumberFormat="1"' : '';
    const applyProtection= (style.locked !== undefined || style.hidden !== undefined) ? ' applyProtection="1"' : '';
    const align = style.alignment ? alignmentXml(style.alignment) : '';
    const prot  = (style.locked !== undefined || style.hidden !== undefined)
      ? `<protection${style.locked !== undefined ? ` locked="${style.locked ? '1' : '0'}"` : ''}${style.hidden !== undefined ? ` hidden="${style.hidden ? '1' : '0'}"` : ''}/>`
      : '';
    const xfId = style.namedStyleId ?? 0;
    const xml = `<xf numFmtId="${numFmtId}" fontId="${fontId}" fillId="${fillId}" borderId="${borderId}" xfId="${xfId}"${applyFont}${applyFill}${applyBorder}${applyAlignment}${applyNumFmt}${applyProtection}>${align}${prot}</xf>`;

    this.xfs.push(xml);
    const idx = this.xfs.length - 1;
    this.styleKey.set(key, idx);
    return idx;
  }


  /**
   * Register a differential format (for conditional formatting).
   * Returns the dxfId (0-based index into <dxfs>).
   * Unlike register(), dxf styles are incremental — only the specified
   * attributes are written; omitted ones are inherited from the cell.
   */
  registerDxf(style: CellStyle): number {
    // OOXML CT_Dxf child order: font, numFmt, fill, alignment, border, protection
    const parts: string[] = [];
    if (style.font)   parts.push(`<font>${fontXml(style.font)}</font>`);
    if (style.numberFormat) {
      const id = this.internNumFmt(style.numberFormat);
      parts.push(`<numFmt numFmtId="${id}" formatCode="${escapeXml(style.numberFormat.formatCode)}"/>`);
    }
    if (style.fill)   parts.push(`<fill>${fillXml(style.fill)}</fill>`);
    if (style.alignment) {
      parts.push(`<alignment${
        style.alignment.horizontal   ? ` horizontal="${style.alignment.horizontal}"` : ''}${
        style.alignment.vertical     ? ` vertical="${style.alignment.vertical}"` : ''}${
        style.alignment.wrapText     ? ' wrapText="1"' : ''}${
        style.alignment.textRotation ? ` textRotation="${style.alignment.textRotation}"` : ''}/>`);
    }
    if (style.border) parts.push(borderXml(style.border));
    const xml = parts.join('');
    this.dxfs.push(xml);
    return this.dxfs.length - 1;
  }

  /**
   * Prepend raw dxf XML strings (already wrapped as inner content of <dxf>).
   * Used to preserve original dxf entries (e.g. table dataDxfId references)
   * before new dxfs are registered during re-serialisation.
   */
  prependRawDxfs(rawInners: string[]): void {
    this.dxfs.unshift(...rawInners);
  }

  /** Custom table style definitions */
  private tableStyleDefs: Array<{
    name: string;
    elements: Array<{ type: string; dxfId: number }>;
  }> = [];

  /** Register a custom table style with DXF-based formatting per element type. */
  registerTableStyle(name: string, def: {
    headerRow?: import('../core/types.js').CellStyle;
    dataRow1?: import('../core/types.js').CellStyle;
    dataRow2?: import('../core/types.js').CellStyle;
    totalRow?: import('../core/types.js').CellStyle;
  }): void {
    const elements: Array<{ type: string; dxfId: number }> = [];
    if (def.headerRow) elements.push({ type: 'headerRow', dxfId: this.registerDxf(def.headerRow) });
    if (def.totalRow)  elements.push({ type: 'totalRow',  dxfId: this.registerDxf(def.totalRow) });
    if (def.dataRow1)  elements.push({ type: 'firstRowStripe', dxfId: this.registerDxf(def.dataRow1) });
    if (def.dataRow2)  elements.push({ type: 'secondRowStripe', dxfId: this.registerDxf(def.dataRow2) });
    this.tableStyleDefs.push({ name, elements });
  }

  /** Produce styles.xml content */
  toXml(): string {
    const numFmtXml = this.numFmts.size
      ? `<numFmts count="${this.numFmts.size}">${
          [...this.numFmts.entries()].map(([fmt, id]) =>
            `<numFmt numFmtId="${id}" formatCode="${escapeXml(fmt)}"/>`
          ).join('')
        }</numFmts>`
      : '';

    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
${numFmtXml}
<fonts count="${this.fonts.length}">${this.fonts.map(f => `<font>${f}</font>`).join('')}</fonts>
<fills count="${this.fills.length}">${this.fills.map(f => `<fill>${f}</fill>`).join('')}</fills>
<borders count="${this.borders.length}">${this.borders.join('')}</borders>
<cellStyleXfs count="${this.cellStyleXfs.length}">${this.cellStyleXfs.join('')}</cellStyleXfs>
<cellXfs count="${this.xfs.length}">${this.xfs.join('')}</cellXfs>
<cellStyles count="${this.cellStyleNames.length}">${this.cellStyleNames.map(cs =>
  `<cellStyle name="${escapeXml(cs.name)}" xfId="${cs.xfId}"${cs.builtinId !== undefined ? ` builtinId="${cs.builtinId}"` : ''}/>`
).join('')}</cellStyles>
${this.dxfs.length ? `<dxfs count="${this.dxfs.length}">${this.dxfs.map(d => `<dxf>${d}</dxf>`).join('')}</dxfs>` : ''}
${this.tableStyleDefs.length ? `<tableStyles count="${this.tableStyleDefs.length}">${this.tableStyleDefs.map(ts =>
  `<tableStyle name="${escapeXml(ts.name)}" count="${ts.elements.length}">${ts.elements.map(el =>
    `<tableStyleElement type="${el.type}" dxfId="${el.dxfId}"/>`
  ).join('')}</tableStyle>`
).join('')}</tableStyles>` : ''}
</styleSheet>`;
  }
}
