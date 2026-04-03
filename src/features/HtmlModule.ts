/**
 * ExcelForge — Enhanced HTML/CSS Export Module (tree-shakeable).
 * Converts worksheets to rich HTML tables with inline CSS styling,
 * number formatting, conditional formatting visualization, charts,
 * column widths, row heights, images at cell anchors, form controls,
 * rich text with superscript/subscript, and MathML formula objects.
 */

import type { Worksheet } from '../core/Worksheet.js';
import type { Workbook } from '../core/Workbook.js';
import type {
  CellStyle, Font, Fill, PatternFill, GradientFill, Border, BorderSide, Alignment,
  ConditionalFormat, Chart, Sparkline, MathElement, Image, CellImage, FormControl,
} from '../core/types.js';
import { escapeXml, colIndexToLetter } from '../utils/helpers.js';

export interface HtmlExportOptions {
  /** Include inline CSS styles (default true) */
  includeStyles?: boolean;
  /** Full HTML document or just the <table> (default true) */
  fullDocument?: boolean;
  /** Document/page title */
  title?: string;
  /** CSS class prefix */
  classPrefix?: string;
  /** Include conditional formatting visualization */
  includeConditionalFormatting?: boolean;
  /** Include chart placeholders */
  includeCharts?: boolean;
  /** Include sparkline visualization (as inline SVG) */
  includeSparklines?: boolean;
  /** Include column widths from worksheet */
  includeColumnWidths?: boolean;
  /** Skip hidden rows/columns */
  skipHidden?: boolean;
  /** Only export the print area */
  printAreaOnly?: boolean;
  /** Sheet name for multi-sheet context */
  sheetName?: string;
}

export interface WorkbookHtmlExportOptions extends HtmlExportOptions {
  /** Export all sheets (default) or specific sheets by name */
  sheets?: string[];
  /** Include sheet navigation tabs */
  includeTabs?: boolean;
}

/* ─── Theme color defaults (Office standard) ──────────────────────────────── */
const THEME_COLORS = [
  '#000000', '#FFFFFF', '#44546A', '#E7E6E6', '#4472C4', '#ED7D31',
  '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47',
];

function colorToCSS(c: string | undefined): string {
  if (!c) return '';
  if (c.startsWith('#')) return c;
  if (c.startsWith('theme:')) {
    const idx = parseInt(c.slice(6), 10);
    return THEME_COLORS[idx] ?? '#000';
  }
  if (c.length === 8 && !c.startsWith('#')) return '#' + c.slice(2);
  return '#' + c;
}

function parseColor(c: string): [number, number, number] {
  const hex = c.replace('#', '');
  return [parseInt(hex.slice(0, 2), 16), parseInt(hex.slice(2, 4), 16), parseInt(hex.slice(4, 6), 16)];
}

function interpolateColor(c1: string, c2: string, t: number): string {
  const [r1, g1, b1] = parseColor(colorToCSS(c1) || '#FFFFFF');
  const [r2, g2, b2] = parseColor(colorToCSS(c2) || '#000000');
  const r = Math.round(r1 + (r2 - r1) * t);
  const g = Math.round(g1 + (g2 - g1) * t);
  const b = Math.round(b1 + (b2 - b1) * t);
  return `#${r.toString(16).padStart(2, '0')}${g.toString(16).padStart(2, '0')}${b.toString(16).padStart(2, '0')}`;
}

/* ─── CSS builders ─────────────────────────────────────────────────────────── */

function fontToCSS(f: Font): string {
  const parts: string[] = [];
  if (f.bold) parts.push('font-weight:bold');
  if (f.italic) parts.push('font-style:italic');
  const decs: string[] = [];
  if (f.underline && f.underline !== 'none') decs.push('underline');
  if (f.strike) decs.push('line-through');
  if (decs.length) parts.push(`text-decoration:${decs.join(' ')}`);
  if (f.size) parts.push(`font-size:${f.size}pt`);
  if (f.color) parts.push(`color:${colorToCSS(f.color)}`);
  if (f.name) parts.push(`font-family:'${f.name}',sans-serif`);
  if (f.vertAlign === 'superscript') parts.push('vertical-align:super;font-size:smaller');
  else if (f.vertAlign === 'subscript') parts.push('vertical-align:sub;font-size:smaller');
  return parts.join(';');
}

function fillToCSS(fill: Fill): string {
  if (fill.type === 'pattern') {
    const pf = fill as PatternFill;
    if (pf.pattern === 'solid' && pf.fgColor) return `background-color:${colorToCSS(pf.fgColor)}`;
  }
  if (fill.type === 'gradient') {
    const gf = fill as GradientFill;
    if (gf.stops && gf.stops.length >= 2) {
      const stops = gf.stops.map(s => `${colorToCSS(s.color)} ${Math.round(s.position * 100)}%`).join(',');
      const deg = gf.degree ?? 0;
      return `background:linear-gradient(${deg}deg,${stops})`;
    }
  }
  return '';
}

function borderSideCSS(side: BorderSide | undefined): string {
  if (!side || !side.style) return '';
  const widthMap: Record<string, string> = {
    thin: '1px', medium: '2px', thick: '3px', dashed: '1px', dotted: '1px',
    double: '3px', hair: '1px', mediumDashed: '2px', dashDot: '1px',
    mediumDashDot: '2px', dashDotDot: '1px', mediumDashDotDot: '2px', slantDashDot: '2px',
  };
  const styleMap: Record<string, string> = {
    thin: 'solid', medium: 'solid', thick: 'solid', dashed: 'dashed', dotted: 'dotted',
    double: 'double', hair: 'solid', mediumDashed: 'dashed', dashDot: 'dashed',
    mediumDashDot: 'dashed', dashDotDot: 'dotted', mediumDashDotDot: 'dotted', slantDashDot: 'dashed',
  };
  const w = widthMap[side.style] ?? '1px';
  const s = styleMap[side.style] ?? 'solid';
  const c = side.color ? colorToCSS(side.color) : '#000';
  return `${w} ${s} ${c}`;
}

function alignmentCSS(a: Alignment): string {
  const parts: string[] = [];
  if (a.horizontal) {
    const hMap: Record<string, string> = { left: 'left', center: 'center', right: 'right', fill: 'justify', justify: 'justify', distributed: 'justify' };
    parts.push(`text-align:${hMap[a.horizontal] ?? a.horizontal}`);
  }
  if (a.vertical) {
    const vMap: Record<string, string> = { top: 'top', center: 'middle', bottom: 'bottom', distributed: 'middle' };
    parts.push(`vertical-align:${vMap[a.vertical] ?? 'bottom'}`);
  }
  if (a.wrapText) parts.push('white-space:normal;word-wrap:break-word');
  if (a.textRotation) parts.push(`transform:rotate(-${a.textRotation}deg)`);
  if (a.indent) parts.push(`padding-left:${a.indent * 8}px`);
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

/* ─── Number formatting ────────────────────────────────────────────────────── */

function formatNumber(value: unknown, fmt: string | undefined): string {
  if (value == null) return '';
  if (!fmt || fmt === 'General') return String(value);
  const num = typeof value === 'number' ? value : parseFloat(String(value));
  if (isNaN(num)) return String(value);

  // Percentage
  if (fmt.includes('%')) {
    const decimals = (fmt.match(/0\.(0+)%/) ?? [])[1]?.length ?? 0;
    return (num * 100).toFixed(decimals) + '%';
  }
  // Currency / Accounting
  const currMatch = fmt.match(/[$€£¥]|"CHF"/);
  if (currMatch) {
    const sym = currMatch[0].replace(/"/g, '');
    const decimals = (fmt.match(/\.(0+)/) ?? [])[1]?.length ?? 2;
    const formatted = Math.abs(num).toFixed(decimals).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
    if (fmt.indexOf(currMatch[0]) < fmt.indexOf('0')) {
      return (num < 0 ? '-' : '') + sym + formatted;
    }
    return (num < 0 ? '-' : '') + formatted + ' ' + sym;
  }
  // Thousands separator
  if (fmt.includes('#,##0') || fmt.includes('#,###')) {
    const decimals = (fmt.match(/\.(0+)/) ?? [])[1]?.length ?? 0;
    return num.toFixed(decimals).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  }
  // Fixed decimals
  const fixedMatch = fmt.match(/^0\.(0+)$/);
  if (fixedMatch) return num.toFixed(fixedMatch[1].length);
  // Date patterns
  if (/[ymdh]/i.test(fmt)) return formatDate(num, fmt);
  // Fraction
  if (fmt.includes('?/?') || fmt.includes('??/??')) return formatFraction(num);
  // Scientific
  if (/0\.0+E\+0+/i.test(fmt)) {
    const decimals = (fmt.match(/0\.(0+)/) ?? [])[1]?.length ?? 2;
    return num.toExponential(decimals).toUpperCase();
  }
  return String(value);
}

function formatDate(serial: number, fmt: string): string {
  const epoch = new Date(1899, 11, 30);
  const d = new Date(epoch.getTime() + serial * 86400000);
  const Y = d.getFullYear(), M = d.getMonth() + 1, D = d.getDate();
  const h = d.getHours(), m = d.getMinutes(), s = d.getSeconds();
  return fmt
    .replace(/yyyy/gi, String(Y))
    .replace(/yy/gi, String(Y).slice(-2))
    .replace(/mmmm/gi, d.toLocaleDateString('en', { month: 'long' }))
    .replace(/mmm/gi, d.toLocaleDateString('en', { month: 'short' }))
    .replace(/mm/gi, String(M).padStart(2, '0'))
    .replace(/m/gi, String(M))
    .replace(/dd/gi, String(D).padStart(2, '0'))
    .replace(/d/gi, String(D))
    .replace(/hh/gi, String(h).padStart(2, '0'))
    .replace(/h/gi, String(h))
    .replace(/ss/gi, String(s).padStart(2, '0'))
    .replace(/nn|MM/g, String(m).padStart(2, '0'));
}

function formatFraction(num: number): string {
  const whole = Math.floor(Math.abs(num));
  const frac = Math.abs(num) - whole;
  if (frac < 0.0001) return String(num < 0 ? -whole : whole);
  let bestN = 0, bestD = 1, bestErr = 1;
  for (let d = 1; d <= 100; d++) {
    const n = Math.round(frac * d);
    const err = Math.abs(frac - n / d);
    if (err < bestErr) { bestN = n; bestD = d; bestErr = err; }
    if (err < 0.0001) break;
  }
  const sign = num < 0 ? '-' : '';
  return whole > 0 ? `${sign}${whole} ${bestN}/${bestD}` : `${sign}${bestN}/${bestD}`;
}

/* ─── Conditional formatting helpers ───────────────────────────────────────── */

function evaluateConditionalFormats(cf: ConditionalFormat, value: number, allValues: number[]): string {
  if (cf.colorScale) {
    const sorted = [...allValues].sort((a, b) => a - b);
    const min = sorted[0], max = sorted[sorted.length - 1];
    const range = max - min || 1;
    const t = (value - min) / range;
    const cs = cf.colorScale;
    if (cs.color.length === 2) return `background-color:${interpolateColor(cs.color[0], cs.color[1], t)}`;
    if (cs.color.length >= 3) {
      if (t <= 0.5) return `background-color:${interpolateColor(cs.color[0], cs.color[1], t * 2)}`;
      return `background-color:${interpolateColor(cs.color[1], cs.color[2], (t - 0.5) * 2)}`;
    }
  }
  if (cf.dataBar) {
    const sorted = [...allValues].sort((a, b) => a - b);
    const min = sorted[0], max = sorted[sorted.length - 1];
    const pct = Math.round(((value - min) / (max - min || 1)) * 100);
    const color = colorToCSS(cf.dataBar.color) || '#638EC6';
    return `background:linear-gradient(90deg,${color} ${pct}%,transparent ${pct}%)`;
  }
  if (cf.iconSet) {
    const sorted = [...allValues].sort((a, b) => a - b);
    const min = sorted[0], max = sorted[sorted.length - 1];
    const t = (value - min) / (max - min || 1);
    const ICON_MAP: Record<string, string[]> = {
      '3Arrows': ['↓', '→', '↑'], '3ArrowsGray': ['⇩', '⇨', '⇧'],
      '3TrafficLights1': ['🔴', '🟡', '🟢'], '3TrafficLights2': ['🔴', '🟡', '🟢'],
      '3Signs': ['⛔', '⚠️', '✅'], '3Symbols': ['✖', '!', '✔'],
      '3Symbols2': ['✖', '!', '✔'], '3Flags': ['🏴', '🏳', '🏁'],
      '3Stars': ['☆', '★', '★'], '4Arrows': ['↓', '↘', '↗', '↑'],
      '4ArrowsGray': ['⇩', '⇘', '⇗', '⇧'], '4Rating': ['◔', '◑', '◕', '●'],
      '4RedToBlack': ['⬤', '⬤', '⬤', '⬤'], '4TrafficLights': ['⬤', '⬤', '⬤', '⬤'],
      '5Arrows': ['↓', '↘', '→', '↗', '↑'], '5ArrowsGray': ['⇩', '⇘', '⇨', '⇗', '⇧'],
      '5Quarters': ['○', '◔', '◑', '◕', '●'], '5Rating': ['◔', '◔', '◑', '◕', '●'],
    };
    const icons = ICON_MAP[cf.iconSet.iconSet ?? '3TrafficLights1'] ?? ['🔴', '🟡', '🟢'];
    const idx = Math.min(Math.floor(t * icons.length), icons.length - 1);
    return `data-icon="${icons[idx]}"`;
  }
  return '';
}

/* ─── Sparkline SVG ────────────────────────────────────────────────────────── */

function sparklineToSvg(sparkline: Sparkline, values: number[]): string {
  if (!values.length) return '';
  const W = 100, H = 20;
  const min = Math.min(...values), max = Math.max(...values);
  const range = max - min || 1;
  const color = colorToCSS(sparkline.color) || '#4472C4';

  if (sparkline.type === 'bar' || sparkline.type === 'stacked') {
    const bw = W / values.length;
    const bars = values.map((v, i) => {
      const barH = sparkline.type === 'stacked'
        ? (v >= 0 ? H / 2 : H / 2)
        : ((v - min) / range) * H;
      const y = sparkline.type === 'stacked'
        ? (v >= 0 ? H / 2 - barH : H / 2)
        : H - barH;
      const fill = v < 0 && sparkline.negativeColor ? colorToCSS(sparkline.negativeColor) : color;
      return `<rect x="${i * bw}" y="${y}" width="${bw * 0.8}" height="${barH}" fill="${fill}"/>`;
    }).join('');
    return `<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}" viewBox="0 0 ${W} ${H}" style="display:inline-block;vertical-align:middle">${bars}</svg>`;
  }

  // Line
  const pts = values.map((v, i) => `${(i / (values.length - 1 || 1)) * W},${H - ((v - min) / range) * H}`).join(' ');
  let markers = '';
  if (sparkline.showMarkers) {
    markers = values.map((v, i) => {
      const x = (i / (values.length - 1 || 1)) * W;
      const y = H - ((v - min) / range) * H;
      return `<circle cx="${x}" cy="${y}" r="1.5" fill="${colorToCSS(sparkline.markersColor) || color}"/>`;
    }).join('');
  }
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}" viewBox="0 0 ${W} ${H}" style="display:inline-block;vertical-align:middle"><polyline points="${pts}" fill="none" stroke="${color}" stroke-width="1.5"/>${markers}</svg>`;
}

/* ─── Chart placeholder ────────────────────────────────────────────────────── */

function chartToHtml(chart: Chart): string {
  const w = 480, h = 300;
  return `<div style="display:inline-block;width:${w}px;height:${h}px;border:1px solid #ccc;background:#f9f9f9;text-align:center;line-height:${h}px;color:#666;font-size:14px;margin:8px 0" data-chart-type="${chart.type}">[Chart: ${escapeXml(chart.title ?? chart.type)}]</div>`;
}

/* ─── Shape rendering ──────────────────────────────────────────────────────── */

function shapeToHtml(shape: { type: string; text?: string; fillColor?: string; lineColor?: string }): string {
  const toHex = (c: string) => { let h = c.replace(/^#/, ''); if (h.length === 8) h = h.substring(2); return '#' + h; };
  const bg = shape.fillColor ? toHex(shape.fillColor) : '#4472C4';
  const border = shape.lineColor ? toHex(shape.lineColor) : '#2F5496';
  const radius = shape.type === 'ellipse' ? '50%' : shape.type === 'roundRect' ? '12px' : '0';
  return `<div style="display:inline-flex;align-items:center;justify-content:center;width:160px;height:80px;background:${bg};border:2px solid ${border};border-radius:${radius};color:white;font-weight:600;font-size:13px;margin:4px;text-align:center">${escapeXml(shape.text ?? shape.type)}</div>`;
}

/* ─── WordArt rendering ────────────────────────────────────────────────────── */

function wordArtToHtml(wa: { text: string; fillColor?: string; outlineColor?: string; font?: any; preset?: string }): string {
  const toHex = (c: string) => { let h = c.replace(/^#/, ''); if (h.length === 8) h = h.substring(2); return '#' + h; };
  const color = wa.fillColor ? toHex(wa.fillColor) : '#333';
  const outline = wa.outlineColor ? toHex(wa.outlineColor) : '';
  const family = wa.font?.name ?? 'Impact';
  const size = wa.font?.size ?? 36;
  const bold = wa.font?.bold !== false ? 'font-weight:bold;' : '';
  const textStroke = outline ? `-webkit-text-stroke:1px ${outline};paint-order:stroke fill;` : '';
  const presetStyle = wordArtPresetCSS(wa.preset);
  return `<div style="display:inline-block;font-family:'${escapeXml(family)}',sans-serif;font-size:${size}px;${bold}color:${color};${textStroke}text-shadow:2px 2px 4px rgba(0,0,0,.3);${presetStyle}margin:8px 0;padding:8px 16px">${escapeXml(wa.text)}</div>`;
}

function wordArtPresetCSS(preset?: string): string {
  if (!preset || preset === 'textPlain') return '';
  if (preset === 'textArchUp') return 'letter-spacing:4px;';
  if (preset === 'textArchDown') return 'letter-spacing:4px;transform:scaleY(-1);';
  if (preset === 'textWave1' || preset === 'textWave2') return 'letter-spacing:2px;font-style:italic;';
  if (preset === 'textInflate') return 'letter-spacing:3px;transform:scaleY(1.3);';
  if (preset === 'textDeflate') return 'letter-spacing:1px;transform:scaleY(0.7);';
  if (preset === 'textSlantUp') return 'transform:perspective(200px) rotateY(-5deg);';
  if (preset === 'textSlantDown') return 'transform:perspective(200px) rotateY(5deg);';
  return '';
}

/* ─── Image rendering ──────────────────────────────────────────────────────── */

function toBase64(bytes: Uint8Array): string {
  let b64 = '';
  for (let i = 0; i < bytes.length; i += 3) {
    const b0 = bytes[i], b1 = bytes[i + 1] ?? 0, b2 = bytes[i + 2] ?? 0;
    const n = (b0 << 16) | (b1 << 8) | b2;
    const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/';
    b64 += chars[(n >> 18) & 63] + chars[(n >> 12) & 63];
    b64 += i + 1 < bytes.length ? chars[(n >> 6) & 63] : '=';
    b64 += i + 2 < bytes.length ? chars[n & 63] : '=';
  }
  return b64;
}

function imageDataUri(data: string | Uint8Array, format?: string): string {
  const mime = formatToMime(format);
  const b64 = typeof data === 'string' ? data : toBase64(data);
  return `data:${mime};base64,${b64}`;
}

/** Render a floating image at its cell anchor position */
function imageToPositionedHtml(img: Image): string {
  const src = imageDataUri(img.data, img.format);
  const alt = img.altText ? ` alt="${escapeXml(img.altText)}"` : '';
  const w = img.width ? `width:${img.width}px;` : 'max-width:400px;';
  const h = img.height ? `height:${img.height}px;` : 'max-height:300px;';
  // Position via data attributes — resolved to CSS in the overlay container
  const fromCol = img.from?.col ?? 0;
  const fromRow = img.from?.row ?? 0;
  return `<img src="${src}"${alt} class="xl-img" data-from-col="${fromCol}" data-from-row="${fromRow}" style="${w}${h}border:1px solid #ddd;border-radius:4px"/>`;
}

/** Render a cell image inline (for in-cell pictures) */
function cellImageToHtml(ci: CellImage): string {
  const src = imageDataUri(ci.data, ci.format);
  const alt = ci.altText ? ` alt="${escapeXml(ci.altText)}"` : '';
  return `<img src="${src}"${alt} style="max-width:100%;max-height:100%;object-fit:contain"/>`;
}

function formatToMime(fmt?: string): string {
  switch (fmt) {
    case 'jpeg': case 'jpg': return 'image/jpeg';
    case 'gif': return 'image/gif';
    case 'svg': return 'image/svg+xml';
    case 'webp': return 'image/webp';
    case 'bmp': return 'image/bmp';
    default: return 'image/png';
  }
}

/* ─── Math equation rendering (MathML) ────────────────────────────────────── */

function mathElementToMathML(el: MathElement): string {
  switch (el.type) {
    case 'text': {
      const text = escapeXml(el.text ?? '');
      // If the text is a known function name or plain text, use <mo> or <mi>
      if (el.font === 'normal' || /^[a-zA-Z]{2,}$/.test(el.text ?? ''))
        return `<mi mathvariant="normal">${text}</mi>`;
      if (/^[0-9.]+$/.test(el.text ?? ''))
        return `<mn>${text}</mn>`;
      if (el.text && el.text.length === 1 && /[+\-*/=<>±×÷≤≥≠∞∈∉∪∩⊂⊃∧∨¬→←↔∀∃∑∏∫]/.test(el.text))
        return `<mo>${text}</mo>`;
      return `<mi>${text}</mi>`;
    }
    case 'frac':
      return `<mfrac><mrow>${(el.base ?? []).map(mathElementToMathML).join('')}</mrow><mrow>${(el.argument ?? []).map(mathElementToMathML).join('')}</mrow></mfrac>`;
    case 'sup':
      return `<msup><mrow>${(el.base ?? []).map(mathElementToMathML).join('')}</mrow><mrow>${(el.argument ?? []).map(mathElementToMathML).join('')}</mrow></msup>`;
    case 'sub':
      return `<msub><mrow>${(el.base ?? []).map(mathElementToMathML).join('')}</mrow><mrow>${(el.argument ?? []).map(mathElementToMathML).join('')}</mrow></msub>`;
    case 'subSup':
      return `<msubsup><mrow>${(el.base ?? []).map(mathElementToMathML).join('')}</mrow><mrow>${(el.subscript ?? []).map(mathElementToMathML).join('')}</mrow><mrow>${(el.superscript ?? []).map(mathElementToMathML).join('')}</mrow></msubsup>`;
    case 'nary':
      return `<munderover><mo>${escapeXml(el.operator ?? '∑')}</mo><mrow>${(el.lower ?? []).map(mathElementToMathML).join('')}</mrow><mrow>${(el.upper ?? []).map(mathElementToMathML).join('')}</mrow></munderover><mrow>${(el.body ?? []).map(mathElementToMathML).join('')}</mrow>`;
    case 'rad':
      if (!el.hideDegree && el.degree?.length)
        return `<mroot><mrow>${(el.body ?? []).map(mathElementToMathML).join('')}</mrow><mrow>${el.degree.map(mathElementToMathML).join('')}</mrow></mroot>`;
      return `<msqrt><mrow>${(el.body ?? []).map(mathElementToMathML).join('')}</mrow></msqrt>`;
    case 'delim':
      return `<mrow><mo>${escapeXml(el.open ?? '(')}</mo>${(el.body ?? []).map(mathElementToMathML).join('')}<mo>${escapeXml(el.close ?? ')')}</mo></mrow>`;
    case 'func':
      return `<mrow><mi mathvariant="normal">${(el.base ?? []).map(e => escapeXml(e.text ?? '')).join('')}</mi><mo>&#x2061;</mo>${(el.argument ?? []).map(mathElementToMathML).join('')}</mrow>`;
    case 'groupChar':
      return `<munder><mrow>${(el.base ?? []).map(mathElementToMathML).join('')}</mrow><mo>${escapeXml(el.operator ?? '⏟')}</mo></munder>`;
    case 'accent':
      return `<mover accent="true"><mrow>${(el.base ?? []).map(mathElementToMathML).join('')}</mrow><mo>${escapeXml(el.operator ?? '̂')}</mo></mover>`;
    case 'bar':
      return `<mover><mrow>${(el.base ?? []).map(mathElementToMathML).join('')}</mrow><mo>¯</mo></mover>`;
    case 'limLow':
      return `<munder><mrow>${(el.base ?? []).map(mathElementToMathML).join('')}</mrow><mrow>${(el.argument ?? []).map(mathElementToMathML).join('')}</mrow></munder>`;
    case 'limUpp':
      return `<mover><mrow>${(el.base ?? []).map(mathElementToMathML).join('')}</mrow><mrow>${(el.argument ?? []).map(mathElementToMathML).join('')}</mrow></mover>`;
    case 'eqArr':
      return `<mtable>${(el.rows ?? []).map(row => `<mtr><mtd>${row.map(mathElementToMathML).join('')}</mtd></mtr>`).join('')}</mtable>`;
    case 'matrix':
      return `<mrow><mo>(</mo><mtable>${(el.rows ?? []).map(row => `<mtr>${row.map(c => `<mtd>${mathElementToMathML(c)}</mtd>`).join('')}</mtr>`).join('')}</mtable><mo>)</mo></mrow>`;
    default:
      return el.text ? `<mi>${escapeXml(el.text)}</mi>` : '<mrow></mrow>';
  }
}

function mathEquationToMathML(eq: { elements: MathElement[]; fontSize?: number; fontName?: string }): string {
  const size = eq.fontSize ?? 11;
  const font = eq.fontName ?? 'Cambria Math';
  return `<div style="display:inline-block;font-family:'${escapeXml(font)}',serif;font-size:${size}pt;margin:8px;padding:4px"><math xmlns="http://www.w3.org/1998/Math/MathML" display="block"><mrow>${eq.elements.map(mathElementToMathML).join('')}</mrow></math></div>`;
}

/* ─── Form control rendering ──────────────────────────────────────────────── */

function formControlToPositionedHtml(fc: FormControl): string {
  const fromCol = fc.from.col;
  const fromRow = fc.from.row;
  const w = fc.width ? `width:${fc.width}px;` : '';
  const h = fc.height ? `height:${fc.height}px;` : '';
  const linked = fc.linkedCell ? ` data-linked-cell="${escapeXml(fc.linkedCell)}"` : '';
  const inputRange = fc.inputRange ? ` data-input-range="${escapeXml(fc.inputRange)}"` : '';
  const macro = fc.macro ? ` data-macro="${escapeXml(fc.macro)}"` : '';
  let inner = '';
  switch (fc.type) {
    case 'button':
    case 'dialog':
      inner = `<button style="padding:6px 16px;font-size:13px;border:1px solid #999;background:#f0f0f0;cursor:pointer;${w}${h}"${macro}>${escapeXml(fc.text ?? fc.type)}</button>`;
      break;
    case 'checkBox': {
      const checked = fc.checked === 'checked' ? ' checked' : '';
      const indeterminate = fc.checked === 'mixed' ? ' data-indeterminate="true"' : '';
      inner = `<label style="font-size:13px;display:inline-flex;align-items:center;gap:4px;${w}"><input type="checkbox"${checked}${indeterminate}${linked}/> ${escapeXml(fc.text ?? 'Checkbox')}</label>`;
      break;
    }
    case 'optionButton': {
      const checked = fc.checked === 'checked' ? ' checked' : '';
      inner = `<label style="font-size:13px;display:inline-flex;align-items:center;gap:4px;${w}"><input type="radio" name="group"${checked}${linked}/> ${escapeXml(fc.text ?? 'Option')}</label>`;
      break;
    }
    case 'comboBox': {
      const lines = fc.dropLines ?? 8;
      inner = `<select style="padding:4px;font-size:13px;${w}"${linked}${inputRange} size="1" data-drop-lines="${lines}"><option>${escapeXml(fc.text ?? 'Select...')}</option></select>`;
      break;
    }
    case 'listBox': {
      const size = fc.dropLines ?? 5;
      const sel = fc.selType ?? 'single';
      const multi = sel === 'multi' || sel === 'extend' ? ' multiple' : '';
      inner = `<select style="padding:4px;font-size:13px;${w}${h}"${linked}${inputRange} size="${size}"${multi}><option>${escapeXml(fc.text ?? 'Item')}</option></select>`;
      break;
    }
    case 'spinner': {
      const min = fc.min ?? 0;
      const max = fc.max ?? 100;
      const step = fc.inc ?? 1;
      const val = fc.val ?? min;
      inner = `<input type="number" value="${val}" min="${min}" max="${max}" step="${step}" style="width:60px;padding:4px;font-size:13px;${h}"${linked}/>`;
      break;
    }
    case 'scrollBar': {
      const min = fc.min ?? 0;
      const max = fc.max ?? 100;
      const step = fc.inc ?? 1;
      const val = fc.val ?? min;
      inner = `<input type="range" value="${val}" min="${min}" max="${max}" step="${step}" style="${w || 'width:120px;'}${h}"${linked}/>`;
      break;
    }
    case 'label':
      inner = `<span style="font-size:13px;${w}">${escapeXml(fc.text ?? 'Label')}</span>`;
      break;
    case 'groupBox':
      inner = `<fieldset style="padding:8px;border:1px solid #999;font-size:13px;${w}${h}"><legend>${escapeXml(fc.text ?? 'Group')}</legend></fieldset>`;
      break;
    default:
      inner = `<span style="font-size:13px">[${escapeXml(fc.type)}]</span>`;
  }
  return `<div class="xl-fc" data-from-col="${fromCol}" data-from-row="${fromRow}" style="position:absolute;margin:4px">${inner}</div>`;
}

/* ─── Column letter → 0-based index ───────────────────────────────────────── */

function colLetterToIdx(letter: string): number {
  let idx = 0;
  for (let i = 0; i < letter.length; i++) {
    idx = idx * 26 + (letter.charCodeAt(i) - 64);
  }
  return idx; // 1-based
}

/* ─── Sparkline data resolver ──────────────────────────────────────────────── */

function resolveSparklineData(ws: Worksheet, dataRange: string): number[] {
  const vals: number[] = [];
  // Try to parse common range formats like "Sheet1!A2:A10" or "A2:A10"
  const ref = dataRange.includes('!') ? dataRange.split('!')[1] : dataRange;
  const m = ref.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
  if (m) {
    const c1 = colLetterToIdx(m[1]), r1 = parseInt(m[2], 10);
    const c2 = colLetterToIdx(m[3]), r2 = parseInt(m[4], 10);
    for (let r = r1; r <= r2; r++) {
      for (let c = c1; c <= c2; c++) {
        const cell = ws.getCell(r, c);
        if (typeof cell.value === 'number') vals.push(cell.value);
      }
    }
  }
  return vals;
}

/* ─── Main worksheet export ────────────────────────────────────────────────── */

/**
 * Convert a worksheet to an HTML table string with rich formatting.
 */
export function worksheetToHtml(ws: Worksheet, options: HtmlExportOptions = {}): string {
  const range = ws.getUsedRange();
  if (!range) {
    return options.fullDocument !== false
      ? `<!DOCTYPE html><html><head><title>${escapeXml(options.title ?? '')}</title></head><body><p>Empty worksheet</p></body></html>`
      : '<table></table>';
  }

  let { startRow, startCol, endRow, endCol } = range;

  // Print area restriction
  if (options.printAreaOnly && ws.printArea) {
    const pa = ws.printArea;
    const m = pa.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
    if (m) {
      startCol = colLetterToIdx(m[1]); startRow = parseInt(m[2], 10);
      endCol = colLetterToIdx(m[3]); endRow = parseInt(m[4], 10);
    }
  }

  const merges = ws.getMerges();
  const conditionalFormats = options.includeConditionalFormatting !== false ? ws.getConditionalFormats() : [];
  const sparklines = options.includeSparklines !== false ? ws.getSparklines() : [];

  // Build sparkline map: "row,col" → Sparkline
  const sparklineMap = new Map<string, Sparkline>();
  for (const sp of sparklines) {
    const m = sp.location.match(/^([A-Z]+)(\d+)$/);
    if (m) sparklineMap.set(`${parseInt(m[2], 10)},${colLetterToIdx(m[1])}`, sp);
  }

  // Collect numeric values per CF sqref for relative evaluation
  const cfValueMap = new Map<ConditionalFormat, number[]>();
  for (const cf of conditionalFormats) {
    if (!cf.colorScale && !cf.dataBar && !cf.iconSet) continue;
    const vals: number[] = [];
    const refs = cf.sqref.split(' ');
    for (const ref of refs) {
      const rm = ref.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
      if (rm) {
        for (let r = parseInt(rm[2], 10); r <= parseInt(rm[4], 10); r++) {
          for (let c = colLetterToIdx(rm[1]); c <= colLetterToIdx(rm[3]); c++) {
            const cell = ws.getCell(r, c);
            if (typeof cell.value === 'number') vals.push(cell.value);
          }
        }
      }
    }
    cfValueMap.set(cf, vals);
  }

  // Build merge map
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

  // Column widths
  const colWidths: string[] = [];
  if (options.includeColumnWidths !== false) {
    for (let c = startCol; c <= endCol; c++) {
      const def = ws.getColumn(c);
      if (options.skipHidden && def?.hidden) continue;
      const w = def?.width ? Math.round(def.width * 7.5) : undefined; // approx px
      colWidths.push(w ? `<col style="width:${w}px">` : '<col>');
    }
  }

  // Build cell image map: "B2" → CellImage
  const cellImageMap = new Map<string, CellImage>();
  const cellImages = ws.getCellImages?.();
  if (cellImages) {
    for (const ci of cellImages) cellImageMap.set(ci.cell, ci);
  }

  const rows: string[] = [];
  for (let r = startRow; r <= endRow; r++) {
    const rowDef = ws.getRow(r);
    if (options.skipHidden && rowDef?.hidden) continue;

    const rowStyle = rowDef?.height ? ` style="height:${rowDef.height}px"` : '';
    const cells: string[] = [];
    for (let c = startCol; c <= endCol; c++) {
      const colDef = ws.getColumn(c);
      if (options.skipHidden && colDef?.hidden) continue;

      const key = `${r},${c}`;
      const merge = mergeMap.get(key);
      if (merge === 'skip') continue;

      const cell = ws.getCell(r, c);
      let val = '';
      // Cell image (in-cell picture) takes priority
      const cellRef = `${colIndexToLetter(c)}${r}`;
      const ci = cellImageMap.get(cellRef);
      if (ci) {
        val = cellImageToHtml(ci);
      } else if (cell.richText) {
        val = cell.richText.map(run => {
          const s = run.font ? fontToCSS(run.font) : '';
          return s ? `<span style="${s}">${escapeXml(run.text)}</span>` : escapeXml(run.text);
        }).join('');
      } else if (cell.value != null) {
        const formatted = cell.style?.numberFormat
          ? formatNumber(cell.value, cell.style.numberFormat.formatCode)
          : String(cell.value);
        val = escapeXml(formatted);
      }

      // Hyperlink
      if (cell.hyperlink) {
        const href = escapeXml(cell.hyperlink.href ?? '');
        const tip = cell.hyperlink.tooltip ? ` title="${escapeXml(cell.hyperlink.tooltip)}"` : '';
        val = `<a href="${href}"${tip} style="color:#0563C1;text-decoration:underline">${val}</a>`;
      }

      // Comment tooltip
      if (cell.comment) {
        const commentText = cell.comment.richText
          ? cell.comment.richText.map(run => run.text).join('')
          : cell.comment.text;
        val = `<span title="${escapeXml(commentText)}" style="cursor:help">${val}</span>`;
      }

      const attrs: string[] = [];
      if (merge && typeof merge !== 'string') {
        if (merge.rowSpan > 1) attrs.push(`rowspan="${merge.rowSpan}"`);
        if (merge.colSpan > 1) attrs.push(`colspan="${merge.colSpan}"`);
      }

      // Cell style + conditional formatting
      const cssParts: string[] = [];
      if (options.includeStyles !== false && cell.style) cssParts.push(styleToCSS(cell.style));

      // Conditional formatting evaluation
      let iconAttr = '';
      if (typeof cell.value === 'number') {
        for (const cf of conditionalFormats) {
          const allVals = cfValueMap.get(cf);
          if (!allVals) continue;
          const result = evaluateConditionalFormats(cf, cell.value, allVals);
          if (result.startsWith('data-icon=')) {
            iconAttr = ` ${result}`;
          } else if (result) {
            cssParts.push(result);
          }
        }
      }

      const css = cssParts.filter(Boolean).join(';');
      if (css) attrs.push(`style="${css}"`);
      attrs.push(`data-cell="${colIndexToLetter(c)}${r}"`);

      // Sparkline
      const sp = sparklineMap.get(key);
      if (sp) val += sparklineToSvg(sp, []); // values would need parsing from dataRange

      const attrStr = attrs.length ? ' ' + attrs.join(' ') : '';
      cells.push(`<td${attrStr}${iconAttr}>${val}</td>`);
    }
    rows.push(`<tr${rowStyle}>${cells.join('')}</tr>`);
  }

  const colGroup = colWidths.length ? `<colgroup>${colWidths.join('')}</colgroup>` : '';
  const tableHtml = `<div class="xl-sheet-wrapper" style="position:relative;display:inline-block"><table border="0" cellpadding="4" cellspacing="0">\n${colGroup}\n${rows.join('\n')}\n</table>`;

  // Charts below table
  let chartsHtml = '';
  if (options.includeCharts !== false) {
    const charts = ws.getCharts();
    if (charts.length) chartsHtml = '\n' + charts.map(chartToHtml).join('\n');
  }

  // Images — positioned as overlays on a wrapper container
  let imagesHtml = '';
  const images = ws.getImages?.();
  if (images?.length) imagesHtml = '\n<div class="xl-images">' + images.map(imageToPositionedHtml).join('\n') + '</div>';

  // Shapes
  let shapesHtml = '';
  const shapes = ws.getShapes?.();
  if (shapes?.length) shapesHtml = '\n<div class="shapes">' + shapes.map(shapeToHtml).join('\n') + '</div>';

  // WordArt
  let wordArtHtml = '';
  const wordArts = ws.getWordArt?.();
  if (wordArts?.length) wordArtHtml = '\n<div class="wordart">' + wordArts.map(wordArtToHtml).join('\n') + '</div>';

  // Math equations (MathML)
  let mathHtml = '';
  const mathEqs = ws.getMathEquations?.();
  if (mathEqs?.length) mathHtml = '\n<div class="math-equations">' + mathEqs.map(mathEquationToMathML).join('\n') + '</div>';

  // Form controls — positioned as overlays
  let formControlsHtml = '';
  const fcs = ws.getFormControls?.();
  if (fcs?.length) formControlsHtml = '\n<div class="xl-form-controls">' + fcs.map(formControlToPositionedHtml).join('\n') + '</div>';

  // Sparklines rendered inline
  let sparklinesHtml = '';
  if (options.includeSparklines !== false) {
    const spks = ws.getSparklines();
    if (spks.length) {
      const spkItems = spks.map(sp => {
        const vals = resolveSparklineData(ws, sp.dataRange);
        return sparklineToSvg(sp, vals);
      }).filter(Boolean);
      if (spkItems.length) sparklinesHtml = '\n<div class="sparklines" style="display:flex;gap:8px;margin:8px 0">' + spkItems.join('\n') + '</div>';
    }
  }

  const extraHtml = chartsHtml + imagesHtml + shapesHtml + wordArtHtml + mathHtml + formControlsHtml + sparklinesHtml;

  const wrapperClose = '</div>'; // close xl-sheet-wrapper
  if (options.fullDocument === false) return tableHtml + extraHtml + wrapperClose;

  const title = escapeXml(options.title ?? options.sheetName ?? 'Export');
  const css = `<style>
  * { box-sizing: border-box; }
  body { font-family: 'Segoe UI', Calibri, sans-serif; margin: 20px; background: #f5f6fa; }
  .xl-sheet-wrapper { position: relative; display: inline-block; }
  table { border-collapse: collapse; background: white; box-shadow: 0 1px 4px rgba(0,0,0,.1); }
  td { padding: 4px 8px; border: 1px solid #d4d4d4; vertical-align: bottom; }
  td[data-icon]::before { content: attr(data-icon); margin-right: 4px; }
  .xl-images { position: absolute; top: 0; left: 0; pointer-events: none; }
  .xl-images .xl-img { pointer-events: auto; position: absolute; z-index: 2; }
  .xl-form-controls { position: absolute; top: 0; left: 0; pointer-events: none; }
  .xl-form-controls .xl-fc { pointer-events: auto; z-index: 3; }
  a { color: #0563C1; }
</style>`;

  const positionScript = `<script>
(function(){
  document.querySelectorAll('.xl-sheet-wrapper').forEach(function(wrapper){
    var table = wrapper.querySelector('table');
    if (!table) return;
    function cellPos(row, col) {
      var tr = table.rows[row]; if (!tr) return {x:0,y:0};
      var td = tr.cells[col]; if (!td) return {x:0,y:0};
      return {x: td.offsetLeft, y: td.offsetTop};
    }
    wrapper.querySelectorAll('[data-from-col][data-from-row]').forEach(function(el){
      var c = parseInt(el.getAttribute('data-from-col'),10);
      var r = parseInt(el.getAttribute('data-from-row'),10);
      var p = cellPos(r, c);
      el.style.left = p.x + 'px';
      el.style.top = p.y + 'px';
    });
  });
})();
</script>`;

  return `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>${title}</title>
${css}
</head>
<body>
${tableHtml}${extraHtml}${wrapperClose}
${positionScript}
</body>
</html>`;
}

/* ─── Multi-sheet workbook export ──────────────────────────────────────────── */

/**
 * Export an entire workbook as a multi-sheet HTML document with tab navigation.
 */
export function workbookToHtml(wb: Workbook, options: WorkbookHtmlExportOptions = {}): string {
  const sheets = wb.getSheets();
  const names = wb.getSheetNames();
  const selected = options.sheets ?? names;
  const includeTabs = options.includeTabs !== false;

  const sheetHtmls: { name: string; html: string }[] = [];
  for (let i = 0; i < sheets.length; i++) {
    if (!selected.includes(names[i])) continue;
    if (sheets[i]._isChartSheet || sheets[i]._isDialogSheet) continue;
    const html = worksheetToHtml(sheets[i], { ...options, fullDocument: false, sheetName: names[i] });
    sheetHtmls.push({ name: names[i], html });
  }

  if (sheetHtmls.length === 1 && !includeTabs) {
    return worksheetToHtml(sheets[0], options);
  }

  const title = escapeXml(options.title ?? 'Workbook Export');
  const tabs = sheetHtmls.map((s, i) =>
    `<button class="tab${i === 0 ? ' active' : ''}" onclick="switchTab(${i})">${escapeXml(s.name)}</button>`
  ).join('');

  const panels = sheetHtmls.map((s, i) =>
    `<div class="panel${i === 0 ? ' active' : ''}" id="panel-${i}">${s.html}</div>`
  ).join('\n');

  return `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>${title}</title>
<style>
  * { box-sizing: border-box; }
  body { font-family: 'Segoe UI', Calibri, sans-serif; margin: 0; background: #f5f6fa; }
  .tab-bar { display: flex; background: #2b579a; padding: 0 16px; gap: 2px; position: sticky; top: 0; z-index: 10; }
  .tab { padding: 8px 20px; border: none; background: rgba(255,255,255,.15); color: white; cursor: pointer;
         font-size: 13px; border-radius: 4px 4px 0 0; margin-top: 4px; transition: background .15s; }
  .tab:hover { background: rgba(255,255,255,.3); }
  .tab.active { background: white; color: #2b579a; font-weight: 600; }
  .panel { display: none; padding: 20px; overflow: auto; }
  .panel.active { display: block; }
  .xl-sheet-wrapper { position: relative; display: inline-block; }
  table { border-collapse: collapse; background: white; box-shadow: 0 1px 4px rgba(0,0,0,.1); }
  td { padding: 4px 8px; border: 1px solid #d4d4d4; vertical-align: bottom; }
  td[data-icon]::before { content: attr(data-icon); margin-right: 4px; }
  .xl-images { position: absolute; top: 0; left: 0; pointer-events: none; }
  .xl-images .xl-img { pointer-events: auto; position: absolute; z-index: 2; }
  .xl-form-controls { position: absolute; top: 0; left: 0; pointer-events: none; }
  .xl-form-controls .xl-fc { pointer-events: auto; z-index: 3; }
  a { color: #0563C1; }
</style>
</head>
<body>
${includeTabs ? `<div class="tab-bar">${tabs}</div>` : ''}
${panels}
<script>
function switchTab(idx) {
  document.querySelectorAll('.tab').forEach((t,i) => t.classList.toggle('active', i===idx));
  document.querySelectorAll('.panel').forEach((p,i) => p.classList.toggle('active', i===idx));
}
(function(){
  document.querySelectorAll('.xl-sheet-wrapper').forEach(function(wrapper){
    var table = wrapper.querySelector('table');
    if (!table) return;
    function cellPos(row, col) {
      var tr = table.rows[row]; if (!tr) return {x:0,y:0};
      var td = tr.cells[col]; if (!td) return {x:0,y:0};
      return {x: td.offsetLeft, y: td.offsetTop};
    }
    wrapper.querySelectorAll('[data-from-col][data-from-row]').forEach(function(el){
      var c = parseInt(el.getAttribute('data-from-col'),10);
      var r = parseInt(el.getAttribute('data-from-row'),10);
      var p = cellPos(r, c);
      el.style.left = p.x + 'px';
      el.style.top = p.y + 'px';
    });
  });
})();
</script>
</body>
</html>`;
}
