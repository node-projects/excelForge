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
  ConditionalFormat, Chart, ChartSeries, Sparkline, MathElement, MathEquation, Image, CellImage, FormControl,
  Shape, WordArt, ChartPosition,
} from '../core/types.js';
import { escapeXml, colIndexToLetter } from '../utils/helpers.js';
import { FormulaEngine } from './FormulaEngine.js';

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
  /** Evaluate formulas before export so calculated cells have values (default false) */
  evaluateFormulas?: boolean;
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
  const maxIdx = values.indexOf(Math.max(...values));
  const minIdx = values.indexOf(Math.min(...values));

  if (sparkline.type === 'bar' || sparkline.type === 'stacked') {
    const bw = W / values.length;
    const bars = values.map((v, i) => {
      const barH = sparkline.type === 'stacked'
        ? (v >= 0 ? H / 2 : H / 2)
        : Math.max(1, ((v - min) / range) * H);
      const y = sparkline.type === 'stacked'
        ? (v >= 0 ? H / 2 - barH : H / 2)
        : H - barH;
      let fill = color;
      if (v < 0 && sparkline.negativeColor) fill = colorToCSS(sparkline.negativeColor);
      else if (sparkline.showHigh && i === maxIdx && sparkline.highColor) fill = colorToCSS(sparkline.highColor);
      else if (sparkline.showLow && i === minIdx && sparkline.lowColor) fill = colorToCSS(sparkline.lowColor);
      else if (sparkline.showFirst && i === 0 && sparkline.firstColor) fill = colorToCSS(sparkline.firstColor);
      else if (sparkline.showLast && i === values.length - 1 && sparkline.lastColor) fill = colorToCSS(sparkline.lastColor);
      return `<rect x="${i * bw + bw * 0.1}" y="${y}" width="${bw * 0.8}" height="${barH}" fill="${fill}" rx="1"/>`;
    }).join('');
    return `<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}" viewBox="0 0 ${W} ${H}" style="display:inline-block;vertical-align:middle">${bars}</svg>`;
  }

  // Line sparkline
  const strokeW = sparkline.lineWidth ?? 1.5;
  const pts = values.map((v, i) => `${(i / (values.length - 1 || 1)) * W},${H - ((v - min) / range) * H}`).join(' ');
  let markers = '';
  if (sparkline.showMarkers) {
    markers = values.map((v, i) => {
      const x = (i / (values.length - 1 || 1)) * W;
      const y = H - ((v - min) / range) * H;
      return `<circle cx="${x}" cy="${y}" r="1.5" fill="${colorToCSS(sparkline.markersColor) || color}"/>`;
    }).join('');
  }
  // Special point markers
  const specialMarkers: string[] = [];
  const addMarker = (idx: number, clr: string) => {
    const x = (idx / (values.length - 1 || 1)) * W;
    const y = H - ((values[idx] - min) / range) * H;
    specialMarkers.push(`<circle cx="${x}" cy="${y}" r="2.5" fill="${clr}" stroke="white" stroke-width="0.5"/>`);
  };
  if (sparkline.showHigh && sparkline.highColor) addMarker(maxIdx, colorToCSS(sparkline.highColor));
  if (sparkline.showLow && sparkline.lowColor) addMarker(minIdx, colorToCSS(sparkline.lowColor));
  if (sparkline.showFirst && sparkline.firstColor) addMarker(0, colorToCSS(sparkline.firstColor));
  if (sparkline.showLast && sparkline.lastColor) addMarker(values.length - 1, colorToCSS(sparkline.lastColor));
  // Negative markers
  if (sparkline.showNegative && sparkline.negativeColor) {
    values.forEach((v, i) => {
      if (v < 0) {
        const x = (i / (values.length - 1 || 1)) * W;
        const y = H - ((v - min) / range) * H;
        specialMarkers.push(`<circle cx="${x}" cy="${y}" r="2" fill="${colorToCSS(sparkline.negativeColor!)}"/>`);
      }
    });
  }
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}" viewBox="0 0 ${W} ${H}" style="display:inline-block;vertical-align:middle"><polyline points="${pts}" fill="none" stroke="${color}" stroke-width="${strokeW}"/>${markers}${specialMarkers.join('')}</svg>`;
}

/* ─── Chart SVG rendering ──────────────────────────────────────────────────── */

const CHART_PALETTE = ['#4472C4','#ED7D31','#A5A5A5','#FFC000','#5B9BD5','#70AD47','#264478','#9B57A0','#636363','#EB7E3A'];

function resolveChartSeriesData(ws: Worksheet, ref: string): (number | null)[] {
  const vals: (number | null)[] = [];
  const part = ref.includes('!') ? ref.split('!')[1] : ref;
  const clean = part.replace(/\$/g, '');
  const m = clean.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
  if (!m) return vals;
  const c1 = colLetterToIdx(m[1]), r1 = parseInt(m[2], 10);
  const c2 = colLetterToIdx(m[3]), r2 = parseInt(m[4], 10);
  for (let r = r1; r <= r2; r++) {
    for (let c = c1; c <= c2; c++) {
      const cell = ws.getCell(r, c);
      vals.push(typeof cell.value === 'number' ? cell.value : null);
    }
  }
  return vals;
}

function resolveChartCategories(ws: Worksheet, ref: string): string[] {
  const cats: string[] = [];
  const part = ref.includes('!') ? ref.split('!')[1] : ref;
  const clean = part.replace(/\$/g, '');
  const m = clean.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
  if (!m) return cats;
  const c1 = colLetterToIdx(m[1]), r1 = parseInt(m[2], 10);
  const c2 = colLetterToIdx(m[3]), r2 = parseInt(m[4], 10);
  for (let r = r1; r <= r2; r++) {
    for (let c = c1; c <= c2; c++) {
      const cell = ws.getCell(r, c);
      cats.push(cell.value != null ? String(cell.value) : '');
    }
  }
  return cats;
}

function chartToSvg(chart: Chart, ws: Worksheet): string {
  const W = 520, H = 340;
  const PAD_T = 45, PAD_B = 55, PAD_L = 65, PAD_R = 20;
  const plotW = W - PAD_L - PAD_R, plotH = H - PAD_T - PAD_B;
  const title = chart.title ?? '';
  const type = chart.type;

  // Resolve all series data
  const allSeries: { name: string; values: number[]; color: string }[] = [];
  let categories: string[] = [];
  for (let si = 0; si < chart.series.length; si++) {
    const s = chart.series[si];
    const rawVals = resolveChartSeriesData(ws, s.values);
    const vals = rawVals.map(v => v ?? 0);
    if (s.categories && !categories.length) categories = resolveChartCategories(ws, s.categories);
    allSeries.push({
      name: s.name ?? `Series ${si + 1}`,
      values: vals,
      color: s.color ? colorToCSS(s.color) : CHART_PALETTE[si % CHART_PALETTE.length],
    });
  }

  if (!allSeries.length || !allSeries[0].values.length) {
    return `<div style="display:inline-block;width:${W}px;height:${H}px;border:1px solid #ccc;background:#f9f9f9;text-align:center;line-height:${H}px;color:#666;font-size:14px" data-chart-type="${type}">[Chart: ${escapeXml(title || type)} — no data]</div>`;
  }

  const numCats = Math.max(...allSeries.map(s => s.values.length));
  if (!categories.length) categories = Array.from({ length: numCats }, (_, i) => String(i + 1));

  // Pie / Doughnut
  if (type === 'pie' || type === 'doughnut') return pieChartSvg(chart, allSeries, categories, W, H);

  // Radar
  if (type === 'radar' || type === 'radarFilled') return radarChartSvg(chart, allSeries, categories, W, H);

  // Scatter / Bubble
  if (type === 'scatter' || type === 'scatterSmooth' || type === 'bubble') return scatterChartSvg(chart, allSeries, ws, W, H);

  // Bar/Column/Line/Area — shared axis charts
  const allVals = allSeries.flatMap(s => s.values);
  let yMin = chart.yAxis?.min ?? Math.min(0, ...allVals);
  let yMax = chart.yAxis?.max ?? Math.max(...allVals);
  if (yMax === yMin) yMax = yMin + 1;
  const yRange = yMax - yMin;

  // Y-axis tick lines
  const numTicks = 5;
  const gridLines: string[] = [];
  const yLabels: string[] = [];
  for (let i = 0; i <= numTicks; i++) {
    const val = yMin + (yRange * i / numTicks);
    const y = PAD_T + plotH - (plotH * (val - yMin) / yRange);
    gridLines.push(`<line x1="${PAD_L}" y1="${y}" x2="${W - PAD_R}" y2="${y}" stroke="#e0e0e0" stroke-width="1"/>`);
    const label = Math.abs(val) >= 1000 ? (val / 1000).toFixed(1) + 'k' : Number.isInteger(val) ? String(val) : val.toFixed(1);
    yLabels.push(`<text x="${PAD_L - 8}" y="${y + 4}" text-anchor="end" font-size="10" fill="#666">${label}</text>`);
  }

  // Zero line
  const zeroY = PAD_T + plotH - (plotH * (0 - yMin) / yRange);

  let dataSvg = '';
  const isBar = type === 'bar' || type === 'barStacked' || type === 'barStacked100';
  const isColumn = type === 'column' || type === 'columnStacked' || type === 'columnStacked100';
  const isArea = type === 'area' || type === 'areaStacked';
  const isLine = type === 'line' || type === 'lineStacked' || type === 'lineMarker';
  const isStacked = type.includes('Stacked');
  const is100 = type.includes('100');

  if (isBar) {
    // Horizontal bars
    const barGroupH = plotH / numCats;
    const barH = isStacked ? barGroupH * 0.7 : (barGroupH * 0.7) / allSeries.length;
    for (let ci = 0; ci < numCats; ci++) {
      if (isStacked) {
        let xAcc = 0;
        const total = is100 ? allSeries.reduce((s, ser) => s + Math.abs(ser.values[ci] ?? 0), 0) || 1 : 1;
        for (let si = 0; si < allSeries.length; si++) {
          let val = allSeries[si].values[ci] ?? 0;
          if (is100) val = (val / total) * (yMax - yMin);
          const barW = (Math.abs(val) / yRange) * plotW;
          const x = PAD_L + (xAcc / yRange) * plotW;
          const y = PAD_T + ci * barGroupH + barGroupH * 0.15;
          dataSvg += `<rect x="${x}" y="${y}" width="${barW}" height="${barH}" fill="${allSeries[si].color}" rx="2"><title>${allSeries[si].name}: ${allSeries[si].values[ci] ?? 0}</title></rect>`;
          xAcc += Math.abs(val);
        }
      } else {
        for (let si = 0; si < allSeries.length; si++) {
          const val = allSeries[si].values[ci] ?? 0;
          const barW = (Math.abs(val - yMin) / yRange) * plotW;
          const y = PAD_T + ci * barGroupH + barGroupH * 0.15 + si * barH;
          dataSvg += `<rect x="${PAD_L}" y="${y}" width="${barW}" height="${barH}" fill="${allSeries[si].color}" rx="2"><title>${allSeries[si].name}: ${val}</title></rect>`;
        }
      }
    }
    // Category labels on Y axis
    for (let ci = 0; ci < numCats; ci++) {
      const y = PAD_T + ci * barGroupH + barGroupH / 2 + 4;
      yLabels.push(`<text x="${PAD_L - 8}" y="${y}" text-anchor="end" font-size="10" fill="#666">${escapeXml(categories[ci] ?? '')}</text>`);
    }
  } else if (isColumn || (!isLine && !isArea)) {
    // Vertical columns (default)
    const groupW = plotW / numCats;
    const barW = isStacked ? groupW * 0.6 : (groupW * 0.6) / allSeries.length;
    for (let ci = 0; ci < numCats; ci++) {
      if (isStacked) {
        let yAcc = 0;
        const total = is100 ? allSeries.reduce((s, ser) => s + Math.abs(ser.values[ci] ?? 0), 0) || 1 : 1;
        for (let si = 0; si < allSeries.length; si++) {
          let val = allSeries[si].values[ci] ?? 0;
          if (is100) val = (val / total) * yRange;
          const bh = (Math.abs(val) / yRange) * plotH;
          const x = PAD_L + ci * groupW + groupW * 0.2;
          const y = zeroY - yAcc - bh;
          dataSvg += `<rect x="${x}" y="${y}" width="${barW}" height="${bh}" fill="${allSeries[si].color}" rx="2"><title>${allSeries[si].name}: ${allSeries[si].values[ci] ?? 0}</title></rect>`;
          yAcc += bh;
        }
      } else {
        for (let si = 0; si < allSeries.length; si++) {
          const val = allSeries[si].values[ci] ?? 0;
          const bh = (Math.abs(val - yMin) / yRange) * plotH;
          const x = PAD_L + ci * groupW + groupW * 0.2 + si * barW;
          const y = PAD_T + plotH - bh;
          dataSvg += `<rect x="${x}" y="${y}" width="${barW}" height="${bh}" fill="${allSeries[si].color}" rx="2"><title>${allSeries[si].name}: ${val}</title></rect>`;
        }
      }
    }
  }

  if (isArea) {
    for (let si = 0; si < allSeries.length; si++) {
      const pts: string[] = [];
      for (let ci = 0; ci < numCats; ci++) {
        const val = allSeries[si].values[ci] ?? 0;
        const x = PAD_L + (ci / (numCats - 1 || 1)) * plotW;
        const y = PAD_T + plotH - ((val - yMin) / yRange) * plotH;
        pts.push(`${x},${y}`);
      }
      const firstX = PAD_L, lastX = PAD_L + ((numCats - 1) / (numCats - 1 || 1)) * plotW;
      const areaPath = `M${firstX},${zeroY} L${pts.join(' L')} L${lastX},${zeroY} Z`;
      dataSvg += `<path d="${areaPath}" fill="${allSeries[si].color}" fill-opacity="0.3"/>`;
      dataSvg += `<polyline points="${pts.join(' ')}" fill="none" stroke="${allSeries[si].color}" stroke-width="2"/>`;
    }
  }

  if (isLine) {
    for (let si = 0; si < allSeries.length; si++) {
      const pts: string[] = [];
      for (let ci = 0; ci < numCats; ci++) {
        const val = allSeries[si].values[ci] ?? 0;
        const x = PAD_L + (ci / (numCats - 1 || 1)) * plotW;
        const y = PAD_T + plotH - ((val - yMin) / yRange) * plotH;
        pts.push(`${x},${y}`);
      }
      dataSvg += `<polyline points="${pts.join(' ')}" fill="none" stroke="${allSeries[si].color}" stroke-width="2.5"/>`;
      // Markers
      if (type === 'lineMarker' || allSeries[si].values.length <= 20) {
        for (let ci = 0; ci < numCats; ci++) {
          const val = allSeries[si].values[ci] ?? 0;
          const x = PAD_L + (ci / (numCats - 1 || 1)) * plotW;
          const y = PAD_T + plotH - ((val - yMin) / yRange) * plotH;
          dataSvg += `<circle cx="${x}" cy="${y}" r="3.5" fill="${allSeries[si].color}" stroke="white" stroke-width="1.5"><title>${allSeries[si].name}: ${val}</title></circle>`;
        }
      }
    }
  }

  // X-axis category labels (for column/line/area)
  let catLabels = '';
  if (!isBar) {
    const step = numCats > 20 ? Math.ceil(numCats / 15) : 1;
    for (let ci = 0; ci < numCats; ci += step) {
      const x = isColumn || (!isLine && !isArea)
        ? PAD_L + ci * (plotW / numCats) + (plotW / numCats) / 2
        : PAD_L + (ci / (numCats - 1 || 1)) * plotW;
      catLabels += `<text x="${x}" y="${PAD_T + plotH + 18}" text-anchor="middle" font-size="10" fill="#666" transform="rotate(-30 ${x} ${PAD_T + plotH + 18})">${escapeXml((categories[ci] ?? '').slice(0, 12))}</text>`;
    }
  }

  // Legend
  let legendSvg = '';
  if (chart.legend !== false && allSeries.length > 1) {
    const ly = H - 12;
    const totalWidth = allSeries.reduce((s, ser) => s + ser.name.length * 7 + 25, 0);
    let lx = (W - totalWidth) / 2;
    for (const ser of allSeries) {
      legendSvg += `<rect x="${lx}" y="${ly - 8}" width="12" height="12" rx="2" fill="${ser.color}"/>`;
      legendSvg += `<text x="${lx + 16}" y="${ly + 2}" font-size="10" fill="#444">${escapeXml(ser.name)}</text>`;
      lx += ser.name.length * 7 + 25;
    }
  }

  // Axis titles
  let axisTitles = '';
  if (chart.xAxis?.title) {
    axisTitles += `<text x="${W / 2}" y="${H - 2}" text-anchor="middle" font-size="11" fill="#444">${escapeXml(chart.xAxis.title)}</text>`;
  }
  if (chart.yAxis?.title) {
    axisTitles += `<text x="14" y="${PAD_T + plotH / 2}" text-anchor="middle" font-size="11" fill="#444" transform="rotate(-90 14 ${PAD_T + plotH / 2})">${escapeXml(chart.yAxis.title)}</text>`;
  }

  const titleSvg = title ? `<text x="${W / 2}" y="22" text-anchor="middle" font-size="14" font-weight="600" fill="#333">${escapeXml(title)}</text>` : '';
  const plotBorder = `<rect x="${PAD_L}" y="${PAD_T}" width="${plotW}" height="${plotH}" fill="none" stroke="#ccc" stroke-width="0.5"/>`;

  return `<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}" viewBox="0 0 ${W} ${H}" style="background:white;border:1px solid #e0e0e0;border-radius:6px;box-shadow:0 1px 4px rgba(0,0,0,.08);margin:4px">
${titleSvg}
${gridLines.join('\n')}
${plotBorder}
${dataSvg}
${catLabels}
${yLabels.join('\n')}
${legendSvg}
${axisTitles}
</svg>`;
}

function pieChartSvg(chart: Chart, allSeries: { name: string; values: number[]; color: string }[], categories: string[], W: number, H: number): string {
  const cx = W / 2, cy = H / 2 + 10;
  const outerR = Math.min(W, H) / 2 - 40;
  const innerR = chart.type === 'doughnut' ? outerR * 0.5 : 0;
  const vals = allSeries[0].values;
  const total = vals.reduce((s, v) => s + Math.abs(v), 0) || 1;
  const title = chart.title ?? '';

  let angle = -Math.PI / 2;
  const slices: string[] = [];
  const labels: string[] = [];
  for (let i = 0; i < vals.length; i++) {
    const pct = Math.abs(vals[i]) / total;
    const sweep = pct * Math.PI * 2;
    const midAngle = angle + sweep / 2;
    const large = sweep > Math.PI ? 1 : 0;

    const x1o = cx + outerR * Math.cos(angle);
    const y1o = cy + outerR * Math.sin(angle);
    const x2o = cx + outerR * Math.cos(angle + sweep);
    const y2o = cy + outerR * Math.sin(angle + sweep);

    let path: string;
    if (innerR > 0) {
      const x1i = cx + innerR * Math.cos(angle);
      const y1i = cy + innerR * Math.sin(angle);
      const x2i = cx + innerR * Math.cos(angle + sweep);
      const y2i = cy + innerR * Math.sin(angle + sweep);
      path = `M${x1o},${y1o} A${outerR},${outerR} 0 ${large} 1 ${x2o},${y2o} L${x2i},${y2i} A${innerR},${innerR} 0 ${large} 0 ${x1i},${y1i} Z`;
    } else {
      path = vals.length === 1
        ? `M${cx - outerR},${cy} A${outerR},${outerR} 0 1 1 ${cx + outerR},${cy} A${outerR},${outerR} 0 1 1 ${cx - outerR},${cy} Z`
        : `M${cx},${cy} L${x1o},${y1o} A${outerR},${outerR} 0 ${large} 1 ${x2o},${y2o} Z`;
    }

    const color = CHART_PALETTE[i % CHART_PALETTE.length];
    slices.push(`<path d="${path}" fill="${color}" stroke="white" stroke-width="1.5"><title>${escapeXml(categories[i] ?? '')}: ${vals[i]} (${(pct * 100).toFixed(1)}%)</title></path>`);

    // Label outside
    if (pct > 0.04) {
      const lr = outerR + 16;
      const lx = cx + lr * Math.cos(midAngle);
      const ly = cy + lr * Math.sin(midAngle);
      const anchor = midAngle > Math.PI / 2 && midAngle < Math.PI * 1.5 ? 'end' : 'start';
      labels.push(`<text x="${lx}" y="${ly + 4}" text-anchor="${anchor}" font-size="10" fill="#444">${escapeXml((categories[i] ?? '').slice(0, 10))} ${(pct * 100).toFixed(0)}%</text>`);
    }
    angle += sweep;
  }

  // Legend
  let legendSvg = '';
  if (chart.legend !== false) {
    const lx = W - 10;
    for (let i = 0; i < Math.min(vals.length, 10); i++) {
      const ly = 40 + i * 18;
      legendSvg += `<rect x="${lx - 80}" y="${ly - 8}" width="10" height="10" rx="2" fill="${CHART_PALETTE[i % CHART_PALETTE.length]}"/>`;
      legendSvg += `<text x="${lx - 65}" y="${ly + 2}" font-size="10" fill="#444">${escapeXml((categories[i] ?? '').slice(0, 10))}</text>`;
    }
  }

  const titleSvg = title ? `<text x="${W / 2}" y="22" text-anchor="middle" font-size="14" font-weight="600" fill="#333">${escapeXml(title)}</text>` : '';

  return `<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}" viewBox="0 0 ${W} ${H}" style="background:white;border:1px solid #e0e0e0;border-radius:6px;box-shadow:0 1px 4px rgba(0,0,0,.08);margin:4px">
${titleSvg}
${slices.join('\n')}
${labels.join('\n')}
${legendSvg}
</svg>`;
}

function radarChartSvg(chart: Chart, allSeries: { name: string; values: number[]; color: string }[], categories: string[], W: number, H: number): string {
  const cx = W / 2, cy = H / 2 + 10;
  const R = Math.min(W, H) / 2 - 50;
  const n = categories.length || 1;
  const isFilled = chart.type === 'radarFilled';

  const allVals = allSeries.flatMap(s => s.values);
  const maxVal = Math.max(...allVals, 1);

  // Grid rings
  const rings: string[] = [];
  for (let ring = 1; ring <= 4; ring++) {
    const r = R * ring / 4;
    const pts = Array.from({ length: n }, (_, i) => {
      const a = -Math.PI / 2 + (2 * Math.PI * i / n);
      return `${cx + r * Math.cos(a)},${cy + r * Math.sin(a)}`;
    });
    rings.push(`<polygon points="${pts.join(' ')}" fill="none" stroke="#e0e0e0" stroke-width="0.5"/>`);
  }

  // Axes
  const axes: string[] = [];
  for (let i = 0; i < n; i++) {
    const a = -Math.PI / 2 + (2 * Math.PI * i / n);
    const x = cx + R * Math.cos(a);
    const y = cy + R * Math.sin(a);
    axes.push(`<line x1="${cx}" y1="${cy}" x2="${x}" y2="${y}" stroke="#ccc" stroke-width="0.5"/>`);
    const lx = cx + (R + 14) * Math.cos(a);
    const ly = cy + (R + 14) * Math.sin(a);
    axes.push(`<text x="${lx}" y="${ly + 4}" text-anchor="middle" font-size="9" fill="#666">${escapeXml((categories[i] ?? '').slice(0, 8))}</text>`);
  }

  // Series polygons
  const seriesSvg: string[] = [];
  for (const ser of allSeries) {
    const pts = ser.values.map((v, i) => {
      const a = -Math.PI / 2 + (2 * Math.PI * i / n);
      const r = (v / maxVal) * R;
      return `${cx + r * Math.cos(a)},${cy + r * Math.sin(a)}`;
    });
    if (isFilled) {
      seriesSvg.push(`<polygon points="${pts.join(' ')}" fill="${ser.color}" fill-opacity="0.2" stroke="${ser.color}" stroke-width="2"/>`);
    } else {
      seriesSvg.push(`<polygon points="${pts.join(' ')}" fill="none" stroke="${ser.color}" stroke-width="2"/>`);
    }
    // Dots
    ser.values.forEach((v, i) => {
      const a = -Math.PI / 2 + (2 * Math.PI * i / n);
      const r = (v / maxVal) * R;
      seriesSvg.push(`<circle cx="${cx + r * Math.cos(a)}" cy="${cy + r * Math.sin(a)}" r="3" fill="${ser.color}" stroke="white" stroke-width="1"><title>${ser.name}: ${v}</title></circle>`);
    });
  }

  const titleSvg = chart.title ? `<text x="${W / 2}" y="20" text-anchor="middle" font-size="14" font-weight="600" fill="#333">${escapeXml(chart.title)}</text>` : '';

  return `<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}" viewBox="0 0 ${W} ${H}" style="background:white;border:1px solid #e0e0e0;border-radius:6px;box-shadow:0 1px 4px rgba(0,0,0,.08);margin:4px">
${titleSvg}
${rings.join('\n')}
${axes.join('\n')}
${seriesSvg.join('\n')}
</svg>`;
}

function scatterChartSvg(chart: Chart, allSeries: { name: string; values: number[]; color: string }[], ws: Worksheet, W: number, H: number): string {
  const PAD_T = 45, PAD_B = 40, PAD_L = 60, PAD_R = 20;
  const plotW = W - PAD_L - PAD_R, plotH = H - PAD_T - PAD_B;

  // For scatter, we need X from categories and Y from values
  const catSeries = chart.series[0]?.categories ? resolveChartSeriesData(ws, chart.series[0].categories) : [];
  const points: { x: number; y: number; name: string; color: string }[] = [];
  for (const ser of allSeries) {
    for (let i = 0; i < ser.values.length; i++) {
      points.push({ x: catSeries[i] ?? i, y: ser.values[i], name: ser.name, color: ser.color });
    }
  }

  if (!points.length) return '';

  const xMin = Math.min(...points.map(p => p.x)), xMax = Math.max(...points.map(p => p.x));
  const yMin = Math.min(0, ...points.map(p => p.y)), yMax = Math.max(...points.map(p => p.y));
  const xRange = xMax - xMin || 1, yRange = yMax - yMin || 1;

  const gridLines: string[] = [];
  for (let i = 0; i <= 4; i++) {
    const y = PAD_T + plotH - (plotH * i / 4);
    gridLines.push(`<line x1="${PAD_L}" y1="${y}" x2="${W - PAD_R}" y2="${y}" stroke="#e0e0e0" stroke-width="0.5"/>`);
    const val = yMin + yRange * i / 4;
    gridLines.push(`<text x="${PAD_L - 8}" y="${y + 4}" text-anchor="end" font-size="10" fill="#666">${val.toFixed(0)}</text>`);
  }

  const dotsSvg = points.map(p => {
    const x = PAD_L + ((p.x - xMin) / xRange) * plotW;
    const y = PAD_T + plotH - ((p.y - yMin) / yRange) * plotH;
    const r = chart.type === 'bubble' ? Math.max(4, Math.min(20, Math.sqrt(Math.abs(p.y)) * 2)) : 4;
    return `<circle cx="${x}" cy="${y}" r="${r}" fill="${p.color}" fill-opacity="${chart.type === 'bubble' ? '0.6' : '1'}" stroke="white" stroke-width="1"><title>${p.name}: (${p.x}, ${p.y})</title></circle>`;
  }).join('\n');

  // Smooth line through points if scatterSmooth
  let lineSvg = '';
  if (chart.type === 'scatterSmooth') {
    for (const ser of allSeries) {
      const pts = ser.values.map((v, i) => {
        const xVal = catSeries[i] ?? i;
        const x = PAD_L + ((xVal - xMin) / xRange) * plotW;
        const y = PAD_T + plotH - ((v - yMin) / yRange) * plotH;
        return `${x},${y}`;
      });
      lineSvg += `<polyline points="${pts.join(' ')}" fill="none" stroke="${ser.color}" stroke-width="2"/>`;
    }
  }

  const titleSvg = chart.title ? `<text x="${W / 2}" y="22" text-anchor="middle" font-size="14" font-weight="600" fill="#333">${escapeXml(chart.title)}</text>` : '';

  return `<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}" viewBox="0 0 ${W} ${H}" style="background:white;border:1px solid #e0e0e0;border-radius:6px;box-shadow:0 1px 4px rgba(0,0,0,.08);margin:4px">
${titleSvg}
${gridLines.join('\n')}
<rect x="${PAD_L}" y="${PAD_T}" width="${plotW}" height="${plotH}" fill="none" stroke="#ccc" stroke-width="0.5"/>
${lineSvg}
${dotsSvg}
</svg>`;
}

function chartToHtml(chart: Chart, ws: Worksheet): string {
  return `<div class="xl-chart" data-from-col="${chart.from.col}" data-from-row="${chart.from.row}" data-to-col="${chart.to.col}" data-to-row="${chart.to.row}" style="position:absolute;z-index:4">${chartToSvg(chart, ws)}</div>`;
}

/* ─── Shape rendering (SVG with positioning) ──────────────────────────────── */

function shapeSvgPath(type: string, w: number, h: number): string {
  switch (type) {
    case 'rect': return `<rect x="0" y="0" width="${w}" height="${h}"/>`;
    case 'roundRect': return `<rect x="0" y="0" width="${w}" height="${h}" rx="${Math.min(w, h) * 0.15}"/>`;
    case 'ellipse': return `<ellipse cx="${w / 2}" cy="${h / 2}" rx="${w / 2}" ry="${h / 2}"/>`;
    case 'triangle': return `<polygon points="${w / 2},0 ${w},${h} 0,${h}"/>`;
    case 'diamond': return `<polygon points="${w / 2},0 ${w},${h / 2} ${w / 2},${h} 0,${h / 2}"/>`;
    case 'pentagon': {
      const pts = Array.from({ length: 5 }, (_, i) => {
        const a = -Math.PI / 2 + (2 * Math.PI * i / 5);
        return `${w / 2 + w / 2 * Math.cos(a)},${h / 2 + h / 2 * Math.sin(a)}`;
      });
      return `<polygon points="${pts.join(' ')}"/>`;
    }
    case 'hexagon': {
      const pts = Array.from({ length: 6 }, (_, i) => {
        const a = -Math.PI / 6 + (2 * Math.PI * i / 6);
        return `${w / 2 + w / 2 * Math.cos(a)},${h / 2 + h / 2 * Math.sin(a)}`;
      });
      return `<polygon points="${pts.join(' ')}"/>`;
    }
    case 'octagon': {
      const pts = Array.from({ length: 8 }, (_, i) => {
        const a = -Math.PI / 8 + (2 * Math.PI * i / 8);
        return `${w / 2 + w / 2 * Math.cos(a)},${h / 2 + h / 2 * Math.sin(a)}`;
      });
      return `<polygon points="${pts.join(' ')}"/>`;
    }
    case 'star5': case 'star6': {
      const n = type === 'star5' ? 5 : 6;
      const pts: string[] = [];
      for (let i = 0; i < n * 2; i++) {
        const a = -Math.PI / 2 + (Math.PI * i / n);
        const r = i % 2 === 0 ? Math.min(w, h) / 2 : Math.min(w, h) / 4.5;
        pts.push(`${w / 2 + r * Math.cos(a)},${h / 2 + r * Math.sin(a)}`);
      }
      return `<polygon points="${pts.join(' ')}"/>`;
    }
    case 'rightArrow':
      return `<polygon points="0,${h * 0.25} ${w * 0.65},${h * 0.25} ${w * 0.65},0 ${w},${h / 2} ${w * 0.65},${h} ${w * 0.65},${h * 0.75} 0,${h * 0.75}"/>`;
    case 'leftArrow':
      return `<polygon points="${w},${h * 0.25} ${w * 0.35},${h * 0.25} ${w * 0.35},0 0,${h / 2} ${w * 0.35},${h} ${w * 0.35},${h * 0.75} ${w},${h * 0.75}"/>`;
    case 'upArrow':
      return `<polygon points="${w * 0.25},${h} ${w * 0.25},${h * 0.35} 0,${h * 0.35} ${w / 2},0 ${w},${h * 0.35} ${w * 0.75},${h * 0.35} ${w * 0.75},${h}"/>`;
    case 'downArrow':
      return `<polygon points="${w * 0.25},0 ${w * 0.75},0 ${w * 0.75},${h * 0.65} ${w},${h * 0.65} ${w / 2},${h} 0,${h * 0.65} ${w * 0.25},${h * 0.65}"/>`;
    case 'heart': {
      const hw = w / 2, hh = h;
      return `<path d="M${hw},${hh * 0.35} C${hw},${hh * 0.15} ${hw * 0.5},0 ${hw * 0.25},0 C0,0 0,${hh * 0.35} 0,${hh * 0.35} C0,${hh * 0.6} ${hw * 0.5},${hh * 0.8} ${hw},${hh} C${hw * 1.5},${hh * 0.8} ${w},${hh * 0.6} ${w},${hh * 0.35} C${w},${hh * 0.35} ${w},0 ${hw * 1.75},0 C${hw * 1.5},0 ${hw},${hh * 0.15} ${hw},${hh * 0.35} Z"/>`;
    }
    case 'lightningBolt':
      return `<polygon points="${w * 0.55},0 ${w * 0.2},${h * 0.45} ${w * 0.45},${h * 0.45} ${w * 0.15},${h} ${w * 0.8},${h * 0.4} ${w * 0.55},${h * 0.4}"/>`;
    case 'sun': return `<circle cx="${w / 2}" cy="${h / 2}" r="${Math.min(w, h) * 0.3}"/>`;
    case 'moon':
      return `<path d="M${w * 0.6},${h * 0.1} A${w * 0.4},${h * 0.4} 0 1 0 ${w * 0.6},${h * 0.9} A${w * 0.3},${h * 0.35} 0 1 1 ${w * 0.6},${h * 0.1} Z"/>`;
    case 'smileyFace':
      return `<circle cx="${w / 2}" cy="${h / 2}" r="${Math.min(w, h) * 0.45}"/>`
        + `<circle cx="${w * 0.35}" cy="${h * 0.38}" r="${Math.min(w, h) * 0.04}" fill="white"/>`
        + `<circle cx="${w * 0.65}" cy="${h * 0.38}" r="${Math.min(w, h) * 0.04}" fill="white"/>`
        + `<path d="M${w * 0.3},${h * 0.58} Q${w / 2},${h * 0.78} ${w * 0.7},${h * 0.58}" fill="none" stroke="white" stroke-width="2"/>`;
    case 'cloud':
      return `<ellipse cx="${w * 0.35}" cy="${h * 0.55}" rx="${w * 0.25}" ry="${h * 0.25}"/>`
        + `<ellipse cx="${w * 0.55}" cy="${h * 0.35}" rx="${w * 0.22}" ry="${h * 0.22}"/>`
        + `<ellipse cx="${w * 0.7}" cy="${h * 0.5}" rx="${w * 0.2}" ry="${h * 0.2}"/>`
        + `<rect x="${w * 0.15}" y="${h * 0.5}" width="${w * 0.7}" height="${h * 0.25}" rx="4"/>`;
    case 'callout1':
      return `<path d="M0,0 L${w},0 L${w},${h * 0.7} L${w * 0.4},${h * 0.7} L${w * 0.25},${h} L${w * 0.3},${h * 0.7} L0,${h * 0.7} Z"/>`;
    case 'callout2':
      return `<path d="M${w * 0.1},0 L${w * 0.9},0 Q${w},0 ${w},${h * 0.1} L${w},${h * 0.6} Q${w},${h * 0.7} ${w * 0.9},${h * 0.7} L${w * 0.4},${h * 0.7} L${w * 0.25},${h} L${w * 0.3},${h * 0.7} L${w * 0.1},${h * 0.7} Q0,${h * 0.7} 0,${h * 0.6} L0,${h * 0.1} Q0,0 ${w * 0.1},0 Z"/>`;
    case 'flowChartProcess':
      return `<rect x="0" y="0" width="${w}" height="${h}" rx="2"/>`;
    case 'flowChartDecision':
      return `<polygon points="${w / 2},0 ${w},${h / 2} ${w / 2},${h} 0,${h / 2}"/>`;
    case 'flowChartTerminator':
      return `<rect x="0" y="0" width="${w}" height="${h}" rx="${h / 2}"/>`;
    case 'flowChartDocument':
      return `<path d="M0,0 L${w},0 L${w},${h * 0.8} Q${w * 0.75},${h * 0.65} ${w * 0.5},${h * 0.8} Q${w * 0.25},${h * 0.95} 0,${h * 0.8} Z"/>`;
    case 'line':
      return `<line x1="0" y1="${h / 2}" x2="${w}" y2="${h / 2}" stroke-width="2"/>`;
    case 'curvedConnector3':
      return `<path d="M0,${h / 2} C${w * 0.3},${h * 0.1} ${w * 0.7},${h * 0.9} ${w},${h / 2}" fill="none" stroke-width="2"/>`;
    default:
      return `<rect x="0" y="0" width="${w}" height="${h}" rx="4"/>`;
  }
}

function shapeToHtml(shape: Shape): string {
  const toHex = (c: string) => { let h = c.replace(/^#/, ''); if (h.length === 8) h = h.substring(2); return '#' + h; };
  const bg = shape.fillColor ? toHex(shape.fillColor) : '#4472C4';
  const border = shape.lineColor ? toHex(shape.lineColor) : '#2F5496';
  const lw = shape.lineWidth ?? 2;
  const w = 160, h = 80;
  const rotation = shape.rotation ? ` transform="rotate(${shape.rotation} ${w / 2} ${h / 2})"` : '';

  const textEl = shape.text ? `<text x="${w / 2}" y="${h / 2 + 5}" text-anchor="middle" fill="white" font-size="13" font-weight="600"${shape.font?.name ? ` font-family="'${escapeXml(shape.font.name)}'"` : ''}>${escapeXml(shape.text)}</text>` : '';

  const isLine = shape.type === 'line' || shape.type === 'curvedConnector3';
  const fillAttr = isLine ? `fill="none" stroke="${bg}"` : `fill="${bg}" stroke="${border}" stroke-width="${lw}"`;

  return `<div class="xl-shape" data-from-col="${shape.from.col}" data-from-row="${shape.from.row}" data-to-col="${shape.to.col}" data-to-row="${shape.to.row}" style="position:absolute;z-index:2">
<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}" viewBox="0 0 ${w} ${h}"${rotation}>
<g ${fillAttr}>${shapeSvgPath(shape.type, w, h)}</g>
${textEl}
</svg></div>`;
}

/* ─── WordArt rendering (positioned) ───────────────────────────────────────── */

function wordArtToHtml(wa: WordArt): string {
  const toHex = (c: string) => { let h = c.replace(/^#/, ''); if (h.length === 8) h = h.substring(2); return '#' + h; };
  const color = wa.fillColor ? toHex(wa.fillColor) : '#333';
  const outline = wa.outlineColor ? toHex(wa.outlineColor) : '';
  const family = wa.font?.name ?? 'Impact';
  const size = wa.font?.size ?? 36;
  const bold = wa.font?.bold !== false ? 'font-weight:bold;' : '';
  const italic = wa.font?.italic ? 'font-style:italic;' : '';
  const textStroke = outline ? `-webkit-text-stroke:1px ${outline};paint-order:stroke fill;` : '';
  const presetStyle = wordArtPresetCSS(wa.preset);
  return `<div class="xl-wordart" data-from-col="${wa.from.col}" data-from-row="${wa.from.row}" data-to-col="${wa.to.col}" data-to-row="${wa.to.row}" style="position:absolute;z-index:2;font-family:'${escapeXml(family)}',sans-serif;font-size:${size}px;${bold}${italic}color:${color};${textStroke}text-shadow:2px 2px 4px rgba(0,0,0,.3);${presetStyle}padding:8px 16px;white-space:nowrap;line-height:1.2">${escapeXml(wa.text)}</div>`;
}

function wordArtPresetCSS(preset?: string): string {
  if (!preset || preset === 'textPlain') return '';
  const presets: Record<string, string> = {
    textArchUp: 'letter-spacing:4px;',
    textArchDown: 'letter-spacing:4px;transform:scaleY(-1);',
    textCircle: 'letter-spacing:6px;',
    textWave1: 'letter-spacing:2px;font-style:italic;transform:skewX(-5deg);',
    textWave2: 'letter-spacing:2px;font-style:italic;transform:skewX(5deg);',
    textInflate: 'letter-spacing:3px;transform:scaleY(1.3);',
    textDeflate: 'letter-spacing:1px;transform:scaleY(0.7);',
    textSlantUp: 'transform:perspective(300px) rotateY(-8deg) rotateX(3deg);',
    textSlantDown: 'transform:perspective(300px) rotateY(8deg) rotateX(-3deg);',
    textFadeUp: 'transform:perspective(200px) rotateX(-8deg);',
    textFadeDown: 'transform:perspective(200px) rotateX(8deg);',
    textFadeLeft: 'transform:perspective(200px) rotateY(8deg);',
    textFadeRight: 'transform:perspective(200px) rotateY(-8deg);',
    textCascadeUp: 'letter-spacing:3px;transform:rotate(-5deg) scaleX(1.1);',
    textCascadeDown: 'letter-spacing:3px;transform:rotate(5deg) scaleX(1.1);',
    textChevron: 'letter-spacing:5px;transform:scaleX(1.15);',
    textRingInside: 'letter-spacing:8px;',
    textRingOutside: 'letter-spacing:6px;',
    textStop: 'letter-spacing:2px;transform:scaleY(0.85) scaleX(1.1);',
  };
  return presets[preset] ?? '';
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

function mathEquationToMathML(eq: MathEquation): string {
  const size = eq.fontSize ?? 11;
  const font = eq.fontName ?? 'Cambria Math';
  return `<div class="xl-math" data-from-col="${eq.from.col}" data-from-row="${eq.from.row}" style="position:absolute;z-index:2;font-family:'${escapeXml(font)}',serif;font-size:${size}pt;padding:4px;background:white"><math xmlns="http://www.w3.org/1998/Math/MathML" display="block"><mrow>${eq.elements.map(mathElementToMathML).join('')}</mrow></math></div>`;
}

/* ─── Form control rendering (positioned with size from anchors) ───────────── */

function formControlToPositionedHtml(fc: FormControl): string {
  const fromCol = fc.from.col;
  const fromRow = fc.from.row;
  const toCol = fc.to?.col;
  const toRow = fc.to?.row;
  const toAttrs = toCol != null && toRow != null ? ` data-to-col="${toCol}" data-to-row="${toRow}"` : '';
  const linked = fc.linkedCell ? ` data-linked-cell="${escapeXml(fc.linkedCell)}"` : '';
  const inputRange = fc.inputRange ? ` data-input-range="${escapeXml(fc.inputRange)}"` : '';
  const macro = fc.macro ? ` data-macro="${escapeXml(fc.macro)}"` : '';
  let inner = '';
  switch (fc.type) {
    case 'button':
    case 'dialog':
      inner = `<button style="width:100%;height:100%;padding:4px 12px;font-size:13px;border:1px outset #ccc;background:linear-gradient(180deg,#f8f8f8,#e0e0e0);cursor:pointer;border-radius:3px;white-space:nowrap"${macro}>${escapeXml(fc.text ?? 'Button')}</button>`;
      break;
    case 'checkBox': {
      const checked = fc.checked === 'checked' ? ' checked' : '';
      const indeterminate = fc.checked === 'mixed' ? ' data-indeterminate="true"' : '';
      inner = `<label style="font-size:13px;display:inline-flex;align-items:center;gap:4px;width:100%;height:100%;padding:2px 4px;cursor:pointer"><input type="checkbox"${checked}${indeterminate}${linked}/> ${escapeXml(fc.text ?? 'Checkbox')}</label>`;
      break;
    }
    case 'optionButton': {
      const checked = fc.checked === 'checked' ? ' checked' : '';
      inner = `<label style="font-size:13px;display:inline-flex;align-items:center;gap:4px;width:100%;height:100%;padding:2px 4px;cursor:pointer"><input type="radio" name="group"${checked}${linked}/> ${escapeXml(fc.text ?? 'Option')}</label>`;
      break;
    }
    case 'comboBox': {
      const lines = fc.dropLines ?? 8;
      inner = `<select style="width:100%;height:100%;padding:2px 4px;font-size:13px;border:1px solid #aaa;background:white"${linked}${inputRange} size="1" data-drop-lines="${lines}"><option>${escapeXml(fc.text ?? 'Select...')}</option></select>`;
      break;
    }
    case 'listBox': {
      const size = fc.dropLines ?? 5;
      const sel = fc.selType ?? 'single';
      const multi = sel === 'multi' || sel === 'extend' ? ' multiple' : '';
      inner = `<select style="width:100%;height:100%;padding:2px;font-size:13px;border:1px solid #aaa;background:white"${linked}${inputRange} size="${size}"${multi}><option>${escapeXml(fc.text ?? 'Item')}</option></select>`;
      break;
    }
    case 'spinner': {
      const min = fc.min ?? 0;
      const max = fc.max ?? 100;
      const step = fc.inc ?? 1;
      const val = fc.val ?? min;
      inner = `<input type="number" value="${val}" min="${min}" max="${max}" step="${step}" style="width:100%;height:100%;padding:2px 4px;font-size:13px;border:1px solid #aaa"${linked}/>`;
      break;
    }
    case 'scrollBar': {
      const min = fc.min ?? 0;
      const max = fc.max ?? 100;
      const step = fc.inc ?? 1;
      const val = fc.val ?? min;
      inner = `<input type="range" value="${val}" min="${min}" max="${max}" step="${step}" style="width:100%;height:100%"${linked}/>`;
      break;
    }
    case 'label':
      inner = `<span style="font-size:13px;display:flex;align-items:center;width:100%;height:100%;padding:2px 4px">${escapeXml(fc.text ?? 'Label')}</span>`;
      break;
    case 'groupBox':
      inner = `<fieldset style="width:100%;height:100%;padding:8px;border:1px solid #999;font-size:13px;margin:0;box-sizing:border-box"><legend>${escapeXml(fc.text ?? 'Group')}</legend></fieldset>`;
      break;
    default:
      inner = `<span style="font-size:13px">[${escapeXml(fc.type)}]</span>`;
  }
  return `<div class="xl-fc" data-from-col="${fromCol}" data-from-row="${fromRow}"${toAttrs} style="position:absolute;overflow:hidden">${inner}</div>`;
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
  // Evaluate formulas if requested
  if (options.evaluateFormulas) {
    new FormulaEngine().calculateSheet(ws);
  }

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

      // Sparkline — resolve data from dataRange
      const sp = sparklineMap.get(key);
      if (sp) val += sparklineToSvg(sp, resolveSparklineData(ws, sp.dataRange));

      const attrStr = attrs.length ? ' ' + attrs.join(' ') : '';
      cells.push(`<td${attrStr}${iconAttr}>${val}</td>`);
    }
    rows.push(`<tr${rowStyle}>${cells.join('')}</tr>`);
  }

  const colGroup = colWidths.length ? `<colgroup>${colWidths.join('')}</colgroup>` : '';
  const tableHtml = `<div class="xl-sheet-wrapper" style="position:relative;display:inline-block"><table border="0" cellpadding="4" cellspacing="0">\n${colGroup}\n${rows.join('\n')}\n</table>`;

  // Charts — positioned overlays with SVG rendering
  let chartsHtml = '';
  if (options.includeCharts !== false) {
    const charts = ws.getCharts();
    if (charts.length) chartsHtml = '\n<div class="xl-charts">' + charts.map(ch => chartToHtml(ch, ws)).join('\n') + '</div>';
  }

  // Images — positioned as overlays on a wrapper container
  let imagesHtml = '';
  const images = ws.getImages?.();
  if (images?.length) imagesHtml = '\n<div class="xl-images">' + images.map(imageToPositionedHtml).join('\n') + '</div>';

  // Shapes — positioned overlays with SVG shapes
  let shapesHtml = '';
  const shapes = ws.getShapes?.();
  if (shapes?.length) shapesHtml = '\n<div class="xl-shapes">' + shapes.map(shapeToHtml).join('\n') + '</div>';

  // WordArt — positioned overlays
  let wordArtHtml = '';
  const wordArts = ws.getWordArt?.();
  if (wordArts?.length) wordArtHtml = '\n<div class="xl-wordarts">' + wordArts.map(wordArtToHtml).join('\n') + '</div>';

  // Math equations (MathML) — positioned overlays
  let mathHtml = '';
  const mathEqs = ws.getMathEquations?.();
  if (mathEqs?.length) mathHtml = '\n<div class="xl-math-equations">' + mathEqs.map(mathEquationToMathML).join('\n') + '</div>';

  // Form controls — positioned overlays with size from anchors
  let formControlsHtml = '';
  const fcs = ws.getFormControls?.();
  if (fcs?.length) formControlsHtml = '\n<div class="xl-form-controls">' + fcs.map(formControlToPositionedHtml).join('\n') + '</div>';

  const extraHtml = chartsHtml + imagesHtml + shapesHtml + wordArtHtml + mathHtml + formControlsHtml;

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
  .xl-charts { position: absolute; top: 0; left: 0; pointer-events: none; }
  .xl-charts .xl-chart { pointer-events: auto; }
  .xl-shapes { position: absolute; top: 0; left: 0; pointer-events: none; }
  .xl-shapes .xl-shape { pointer-events: auto; }
  .xl-wordarts { position: absolute; top: 0; left: 0; pointer-events: none; }
  .xl-wordarts .xl-wordart { pointer-events: auto; }
  .xl-math-equations { position: absolute; top: 0; left: 0; pointer-events: none; }
  .xl-math-equations .xl-math { pointer-events: auto; }
  .xl-form-controls { position: absolute; top: 0; left: 0; pointer-events: none; }
  .xl-form-controls .xl-fc { pointer-events: auto; z-index: 3; }
  a { color: #0563C1; }
</style>`;

  const positionScript = `<script>
(function(){
  document.querySelectorAll('.xl-sheet-wrapper').forEach(function(wrapper){
    var table = wrapper.querySelector('table');
    if (!table) return;
    var startRow = parseInt((table.rows[0] && table.rows[0].cells[0] || {}).getAttribute && table.rows[0].cells[0].getAttribute('data-cell') ? table.rows[0].cells[0].getAttribute('data-cell').replace(/[A-Z]+/,'') : '1', 10) || 1;
    var startCol = 1;
    var firstCell = table.rows[0] && table.rows[0].cells[0] ? table.rows[0].cells[0].getAttribute('data-cell') : null;
    if (firstCell) {
      var colStr = firstCell.replace(/[0-9]+/g,'');
      startCol = 0;
      for (var ci = 0; ci < colStr.length; ci++) startCol = startCol * 26 + (colStr.charCodeAt(ci) - 64);
    }
    function cellRect(sheetRow, sheetCol) {
      var ri = sheetRow - startRow;
      var colIdx = sheetCol - startCol;
      if (ri < 0) ri = 0;
      if (colIdx < 0) colIdx = 0;
      var tr = table.rows[ri];
      if (!tr) tr = table.rows[table.rows.length - 1] || table.rows[0];
      if (!tr) return {x:0, y:0, w:0, h:0};
      var td = tr.cells[colIdx];
      if (!td) td = tr.cells[tr.cells.length - 1] || tr.cells[0];
      if (!td) return {x:0, y:0, w:0, h:0};
      return {x: td.offsetLeft, y: td.offsetTop, w: td.offsetWidth, h: td.offsetHeight};
    }
    wrapper.querySelectorAll('[data-from-col][data-from-row]').forEach(function(el){
      var fc = parseInt(el.getAttribute('data-from-col'),10);
      var fr = parseInt(el.getAttribute('data-from-row'),10);
      var from = cellRect(fr, fc);
      el.style.left = from.x + 'px';
      el.style.top = from.y + 'px';
      var tc = el.getAttribute('data-to-col');
      var tr2 = el.getAttribute('data-to-row');
      if (tc !== null && tr2 !== null) {
        var to = cellRect(parseInt(tr2,10), parseInt(tc,10));
        var w = to.x + to.w - from.x;
        var h = to.y + to.h - from.y;
        if (w > 0) el.style.width = w + 'px';
        if (h > 0) el.style.height = h + 'px';
      }
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

  // Evaluate formulas across workbook if requested
  if (options.evaluateFormulas) {
    new FormulaEngine().calculateWorkbook(wb);
  }

  const sheetHtmls: { name: string; html: string }[] = [];
  for (let i = 0; i < sheets.length; i++) {
    if (!selected.includes(names[i])) continue;
    if (sheets[i]._isChartSheet) {
      // Render chart sheet as SVG chart
      const charts = sheets[i].getCharts();
      if (charts.length) {
        const chartHtml = chartToSvg(charts[0], sheets[i]);
        sheetHtmls.push({ name: names[i], html: `<div class="xl-sheet-wrapper" style="position:relative;display:inline-block">${chartHtml}</div>` });
      }
      continue;
    }
    if (sheets[i]._isDialogSheet) {
      // Render dialog sheet with form controls
      const fcs = sheets[i].getFormControls?.() ?? [];
      if (fcs.length) {
        const fcHtml = fcs.map(formControlToPositionedHtml).join('\n');
        sheetHtmls.push({ name: names[i], html: `<div class="xl-sheet-wrapper" style="position:relative;display:inline-block;min-width:400px;min-height:300px"><div class="xl-form-controls">${fcHtml}</div></div>` });
      } else {
        sheetHtmls.push({ name: names[i], html: '<div class="xl-sheet-wrapper"><p>Dialog Sheet</p></div>' });
      }
      continue;
    }
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
  .xl-charts { position: absolute; top: 0; left: 0; pointer-events: none; }
  .xl-charts .xl-chart { pointer-events: auto; }
  .xl-shapes { position: absolute; top: 0; left: 0; pointer-events: none; }
  .xl-shapes .xl-shape { pointer-events: auto; }
  .xl-wordarts { position: absolute; top: 0; left: 0; pointer-events: none; }
  .xl-wordarts .xl-wordart { pointer-events: auto; }
  .xl-math-equations { position: absolute; top: 0; left: 0; pointer-events: none; }
  .xl-math-equations .xl-math { pointer-events: auto; }
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
  // Re-run positioning after tab switch since hidden panels may have 0 dimensions
  setTimeout(positionOverlays, 50);
}
function positionOverlays() {
  document.querySelectorAll('.xl-sheet-wrapper').forEach(function(wrapper){
    var table = wrapper.querySelector('table');
    if (!table) return;
    var startRow = 1, startCol = 1;
    var firstCell = table.rows[0] && table.rows[0].cells[0] ? table.rows[0].cells[0].getAttribute('data-cell') : null;
    if (firstCell) {
      startRow = parseInt(firstCell.replace(/[A-Z]+/,''), 10) || 1;
      var colStr = firstCell.replace(/[0-9]+/g,'');
      startCol = 0;
      for (var ci = 0; ci < colStr.length; ci++) startCol = startCol * 26 + (colStr.charCodeAt(ci) - 64);
    }
    function cellRect(sheetRow, sheetCol) {
      var ri = sheetRow - startRow;
      var colIdx = sheetCol - startCol;
      if (ri < 0) ri = 0;
      if (colIdx < 0) colIdx = 0;
      var tr = table.rows[ri];
      if (!tr) tr = table.rows[table.rows.length - 1] || table.rows[0];
      if (!tr) return {x:0, y:0, w:0, h:0};
      var td = tr.cells[colIdx];
      if (!td) td = tr.cells[tr.cells.length - 1] || tr.cells[0];
      if (!td) return {x:0, y:0, w:0, h:0};
      return {x: td.offsetLeft, y: td.offsetTop, w: td.offsetWidth, h: td.offsetHeight};
    }
    wrapper.querySelectorAll('[data-from-col][data-from-row]').forEach(function(el){
      var fc = parseInt(el.getAttribute('data-from-col'),10);
      var fr = parseInt(el.getAttribute('data-from-row'),10);
      var from = cellRect(fr, fc);
      el.style.left = from.x + 'px';
      el.style.top = from.y + 'px';
      var tc = el.getAttribute('data-to-col');
      var tr2 = el.getAttribute('data-to-row');
      if (tc !== null && tr2 !== null) {
        var to = cellRect(parseInt(tr2,10), parseInt(tc,10));
        var w = to.x + to.w - from.x;
        var h = to.y + to.h - from.y;
        if (w > 0) el.style.width = w + 'px';
        if (h > 0) el.style.height = h + 'px';
      }
    });
  });
}
positionOverlays();
</script>
</body>
</html>`;
}
