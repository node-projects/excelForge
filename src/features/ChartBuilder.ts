import type { Chart, ChartType, ChartSeries, ChartAxis, ChartDataLabels, ChartTemplate, ChartModernStyle, ChartColorPalette } from '../core/types.js';
import { escapeXml } from '../utils/helpers.js';

const COLORS = [
  'FF4472C4','FFED7D31','FFA5A5A5','FFFFC000','FF5B9BD5',
  'FF70AD47','FF264478','FF9E480E','FF636363','FF997300',
];

// ── Color palettes ───────────────────────────────────────────────────────────
const PALETTES: Record<string, string[]> = {
  office:      ['4472C4','ED7D31','A5A5A5','FFC000','5B9BD5','70AD47','264478','9E480E','636363','997300'],
  blue:        ['5B9BD5','8FAADC','B4C7E7','D6DCE4','4472C4','2F5597','1F4E79','002060','00B0F0','0070C0'],
  orange:      ['ED7D31','F4B183','F8CBAD','FCE4D6','C55A11','843C0C','BF8F00','806000','FFD966','FFC000'],
  green:       ['70AD47','A9D18E','C5E0B4','E2F0D9','548235','375623','92D050','00B050','00B0F0','00FF00'],
  red:         ['FF0000','FF6161','FF9F9F','FFD2D2','C00000','8B0000','FFC000','FF6600','FF3300','CC0000'],
  purple:      ['7030A0','9B59B6','BB8FCE','D7BDE2','5B2C6F','4A235A','8E44AD','6C3483','A569BD','D2B4DE'],
  teal:        ['00B0F0','00BCD4','26C6DA','80DEEA','0097A7','006064','009688','00796B','4DB6AC','80CBC4'],
  gray:        ['595959','808080','A6A6A6','D9D9D9','404040','262626','BFBFBF','F2F2F2','7F7F7F','C0C0C0'],
  gold:        ['FFC000','FFD966','FFE699','FFF2CC','BF8F00','806000','ED7D31','F4B183','DBA13A','C09100'],
  blueWarm:    ['4472C4','5B9BD5','9DC3E6','BDD7EE','DEEBF7','2E75B6','1F4E79','2F5597','8FAADC','D6DCE4'],
  blueGreen:   ['00B0F0','00B050','70AD47','5B9BD5','4472C4','2E75B6','548235','00B0F0','009688','00796B'],
  greenYellow: ['70AD47','92D050','C9E265','FFD966','FFC000','548235','BF8F00','A9D18E','E2F0D9','C5E0B4'],
  redOrange:   ['FF0000','ED7D31','FFC000','FF6600','C00000','843C0C','BF8F00','FF3300','F4B183','FFD966'],
  redViolet:   ['FF0000','7030A0','ED7D31','C00000','5B2C6F','843C0C','FF6161','9B59B6','F4B183','BB8FCE'],
  yellowOrange:['FFC000','ED7D31','FFD966','F4B183','BF8F00','C55A11','FFE699','F8CBAD','806000','843C0C'],
  slipstream:  ['4472C4','ED7D31','A5A5A5','FFC000','5B9BD5','70AD47','255E91','9E480E','636363','997300'],
  marquee:     ['ED7D31','4472C4','70AD47','FFC000','5B9BD5','A5A5A5','264478','9E480E','636363','997300'],
  aspect:      ['5B9BD5','4472C4','ED7D31','FFC000','70AD47','A5A5A5','9E480E','264478','636363','997300'],
};

// ── Modern style → <c:style val="N"/> mapping ────────────────────────────────
const MODERN_STYLE_MAP: Record<string, number> = {
  colorful1: 102, colorful2: 103, colorful3: 104, colorful4: 105,
  monochromatic1: 201, monochromatic2: 202, monochromatic3: 203, monochromatic4: 204,
  monochromatic5: 205, monochromatic6: 206, monochromatic7: 207, monochromatic8: 208,
  monochromatic9: 209, monochromatic10: 210, monochromatic11: 211, monochromatic12: 212,
};

function spPr(color: string, series?: ChartSeries): string {
  const aNs = ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';
  const hex = color.replace(/^FF/, '');
  let fill: string;
  if (series?.fillType === 'gradient' && series.gradientStops?.length) {
    const stops = series.gradientStops.map(s =>
      `<a:gs pos="${s.pos * 1000}"><a:srgbClr val="${s.color.replace(/^FF/,'').replace(/^#/,'')}"/></a:gs>`
    ).join('');
    fill = `<a:gradFill><a:gsLst>${stops}</a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>`;
  } else {
    fill = `<a:solidFill><a:srgbClr val="${hex}"/></a:solidFill>`;
  }
  const lnWidth = series?.lineWidth ? ` w="${Math.round(series.lineWidth * 12700)}"` : '';
  const ln = lnWidth ? `<a:ln${lnWidth}><a:solidFill><a:srgbClr val="${hex}"/></a:solidFill></a:ln>` : '';
  return `<c:spPr${aNs}>${fill}${ln}</c:spPr>`;
}

function spPrSimple(color: string): string {
  return `<c:spPr><a:solidFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:srgbClr val="${color.replace(/^FF/,'')}"/></a:solidFill></c:spPr>`;
}

function dataLabelsXml(dl: ChartDataLabels | undefined): string {
  if (!dl) return '';
  // Schema order: numFmt, spPr, txPr, dLblPos, showLegendKey, showVal, showCatName, showSerName, showPercent, showBubbleSize
  const parts = [
    dl.numFmt ? `<c:numFmt formatCode="${escapeXml(dl.numFmt)}" sourceLinked="0"/>` : '',
    dl.position ? `<c:dLblPos val="${dl.position}"/>` : '',
    `<c:showLegendKey val="0"/>`,
    `<c:showVal val="${dl.showValue ? '1' : '0'}"/>`,
    `<c:showCatName val="${dl.showCategory ? '1' : '0'}"/>`,
    `<c:showSerName val="${dl.showSeriesName ? '1' : '0'}"/>`,
    `<c:showPercent val="${dl.showPercent ? '1' : '0'}"/>`,
  ];
  return `<c:dLbls>${parts.join('')}</c:dLbls>`;
}

function seriesXml(type: ChartType, series: ChartSeries[], idx: number, palette?: string[]): string {
  return series.map((s, i) => {
    const color = s.color
      ? s.color.startsWith('#') ? 'FF' + s.color.slice(1) : s.color
      : palette ? 'FF' + palette[i % palette.length] : COLORS[i % COLORS.length];
    const catXml = s.categories
      ? `<c:cat><c:strRef><c:f>${escapeXml(s.categories)}</c:f></c:strRef></c:cat>`
      : '';
    const valXml = `<c:val><c:numRef><c:f>${escapeXml(s.values)}</c:f></c:numRef></c:val>`;
    const nameXml = s.name
      ? `<c:tx><c:strRef><c:f>"${escapeXml(s.name)}"</c:f></c:strRef></c:tx>`
      : '';
    const marker = type.startsWith('line') || type === 'scatter'
      ? `<c:marker><c:symbol val="none"/></c:marker>` : '';
    const dlXml = dataLabelsXml(s.dataLabels);
    return `<c:ser><c:idx val="${i}"/><c:order val="${i}"/>${nameXml}${spPr(color, s)}${marker}${dlXml}${catXml}${valXml}</c:ser>`;
  }).join('');
}

function axisXml(id: number, crossId: number, axis?: ChartAxis, delete_ = false): string {
  if (delete_) return `<c:valAx><c:axId val="${id}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="1"/><c:axPos val="b"/><c:crossAx val="${crossId}"/></c:valAx>`;
  const title = axis?.title
    ? `<c:title><c:tx><c:rich><a:bodyPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/><a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:r><a:t>${escapeXml(axis.title)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>` : '';
  const minMax = [
    axis?.min !== undefined ? `<c:min val="${axis.min}"/>` : '',
    axis?.max !== undefined ? `<c:max val="${axis.max}"/>` : '',
  ].join('');
  const numFmt = axis?.numFmt ? `<c:numFmt formatCode="${escapeXml(axis.numFmt)}" sourceLinked="0"/>` : '';
  const gridLines = axis?.gridLines !== false ? `<c:majorGridlines/>` : '';
  return `<c:valAx>
  <c:axId val="${id}"/>
  <c:scaling><c:orientation val="minMax"/>${minMax}</c:scaling>
  <c:delete val="0"/>
  <c:axPos val="l"/>
  ${gridLines}${title}${numFmt}
  <c:crossAx val="${crossId}"/>
</c:valAx>`;
}

function catAxisXml(id: number, crossId: number, axis?: ChartAxis): string {
  const title = axis?.title
    ? `<c:title><c:tx><c:rich><a:bodyPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/><a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:r><a:t>${escapeXml(axis.title)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>` : '';
  const gridLines = axis?.gridLines ? `<c:majorGridlines/>` : '';
  return `<c:catAx>
  <c:axId val="${id}"/>
  <c:scaling><c:orientation val="minMax"/></c:scaling>
  <c:delete val="0"/>
  <c:axPos val="b"/>
  ${gridLines}${title}
  <c:crossAx val="${crossId}"/>
</c:catAx>`;
}

function legendXml(legend: Chart['legend']): string {
  if (legend === false) return '';
  const raw = typeof legend === 'string' ? legend : 'b';
  // Normalize long form to short form for OOXML compliance
  const posMap: Record<string, string> = { bottom: 'b', top: 't', left: 'l', right: 'r', corner: 'tr' };
  const pos = posMap[raw] ?? raw;
  return `<c:legend><c:legendPos val="${pos}"/></c:legend>`;
}

export function buildChartXml(chart: Chart): string {
  const type = chart.type;
  const series = chart.series;
  const palette = chart.colorPalette ? PALETTES[chart.colorPalette] : undefined;
  const title = chart.title
    ? `<c:title><c:tx><c:rich>
    <a:bodyPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>
    <a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <a:r><a:t>${escapeXml(chart.title)}</a:t></a:r>
    </a:p>
  </c:rich></c:tx><c:overlay val="0"/></c:title>` : '';

  const varyColors = chart.varyColors ? `<c:varyColors val="1"/>` : '';
  const globalDL = dataLabelsXml(chart.dataLabels);

  let plotXml = '';

  const isPie = type === 'pie' || type === 'doughnut';
  const isBar = type.startsWith('bar') || type.startsWith('column');
  const isLine = type.startsWith('line');
  const isScatter = type.startsWith('scatter') || type === 'bubble';
  const isArea = type.startsWith('area');
  const isRadar = type.startsWith('radar');
  const isStock = type === 'stock';

  const grouping = chart.grouping ?? (
    type.endsWith('Stacked100') ? 'percentStacked' :
    type.endsWith('Stacked')    ? 'stacked' :
    isBar ? 'clustered' : isLine || isArea ? 'standard' : undefined
  );

  if (isPie) {
    const tag = type === 'doughnut' ? 'doughnutChart' : 'pieChart';
    const hole = type === 'doughnut' ? `<c:holeSize val="50"/>` : '';
    plotXml = `<c:${tag}>${varyColors}${seriesXml(type, series, 0, palette)}${globalDL}${hole}</c:${tag}>`;
  } else if (isBar) {
    const barDir = type.startsWith('bar') ? 'bar' : 'col';
    plotXml = `<c:barChart>
  <c:barDir val="${barDir}"/>
  ${grouping ? `<c:grouping val="${grouping}"/>` : ''}
  ${varyColors}
  ${seriesXml(type, series, 0, palette)}
  ${globalDL}
  <c:axId val="1"/><c:axId val="2"/>
</c:barChart>`;
  } else if (isLine) {
    plotXml = `<c:lineChart>
  ${grouping ? `<c:grouping val="${grouping}"/>` : ''}
  ${varyColors}
  ${seriesXml(type, series, 0, palette)}
  ${globalDL}
  <c:axId val="1"/><c:axId val="2"/>
</c:lineChart>`;
  } else if (isScatter) {
    plotXml = `<c:scatterChart>
  <c:scatterStyle val="${type === 'scatterSmooth' ? 'smoothMarker' : 'marker'}"/>
  ${varyColors}
  ${seriesXml(type, series, 0, palette)}
  ${globalDL}
  <c:axId val="1"/><c:axId val="2"/>
</c:scatterChart>`;
  } else if (isArea) {
    plotXml = `<c:areaChart>
  ${grouping ? `<c:grouping val="${grouping}"/>` : ''}
  ${varyColors}
  ${seriesXml(type, series, 0, palette)}
  ${globalDL}
  <c:axId val="1"/><c:axId val="2"/>
</c:areaChart>`;
  } else if (isRadar) {
    plotXml = `<c:radarChart>
  <c:radarStyle val="${type === 'radarFilled' ? 'filled' : 'marker'}"/>
  ${varyColors}
  ${seriesXml(type, series, 0, palette)}
  ${globalDL}
  <c:axId val="1"/><c:axId val="2"/>
</c:radarChart>`;
  } else {
    // fallback bar
    plotXml = `<c:barChart>
  <c:barDir val="col"/>
  <c:grouping val="clustered"/>
  ${varyColors}
  ${seriesXml(type, series, 0, palette)}
  ${globalDL}
  <c:axId val="1"/><c:axId val="2"/>
</c:barChart>`;
  }

  const needsAxes = !isPie;
  const axesXml = needsAxes
    ? catAxisXml(1, 2, chart.xAxis) + axisXml(2, 1, chart.yAxis)
    : '';

  const legXml = legendXml(chart.legend ?? 'b');

  // Determine style: modern styles > 48 are not valid in c:style, cap at 48
  const styleVal = chart.modernStyle ? Math.min(MODERN_STYLE_MAP[chart.modernStyle], 48) : chart.style;
  const styleXml = styleVal ? `<c:style val="${styleVal}"/>` : '';

  // Chart-area shape properties (fill, shadow, rounded corners)
  let chartAreaSpPr = '';
  if (chart.chartFill || chart.shadow) {
    const fillXml = chart.chartFill === 'gradient'
      ? `<a:gradFill><a:gsLst><a:gs pos="0"><a:schemeClr val="lt1"/></a:gs><a:gs pos="100000"><a:schemeClr val="bg1"><a:lumMod val="85000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>`
      : chart.chartFill === 'white' ? `<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>`
      : chart.chartFill === 'none' ? `<a:noFill/>` : '';
    const shadowXml = chart.shadow
      ? `<a:effectLst><a:outerShdw blurRad="50800" dist="38100" dir="5400000" algn="t" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="40000"/></a:srgbClr></a:outerShdw></a:effectLst>`
      : '';
    chartAreaSpPr = `<c:spPr>${fillXml}${shadowXml}</c:spPr>`;
  }

  const roundedXml = chart.roundedCorners ? `<c:roundedCorners val="1"/>` : '';

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
${roundedXml}
${styleXml}
<c:chart>
  ${title}
  <c:autoTitleDeleted val="${chart.title ? '0' : '1'}"/>
  <c:plotArea>
    ${plotXml}
    ${axesXml}
  </c:plotArea>
  ${legXml}
  <c:plotVisOnly val="1"/>
</c:chart>
${chartAreaSpPr}
</c:chartSpace>`;
}

// ── Chart Template Support (.crtx) ──────────────────────────────────────────

/**
 * Create a chart template (.crtx) binary from a Chart config.
 * A .crtx file is an OPC (ZIP) package containing a chart definition.
 */
export function saveChartTemplate(chart: Chart): ChartTemplate {
  return {
    type: chart.type,
    style: chart.style,
    modernStyle: chart.modernStyle,
    colorPalette: chart.colorPalette,
    legend: chart.legend,
    xAxis: chart.xAxis,
    yAxis: chart.yAxis,
    dataLabels: chart.dataLabels,
    chartFill: chart.chartFill,
    roundedCorners: chart.roundedCorners,
    shadow: chart.shadow,
    varyColors: chart.varyColors,
    grouping: chart.grouping,
  };
}

/**
 * Apply a chart template to a partial chart definition.
 * The template provides defaults that the chart can override.
 */
export function applyChartTemplate(template: ChartTemplate, chart: Partial<Chart> & { series: ChartSeries[]; from: Chart['from']; to: Chart['to'] }): Chart {
  return {
    type: chart.type ?? template.type,
    title: chart.title,
    series: chart.series,
    from: chart.from,
    to: chart.to,
    xAxis: chart.xAxis ?? template.xAxis,
    yAxis: chart.yAxis ?? template.yAxis,
    legend: chart.legend ?? template.legend,
    style: chart.style ?? template.style,
    modernStyle: chart.modernStyle ?? template.modernStyle,
    colorPalette: chart.colorPalette ?? template.colorPalette,
    dataLabels: chart.dataLabels ?? template.dataLabels,
    chartFill: chart.chartFill ?? template.chartFill,
    roundedCorners: chart.roundedCorners ?? template.roundedCorners,
    shadow: chart.shadow ?? template.shadow,
    varyColors: chart.varyColors ?? template.varyColors,
    grouping: chart.grouping ?? template.grouping,
  };
}

/** Serialize a chart template to JSON for saving/loading */
export function serializeChartTemplate(template: ChartTemplate): string {
  return JSON.stringify(template);
}

/** Deserialize a chart template from JSON */
export function deserializeChartTemplate(json: string): ChartTemplate {
  return JSON.parse(json) as ChartTemplate;
}
