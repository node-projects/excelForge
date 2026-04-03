import type { Chart, ChartType, ChartSeries, ChartAxis } from '../core/types.js';
import { escapeXml } from '../utils/helpers.js';

const COLORS = [
  'FF4472C4','FFED7D31','FFA5A5A5','FFFFC000','FF5B9BD5',
  'FF70AD47','FF264478','FF9E480E','FF636363','FF997300',
];

function spPr(color: string): string {
  return `<c:spPr><a:solidFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:srgbClr val="${color.replace('FF','')}"/></a:solidFill></c:spPr>`;
}

function seriesXml(type: ChartType, series: ChartSeries[], idx: number): string {
  return series.map((s, i) => {
    const color = s.color
      ? s.color.startsWith('#') ? 'FF' + s.color.slice(1) : s.color
      : COLORS[i % COLORS.length];
    const catXml = s.categories
      ? `<c:cat><c:strRef><c:f>${escapeXml(s.categories)}</c:f></c:strRef></c:cat>`
      : '';
    const valXml = `<c:val><c:numRef><c:f>${escapeXml(s.values)}</c:f></c:numRef></c:val>`;
    const nameXml = s.name
      ? `<c:tx><c:strRef><c:f>"${escapeXml(s.name)}"</c:f></c:strRef></c:tx>`
      : '';
    const marker = type.startsWith('line') || type === 'scatter'
      ? `<c:marker><c:symbol val="none"/></c:marker>` : '';
    return `<c:ser><c:idx val="${i}"/><c:order val="${i}"/>${nameXml}${spPr(color)}${marker}${catXml}${valXml}</c:ser>`;
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
  const title = chart.title
    ? `<c:title><c:tx><c:rich>
    <a:bodyPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>
    <a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <a:r><a:t>${escapeXml(chart.title)}</a:t></a:r>
    </a:p>
  </c:rich></c:tx><c:overlay val="0"/></c:title>` : '';

  const varyColors = chart.varyColors ? `<c:varyColors val="1"/>` : '';

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
    plotXml = `<c:${tag}>${varyColors}${seriesXml(type, series, 0)}${hole}</c:${tag}>`;
  } else if (isBar) {
    const barDir = type.startsWith('bar') ? 'bar' : 'col';
    plotXml = `<c:barChart>
  <c:barDir val="${barDir}"/>
  ${grouping ? `<c:grouping val="${grouping}"/>` : ''}
  ${varyColors}
  ${seriesXml(type, series, 0)}
  <c:axId val="1"/><c:axId val="2"/>
</c:barChart>`;
  } else if (isLine) {
    plotXml = `<c:lineChart>
  ${grouping ? `<c:grouping val="${grouping}"/>` : ''}
  ${varyColors}
  ${seriesXml(type, series, 0)}
  <c:axId val="1"/><c:axId val="2"/>
</c:lineChart>`;
  } else if (isScatter) {
    plotXml = `<c:scatterChart>
  <c:scatterStyle val="${type === 'scatterSmooth' ? 'smoothMarker' : 'marker'}"/>
  ${varyColors}
  ${seriesXml(type, series, 0)}
  <c:axId val="1"/><c:axId val="2"/>
</c:scatterChart>`;
  } else if (isArea) {
    plotXml = `<c:areaChart>
  ${grouping ? `<c:grouping val="${grouping}"/>` : ''}
  ${varyColors}
  ${seriesXml(type, series, 0)}
  <c:axId val="1"/><c:axId val="2"/>
</c:areaChart>`;
  } else if (isRadar) {
    plotXml = `<c:radarChart>
  <c:radarStyle val="${type === 'radarFilled' ? 'filled' : 'marker'}"/>
  ${varyColors}
  ${seriesXml(type, series, 0)}
  <c:axId val="1"/><c:axId val="2"/>
</c:radarChart>`;
  } else {
    // fallback bar
    plotXml = `<c:barChart>
  <c:barDir val="col"/>
  <c:grouping val="clustered"/>
  ${varyColors}
  ${seriesXml(type, series, 0)}
  <c:axId val="1"/><c:axId val="2"/>
</c:barChart>`;
  }

  const needsAxes = !isPie;
  const axesXml = needsAxes
    ? catAxisXml(1, 2, chart.xAxis) + axisXml(2, 1, chart.yAxis)
    : '';

  const legXml = legendXml(chart.legend ?? 'b');

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
${chart.style ? `<c:style val="${chart.style}"/>` : ''}
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
</c:chartSpace>`;
}
