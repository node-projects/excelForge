/**
 * WorkbookReader — reads an existing .xlsx file into the ExcelForge object model.
 *
 * Strategy: parse what we understand, preserve everything else verbatim so
 * that on write we can patch only the parts we changed.
 */

import { readZip, entryText, type ZipReadEntry } from '../utils/zipReader.js';
import { parseXml, child, children, attr, localName, nodeToXml, type XmlNode } from '../utils/xmlParser.js';
import {
  parseCoreXml, parseAppXml, parseCustomXml,
  type CoreProperties, type ExtendedProperties, type CustomProperty,
} from './properties.js';
import type {
  CellStyle, Font, Fill, PatternFill, GradientFill, Border, BorderSide,
  Alignment, NumberFormat, MergeRange, FreezePane, SheetView,
  PageSetup, PageMargins, HeaderFooter, PrintOptions, SheetProtection,
  AutoFilter, ConditionalFormat, DataValidation, Cell, CellValue,
  RichTextRun, Comment, Hyperlink, Table, TableColumn, ValidationType,
} from './types.js';
import { Worksheet } from './Worksheet.js';
import { SharedStrings } from './SharedStrings.js';
import { StyleRegistry } from '../styles/StyleRegistry.js';
import { cellRefToIndices, colLetterToIndex, dateToSerial } from '../utils/helpers.js';

// ─── Raw file store (for unknown parts) ───────────────────────────────────────

export interface RawPart {
  path:        string;
  data:        Uint8Array;
  contentType: string;
}

// ─── Relationship map ─────────────────────────────────────────────────────────

type RelEntry = { type: string; target: string; targetMode?: string };
type RelMap = Map<string, RelEntry>;

function parseRels(xml: string): RelMap {
  const map: RelMap = new Map();
  try {
    const root = parseXml(xml);
    for (const c of root.children) {
      if (localName(c.tag) === 'Relationship') {
        const entry: RelEntry = {
          type:   c.attrs['Type'] ?? '',
          target: c.attrs['Target'] ?? '',
        };
        if (c.attrs['TargetMode']) entry.targetMode = c.attrs['TargetMode'];
        map.set(c.attrs['Id'] ?? '', entry);
      }
    }
  } catch {}
  return map;
}

// ─── Content-type map ─────────────────────────────────────────────────────────

type CTMap = Map<string, string>; // partName → contentType

function parseContentTypes(xml: string): CTMap {
  const map: CTMap = new Map();
  try {
    const root = parseXml(xml);
    for (const c of root.children) {
      const ln = localName(c.tag);
      if (ln === 'Override') {
        const part = c.attrs['PartName'] ?? '';
        const ct   = c.attrs['ContentType'] ?? '';
        map.set(part.startsWith('/') ? part.slice(1) : part, ct);
      }
    }
  } catch {}
  return map;
}

// ─── Style parsing ────────────────────────────────────────────────────────────

interface ParsedStyles {
  /** xf index → CellStyle */
  xfs: CellStyle[];
  /** custom numFmtId → formatCode */
  numFmts: Map<number, string>;
  /** dxf index → CellStyle (differential formats for conditional formatting) */
  dxfs: CellStyle[];
}

function parseStyles(xml: string): ParsedStyles {
  const root = parseXml(xml);
  const numFmts = new Map<number, string>();
  const xfs: CellStyle[] = [];

  // Built-in numfmt IDs
  const builtinFmts: Record<number, string> = {
    0:'General',1:'0',2:'0.00',3:'#,##0',4:'#,##0.00',
    9:'0%',10:'0.00%',11:'0.00E+00',12:'# ?/?',13:'# ??/??',
    14:'mm-dd-yy',15:'d-mmm-yy',16:'d-mmm',17:'mmm-yy',
    18:'h:mm AM/PM',19:'h:mm:ss AM/PM',20:'h:mm',21:'h:mm:ss',
    22:'m/d/yy h:mm',37:'#,##0 ;(#,##0)',38:'#,##0 ;[Red](#,##0)',
    39:'#,##0.00;(#,##0.00)',40:'#,##0.00;[Red](#,##0.00)',
    45:'mm:ss',46:'[h]:mm:ss',47:'mmss.0',48:'##0.0E+0',49:'@',
  };

  // Parse custom numFmts
  const numFmtNode = find(root, 'numFmts');
  if (numFmtNode) {
    for (const n of children(numFmtNode, 'numFmt')) {
      const id = parseInt(n.attrs['numFmtId'] ?? '0', 10);
      const code = n.attrs['formatCode'] ?? '';
      numFmts.set(id, code);
    }
  }

  // Parse fonts
  const fontsNode = find(root, 'fonts');
  const fonts: Font[] = [];
  if (fontsNode) {
    for (const fn of children(fontsNode, 'font')) {
      fonts.push(parseFont(fn));
    }
  }

  // Parse fills
  const fillsNode = find(root, 'fills');
  const fills: Fill[] = [];
  if (fillsNode) {
    for (const fn of children(fillsNode, 'fill')) {
      fills.push(parseFill(fn));
    }
  }

  // Parse borders
  const bordersNode = find(root, 'borders');
  const borders: Border[] = [];
  if (bordersNode) {
    for (const bn of children(bordersNode, 'border')) {
      borders.push(parseBorder(bn));
    }
  }

  // Parse cellXfs
  const cellXfsNode = find(root, 'cellXfs');
  if (cellXfsNode) {
    for (const xf of children(cellXfsNode, 'xf')) {
      const fontId   = parseInt(xf.attrs['fontId']   ?? '0', 10);
      const fillId   = parseInt(xf.attrs['fillId']   ?? '0', 10);
      const borderId = parseInt(xf.attrs['borderId'] ?? '0', 10);
      const numFmtId = parseInt(xf.attrs['numFmtId'] ?? '0', 10);
      const applyFont      = xf.attrs['applyFont']      === '1';
      const applyFill      = xf.attrs['applyFill']      === '1';
      const applyBorder    = xf.attrs['applyBorder']    === '1';
      const applyAlignment = xf.attrs['applyAlignment'] === '1';
      const applyNumFmt    = xf.attrs['applyNumberFormat'] === '1';
      const applyProt      = xf.attrs['applyProtection'] === '1';

      const style: CellStyle = {};
      if (applyFont    && fonts[fontId])   style.font   = fonts[fontId];
      if (applyFill    && fills[fillId])   style.fill   = fills[fillId];
      if (applyBorder  && borders[borderId]) style.border = borders[borderId];

      if (applyNumFmt && numFmtId !== 0) {
        if (numFmts.has(numFmtId)) {
          style.numberFormat = { formatCode: numFmts.get(numFmtId)! };
        } else if (builtinFmts[numFmtId]) {
          style.numFmtId = numFmtId;
        }
      }

      const alignNode = child(xf, 'alignment');
      if (applyAlignment && alignNode) {
        style.alignment = parseAlignment(alignNode);
      }

      const protNode = child(xf, 'protection');
      if (applyProt && protNode) {
        if (protNode.attrs['locked'] !== undefined) style.locked = protNode.attrs['locked'] !== '0';
        if (protNode.attrs['hidden'] !== undefined) style.hidden = protNode.attrs['hidden'] !== '0';
      }

      xfs.push(style);
    }
  }

  // Parse dxfs (differential formats for conditional formatting)
  const dxfs: CellStyle[] = [];
  const dxfsNode = find(root, 'dxfs');
  if (dxfsNode) {
    for (const dxf of children(dxfsNode, 'dxf')) {
      const style: CellStyle = {};
      const fontNode = child(dxf, 'font');
      if (fontNode) style.font = parseFont(fontNode);
      const fillNode = child(dxf, 'fill');
      if (fillNode) style.fill = parseFill(fillNode);
      const borderNode = child(dxf, 'border');
      if (borderNode) style.border = parseBorder(borderNode);
      const numFmtDxf = child(dxf, 'numFmt');
      if (numFmtDxf) {
        const code = numFmtDxf.attrs['formatCode'] ?? '';
        if (code) style.numberFormat = { formatCode: code };
      }
      const alignDxf = child(dxf, 'alignment');
      if (alignDxf) style.alignment = parseAlignment(alignDxf);
      dxfs.push(style);
    }
  }

  return { xfs, numFmts, dxfs };
}

function find(node: XmlNode, localTag: string): XmlNode | undefined {
  if (localName(node.tag) === localTag) return node;
  for (const c of node.children) {
    const r = find(c, localTag);
    if (r) return r;
  }
  return undefined;
}

function parseFont(node: XmlNode): Font {
  const f: Font = {};
  for (const c of node.children) {
    switch (localName(c.tag)) {
      case 'b':       f.bold   = c.attrs['val'] !== '0'; break;
      case 'i':       f.italic = c.attrs['val'] !== '0'; break;
      case 'strike':  f.strike = c.attrs['val'] !== '0'; break;
      case 'u':       f.underline = (c.attrs['val'] as Font['underline']) ?? 'single'; break;
      case 'sz':      f.size   = parseFloat(c.attrs['val'] ?? '11'); break;
      case 'name':    f.name   = c.attrs['val']; break;
      case 'family':  f.family = parseInt(c.attrs['val'] ?? '0', 10); break;
      case 'scheme':  f.scheme = c.attrs['val'] as Font['scheme']; break;
      case 'charset': f.charset = parseInt(c.attrs['val'] ?? '0', 10); break;
      case 'vertAlign': f.vertAlign = c.attrs['val'] as Font['vertAlign']; break;
      case 'color': {
        const rgb = c.attrs['rgb'];
        const theme = c.attrs['theme'];
        if (rgb) f.color = rgb;
        else if (theme) f.color = `theme:${theme}`;
        break;
      }
    }
  }
  return f;
}

function parseFill(node: XmlNode): Fill {
  const pattern = child(node, 'patternFill');
  if (pattern) {
    const fg = child(pattern, 'fgColor');
    const bg = child(pattern, 'bgColor');
    return {
      type: 'pattern',
      pattern: (pattern.attrs['patternType'] ?? 'none') as any,
      fgColor: fg?.attrs['rgb'] ?? (fg?.attrs['theme'] ? `theme:${fg.attrs['theme']}` : undefined),
      bgColor: bg?.attrs['rgb'] ?? (bg?.attrs['theme'] ? `theme:${bg.attrs['theme']}` : undefined),
    } as PatternFill;
  }
  const gradient = child(node, 'gradientFill');
  if (gradient) {
    const stops = children(gradient, 'stop').concat(children(gradient, 'gradientStop')).map(s => {
      const colorNode = child(s, 'color');
      return {
        position: parseFloat(s.attrs['position'] ?? '0'),
        color: colorNode?.attrs['rgb'] ?? (colorNode?.attrs['theme'] ? `theme:${colorNode.attrs['theme']}` : 'FF000000'),
      };
    });
    return {
      type: 'gradient',
      gradientType: gradient.attrs['type'] as any,
      degree: gradient.attrs['degree'] ? parseFloat(gradient.attrs['degree']) : undefined,
      stops,
    } as GradientFill;
  }
  return { type: 'pattern', pattern: 'none' } as PatternFill;
}

function parseBorder(node: XmlNode): Border {
  const parseSide = (tag: string): BorderSide | undefined => {
    const n = child(node, tag);
    if (!n) return undefined;
    const style = n.attrs['style'];
    const color = child(n, 'color');
    if (!style && !color) return undefined;
    const colorVal = color?.attrs['rgb'] ?? (color?.attrs['theme'] ? `theme:${color.attrs['theme']}` : undefined);
    return { style: style as any, color: colorVal };
  };
  return {
    left:     parseSide('left'),
    right:    parseSide('right'),
    top:      parseSide('top'),
    bottom:   parseSide('bottom'),
    diagonal: parseSide('diagonal'),
    diagonalUp:   node.attrs['diagonalUp']   === '1',
    diagonalDown: node.attrs['diagonalDown'] === '1',
  };
}

function parseAlignment(node: XmlNode): Alignment {
  const a: Alignment = {};
  if (node.attrs['horizontal'])   a.horizontal   = node.attrs['horizontal'] as any;
  if (node.attrs['vertical'])     a.vertical     = node.attrs['vertical'] as any;
  if (node.attrs['wrapText'])     a.wrapText     = node.attrs['wrapText'] !== '0';
  if (node.attrs['shrinkToFit'])  a.shrinkToFit  = node.attrs['shrinkToFit'] !== '0';
  if (node.attrs['textRotation']) a.textRotation = parseInt(node.attrs['textRotation'], 10);
  if (node.attrs['indent'])       a.indent       = parseInt(node.attrs['indent'], 10);
  if (node.attrs['readingOrder']) a.readingOrder = parseInt(node.attrs['readingOrder'], 10) as any;
  return a;
}

// ─── Shared strings parsing ───────────────────────────────────────────────────

function parseSharedStrings(xml: string): string[] {
  const root = parseXml(xml);
  return children(root, 'si').map(si => {
    // Simple string
    const t = child(si, 't');
    if (t && !child(si, 'r')) return t.text ?? '';
    // Rich text — concatenate all runs
    return children(si, 'r').map(r => child(r, 't')?.text ?? '').join('');
  });
}

// ─── Worksheet parsing ────────────────────────────────────────────────────────

interface ParsedSheet {
  ws: Worksheet;
  /** The original XML, used for patching */
  originalXml: string;
  /** Unknown top-level elements (pivot tables, VML, etc.) — preserved verbatim */
  unknownParts: string[];
  /** Relationship IDs of tables referenced by <tableParts> */
  tableRIds: string[];
}

function parseWorksheet(
  xml: string,
  name: string,
  styles: ParsedStyles,
  sharedStrings: string[],
): ParsedSheet {
  const ws = new Worksheet(name);
  const root = parseXml(xml);
  const unknownParts: string[] = [];
  const tableRIds: string[] = [];

  const knownTags = new Set([
    'sheetPr','dimension','sheetViews','sheetFormatPr','cols',
    'sheetData','mergeCells','conditionalFormatting','dataValidations',
    'sheetProtection','printOptions','pageMargins','pageSetup',
    'headerFooter','drawing','tableParts','autoFilter',
    'rowBreaks','colBreaks','picture','oleObjects','ctrlProps',
  ]);

  for (const node of root.children) {
    const tag = localName(node.tag);
    switch (tag) {
      case 'sheetViews':    parseSheetViews(node, ws);   break;
      case 'cols':          parseCols(node, ws, styles); break;
      case 'sheetData':     parseSheetData(node, ws, styles, sharedStrings); break;
      case 'mergeCells':    parseMerges(node, ws);       break;
      case 'autoFilter':    ws.autoFilter = { ref: node.attrs['ref'] ?? '' }; break;
      case 'tableParts':
        for (const tp of children(node, 'tablePart')) {
          const rid = tp.attrs['r:id'] ?? '';
          if (rid) tableRIds.push(rid);
        }
        break;
      case 'sheetProtection': parseProtection(node, ws); break;
      case 'pageMargins':   parsePageMargins(node, ws); break;
      case 'pageSetup':     parsePageSetup(node, ws);   break;
      case 'headerFooter':  parseHeaderFooter(node, ws); break;
      case 'printOptions':  parsePrintOptions(node, ws); break;
      case 'conditionalFormatting':
        parseConditionalFormatting(node, ws, styles);
        break;
      case 'dataValidations':
        parseDataValidations(node, ws);
        break;
      case 'rowBreaks':
        for (const brk of children(node, 'brk')) {
          const id = parseInt(brk.attrs['id'] ?? '0', 10);
          if (id > 0) ws.addRowBreak(id, brk.attrs['man'] === '1');
        }
        break;
      case 'colBreaks':
        for (const brk of children(node, 'brk')) {
          const id = parseInt(brk.attrs['id'] ?? '0', 10);
          if (id > 0) ws.addColBreak(id, brk.attrs['man'] === '1');
        }
        break;
      default:
        if (!knownTags.has(tag)) {
          unknownParts.push(nodeToXml(node));
        }
        break;
    }
  }

  return { ws, originalXml: xml, unknownParts, tableRIds };
}

function parseSheetViews(node: XmlNode, ws: Worksheet): void {
  const sv = children(node, 'sheetView')[0];
  if (!sv) return;

  ws.view = {
    showGridLines:     sv.attrs['showGridLines']     !== '0',
    showRowColHeaders: sv.attrs['showRowColHeaders'] !== '0',
    zoomScale:         sv.attrs['zoomScale'] ? parseInt(sv.attrs['zoomScale'], 10) : undefined,
    rightToLeft:       sv.attrs['rightToLeft'] === '1',
    tabSelected:       sv.attrs['tabSelected'] === '1',
    view:              sv.attrs['view'] as any,
  };

  const pane = child(sv, 'pane');
  if (pane && pane.attrs['state'] === 'frozen') {
    ws.freezePane = {
      col: pane.attrs['xSplit'] ? parseInt(pane.attrs['xSplit'], 10) : undefined,
      row: pane.attrs['ySplit'] ? parseInt(pane.attrs['ySplit'], 10) : undefined,
    };
  }
}

function parseCols(node: XmlNode, ws: Worksheet, styles: ParsedStyles): void {
  for (const col of children(node, 'col')) {
    const min = parseInt(col.attrs['min'] ?? '1', 10);
    const max = parseInt(col.attrs['max'] ?? '1', 10);
    const def = {
      width:        col.attrs['width']  ? parseFloat(col.attrs['width']) : undefined,
      hidden:       col.attrs['hidden'] === '1',
      customWidth:  col.attrs['customWidth'] === '1',
      outlineLevel: col.attrs['outlineLevel'] ? parseInt(col.attrs['outlineLevel'], 10) : undefined,
      style:        col.attrs['style']  ? styles.xfs[parseInt(col.attrs['style'], 10)] : undefined,
    };
    for (let c = min; c <= max; c++) ws.setColumn(c, def);
  }
}

function parseSheetData(
  node: XmlNode,
  ws: Worksheet,
  styles: ParsedStyles,
  sharedStrings: string[],
): void {
  for (const rowNode of children(node, 'row')) {
    const rowIdx = parseInt(rowNode.attrs['r'] ?? '0', 10);
    if (!rowIdx) continue;

    const rowDef: any = {};
    if (rowNode.attrs['ht'])        rowDef.height       = parseFloat(rowNode.attrs['ht']);
    if (rowNode.attrs['hidden'])    rowDef.hidden        = rowNode.attrs['hidden'] === '1';
    if (rowNode.attrs['outlineLevel']) rowDef.outlineLevel = parseInt(rowNode.attrs['outlineLevel'], 10);
    if (rowNode.attrs['collapsed']) rowDef.collapsed     = rowNode.attrs['collapsed'] === '1';
    if (rowNode.attrs['s'])         rowDef.style         = styles.xfs[parseInt(rowNode.attrs['s'], 10)];
    if (Object.keys(rowDef).length) ws.setRow(rowIdx, rowDef);

    for (const cNode of children(rowNode, 'c')) {
      const ref = cNode.attrs['r'] ?? '';
      if (!ref) continue;
      const { row, col } = cellRefToIndices(ref);
      const styleIdx = cNode.attrs['s'] ? parseInt(cNode.attrs['s'], 10) : 0;
      const cellStyle = styleIdx > 0 ? styles.xfs[styleIdx] : undefined;
      const t = cNode.attrs['t'] ?? '';
      const fNode = child(cNode, 'f');
      const vNode = child(cNode, 'v');

      const cell: Cell = {};
      if (cellStyle) cell.style = cellStyle;

      if (fNode) {
        if (fNode.attrs['t'] === 'array') {
          cell.arrayFormula = fNode.text ?? '';
        } else {
          cell.formula = fNode.text ?? '';
        }
      } else if (vNode) {
        const raw = vNode.text ?? '';
        switch (t) {
          case 's': {
            const idx = parseInt(raw, 10);
            cell.value = sharedStrings[idx] ?? '';
            break;
          }
          case 'b':
            cell.value = raw === '1' || raw === 'true';
            break;
          case 'str':
          case 'inlineStr': {
            const is = child(cNode, 'is');
            cell.value = is ? (child(is, 't')?.text ?? raw) : raw;
            break;
          }
          case 'e':
            cell.value = raw; // error value
            break;
          default: {
            const n = parseFloat(raw);
            cell.value = isNaN(n) ? raw : n;
            break;
          }
        }
      }

      if (Object.keys(cell).length || cell.value !== undefined) {
        ws.setCell(row, col, cell);
      }
    }
  }
}

function parseMerges(node: XmlNode, ws: Worksheet): void {
  for (const m of children(node, 'mergeCell')) {
    const ref = m.attrs['ref'] ?? '';
    if (ref.includes(':')) ws.mergeByRef(ref);
  }
}

function parseProtection(node: XmlNode, ws: Worksheet): void {
  ws.protection = {
    sheet:                node.attrs['sheet']               !== '0',
    password:             undefined, // hash only, can't reverse
    selectLockedCells:    node.attrs['selectLockedCells']   !== '0',
    selectUnlockedCells:  node.attrs['selectUnlockedCells'] !== '0',
    formatCells:          node.attrs['formatCells']         === '0',
    formatColumns:        node.attrs['formatColumns']       === '0',
    formatRows:           node.attrs['formatRows']          === '0',
    insertColumns:        node.attrs['insertColumns']       === '0',
    insertRows:           node.attrs['insertRows']          === '0',
    insertHyperlinks:     node.attrs['insertHyperlinks']    === '0',
    deleteColumns:        node.attrs['deleteColumns']       === '0',
    deleteRows:           node.attrs['deleteRows']          === '0',
    sort:                 node.attrs['sort']                === '0',
    autoFilter:           node.attrs['autoFilter']          === '0',
    pivotTables:          node.attrs['pivotTables']         === '0',
  };
}

function parsePageMargins(node: XmlNode, ws: Worksheet): void {
  ws.pageMargins = {
    left:   parseFloat(node.attrs['left']   ?? '0.7'),
    right:  parseFloat(node.attrs['right']  ?? '0.7'),
    top:    parseFloat(node.attrs['top']    ?? '0.75'),
    bottom: parseFloat(node.attrs['bottom'] ?? '0.75'),
    header: parseFloat(node.attrs['header'] ?? '0.3'),
    footer: parseFloat(node.attrs['footer'] ?? '0.3'),
  };
}

function parsePageSetup(node: XmlNode, ws: Worksheet): void {
  ws.pageSetup = {
    paperSize:       node.attrs['paperSize']     ? parseInt(node.attrs['paperSize'], 10) as any : undefined,
    orientation:     node.attrs['orientation']   as any,
    fitToPage:       node.attrs['fitToPage']     === '1',
    fitToWidth:      node.attrs['fitToWidth']    ? parseInt(node.attrs['fitToWidth'], 10) : undefined,
    fitToHeight:     node.attrs['fitToHeight']   ? parseInt(node.attrs['fitToHeight'], 10) : undefined,
    scale:           node.attrs['scale']         ? parseInt(node.attrs['scale'], 10) : undefined,
    horizontalDpi:   node.attrs['horizontalDpi'] ? parseInt(node.attrs['horizontalDpi'], 10) : undefined,
    verticalDpi:     node.attrs['verticalDpi']   ? parseInt(node.attrs['verticalDpi'], 10) : undefined,
  };
}

function parseHeaderFooter(node: XmlNode, ws: Worksheet): void {
  ws.headerFooter = {
    oddHeader:         child(node, 'oddHeader')?.text,
    oddFooter:         child(node, 'oddFooter')?.text,
    evenHeader:        child(node, 'evenHeader')?.text,
    evenFooter:        child(node, 'evenFooter')?.text,
    firstHeader:       child(node, 'firstHeader')?.text,
    firstFooter:       child(node, 'firstFooter')?.text,
    differentOddEven:  node.attrs['differentOddEven'] === '1',
    differentFirst:    node.attrs['differentFirst']   === '1',
  };
}

function parsePrintOptions(node: XmlNode, ws: Worksheet): void {
  ws.printOptions = {
    gridLines:         node.attrs['gridLines']           === '1',
    gridLinesSet:      node.attrs['gridLinesSet']        === '1',
    headings:          node.attrs['headings']            === '1',
    centerHorizontal:  node.attrs['horizontalCentered']  === '1',
    centerVertical:    node.attrs['verticalCentered']    === '1',
  };
}

// ─── Conditional formatting parsing ──────────────────────────────────────────

function parseConditionalFormatting(node: XmlNode, ws: Worksheet, styles: ParsedStyles): void {
  const sqref = node.attrs['sqref'] ?? '';
  for (const rule of children(node, 'cfRule')) {
    const type = (rule.attrs['type'] ?? 'cellIs') as ConditionalFormat['type'];
    const cf: ConditionalFormat = { sqref, type };

    if (rule.attrs['operator']) cf.operator = rule.attrs['operator'] as any;
    if (rule.attrs['priority'])  cf.priority = parseInt(rule.attrs['priority'], 10);
    if (rule.attrs['text'])      cf.text = rule.attrs['text'];
    if (rule.attrs['aboveAverage'] === '0') cf.aboveAverage = false;
    if (rule.attrs['percent'] === '1') cf.percent = true;
    if (rule.attrs['rank'])     cf.rank = parseInt(rule.attrs['rank'], 10);
    if (rule.attrs['timePeriod']) cf.timePeriod = rule.attrs['timePeriod'];

    // Resolve dxfId to CellStyle
    if (rule.attrs['dxfId'] !== undefined) {
      const dxfId = parseInt(rule.attrs['dxfId'], 10);
      if (styles.dxfs[dxfId]) cf.style = styles.dxfs[dxfId];
    }

    // Parse formulas
    const formulas = children(rule, 'formula');
    if (formulas[0]?.text) cf.formula = formulas[0].text;
    if (formulas[1]?.text) cf.formula2 = formulas[1].text;

    // Parse colorScale
    const csNode = child(rule, 'colorScale');
    if (csNode) {
      const cfvos = children(csNode, 'cfvo').map(c => ({
        type: (c.attrs['type'] ?? 'min') as any,
        val: c.attrs['val'],
      }));
      const colors = children(csNode, 'color').map(c => c.attrs['rgb'] ?? c.attrs['theme'] ?? '');
      cf.colorScale = { type: 'colorScale', cfvo: cfvos, color: colors };
    }

    // Parse dataBar
    const dbNode = child(rule, 'dataBar');
    if (dbNode) {
      const cfvos = children(dbNode, 'cfvo');
      const colorNode = child(dbNode, 'color');
      cf.dataBar = {
        type: 'dataBar',
        showValue: dbNode.attrs['showValue'] !== '0' ? undefined : false,
        minType: cfvos[0]?.attrs['type'] as any,
        minVal: cfvos[0]?.attrs['val'],
        maxType: cfvos[1]?.attrs['type'] as any,
        maxVal: cfvos[1]?.attrs['val'],
        color: colorNode?.attrs['rgb'],
      };
    }

    // Parse iconSet
    const isNode = child(rule, 'iconSet');
    if (isNode) {
      const cfvos = children(isNode, 'cfvo').map(c => ({
        type: c.attrs['type'] ?? 'percent',
        val: c.attrs['val'],
      }));
      cf.iconSet = {
        type: 'iconSet',
        iconSet: (isNode.attrs['iconSet'] ?? '3TrafficLights1') as any,
        cfvo: cfvos,
        showValue: isNode.attrs['showValue'] === '0' ? false : undefined,
        reverse: isNode.attrs['reverse'] === '1' ? true : undefined,
      };
    }

    ws.addConditionalFormat(cf);
  }
}

// ─── Data validation parsing ─────────────────────────────────────────────────

function parseDataValidations(node: XmlNode, ws: Worksheet): void {
  for (const dv of children(node, 'dataValidation')) {
    const sqref = dv.attrs['sqref'] ?? '';
    if (!sqref) continue;

    const type = (dv.attrs['type'] ?? 'whole') as DataValidation['type'];
    const val: DataValidation = { type };

    if (dv.attrs['operator']) val.operator = dv.attrs['operator'] as any;
    if (dv.attrs['allowBlank'] === '1') val.allowBlank = true;
    if (dv.attrs['showErrorMessage'] === '1') val.showErrorAlert = true;
    if (dv.attrs['errorTitle']) val.errorTitle = dv.attrs['errorTitle'];
    if (dv.attrs['error'])      val.error = dv.attrs['error'];
    if (dv.attrs['showInputMessage'] === '1') val.showInputMessage = true;
    if (dv.attrs['promptTitle']) val.promptTitle = dv.attrs['promptTitle'];
    if (dv.attrs['prompt'])      val.prompt = dv.attrs['prompt'];
    // showDropDown in OOXML means "suppress dropdown" (inverted semantics)
    if (dv.attrs['showDropDown'] === '1') val.showDropDown = false;

    const f1 = child(dv, 'formula1');
    const f2 = child(dv, 'formula2');
    if (f1?.text) {
      if (type === 'list' && f1.text.startsWith('"') && f1.text.endsWith('"')) {
        val.list = f1.text.slice(1, -1).split(',');
      } else {
        val.formula1 = f1.text;
      }
    }
    if (f2?.text) val.formula2 = f2.text;

    ws.addDataValidation(sqref, val);
  }
}

// ─── Table XML parsing ────────────────────────────────────────────────────────

function parseTableXml(xml: string): Table | null {
  try {
    const root = parseXml(xml);
    const tag = localName(root.tag);
    if (tag !== 'table') return null;

    const name        = root.attrs['name'] ?? '';
    const displayName = root.attrs['displayName'] ?? name;
    const ref         = root.attrs['ref'] ?? '';
    const totalsCount = parseInt(root.attrs['totalsRowCount'] ?? '0', 10);

    const columns: TableColumn[] = [];
    const colsNode = find(root, 'tableColumns');
    if (colsNode) {
      for (const col of children(colsNode, 'tableColumn')) {
        const tc: TableColumn = { name: col.attrs['name'] ?? '' };
        if (col.attrs['totalsRowFunction']) tc.totalsRowFunction = col.attrs['totalsRowFunction'] as any;
        if (col.attrs['totalsRowFormula']) tc.totalsRowFormula = col.attrs['totalsRowFormula'];
        if (col.attrs['totalsRowLabel'])  tc.totalsRowLabel = col.attrs['totalsRowLabel'];
        columns.push(tc);
      }
    }

    const table: Table = { name, ref, columns };
    if (displayName && displayName !== name) table.displayName = displayName;
    if (totalsCount > 0) table.totalsRow = true;

    const styleNode = find(root, 'tableStyleInfo');
    if (styleNode) {
      if (styleNode.attrs['name'])              table.style = styleNode.attrs['name'] as any;
      if (styleNode.attrs['showFirstColumn'] === '1')   table.showFirstColumn = true;
      if (styleNode.attrs['showLastColumn'] === '1')    table.showLastColumn = true;
      if (styleNode.attrs['showRowStripes'] === '1')    table.showRowStripes = true;
      if (styleNode.attrs['showColumnStripes'] === '1') table.showColumnStripes = true;
    }

    return table;
  } catch {
    return null;
  }
}

/** Resolve a relative path (e.g. "../tables/table1.xml") against a base directory */
function resolvePath(base: string, relative: string): string {
  const parts = base.replace(/\/$/, '').split('/');
  for (const seg of relative.split('/')) {
    if (seg === '..') parts.pop();
    else if (seg !== '.') parts.push(seg);
  }
  return parts.join('/');
}

// ─── Main reader ──────────────────────────────────────────────────────────────

export interface ReadResult {
  sheets: Array<{
    ws: Worksheet;
    sheetId: string;
    rId: string;
    originalXml: string;
    unknownParts: string[];
    /** Resolved paths of table XML files belonging to this sheet */
    tablePaths: string[];
    /** Original table XML strings (parallel to tablePaths) for verbatim round-trip */
    tableXmls: string[];
  }>;
  styles:         ParsedStyles;
  stylesXml:      string;       // original — for patching
  sharedStrings:  string[];
  sharedXml:      string;       // original
  workbookXml:    string;       // original
  workbookRels:   RelMap;
  contentTypes:   CTMap;
  contentTypesXml: string;
  core:           CoreProperties;
  extended:       ExtendedProperties;
  extendedUnknownRaw: string;
  custom:         CustomProperty[];
  hasCustomProps: boolean;
  /** All files from the ZIP that we don't otherwise handle — preserved verbatim */
  unknownParts:   Map<string, Uint8Array>;
  /** All relationship files (we need them to route images/charts/etc) */
  allRels:        Map<string, RelMap>;
}

export async function readWorkbook(data: Uint8Array): Promise<ReadResult> {
  const zip = await readZip(data);

  const get = (path: string) => {
    // Try with and without leading slash
    return zip.get(path) ?? zip.get(path.replace(/^\//, ''));
  };

  const getText = (path: string) => {
    const e = get(path);
    return e ? entryText(e) : undefined;
  };

  // Content types
  const ctXml = getText('[Content_Types].xml') ?? '<Types/>';
  const contentTypes = parseContentTypes(ctXml);

  // Workbook rels
  const wbRelsXml = getText('xl/_rels/workbook.xml.rels') ?? '<Relationships/>';
  const workbookRels = parseRels(wbRelsXml);

  // Workbook XML
  const wbXml = getText('xl/workbook.xml') ?? '<workbook/>';

  // Styles
  const stylesXml = getText('xl/styles.xml') ?? '<styleSheet/>';
  const styles = parseStyles(stylesXml);

  // Shared strings
  const ssXml = getText('xl/sharedStrings.xml') ?? '<sst/>';
  const sharedStrings = ssXml !== '<sst/>' ? parseSharedStrings(ssXml) : [];

  // Properties
  const coreXml = getText('docProps/core.xml') ?? '';
  const core: CoreProperties = coreXml ? parseCoreXml(coreXml) : {};

  const appXml = getText('docProps/app.xml') ?? '';
  let extended: ExtendedProperties = {};
  let extendedUnknownRaw = '';
  if (appXml) {
    const r = parseAppXml(appXml);
    extended = r.props;
    extendedUnknownRaw = r.unknownRaw;
  }

  const customXml = getText('docProps/custom.xml') ?? '';
  const custom: CustomProperty[] = customXml ? parseCustomXml(customXml) : [];

  // Parse sheet list from workbook.xml
  const wbRoot = parseXml(wbXml);
  const sheetsNode = find(wbRoot, 'sheets')!;
  const sheetNodes = sheetsNode ? children(sheetsNode, 'sheet') : [];

  // All rels files
  const allRels = new Map<string, RelMap>();
  for (const [path, entry] of zip) {
    if (path.includes('_rels/')) {
      allRels.set(path, parseRels(entryText(entry)));
    }
  }

  // Parse each sheet
  const sheets: ReadResult['sheets'] = [];
  for (const sn of sheetNodes) {
    const rId     = sn.attrs['r:id'] ?? Object.values(sn.attrs).find(v => v.startsWith('rId')) ?? '';
    const sheetId = sn.attrs['sheetId'] ?? '';
    const name    = sn.attrs['name'] ?? `Sheet${sheetId}`;
    const rel     = workbookRels.get(rId);
    if (!rel) continue;

    // Target is relative to xl/
    const target = rel.target.startsWith('/') ? rel.target.slice(1) : `xl/${rel.target}`;
    const sheetXml = getText(target) ?? '';
    if (!sheetXml) continue;

    const { ws, originalXml, unknownParts: sheetUnknown, tableRIds } = parseWorksheet(
      sheetXml, name, styles, sharedStrings,
    );
    ws.sheetIndex = sheets.length + 1;
    ws.rId = rId;

    // Resolve table references and parse table XML files
    const tablePaths: string[] = [];
    const tableXmls: string[] = [];
    if (tableRIds.length) {
      // Sheet rels file path: xl/worksheets/_rels/sheet<N>.xml.rels
      const sheetFileName = target.split('/').pop() ?? '';
      const sheetDir = target.substring(0, target.lastIndexOf('/') + 1);
      const sheetRelsPath = `${sheetDir}_rels/${sheetFileName}.rels`;
      const sheetRels = allRels.get(sheetRelsPath);
      if (sheetRels) {
        for (const tblRId of tableRIds) {
          const tblRel = sheetRels.get(tblRId);
          if (!tblRel) continue;
          // Resolve relative path (e.g. "../tables/table1.xml" relative to xl/worksheets/)
          const tblTarget = tblRel.target.startsWith('/')
            ? tblRel.target.slice(1)
            : resolvePath(sheetDir, tblRel.target);
          const tblXml = getText(tblTarget);
          if (tblXml) {
            const table = parseTableXml(tblXml);
            if (table) ws.addTable(table);
            tablePaths.push(tblTarget);
            tableXmls.push(tblXml);
          }
        }
        ws.tableRIds = tableRIds;
      }
    }

    sheets.push({ ws, sheetId, rId, originalXml, unknownParts: sheetUnknown, tablePaths, tableXmls });
  }

  // Collect truly unknown parts (not sheets, styles, strings, rels, content-types, props)
  const handledPrefixes = new Set([
    'xl/workbook.xml', 'xl/styles.xml', 'xl/sharedStrings.xml',
    'xl/worksheets/', 'docProps/', '[Content_Types].xml',
    '_rels/', 'xl/_rels/',
  ]);

  const unknownParts = new Map<string, Uint8Array>();
  for (const [path, entry] of zip) {
    if (path.endsWith('_rels/') || path === '[Content_Types].xml') continue;
    const isHandled = [...handledPrefixes].some(p => path.startsWith(p));
    if (!isHandled) {
      unknownParts.set(path, entry.data);
    }
  }

  return {
    sheets, styles, stylesXml, sharedStrings, sharedXml: ssXml,
    workbookXml: wbXml, workbookRels,
    contentTypes, contentTypesXml: ctXml,
    core, extended, extendedUnknownRaw, custom, hasCustomProps: custom.length > 0,
    unknownParts, allRels,
  };
}

