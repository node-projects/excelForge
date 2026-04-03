/**
 * ExcelForge — A complete TypeScript library for creating Excel (.xlsx) files.
 * Zero external dependencies. Works in browsers, Node.js, Deno, Bun, and edge runtimes.
 *
 * @example
 * ```typescript
 * import { Workbook, style, Colors, NumFmt } from './index.js';
 *
 * const wb = new Workbook();
 * wb.properties.title = 'My Report';
 *
 * const ws = wb.addSheet('Sales');
 * ws.setValue(1, 1, 'Product').setStyle(1, 1, style().bold().bg(Colors.ExcelBlue).fontColor(Colors.White).build());
 * ws.setValue(1, 2, 'Revenue').setStyle(1, 2, style().bold().bg(Colors.ExcelBlue).fontColor(Colors.White).build());
 * ws.setValue(2, 1, 'Widget A');
 * ws.setValue(2, 2, 1234.56).setStyle(2, 2, style().numFmt(NumFmt.Currency).build());
 * ws.setFormula(3, 2, 'SUM(B2:B2)');
 *
 * // Browser download
 * await wb.download('report.xlsx');
 *
 * // Node.js
 * await wb.writeFile('./report.xlsx');
 * ```
 */

// ── Core ────────────────────────────────────────────────────────────────────
export { Workbook }     from './core/Workbook.js';
export { Worksheet }    from './core/Worksheet.js';
export { SharedStrings } from './core/SharedStrings.js';
export { StyleRegistry } from './styles/StyleRegistry.js';
export { CellError }    from './core/types.js';

// ── Builders & helpers ──────────────────────────────────────────────────────
export { style, StyleBuilder, Styles, Colors, NumFmt } from './styles/builders.js';
export {
  a1ToR1C1, r1c1ToA1, formulaToR1C1, formulaFromR1C1,
} from './utils/helpers.js';

// ── CSV & JSON modules (tree-shakeable) ─────────────────────────────────────
export { worksheetToCsv, csvToWorkbook } from './features/CsvModule.js';
export type { CsvOptions } from './features/CsvModule.js';
export { worksheetToJson, workbookToJson } from './features/JsonModule.js';
export type { JsonExportOptions } from './features/JsonModule.js';

// ── HTML export module (tree-shakeable) ─────────────────────────────────────
export { worksheetToHtml, workbookToHtml } from './features/HtmlModule.js';
export type { HtmlExportOptions, WorkbookHtmlExportOptions } from './features/HtmlModule.js';

// ── Formula Calculation Engine (tree-shakeable) ──────────────────────────────
export { FormulaEngine } from './features/FormulaEngine.js';

// ── Types ───────────────────────────────────────────────────────────────────
export type {
  // Values
  CellValue,
  ErrorValue,
  Cell,
  RichTextRun,

  // Styles
  CellStyle,
  Font,
  Fill,
  PatternFill,
  GradientFill,
  GradientStop,
  Border,
  BorderSide,
  BorderStyle,
  Alignment,
  NumberFormat,
  Color,
  FillPattern,
  HorizontalAlign,
  VerticalAlign,

  // Sheet features
  MergeRange,
  Image,
  CellImage,
  ImageFormat,
  ImagePosition,
  Chart,
  ChartType,
  ChartSeries,
  ChartAxis,
  ChartPosition,
  ConditionalFormat,
  CFType,
  CFColorScale,
  CFDataBar,
  CFIconSet,
  IconSet,
  Table,
  TableColumn,
  TableStyle,
  PivotTable,
  PivotDataField,
  PivotFunction,
  Sparkline,
  SparklineType,
  DataValidation,
  ValidationType,
  ValidationOperator,
  AutoFilter,
  Comment,
  Hyperlink,
  FormControl,
  FormControlType,
  FormControlAnchor,
  NamedRange,
  Connection,
  ConnectionType,
  CommandType,
  PowerQuery,

  // Sheet layout
  ColumnDef,
  RowDef,
  FreezePane,
  SplitPane,
  SheetProtection,
  PageSetup,
  PageMargins,
  PageBreak,
  HeaderFooter,
  PrintOptions,
  SheetView,
  PaperSize,
  Orientation,

  // Workbook
  WorkbookProperties,
  WorksheetOptions,

  // New features
  Shape,
  ShapeType,
  WordArt,
  WordArtPreset,
  Theme,
  ThemeColor,
  CustomIcon,
  CFCustomIconSet,
  TableSlicer,
  PivotSlicer,
  CustomPivotStyle,
  PivotStyleElement,
  QueryTable,
  ExternalLink,
  LocaleSettings,
} from './core/types.js';

// ── VBA ──────────────────────────────────────────────────────────────────────
export { VbaProject } from './vba/VbaProject.js';
export type { VbaModule, VbaModuleType } from './vba/VbaProject.js';

// ── Utility functions (re-exported for advanced users) ──────────────────────
export {
  colIndexToLetter,
  colLetterToIndex,
  cellRefToIndices,
  indicesToCellRef,
  parseRange,
  dateToSerial,
  pxToEmu,
  base64ToBytes,
  bytesToBase64,
} from './utils/helpers.js';
