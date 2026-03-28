// ─── Color ────────────────────────────────────────────────────────────────────

/** ARGB hex string, e.g. "FF0070C0" */
export type Color = string;

// ─── Border ───────────────────────────────────────────────────────────────────

export type BorderStyle =
  | 'thin' | 'medium' | 'thick' | 'dashed' | 'dotted'
  | 'double' | 'hair' | 'mediumDashed' | 'dashDot'
  | 'mediumDashDot' | 'dashDotDot' | 'mediumDashDotDot' | 'slantDashDot';

export interface BorderSide {
  style?: BorderStyle;
  color?: Color;
}

export interface Border {
  left?:     BorderSide;
  right?:    BorderSide;
  top?:      BorderSide;
  bottom?:   BorderSide;
  diagonal?: BorderSide;
  diagonalUp?:   boolean;
  diagonalDown?: boolean;
}

// ─── Font ─────────────────────────────────────────────────────────────────────

export interface Font {
  name?:      string;
  size?:      number;
  bold?:      boolean;
  italic?:    boolean;
  underline?: 'single' | 'double' | 'singleAccounting' | 'doubleAccounting' | 'none';
  strike?:    boolean;
  color?:     Color;
  scheme?:    'minor' | 'major' | 'none';
  charset?:   number;
  family?:    number;
  vertAlign?: 'superscript' | 'subscript' | 'baseline';
}

// ─── Fill ─────────────────────────────────────────────────────────────────────

export type FillPattern =
  | 'solid' | 'none' | 'gray125' | 'gray0625'
  | 'darkGray' | 'mediumGray' | 'lightGray'
  | 'darkHorizontal' | 'darkVertical' | 'darkDown' | 'darkUp'
  | 'darkGrid' | 'darkTrellis'
  | 'lightHorizontal' | 'lightVertical' | 'lightDown' | 'lightUp'
  | 'lightGrid' | 'lightTrellis';

export interface PatternFill {
  type: 'pattern';
  pattern: FillPattern;
  fgColor?: Color;
  bgColor?: Color;
}

export type GradientType = 'linear' | 'path';

export interface GradientStop { position: number; color: Color; }

export interface GradientFill {
  type: 'gradient';
  gradientType?: GradientType;
  degree?: number;
  left?: number; right?: number; top?: number; bottom?: number;
  stops: GradientStop[];
}

export type Fill = PatternFill | GradientFill;

// ─── Alignment ────────────────────────────────────────────────────────────────

export type HorizontalAlign = 'general' | 'left' | 'center' | 'right' | 'fill' | 'justify' | 'centerContinuous' | 'distributed';
export type VerticalAlign   = 'top' | 'center' | 'bottom' | 'justify' | 'distributed';

export interface Alignment {
  horizontal?:   HorizontalAlign;
  vertical?:     VerticalAlign;
  wrapText?:     boolean;
  shrinkToFit?:  boolean;
  textRotation?: number;   // 0–180 (90 = upward, 255 = stacked)
  indent?:       number;
  readingOrder?: 0 | 1 | 2;
}

// ─── Number Format ────────────────────────────────────────────────────────────

export interface NumberFormat {
  /** Custom format string, e.g. '#,##0.00' */
  formatCode: string;
}

// ─── Cell Style ───────────────────────────────────────────────────────────────

export interface CellStyle {
  font?:         Font;
  fill?:         Fill;
  border?:       Border;
  alignment?:    Alignment;
  numberFormat?: NumberFormat;
  /** Built-in number format ID (0-49 range) */
  numFmtId?:     number;
  locked?:       boolean;
  hidden?:       boolean;
}

// ─── Cell Value ───────────────────────────────────────────────────────────────

export type CellValue = string | number | boolean | Date | null | undefined;

// ─── Rich Text ────────────────────────────────────────────────────────────────

export interface RichTextRun {
  text:   string;
  font?:  Font;
}

// ─── Cell ─────────────────────────────────────────────────────────────────────

export interface Cell {
  value?:       CellValue;
  formula?:     string;        // e.g. "SUM(A1:A10)"
  arrayFormula?: string;       // array formula
  richText?:    RichTextRun[];
  style?:       CellStyle;
  comment?:     Comment;
  hyperlink?:   Hyperlink;
  /** Data validation rule on this cell */
  validation?:  DataValidation;
}

// ─── Comment ──────────────────────────────────────────────────────────────────

export interface Comment {
  text:       string;
  author?:    string;
}

// ─── Hyperlink ────────────────────────────────────────────────────────────────

export interface Hyperlink {
  href:     string;
  tooltip?: string;
}

// ─── Data Validation ──────────────────────────────────────────────────────────

export type ValidationOperator =
  | 'between' | 'notBetween' | 'equal' | 'notEqual'
  | 'lessThan' | 'lessThanOrEqual' | 'greaterThan' | 'greaterThanOrEqual';

export type ValidationType = 'whole' | 'decimal' | 'list' | 'date' | 'time' | 'textLength' | 'custom';

export interface DataValidation {
  type:           ValidationType;
  operator?:      ValidationOperator;
  formula1?:      string;
  formula2?:      string;
  /** Comma-separated list of values for 'list' type */
  list?:          string[];
  showDropDown?:  boolean;
  showErrorAlert?: boolean;
  errorTitle?:    string;
  error?:         string;
  showInputMessage?: boolean;
  promptTitle?:   string;
  prompt?:        string;
  allowBlank?:    boolean;
}

// ─── Merge ────────────────────────────────────────────────────────────────────

export interface MergeRange {
  startRow: number; startCol: number;
  endRow:   number; endCol:   number;
}

// ─── Image ────────────────────────────────────────────────────────────────────

export type ImageFormat = 'png' | 'jpeg' | 'gif' | 'emf' | 'wmf' | 'tiff';

export interface ImagePosition {
  col:      number;
  row:      number;
  colOff?:  number;  // EMU offset
  rowOff?:  number;
}

export interface Image {
  data:      Uint8Array | string;  // raw bytes or base64
  format:    ImageFormat;
  from:      ImagePosition;
  to?:       ImagePosition;
  /** Absolute size in pixels (used when 'to' is omitted) */
  width?:    number;
  height?:   number;
  altText?:  string;
}

// ─── Chart ────────────────────────────────────────────────────────────────────

export type ChartType =
  | 'bar' | 'barStacked' | 'barStacked100'
  | 'column' | 'columnStacked' | 'columnStacked100'
  | 'line' | 'lineStacked' | 'lineMarker'
  | 'area' | 'areaStacked'
  | 'pie' | 'doughnut'
  | 'scatter' | 'scatterSmooth'
  | 'bubble' | 'radar' | 'radarFilled'
  | 'stock';

export interface ChartSeries {
  name?:    string;
  /** Sheet ref like "Sheet1!$A$2:$A$10" */
  values:   string;
  /** Category ref */
  categories?: string;
  color?:   Color;
}

export interface ChartAxis {
  title?:    string;
  min?:      number;
  max?:      number;
  gridLines?: boolean;
  numFmt?:   string;
}

export interface ChartPosition {
  col: number; row: number;
  colOff?: number; rowOff?: number;
}

export interface Chart {
  type:      ChartType;
  title?:    string;
  series:    ChartSeries[];
  from:      ChartPosition;
  to:        ChartPosition;
  xAxis?:    ChartAxis;
  yAxis?:    ChartAxis;
  legend?:   boolean | 'top' | 'bottom' | 'left' | 'right' | 'b' | 't' | 'l' | 'r';
  style?:    number;   // built-in chart style 1-48
  varyColors?: boolean;
  grouping?:  string;
}

// ─── Conditional Formatting ───────────────────────────────────────────────────

export type CFType =
  | 'cellIs' | 'containsText' | 'notContainsText' | 'beginsWith' | 'endsWith'
  | 'expression' | 'colorScale' | 'dataBar' | 'iconSet'
  | 'top10' | 'aboveAverage' | 'duplicateValues' | 'uniqueValues'
  | 'containsBlanks' | 'notContainsBlanks' | 'containsErrors' | 'notContainsErrors'
  | 'timePeriod';

export interface CFColorScale {
  type: 'colorScale';
  cfvo: Array<{ type: 'min'|'max'|'percent'|'num'|'formula'; val?: string }>;
  color: Color[];
}

export interface CFDataBar {
  type: 'dataBar';
  /** Bar fill color (preferred) */
  color?: Color;
  /** @deprecated use color */
  minColor?: Color;
  /** @deprecated use color */
  maxColor?: Color;
  /** cfvo type for min bound (default: 'min') */
  minType?: 'min' | 'max' | 'percent' | 'num' | 'formula';
  /** cfvo value for min bound */
  minVal?: string | number;
  /** cfvo type for max bound (default: 'max') */
  maxType?: 'min' | 'max' | 'percent' | 'num' | 'formula';
  /** cfvo value for max bound */
  maxVal?: string | number;
  showValue?: boolean;
}

export type IconSet = '3Arrows'|'3ArrowsGray'|'3Flags'|'3TrafficLights1'|'3TrafficLights2'|
  '3Signs'|'3Symbols'|'3Symbols2'|'4Arrows'|'4ArrowsGray'|'4RedToBlack'|
  '4Rating'|'4TrafficLights'|'5Arrows'|'5ArrowsGray'|'5Rating'|'5Quarters';

export interface CFIconSet {
  type: 'iconSet';
  iconSet: IconSet;
  cfvo: Array<{ type: string; val?: string }>;
  showValue?: boolean;
  reverse?: boolean;
}

export interface ConditionalFormat {
  sqref:     string;   // e.g. "A1:A10"
  type:      CFType;
  operator?: ValidationOperator;
  formula?:  string;
  formula2?: string;
  text?:     string;
  priority?: number;
  style?:    CellStyle;
  colorScale?: CFColorScale;
  dataBar?:   CFDataBar;
  iconSet?:   CFIconSet;
  aboveAverage?: boolean;
  percent?:   boolean;
  rank?:      number;
  timePeriod?: string;
}

// ─── Table ────────────────────────────────────────────────────────────────────

export type TableStyle =
  | 'TableStyleLight1' | 'TableStyleLight2' | 'TableStyleLight3'
  | 'TableStyleLight4' | 'TableStyleLight5' | 'TableStyleLight6'
  | 'TableStyleMedium1' | 'TableStyleMedium2' | 'TableStyleMedium3'
  | 'TableStyleMedium4' | 'TableStyleMedium5' | 'TableStyleMedium6'
  | 'TableStyleMedium7' | 'TableStyleMedium8' | 'TableStyleMedium9'
  | 'TableStyleDark1' | 'TableStyleDark2' | 'TableStyleDark3'
  | 'TableStyleDark4' | 'TableStyleDark5' | 'TableStyleDark6'
  | 'TableStyleDark7' | 'TableStyleDark8' | 'TableStyleDark9'
  | string;

export interface TableColumn {
  name:          string;
  totalsRowFunction?: 'sum'|'count'|'average'|'max'|'min'|'stdDev'|'var'|'vars'|'countNums'|'custom'|'none';
  totalsRowFormula?: string;
  totalsRowLabel?: string;
  filterButton?:  boolean;
  style?:         CellStyle;
  numFmt?:        string;
}

export interface Table {
  name:            string;
  displayName?:    string;
  ref:             string;   // e.g. "A1:D10"
  style?:          TableStyle;
  showFirstColumn?: boolean;
  showLastColumn?:  boolean;
  showRowStripes?:  boolean;
  showColumnStripes?: boolean;
  totalsRow?:       boolean;
  columns:          TableColumn[];
}

// ─── Named Range ──────────────────────────────────────────────────────────────

export interface NamedRange {
  name:   string;
  /** e.g. "Sheet1!$A$1:$D$10" */
  ref:    string;
  scope?: string;  // sheet name for local scope
  comment?: string;
}

// ─── Page Setup ───────────────────────────────────────────────────────────────

export type PaperSize = 1|2|3|4|5|6|7|8|9|10|11|12|13|14|15|16|17;

export type Orientation = 'portrait' | 'landscape';

export interface PageSetup {
  paperSize?:      PaperSize;
  orientation?:    Orientation;
  fitToPage?:      boolean;
  fitToWidth?:     number;
  fitToHeight?:    number;
  scale?:          number;
  horizontalDpi?:  number;
  verticalDpi?:    number;
  firstPageNumber?: number;
}

export interface PageMargins {
  left?:   number;
  right?:  number;
  top?:    number;
  bottom?: number;
  header?: number;
  footer?: number;
}

// ─── Page Breaks ─────────────────────────────────────────────────────────────

export interface PageBreak {
  /** Row (for rowBreaks) or column (for colBreaks) index, 1-based */
  id:     number;
  /** Whether this is a manual break (default true) */
  manual?: boolean;
}

// ─── Header/Footer ────────────────────────────────────────────────────────────

export interface HeaderFooter {
  oddHeader?:   string;
  oddFooter?:   string;
  evenHeader?:  string;
  evenFooter?:  string;
  firstHeader?: string;
  firstFooter?: string;
  differentOddEven?: boolean;
  differentFirst?:   boolean;
}

// ─── Print Options ────────────────────────────────────────────────────────────

export interface PrintOptions {
  gridLines?:       boolean;
  gridLinesSet?:    boolean;
  headings?:        boolean;
  centerHorizontal?: boolean;
  centerVertical?:   boolean;
}

// ─── Column / Row ─────────────────────────────────────────────────────────────

export interface ColumnDef {
  width?:        number;
  hidden?:       boolean;
  outlineLevel?: number;
  collapsed?:    boolean;
  style?:        CellStyle;
  bestFit?:      boolean;
  customWidth?:  boolean;
}

export interface RowDef {
  height?:       number;
  hidden?:       boolean;
  outlineLevel?: number;
  collapsed?:    boolean;
  style?:        CellStyle;
  thickTop?:     boolean;
  thickBot?:     boolean;
}

// ─── Freeze / Split Pane ──────────────────────────────────────────────────────

export interface FreezePane {
  col?: number;
  row?: number;
}

export interface SplitPane {
  xSplit?: number;
  ySplit?: number;
  topLeftCell?: string;
  activePane?: 'topLeft'|'topRight'|'bottomLeft'|'bottomRight';
  state?: 'split'|'frozen'|'frozenSplit';
}

// ─── Sheet Protection ─────────────────────────────────────────────────────────

export interface SheetProtection {
  password?:         string;
  sheet?:            boolean;
  selectLockedCells?: boolean;
  selectUnlockedCells?: boolean;
  formatCells?:      boolean;
  formatColumns?:    boolean;
  formatRows?:       boolean;
  insertColumns?:    boolean;
  insertRows?:       boolean;
  insertHyperlinks?: boolean;
  deleteColumns?:    boolean;
  deleteRows?:       boolean;
  sort?:             boolean;
  autoFilter?:       boolean;
  pivotTables?:      boolean;
}

// ─── Auto Filter ──────────────────────────────────────────────────────────────

export interface AutoFilter {
  ref: string;  // e.g. "A1:D1"
}

// ─── Sheet View ───────────────────────────────────────────────────────────────

export interface SheetView {
  showGridLines?:     boolean;
  showRowColHeaders?: boolean;
  zoomScale?:         number;
  rightToLeft?:       boolean;
  tabSelected?:       boolean;
  showRuler?:         boolean;
  view?:              'normal' | 'pageLayout' | 'pageBreakPreview';
}

// ─── Sparklines ───────────────────────────────────────────────────────────────

export type SparklineType = 'line' | 'bar' | 'stacked';

export interface Sparkline {
  type:       SparklineType;
  dataRange:  string;   // source data ref
  location:   string;   // single cell ref
  color?:     Color;
  highColor?: Color;
  lowColor?:  Color;
  firstColor?: Color;
  lastColor?:  Color;
  negativeColor?: Color;
  markersColor?:  Color;
  showMarkers?:   boolean;
  showFirst?:     boolean;
  showLast?:      boolean;
  showHigh?:      boolean;
  showLow?:       boolean;
  showNegative?:  boolean;
  minAxisType?:   'individual' | 'custom' | 'group';
  maxAxisType?:   'individual' | 'custom' | 'group';
  lineWidth?:     number;
}

// ─── Pivot Table ──────────────────────────────────────────────────────────────

export type PivotFunction =
  | 'sum' | 'count' | 'average' | 'max' | 'min'
  | 'product' | 'countNums' | 'stdDev' | 'stdDevp' | 'var' | 'varp';

export interface PivotDataField {
  /** Source field name (must match a column header in the source range) */
  field: string;
  /** Display name shown in the pivot table (defaults to "Sum of <field>") */
  name?: string;
  /** Aggregation function (default: 'sum') */
  func?: PivotFunction;
}

export interface PivotTable {
  /** Unique name for the pivot table, e.g. "PivotTable1" */
  name: string;
  /** Name of the sheet containing the source data */
  sourceSheet: string;
  /** Source data range including header row, e.g. "A1:D10" */
  sourceRef: string;
  /** Cell address of the pivot table's top-left corner, e.g. "F1" */
  targetCell: string;
  /** Field names to display as row labels (in order) */
  rowFields: string[];
  /** Field names to display as column labels (in order) */
  colFields: string[];
  /** Fields to aggregate in the values area */
  dataFields: PivotDataField[];
  /** Pivot table style name (default: "PivotStyleMedium9") */
  style?: string;
  /** Show grand totals for rows (default: true) */
  rowGrandTotals?: boolean;
  /** Show grand totals for columns (default: true) */
  colGrandTotals?: boolean;
}

// ─── Connections & Power Query ────────────────────────────────────────────────

/** OOXML connection type */
export type ConnectionType = 'odbc' | 'dao' | 'file' | 'web' | 'oledb' | 'text' | 'dsp';

/** OOXML command type for database connections */
export type CommandType = 'sql' | 'table' | 'default' | 'web' | 'oledb';

export interface Connection {
  /** Unique connection ID */
  id:            number;
  /** Display name */
  name:          string;
  /** Connection type */
  type:          ConnectionType;
  /** Connection string (for OLEDB/ODBC) */
  connectionString?: string;
  /** SQL command text */
  command?:      string;
  /** Command type (default 'table') */
  commandType?:  CommandType;
  /** Description */
  description?:  string;
  /** Refresh on open? */
  refreshOnLoad?: boolean;
  /** Background refresh? */
  background?:   boolean;
  /** Save cached data with the workbook? */
  saveData?:     boolean;
  /** Keep connection alive between refreshes? */
  keepAlive?:    boolean;
  /** Interval between auto-refreshes in minutes (0 = no auto-refresh) */
  interval?:     number;
  /** Raw XML string for preserving unrecognized attributes during round-trip */
  _rawXml?:      string;
}

export interface PowerQuery {
  /** Query name (appears in Excel's Queries pane) */
  name: string;
  /** Power Query M formula code */
  formula: string;
}

// ─── Workbook Options ─────────────────────────────────────────────────────────

export interface WorkbookProperties {
  title?:   string;
  author?:  string;
  company?: string;
  subject?: string;
  description?: string;
  keywords?: string;
  created?:  Date;
  modified?: Date;
  lastModifiedBy?: string;
  category?: string;
  status?:   string;
  /** 1904 date system */
  date1904?: boolean;
}

// ─── Worksheet Options ────────────────────────────────────────────────────────

export interface WorksheetOptions {
  name?:          string;
  tabColor?:      Color;
  state?:         'visible' | 'hidden' | 'veryHidden';
  codeName?:      string;
  defaultRowHeight?: number;
  defaultColWidth?:  number;
  outlineSummaryBelow?: boolean;
  outlineSummaryRight?: boolean;
}
