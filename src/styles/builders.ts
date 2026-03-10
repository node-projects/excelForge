/**
 * Fluent style builder helpers — composable utilities for common patterns.
 */

import type {
  CellStyle, Font, Fill, Border, Alignment, NumberFormat,
  Color, BorderStyle, FillPattern, HorizontalAlign, VerticalAlign,
} from '../core/types.js';

// ─── Style Builder ─────────────────────────────────────────────────────────

export class StyleBuilder {
  private s: CellStyle = {};

  font(props: Font): this {
    this.s.font = { ...(this.s.font ?? {}), ...props };
    return this;
  }

  fontSize(size: number): this { return this.font({ size }); }
  fontName(name: string): this { return this.font({ name }); }
  bold(v = true): this { return this.font({ bold: v }); }
  italic(v = true): this { return this.font({ italic: v }); }
  strike(v = true): this { return this.font({ strike: v }); }
  underline(style: Font['underline'] = 'single'): this { return this.font({ underline: style }); }
  fontColor(color: Color): this { return this.font({ color }); }

  fill(pattern: FillPattern, fgColor: Color, bgColor?: Color): this {
    this.s.fill = { type: 'pattern', pattern, fgColor, bgColor };
    return this;
  }

  bg(color: Color): this { return this.fill('solid', color); }

  border(style: BorderStyle, color?: Color): this {
    const side = { style, color };
    this.s.border = { left: side, right: side, top: side, bottom: side };
    return this;
  }

  borderLeft(style: BorderStyle, color?: Color): this {
    this.s.border = { ...(this.s.border ?? {}), left: { style, color } };
    return this;
  }
  borderRight(style: BorderStyle, color?: Color): this {
    this.s.border = { ...(this.s.border ?? {}), right: { style, color } };
    return this;
  }
  borderTop(style: BorderStyle, color?: Color): this {
    this.s.border = { ...(this.s.border ?? {}), top: { style, color } };
    return this;
  }
  borderBottom(style: BorderStyle, color?: Color): this {
    this.s.border = { ...(this.s.border ?? {}), bottom: { style, color } };
    return this;
  }

  align(horizontal: HorizontalAlign, vertical?: VerticalAlign): this {
    this.s.alignment = { ...(this.s.alignment ?? {}), horizontal, vertical };
    return this;
  }
  center(): this { return this.align('center', 'center'); }
  wrap(v = true): this {
    this.s.alignment = { ...(this.s.alignment ?? {}), wrapText: v };
    return this;
  }
  rotate(degrees: number): this {
    this.s.alignment = { ...(this.s.alignment ?? {}), textRotation: degrees };
    return this;
  }
  indent(n: number): this {
    this.s.alignment = { ...(this.s.alignment ?? {}), indent: n };
    return this;
  }
  shrink(v = true): this {
    this.s.alignment = { ...(this.s.alignment ?? {}), shrinkToFit: v };
    return this;
  }

  numFmt(formatCode: string): this {
    this.s.numberFormat = { formatCode };
    return this;
  }
  numFmtId(id: number): this {
    this.s.numFmtId = id;
    return this;
  }

  locked(v = true): this { this.s.locked = v; return this; }
  hidden(v = true): this { this.s.hidden = v; return this; }

  build(): CellStyle { return { ...this.s }; }
}

export function style(): StyleBuilder { return new StyleBuilder(); }

// ─── Pre-built common styles ─────────────────────────────────────────────

export const Styles = {
  /** Bold, centered header with blue background */
  headerBlue: style().bold().bg('FF4472C4').fontColor('FFFFFFFF').center().build(),
  /** Bold, centered header with dark gray background */
  headerGray: style().bold().bg('FF595959').fontColor('FFFFFFFF').center().build(),
  /** Bold, green header */
  headerGreen: style().bold().bg('FF70AD47').fontColor('FFFFFFFF').center().build(),
  /** Currency with 2 decimal places */
  currency: style().numFmt('"$"#,##0.00').build(),
  /** Percentage */
  percent: style().numFmt('0.00%').build(),
  /** Integer with thousands separator */
  integer: style().numFmt('#,##0').build(),
  /** Short date */
  date: style().numFmtId(14).build(),
  /** DateTime */
  dateTime: style().numFmtId(22).build(),
  /** Bold */
  bold: style().bold().build(),
  /** Centered */
  centered: style().center().build(),
  /** Wrapped text */
  wrapped: style().wrap().build(),
  /** All thin borders */
  bordered: style().border('thin').build(),
  /** Thin all-black borders */
  borderedBlack: style().border('thin', 'FF000000').build(),
  /** Light yellow fill */
  highlight: style().bg('FFFFFF00').build(),
  /** Light blue fill */
  lightBlue: style().bg('FFDCE6F1').build(),
  /** Red text */
  redText: style().fontColor('FFFF0000').build(),
  /** Blue text */
  blueText: style().fontColor('FF0000FF').build(),
  /** Green text */
  greenText: style().fontColor('FF00B050').build(),
  /** Bold + border + centered */
  tableHeader: style().bold().border('thin').center().bg('FFD9E1F2').build(),
  /** Strikethrough for deleted items */
  deleted: style().strike().fontColor('FF808080').build(),
};

// ─── Number format strings ───────────────────────────────────────────────

export const NumFmt = {
  General:       'General',
  Integer:       '#,##0',
  Decimal2:      '#,##0.00',
  Currency:      '"$"#,##0.00',
  CurrencyNeg:   '"$"#,##0.00;[Red]("$"#,##0.00)',
  Percent:       '0%',
  Percent2:      '0.00%',
  Scientific:    '0.00E+00',
  Fraction:      '# ?/?',
  ShortDate:     'mm/dd/yyyy',
  LongDate:      'mmmm d, yyyy',
  Time:          'h:mm AM/PM',
  DateTime:      'mm/dd/yyyy h:mm',
  Text:          '@',
  Accounting:    '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)',
  FinancialPos:  '#,##0;(#,##0)',
  ZeroDash:      '#,##0;(#,##0);"-"',
  Multiple:      '0.0"x"',
} as const;

// ─── Color palette ────────────────────────────────────────────────────────

export const Colors = {
  White:       'FFFFFFFF',
  Black:       'FF000000',
  Red:         'FFFF0000',
  Green:       'FF00B050',
  Blue:        'FF0070C0',
  Yellow:      'FFFFFF00',
  Orange:      'FFFFA500',
  Gray:        'FF808080',
  LightGray:   'FFD3D3D3',
  DarkBlue:    'FF003366',
  ExcelBlue:   'FF4472C4',
  ExcelOrange: 'FFED7D31',
  ExcelGreen:  'FF70AD47',
  ExcelRed:    'FFFF0000',
  ExcelPurple: 'FF7030A0',
  Transparent: '00FFFFFF',
} as const;
