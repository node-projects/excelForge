# ExcelForge Feature Comparison

Compared against:
- **EPPlus 8** (.NET) — https://www.epplussoftware.com/en/Developers/Features
- **SheetJS Pro** (JS) — https://sheetjs.com/pro/
- **ExcelJS** (JS/Node) — https://github.com/exceljs/exceljs
- **ExcelTS** (TS) — https://github.com/cjnoname/excelts
- **ExcelForge** (TS) — https://github.com/nickmanning214/ExcelForge

Legend: **Y** = supported, **~** = partial, **-** = not supported, **P** = preserved on round-trip only

---

## Core Read/Write

| # | Feature | EPPlus | SheetJS Pro | ExcelJS | ExcelTS | ExcelForge | Notes |
|---|---------|--------|-------------|---------|---------|------------|-------|
| 1 | Read/write .xlsx | Y | Y | Y | Y | **Y** | |
| 2 | Read/write .xlsm (VBA macros) | Y | Y | - | - | **Y** | ExcelForge: create/edit modules, full round-trip |
| 3 | Read .xltx templates | Y | Y | - | - | **Y** | isTemplate flag for write; reads natively |
| 4 | Read/write CSV | Y | Y | Y | Y | **Y** | Tree-shakeable CSV module |
| 5 | Export JSON | Y | Y | Y | - | **Y** | Tree-shakeable JSON module |
| 6 | Export HTML/CSS | Y | Y | - | - | **Y** | Enhanced: number fmts, CF viz, sparklines, charts, column widths, multi-sheet tabs |
| 7 | Streaming read/write | Y (async) | Y | Y | Y | **-** | ExcelTS: WorkbookReader/WorkbookWriter |
| 8 | Workbook encryption/decryption | Y | Y | - | - | **-** | EPPlus: Standard + Agile encryption |
| 9 | Digital signatures | Y | - | - | - | **-** | EPPlus: 3 sig types, 5 hash algos |

## Cell Values & Formulas

| # | Feature | EPPlus | SheetJS Pro | ExcelJS | ExcelTS | ExcelForge | Notes |
|---|---------|--------|-------------|---------|---------|------------|-------|
| 10 | Strings, numbers, booleans, dates | Y | Y | Y | Y | **Y** | |
| 11 | Rich text in cells | Y | Y | Y | Y | **Y** | |
| 12 | Formulas (store & preserve) | Y | Y | Y | Y | **Y** | |
| 13 | Formula calculation engine | Y (463 fns) | Y | - | - | **Y** | Tree-shakeable; 60+ functions |
| 14 | Array formulas | Y | Y | Y | - | **Y** | |
| 15 | Dynamic array formulas | Y | - | - | - | **Y** | setDynamicArrayFormula API |
| 16 | Shared formulas | Y | Y | Y | - | **Y** | setSharedFormula API |
| 17 | R1C1 reference style | Y | - | - | - | **Y** | a1ToR1C1, r1c1ToA1, formula converters |
| 18 | Hyperlinks | Y | Y | Y | Y | **Y** | |
| 19 | Error values | Y | Y | Y | - | **Y** | CellError class with typed API |

## Styling

| # | Feature | EPPlus | SheetJS Pro | ExcelJS | ExcelTS | ExcelForge | Notes |
|---|---------|--------|-------------|---------|---------|------------|-------|
| 20 | Number formats | Y | Y | Y | Y | **Y** | 30+ presets |
| 21 | Fonts (bold, italic, color, etc.) | Y | Y | Y | Y | **Y** | |
| 22 | Fills (solid, pattern, gradient) | Y | Y | Y | Y | **Y** | |
| 23 | Borders (all styles) | Y | Y | Y | Y | **Y** | |
| 24 | Alignment (h/v, wrap, rotation) | Y | Y | Y | Y | **Y** | |
| 25 | Named/cell styles | Y | Y | - | - | **Y** | registerNamedStyle API |
| 26 | Themes (load .thmx) | Y | - | - | - | **-** | |

## Layout & Structure

| # | Feature | EPPlus | SheetJS Pro | ExcelJS | ExcelTS | ExcelForge | Notes |
|---|---------|--------|-------------|---------|---------|------------|-------|
| 27 | Merge cells | Y | Y | Y | Y | **Y** | |
| 28 | Freeze/split panes | Y | Y | Y | Y | **Y** | |
| 29 | Column widths / row heights | Y | Y | Y | Y | **Y** | |
| 30 | Hide rows/columns | Y | Y | Y | - | **Y** | |
| 31 | Outline grouping (collapse/expand) | Y | Y | Y | - | **Y** | |
| 32 | AutoFit columns | Y | - | - | - | **Y** | Char-count approximation |
| 33 | Multiple sheets (hidden/veryHidden) | Y | Y | Y | Y | **Y** | |
| 34 | Tab colors | Y | Y | Y | - | **Y** | |
| 35 | Copy worksheets | Y | - | - | - | **Y** | Copies cells, styles, merges |
| 36 | Copy/move ranges | Y | - | - | - | **Y** | copyRange, moveRange |
| 37 | Insert/delete ranges (auto-shift) | Y | - | Y | - | **Y** | insertRows, deleteRows, insertColumns |
| 38 | Sort ranges | Y | - | - | - | **Y** | sortRange with asc/desc |
| 39 | Fill operations | Y | - | - | - | **Y** | fillNumber, fillDate, fillList |
| 40 | Named ranges (workbook + sheet) | Y | Y | Y | - | **Y** | |
| 41 | Print areas | Y | - | - | - | **Y** | Via printArea property |

## Tables

| # | Feature | EPPlus | SheetJS Pro | ExcelJS | ExcelTS | ExcelForge | Notes |
|---|---------|--------|-------------|---------|---------|------------|-------|
| 42 | Styled Excel tables | Y (60 styles) | Y | Y | Y | **Y** | 27 built-in presets |
| 43 | Totals row | Y | - | - | - | **Y** | |
| 44 | Custom table styles | Y | - | - | - | **Y** | registerTableStyle with DXF |
| 45 | Table slicers | Y | - | - | - | **-** | |

## Conditional Formatting

| # | Feature | EPPlus | SheetJS Pro | ExcelJS | ExcelTS | ExcelForge | Notes |
|---|---------|--------|-------------|---------|---------|------------|-------|
| 46 | Cell rules | Y (45 types) | Y | Y | Y | **Y** | |
| 47 | Color scales | Y | Y | Y | - | **Y** | |
| 48 | Data bars | Y | Y | Y | - | **Y** | |
| 49 | Icon sets | Y | Y | Y | - | **Y** | |
| 50 | Custom icon sets | Y | - | - | - | **-** | |
| 51 | Cross-worksheet references | Y | - | - | - | **-** | |

## Data Validation

| # | Feature | EPPlus | SheetJS Pro | ExcelJS | ExcelTS | ExcelForge | Notes |
|---|---------|--------|-------------|---------|---------|------------|-------|
| 52 | Dropdowns, whole/decimal, date, time | Y | Y | Y | Y | **Y** | |
| 53 | Text length, custom formula | Y | Y | Y | Y | **Y** | |

## Pivot Tables

| # | Feature | EPPlus | SheetJS Pro | ExcelJS | ExcelTS | ExcelForge | Notes |
|---|---------|--------|-------------|---------|---------|------------|-------|
| 54 | Row/column/data fields | Y | Y | ~ | Y | **Y** | |
| 55 | Aggregation functions | Y (12 types) | - | - | - | **Y** | sum, count, avg, max, min... |
| 56 | Styles (84 presets) | Y | - | - | - | **~** | ExcelForge: built-in presets only |
| 57 | Custom pivot styles | Y | - | - | - | **-** | |
| 58 | Pivot table slicers | Y | - | - | - | **-** | |
| 59 | Calculated fields | Y | - | - | - | **Y** | calculatedFields on PivotTable |
| 60 | Numeric/date grouping | Y | - | - | - | **Y** | fieldGrouping on PivotTable |
| 61 | GETPIVOTDATA function | Y | - | - | - | **-** | |
| 62 | Pivot area styling | Y | - | - | - | **-** | |

## Charts

| # | Feature | EPPlus | SheetJS Pro | ExcelJS | ExcelTS | ExcelForge | Notes |
|---|---------|--------|-------------|---------|---------|------------|-------|
| 63 | Bar, column, line, area, pie, etc. | Y (all 2019) | Y | - | Y | **Y** | 10 chart types |
| 64 | Scatter, radar, bubble, doughnut | Y | Y | - | - | **Y** | |
| 65 | Chart sheets | Y | Y | - | - | **Y** | addChartSheet API |
| 66 | Chart templates (.crtx) | Y | - | - | - | **-** | |
| 67 | Modern chart styling (Excel 2019) | Y | - | - | - | **-** | |
| 68 | WordArt | - | Y | - | - | **-** | |

## Images & Drawings

| # | Feature | EPPlus | SheetJS Pro | ExcelJS | ExcelTS | ExcelForge | Notes |
|---|---------|--------|-------------|---------|---------|------------|-------|
| 69 | PNG, JPEG, GIF, TIFF | Y | Y | Y | ~ | **Y** | ExcelTS: JPEG, PNG only |
| 70 | BMP, SVG, WebP, ICO, EMF, WMF | Y | ~ | - | - | **Y** | |
| 71 | Two-cell anchor | Y | Y | Y | Y | **Y** | |
| 72 | One-cell anchor (from + size) | Y | - | Y | - | **Y** | |
| 73 | Absolute anchor (no cell ref) | - | - | - | - | **Y** | ExcelForge unique |
| 74 | In-cell pictures (richData) | Y | - | - | - | **Y** | Excel 365+ |
| 75 | Shapes (187 types) | Y | Y | - | - | **P** | Preserved on round-trip |
| 76 | Shape text, effects, gradients | Y | ~ | - | - | **-** | |

## Comments

| # | Feature | EPPlus | SheetJS Pro | ExcelJS | ExcelTS | ExcelForge | Notes |
|---|---------|--------|-------------|---------|---------|------------|-------|
| 77 | Cell comments with author | Y | Y | Y | Y | **Y** | |
| 78 | Rich-text comments | Y | - | - | - | **Y** | Comment.richText with Font runs |
| 79 | Threaded comments (mentions) | Y | - | - | - | **Y** | Via rich-text comments with author prefixes |

## Form Controls

| # | Feature | EPPlus | SheetJS Pro | ExcelJS | ExcelTS | ExcelForge | Notes |
|---|---------|--------|-------------|---------|---------|------------|-------|
| 80 | Button, checkbox, radio, etc. | Y (9 types) | Y | - | - | **Y** | All 9 types |
| 81 | Macro assignment | Y | - | - | - | **Y** | |
| 82 | Linked cells, input ranges | Y | - | - | - | **Y** | |
| 83 | Width/height sizing | Y | - | - | - | **Y** | |

## Page Setup & Printing

| # | Feature | EPPlus | SheetJS Pro | ExcelJS | ExcelTS | ExcelForge | Notes |
|---|---------|--------|-------------|---------|---------|------------|-------|
| 84 | Paper size, orientation, margins | Y | - | Y | Y | **Y** | |
| 85 | Headers/footers (odd/even/first) | Y | - | Y | Y | **Y** | |
| 86 | Page breaks | Y | - | Y | - | **Y** | |
| 87 | Print areas | Y | - | Y | - | **Y** | Via printArea + defined names |
| 88 | Scaling / fit-to-page | Y | - | Y | - | **Y** | fitToPage, fitToWidth, fitToHeight, scale |

## Protection & Security

| # | Feature | EPPlus | SheetJS Pro | ExcelJS | ExcelTS | ExcelForge | Notes |
|---|---------|--------|-------------|---------|---------|------------|-------|
| 89 | Sheet protection with password | Y | Y | Y | Y | **Y** | |
| 90 | Cell locking/hiding | Y | - | Y | - | **Y** | |
| 91 | Workbook encryption | Y | Y | - | - | **-** | |
| 92 | VBA code signing | Y | - | - | - | **-** | |

## Connections & External Data

| # | Feature | EPPlus | SheetJS Pro | ExcelJS | ExcelTS | ExcelForge | Notes |
|---|---------|--------|-------------|---------|---------|------------|-------|
| 93 | OLEDB, ODBC, text, web connections | Y | - | - | - | **Y** | |
| 94 | Power Query (M formulas) | Y | - | - | - | **Y** | Read + round-trip |
| 95 | Query tables | Y | - | - | - | **-** | |
| 96 | External links (cross-workbook) | Y | - | - | - | **-** | |

## Auto Filters

| # | Feature | EPPlus | SheetJS Pro | ExcelJS | ExcelTS | ExcelForge | Notes |
|---|---------|--------|-------------|---------|---------|------------|-------|
| 97 | Basic column filters | Y | Y | Y | Y | **Y** | |
| 98 | Value/date/custom/top-10/dynamic | Y | - | - | - | **Y** | setAutoFilter with column criteria |

## Sparklines

| # | Feature | EPPlus | SheetJS Pro | ExcelJS | ExcelTS | ExcelForge | Notes |
|---|---------|--------|-------------|---------|---------|------------|-------|
| 99 | Line, bar/column, win/loss | Y | - | - | - | **Y** | |
| 100 | Colors (high/low/first/last/neg) | Y | - | - | - | **Y** | |

## VBA Macros

| # | Feature | EPPlus | SheetJS Pro | ExcelJS | ExcelTS | ExcelForge | Notes |
|---|---------|--------|-------------|---------|---------|------------|-------|
| 101 | Create/read/edit modules | Y | Y | - | - | **Y** | Standard, class, document modules |
| 102 | VBA code signing | Y | - | - | - | **-** | |
| 103 | VBA UserForms | Y | Y | - | - | **-** | |

## Properties

| # | Feature | EPPlus | SheetJS Pro | ExcelJS | ExcelTS | ExcelForge | Notes |
|---|---------|--------|-------------|---------|---------|------------|-------|
| 104 | Core properties | Y | Y | Y | - | **Y** | title, author, subject... |
| 105 | Extended properties | Y | - | - | - | **Y** | company, manager... |
| 106 | Custom properties | Y | - | - | - | **Y** | typed key-value store |

## Miscellaneous

| # | Feature | EPPlus | SheetJS Pro | ExcelJS | ExcelTS | ExcelForge | Notes |
|---|---------|--------|-------------|---------|---------|------------|-------|
| 107 | OLE objects | Y | - | - | - | **-** | |
| 108 | Ignore error rules | Y | - | - | - | **Y** | addIgnoredError API |
| 109 | Locale/international support | - | Y | - | - | **-** | |
| 110 | PDF/Canvas/SVG rendering | - | Y | - | Y | **-** | ExcelTS: PDF export module |
| 111 | Row duplicate/splice | - | - | Y | - | **Y** | duplicateRow, spliceRows |

---

## Summary Counts

| Library | Features Supported | Partial/Preserved | Not Supported |
|---------|-------------------|-------------------|---------------|
| **EPPlus 8** | 106 | 0 | 5 |
| **SheetJS Pro** | 55 | 2 | 54 |
| **ExcelJS** | 46 | 1 | 64 |
| **ExcelTS** | 33 | 1 | 77 |
| **ExcelForge** | 84 | 1 | 26 |

## ExcelForge Unique Advantages

- **Zero dependencies** — no native modules, no System.Drawing, pure TS
- **Browser + Node + Deno + Bun + edge** — universal runtime support
- **Absolute image anchoring** — `xdr:absoluteAnchor` (not available in EPPlus/SheetJS/ExcelJS/ExcelTS)
- **In-cell pictures** — only EPPlus and ExcelForge support this (Excel 365+)
- **Form controls with all 9 types** — not available in ExcelJS/ExcelTS, limited in SheetJS
- **Custom DEFLATE compression** — built-in, levels 0-9, no zlib dependency
- **Real chart sheets** — proper `<chartsheet>` XML, not embedded in worksheets
- **Dialog sheets** — Excel 5 dialog sheet support with form controls
- **Multi-sheet HTML export** — tabbed workbook HTML with CF visualization, sparklines, charts
- **.xltx template support** — read and write Excel template files

## Key Missing Features (prioritized)

### High Impact
| # | Feature | Available In | Effort |
|---|---------|-------------|--------|
| 8 | Workbook encryption | EPPlus, SheetJS | High |
| 7 | Streaming read/write | EPPlus, SheetJS, ExcelJS, ExcelTS | High |

### Medium Impact
| # | Feature | Available In | Effort |
|---|---------|-------------|--------|
| 110 | PDF export | SheetJS, ExcelTS | Medium |
| 26 | Themes (.thmx) | EPPlus | Medium |

### Lower Impact
| # | Feature | Available In | Effort |
|---|---------|-------------|--------|
| 96 | External links | EPPlus | Medium |
| 75 | Shapes (creation API) | EPPlus, SheetJS | High |
| 9 | Digital signatures | EPPlus | High |
| 107 | OLE objects | EPPlus | Medium |
| 36 | Copy/move ranges | EPPlus | Medium |

### Recently Implemented (v3.1)
| # | Feature | Notes |
|---|---------|-------|
| 4 | CSV read/write | Tree-shakeable module |
| 5 | JSON export | Tree-shakeable module |
| 6 | HTML/CSS export | Tree-shakeable module |
| 13 | Formula calculation engine | Tree-shakeable, 60+ functions |
| 17 | R1C1 reference style | Full A1↔R1C1 and formula conversion |
| 19 | Error values typed API | CellError class with constants |
| 25 | Named/cell styles | registerNamedStyle API |
| 32 | AutoFit columns | Character-count approximation |
| 35 | Copy worksheets | Cells, styles, merges |
| 37 | Insert/delete ranges | insertRows, deleteRows, insertColumns |
| 38 | Sort ranges | sortRange with asc/desc |
| 39 | Fill operations | fillNumber, fillDate, fillList |
| 41/87 | Print areas | printArea property + defined names |
| 44 | Custom table styles | registerTableStyle with DXF |
| 65 | Chart sheets | addChartSheet API |
| 78 | Rich-text comments | Comment.richText with Font runs |
| 79 | Threaded comments | Rich-text with author prefixes |
| 88 | Scaling / fit-to-page | fitToPage, scale, fitToWidth/Height |
| 98 | Advanced filter types | custom, top10, value, dynamic |
| 108 | Ignore error rules | addIgnoredError API |
| 111 | Row duplicate/splice | duplicateRow, spliceRows |
