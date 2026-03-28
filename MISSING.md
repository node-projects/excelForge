# ExcelForge Feature Comparison

Compared against: **EPPlus 8** (.NET), **SheetJS Pro** (JS), **ExcelJS** (JS/Node)

Legend: **Y** = supported, **~** = partial, **-** = not supported, **P** = preserved on round-trip only

---

## Core Read/Write

| # | Feature | EPPlus | SheetJS Pro | ExcelJS | ExcelForge | Notes |
|---|---------|--------|-------------|---------|------------|-------|
| 1 | Read/write .xlsx | Y | Y | Y | **Y** | |
| 2 | Read/write .xlsm (VBA macros) | Y | Y | - | **Y** | ExcelForge: create/edit modules, full round-trip |
| 3 | Read .xltx templates | Y | Y | - | **-** | |
| 4 | Read/write CSV | Y | Y | Y | **-** | EPPlus: LoadFromText/SaveToText |
| 5 | Export JSON | Y | Y | Y | **-** | |
| 6 | Export HTML/CSS | Y | Y | - | **-** | |
| 7 | Streaming read/write | Y (async) | Y | Y | **-** | ExcelJS: streaming XLSX for large files |
| 8 | Workbook encryption/decryption | Y | Y | - | **-** | EPPlus: Standard + Agile encryption |
| 9 | Digital signatures | Y | - | - | **-** | EPPlus: 3 sig types, 5 hash algos |

## Cell Values & Formulas

| # | Feature | EPPlus | SheetJS Pro | ExcelJS | ExcelForge | Notes |
|---|---------|--------|-------------|---------|------------|-------|
| 10 | Strings, numbers, booleans, dates | Y | Y | Y | **Y** | |
| 11 | Rich text in cells | Y | Y | Y | **Y** | |
| 12 | Formulas (store & preserve) | Y | Y | Y | **Y** | |
| 13 | Formula calculation engine | Y (463 fns) | Y | - | **-** | EPPlus: LAMBDA, dynamic arrays |
| 14 | Array formulas | Y | Y | Y | **Y** | |
| 15 | Dynamic array formulas | Y | - | - | **P** | Preserved on round-trip |
| 16 | Shared formulas | Y | Y | Y | **P** | Preserved on round-trip |
| 17 | R1C1 reference style | Y | - | - | **-** | |
| 18 | Hyperlinks | Y | Y | Y | **Y** | |
| 19 | Error values | Y | Y | Y | **~** | Preserved, no typed API |

## Styling

| # | Feature | EPPlus | SheetJS Pro | ExcelJS | ExcelForge | Notes |
|---|---------|--------|-------------|---------|------------|-------|
| 20 | Number formats | Y | Y | Y | **Y** | 30+ presets |
| 21 | Fonts (bold, italic, color, etc.) | Y | Y | Y | **Y** | |
| 22 | Fills (solid, pattern, gradient) | Y | Y | Y | **Y** | |
| 23 | Borders (all styles) | Y | Y | Y | **Y** | |
| 24 | Alignment (h/v, wrap, rotation) | Y | Y | Y | **Y** | |
| 25 | Named/cell styles | Y | Y | - | **-** | |
| 26 | Themes (load .thmx) | Y | - | - | **-** | |

## Layout & Structure

| # | Feature | EPPlus | SheetJS Pro | ExcelJS | ExcelForge | Notes |
|---|---------|--------|-------------|---------|------------|-------|
| 27 | Merge cells | Y | Y | Y | **Y** | |
| 28 | Freeze/split panes | Y | Y | Y | **Y** | |
| 29 | Column widths / row heights | Y | Y | Y | **Y** | |
| 30 | Hide rows/columns | Y | Y | Y | **Y** | |
| 31 | Outline grouping (collapse/expand) | Y | Y | Y | **Y** | |
| 32 | AutoFit columns | Y | - | - | **-** | Requires font metrics |
| 33 | Multiple sheets (hidden/veryHidden) | Y | Y | Y | **Y** | |
| 34 | Tab colors | Y | Y | Y | **Y** | |
| 35 | Copy worksheets | Y | - | - | **-** | EPPlus: with style + reference shifting |
| 36 | Copy/move ranges | Y | - | - | **-** | |
| 37 | Insert/delete ranges (auto-shift) | Y | - | Y | **-** | ExcelJS: row insert/delete/splice |
| 38 | Sort ranges | Y | - | - | **-** | EPPlus: multi-column, custom lists |
| 39 | Fill operations | Y | - | - | **-** | EPPlus: FillNumber, FillDateTime, FillList |
| 40 | Named ranges (workbook + sheet) | Y | Y | Y | **Y** | |
| 41 | Print areas | Y | - | - | **-** | |

## Tables

| # | Feature | EPPlus | SheetJS Pro | ExcelJS | ExcelForge | Notes |
|---|---------|--------|-------------|---------|------------|-------|
| 42 | Styled Excel tables | Y (60 styles) | Y | Y | **Y** | 27 built-in presets |
| 43 | Totals row | Y | - | - | **Y** | |
| 44 | Custom table styles | Y | - | - | **-** | |
| 45 | Table slicers | Y | - | - | **-** | |

## Conditional Formatting

| # | Feature | EPPlus | SheetJS Pro | ExcelJS | ExcelForge | Notes |
|---|---------|--------|-------------|---------|------------|-------|
| 46 | Cell rules | Y (45 types) | Y | Y | **Y** | |
| 47 | Color scales | Y | Y | Y | **Y** | |
| 48 | Data bars | Y | Y | Y | **Y** | |
| 49 | Icon sets | Y | Y | Y | **Y** | |
| 50 | Custom icon sets | Y | - | - | **-** | |
| 51 | Cross-worksheet references | Y | - | - | **-** | |

## Data Validation

| # | Feature | EPPlus | SheetJS Pro | ExcelJS | ExcelForge | Notes |
|---|---------|--------|-------------|---------|------------|-------|
| 52 | Dropdowns, whole/decimal, date, time | Y | Y | Y | **Y** | |
| 53 | Text length, custom formula | Y | Y | Y | **Y** | |

## Pivot Tables

| # | Feature | EPPlus | SheetJS Pro | ExcelJS | ExcelForge | Notes |
|---|---------|--------|-------------|---------|------------|-------|
| 54 | Row/column/data fields | Y | Y | ~ | **Y** | ExcelJS: limited pivot support |
| 55 | Aggregation functions | Y (12 types) | - | - | **Y** | sum, count, avg, max, min... |
| 56 | Styles (84 presets) | Y | - | - | **~** | ExcelForge: built-in presets only |
| 57 | Custom pivot styles | Y | - | - | **-** | |
| 58 | Pivot table slicers | Y | - | - | **-** | |
| 59 | Calculated fields | Y | - | - | **-** | |
| 60 | Numeric/date grouping | Y | - | - | **-** | |
| 61 | GETPIVOTDATA function | Y | - | - | **-** | |
| 62 | Pivot area styling | Y | - | - | **-** | |

## Charts

| # | Feature | EPPlus | SheetJS Pro | ExcelJS | ExcelForge | Notes |
|---|---------|--------|-------------|---------|------------|-------|
| 63 | Bar, column, line, area, pie, etc. | Y (all 2019) | Y | - | **Y** | 10 chart types |
| 64 | Scatter, radar, bubble, doughnut | Y | Y | - | **Y** | |
| 65 | Chart sheets | Y | Y | - | **-** | Dedicated sheet that IS a chart |
| 66 | Chart templates (.crtx) | Y | - | - | **-** | |
| 67 | Modern chart styling (Excel 2019) | Y | - | - | **-** | |
| 68 | WordArt | - | Y | - | **-** | |

## Images & Drawings

| # | Feature | EPPlus | SheetJS Pro | ExcelJS | ExcelForge | Notes |
|---|---------|--------|-------------|---------|------------|-------|
| 69 | PNG, JPEG, GIF, TIFF | Y | Y | Y | **Y** | |
| 70 | BMP, SVG, WebP, ICO, EMF, WMF | Y | ~ | - | **Y** | |
| 71 | Two-cell anchor | Y | Y | Y | **Y** | |
| 72 | One-cell anchor (from + size) | Y | - | Y | **Y** | |
| 73 | Absolute anchor (no cell ref) | - | - | - | **Y** | ExcelForge unique |
| 74 | In-cell pictures (richData) | Y | - | - | **Y** | Excel 365+ |
| 75 | Shapes (187 types) | Y | Y | - | **P** | Preserved on round-trip |
| 76 | Shape text, effects, gradients | Y | ~ | - | **-** | |

## Comments

| # | Feature | EPPlus | SheetJS Pro | ExcelJS | ExcelForge | Notes |
|---|---------|--------|-------------|---------|------------|-------|
| 77 | Cell comments with author | Y | Y | Y | **Y** | |
| 78 | Rich-text comments | Y | - | - | **-** | |
| 79 | Threaded comments (mentions) | Y | - | - | **-** | |

## Form Controls

| # | Feature | EPPlus | SheetJS Pro | ExcelJS | ExcelForge | Notes |
|---|---------|--------|-------------|---------|------------|-------|
| 80 | Button, checkbox, radio, etc. | Y (9 types) | Y | - | **Y** | All 9 types |
| 81 | Macro assignment | Y | - | - | **Y** | |
| 82 | Linked cells, input ranges | Y | - | - | **Y** | |
| 83 | Width/height sizing | Y | - | - | **Y** | |

## Page Setup & Printing

| # | Feature | EPPlus | SheetJS Pro | ExcelJS | ExcelForge | Notes |
|---|---------|--------|-------------|---------|------------|-------|
| 84 | Paper size, orientation, margins | Y | - | Y | **Y** | |
| 85 | Headers/footers (odd/even/first) | Y | - | Y | **Y** | |
| 86 | Page breaks | Y | - | Y | **Y** | |
| 87 | Print areas | Y | - | Y | **-** | |
| 88 | Scaling / fit-to-page | Y | - | Y | **~** | Basic via page setup |

## Protection & Security

| # | Feature | EPPlus | SheetJS Pro | ExcelJS | ExcelForge | Notes |
|---|---------|--------|-------------|---------|------------|-------|
| 89 | Sheet protection with password | Y | Y | Y | **Y** | |
| 90 | Cell locking/hiding | Y | - | Y | **Y** | |
| 91 | Workbook encryption | Y | Y | - | **-** | |
| 92 | VBA code signing | Y | - | - | **-** | |

## Connections & External Data

| # | Feature | EPPlus | SheetJS Pro | ExcelJS | ExcelForge | Notes |
|---|---------|--------|-------------|---------|------------|-------|
| 93 | OLEDB, ODBC, text, web connections | Y | - | - | **Y** | |
| 94 | Power Query (M formulas) | Y | - | - | **Y** | Read + round-trip |
| 95 | Query tables | Y | - | - | **-** | |
| 96 | External links (cross-workbook) | Y | - | - | **-** | |

## Auto Filters

| # | Feature | EPPlus | SheetJS Pro | ExcelJS | ExcelForge | Notes |
|---|---------|--------|-------------|---------|------------|-------|
| 97 | Basic column filters | Y | Y | Y | **Y** | |
| 98 | Value/date/custom/top-10/dynamic | Y | - | - | **-** | |

## Sparklines

| # | Feature | EPPlus | SheetJS Pro | ExcelJS | ExcelForge | Notes |
|---|---------|--------|-------------|---------|------------|-------|
| 99 | Line, bar/column, win/loss | Y | - | - | **Y** | |
| 100 | Colors (high/low/first/last/neg) | Y | - | - | **Y** | |

## VBA Macros

| # | Feature | EPPlus | SheetJS Pro | ExcelJS | ExcelForge | Notes |
|---|---------|--------|-------------|---------|------------|-------|
| 101 | Create/read/edit modules | Y | Y | - | **Y** | Standard, class, document modules |
| 102 | VBA code signing | Y | - | - | **-** | |
| 103 | VBA UserForms | Y | Y | - | **-** | |

## Properties

| # | Feature | EPPlus | SheetJS Pro | ExcelJS | ExcelForge | Notes |
|---|---------|--------|-------------|---------|------------|-------|
| 104 | Core properties | Y | Y | Y | **Y** | title, author, subject... |
| 105 | Extended properties | Y | - | - | **Y** | company, manager... |
| 106 | Custom properties | Y | - | - | **Y** | typed key-value store |

## Miscellaneous

| # | Feature | EPPlus | SheetJS Pro | ExcelJS | ExcelForge | Notes |
|---|---------|--------|-------------|---------|------------|-------|
| 107 | OLE objects | Y | - | - | **-** | |
| 108 | Ignore error rules | Y | - | - | **-** | Suppress green triangles |
| 109 | Locale/international support | - | Y | - | **-** | |
| 110 | PDF/Canvas/SVG rendering | - | Y | - | **-** | SheetJS Renderer component |
| 111 | Row duplicate/splice | - | - | Y | **-** | ExcelJS-specific |

---

## Summary Counts

| Library | Features Supported | Partial/Preserved | Not Supported |
|---------|-------------------|-------------------|---------------|
| **EPPlus 8** | 106 | 0 | 5 |
| **SheetJS Pro** | 55 | 2 | 54 |
| **ExcelJS** | 46 | 1 | 64 |
| **ExcelForge** | 57 | 6 | 48 |

## ExcelForge Unique Advantages

- **Zero dependencies** — no native modules, no System.Drawing, pure TS
- **Browser + Node + Deno + Bun + edge** — universal runtime support
- **Absolute image anchoring** — `xdr:absoluteAnchor` (not available in EPPlus/SheetJS/ExcelJS)
- **In-cell pictures** — only EPPlus and ExcelForge support this (Excel 365+)
- **Form controls with all 9 types** — not available in ExcelJS, limited in SheetJS
- **Custom DEFLATE compression** — built-in, levels 0-9, no zlib dependency

## Key Missing Features (prioritized)

### High Impact
| # | Feature | Available In | Effort |
|---|---------|-------------|--------|
| 13 | Formula calculation engine | EPPlus, SheetJS | Very High |
| 4 | CSV read/write | EPPlus, SheetJS, ExcelJS | Low |
| 5 | JSON export | EPPlus, SheetJS, ExcelJS | Low |
| 8 | Workbook encryption | EPPlus, SheetJS | High |
| 7 | Streaming read/write | EPPlus, SheetJS, ExcelJS | High |

### Medium Impact
| # | Feature | Available In | Effort |
|---|---------|-------------|--------|
| 6 | HTML/CSS export | EPPlus, SheetJS | Medium |
| 32 | AutoFit columns | EPPlus | Medium (font metrics) |
| 35 | Copy worksheets | EPPlus | Medium |
| 37 | Insert/delete ranges (auto-shift) | EPPlus, ExcelJS | Medium |
| 65 | Chart sheets | EPPlus, SheetJS | Low |
| 44 | Custom table styles | EPPlus | Medium |
| 79 | Threaded comments | EPPlus | Medium |
| 41 | Print areas | EPPlus, ExcelJS | Low |

### Lower Impact
| # | Feature | Available In | Effort |
|---|---------|-------------|--------|
| 96 | External links | EPPlus | Medium |
| 75 | Shapes (creation API) | EPPlus, SheetJS | High |
| 98 | Advanced filter types | EPPlus | Medium |
| 38 | Sort ranges | EPPlus | Medium |
| 25 | Named/cell styles | EPPlus, SheetJS | Low |
| 9 | Digital signatures | EPPlus | High |
| 107 | OLE objects | EPPlus | Medium |
