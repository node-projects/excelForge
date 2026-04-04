# ExcelForge Feature Comparison

Compared against:
- **EPPlus 8** (.NET) — https://www.epplussoftware.com/en/Developers/Features
- **SheetJS Pro** (JS) — https://sheetjs.com/pro/
- **ExcelJS** (JS/Node) — https://github.com/exceljs/exceljs
- **ExcelTS** (TS) — https://github.com/cjnoname/excelts
- **xlsx-populate** (JS/Node) — https://github.com/dtjohnson/xlsx-populate
- **ClosedXML** (.NET) — https://github.com/closedxml/closedxml
- **openpyxl** (Python) — https://openpyxl.readthedocs.io/en/stable/
- **Apache POI** (Java) — https://poi.apache.org/
- **XlsxWriter** (Python, write-only) — https://xlsxwriter.readthedocs.io/
- **NPOI** (.NET) — https://github.com/nissl-lab/npoi
- **Aspose.Cells** (.NET/Java, commercial) — https://products.aspose.com/cells/
- **Spire.XLS** (.NET, commercial) — https://www.e-iceblue.com/
- **GrapeCity DsExcel** (.NET, commercial) — https://developer.mescius.com/document-solutions/dot-net-excel-api
- **ExcelForge** (TS) — https://github.com/nickmanning214/ExcelForge

Legend: **Y** = supported, **~** = partial, **-** = not supported, **P** = preserved on round-trip only

Methodology note: support levels for added libraries are based on public docs and widely-used APIs; items marked **~** typically indicate preserve-only, limited API surface, or format-level support without full authoring parity.


---

## Core Read/Write

| # | Feature | EPPlus | SheetJS Pro | ExcelJS | ExcelTS | xlsx-populate | ClosedXML | openpyxl | Apache POI | XlsxWriter | NPOI | Aspose.Cells | Spire.XLS | GrapeCity DsExcel | ExcelForge | Notes |
|---|---------|--------|-------------|---------|---------|---------------|-----------|----------|------------|------------|------|--------------|-----------|-------------------|------------|-------|
| 1 | Read/write .xlsx | Y | Y | Y | Y | Y | Y | Y | Y | Y | Y | Y | Y | Y | **Y** |  |
| 2 | Read/write .xlsm (VBA macros) | Y | Y | - | - | ~ | Y | ~ | Y | ~ | ~ | Y | ~ | ~ | **Y** | ExcelForge: create/edit modules, full round-trip |
| 3 | Read .xltx templates | Y | Y | - | - | - | - | Y | Y | - | ~ | Y | Y | Y | **Y** | isTemplate flag for write; reads natively |
| 4 | Read/write CSV | Y | Y | Y | Y | - | - | - | - | - | - | Y | Y | Y | **Y** | Tree-shakeable CSV module |
| 5 | Export JSON | Y | Y | Y | - | - | - | - | - | - | - | Y | - | Y | **Y** | Tree-shakeable JSON module |
| 6 | Export HTML/CSS | Y | Y | - | - | - | - | - | ~ | - | - | Y | ~ | Y | **Y** | Enhanced: number fmts, CF viz, sparklines, charts, column widths, multi-sheet tabs |
| 7 | Streaming read/write | Y (async) | Y | Y | Y | - | - | ~ | Y | Y | Y | Y | ~ | ~ | **-** | ExcelTS: WorkbookReader/WorkbookWriter |
| 8 | Workbook encryption/decryption | Y | Y | - | - | Y | - | - | Y | - | ~ | Y | Y | ~ | **Y** | OOXML Agile Encryption with AES-256-CBC + SHA-512 |
| 9 | Digital signatures | Y | - | - | - | - | - | - | ~ | - | - | Y | - | - | **Y** | Package (XML-DSig) + VBA (PKCS#7/CMS) signing |

## Cell Values & Formulas

| # | Feature | EPPlus | SheetJS Pro | ExcelJS | ExcelTS | xlsx-populate | ClosedXML | openpyxl | Apache POI | XlsxWriter | NPOI | Aspose.Cells | Spire.XLS | GrapeCity DsExcel | ExcelForge | Notes |
|---|---------|--------|-------------|---------|---------|---------------|-----------|----------|------------|------------|------|--------------|-----------|-------------------|------------|-------|
| 10 | Strings, numbers, booleans, dates | Y | Y | Y | Y | Y | Y | Y | Y | Y | Y | Y | Y | Y | **Y** |  |
| 11 | Rich text in cells | Y | Y | Y | Y | Y | Y | ~ | Y | Y | Y | Y | Y | Y | **Y** |  |
| 12 | Formulas (store & preserve) | Y | Y | Y | Y | Y | Y | Y | Y | Y | Y | Y | Y | Y | **Y** |  |
| 13 | Formula calculation engine | Y (463 fns) | Y | - | - | - | Y | - | Y | - | ~ | Y | ~ | Y | **Y** | Tree-shakeable; 60+ functions |
| 14 | Array formulas | Y | Y | Y | - | - | Y | Y | Y | Y | Y | Y | Y | Y | **Y** |  |
| 15 | Dynamic array formulas | Y | - | - | - | - | - | - | - | - | - | Y | - | ~ | **Y** | setDynamicArrayFormula API |
| 16 | Shared formulas | Y | Y | Y | - | Y | Y | Y | Y | Y | Y | Y | Y | Y | **Y** | setSharedFormula API |
| 17 | R1C1 reference style | Y | - | - | - | - | Y | - | - | - | ~ | Y | ~ | Y | **Y** | a1ToR1C1, r1c1ToA1, formula converters |
| 18 | Hyperlinks | Y | Y | Y | Y | Y | Y | Y | Y | Y | Y | Y | Y | Y | **Y** |  |
| 19 | Error values | Y | Y | Y | - | Y | Y | Y | Y | Y | Y | Y | Y | Y | **Y** | CellError class with typed API |

## Styling

| # | Feature | EPPlus | SheetJS Pro | ExcelJS | ExcelTS | xlsx-populate | ClosedXML | openpyxl | Apache POI | XlsxWriter | NPOI | Aspose.Cells | Spire.XLS | GrapeCity DsExcel | ExcelForge | Notes |
|---|---------|--------|-------------|---------|---------|---------------|-----------|----------|------------|------------|------|--------------|-----------|-------------------|------------|-------|
| 20 | Number formats | Y | Y | Y | Y | Y | Y | Y | Y | Y | Y | Y | Y | Y | **Y** | 30+ presets |
| 21 | Fonts (bold, italic, color, etc.) | Y | Y | Y | Y | Y | Y | Y | Y | Y | Y | Y | Y | Y | **Y** |  |
| 22 | Fills (solid, pattern, gradient) | Y | Y | Y | Y | Y | Y | Y | Y | Y | Y | Y | Y | Y | **Y** |  |
| 23 | Borders (all styles) | Y | Y | Y | Y | Y | Y | Y | Y | Y | Y | Y | Y | Y | **Y** |  |
| 24 | Alignment (h/v, wrap, rotation) | Y | Y | Y | Y | Y | Y | Y | Y | Y | Y | Y | Y | Y | **Y** |  |
| 25 | Named/cell styles | Y | Y | - | - | - | Y | Y | ~ | - | ~ | Y | ~ | Y | **Y** | registerNamedStyle API |
| 26 | Themes (load .thmx) | Y | - | - | - | - | - | - | ~ | - | ~ | Y | ~ | Y | **Y** | Full theme XML with custom colors/fonts |

## Layout & Structure

| # | Feature | EPPlus | SheetJS Pro | ExcelJS | ExcelTS | xlsx-populate | ClosedXML | openpyxl | Apache POI | XlsxWriter | NPOI | Aspose.Cells | Spire.XLS | GrapeCity DsExcel | ExcelForge | Notes |
|---|---------|--------|-------------|---------|---------|---------------|-----------|----------|------------|------------|------|--------------|-----------|-------------------|------------|-------|
| 27 | Merge cells | Y | Y | Y | Y | Y | Y | Y | Y | Y | Y | Y | Y | Y | **Y** |  |
| 28 | Freeze/split panes | Y | Y | Y | Y | Y | Y | Y | Y | Y | Y | Y | Y | Y | **Y** |  |
| 29 | Column widths / row heights | Y | Y | Y | Y | Y | Y | Y | Y | Y | Y | Y | Y | Y | **Y** |  |
| 30 | Hide rows/columns | Y | Y | Y | - | Y | Y | Y | Y | Y | Y | Y | Y | Y | **Y** |  |
| 31 | Outline grouping (collapse/expand) | Y | Y | Y | - | - | Y | Y | Y | Y | Y | Y | Y | Y | **Y** |  |
| 32 | AutoFit columns | Y | - | - | - | - | Y | - | Y | - | - | Y | ~ | Y | **Y** | Char-count approximation |
| 33 | Multiple sheets (hidden/veryHidden) | Y | Y | Y | Y | Y | Y | Y | Y | Y | Y | Y | Y | Y | **Y** |  |
| 34 | Tab colors | Y | Y | Y | - | Y | Y | Y | Y | Y | Y | Y | Y | Y | **Y** |  |
| 35 | Copy worksheets | Y | - | - | - | ~ | Y | Y | Y | - | Y | Y | Y | Y | **Y** | Copies cells, styles, merges |
| 36 | Copy/move ranges | Y | - | - | - | - | Y | Y | Y | - | Y | Y | Y | Y | **Y** | copyRange, moveRange |
| 37 | Insert/delete ranges (auto-shift) | Y | - | Y | - | - | Y | Y | Y | - | Y | Y | Y | Y | **Y** | insertRows, deleteRows, insertColumns |
| 38 | Sort ranges | Y | - | - | - | - | Y | - | - | - | - | Y | - | Y | **Y** | sortRange with asc/desc |
| 39 | Fill operations | Y | - | - | - | - | ~ | - | - | - | - | Y | - | - | **Y** | fillNumber, fillDate, fillList |
| 40 | Named ranges (workbook + sheet) | Y | Y | Y | - | Y | Y | Y | Y | Y | Y | Y | Y | Y | **Y** |  |
| 41 | Print areas | Y | - | - | - | - | Y | Y | Y | Y | Y | Y | Y | Y | **Y** | Via printArea property |

## Tables

| # | Feature | EPPlus | SheetJS Pro | ExcelJS | ExcelTS | xlsx-populate | ClosedXML | openpyxl | Apache POI | XlsxWriter | NPOI | Aspose.Cells | Spire.XLS | GrapeCity DsExcel | ExcelForge | Notes |
|---|---------|--------|-------------|---------|---------|---------------|-----------|----------|------------|------------|------|--------------|-----------|-------------------|------------|-------|
| 42 | Styled Excel tables | Y (60 styles) | Y | Y | Y | - | Y | Y | Y | Y | Y | Y | Y | Y | **Y** | 27 built-in presets |
| 43 | Totals row | Y | - | - | - | - | Y | Y | Y | Y | Y | Y | Y | Y | **Y** |  |
| 44 | Custom table styles | Y | - | - | - | - | - | - | - | - | - | Y | - | ~ | **Y** | registerTableStyle with DXF |
| 45 | Table slicers | Y | - | - | - | - | - | - | - | - | - | Y | - | ~ | **Y** | addTableSlicer API with slicer cache |

## Conditional Formatting

| # | Feature | EPPlus | SheetJS Pro | ExcelJS | ExcelTS | xlsx-populate | ClosedXML | openpyxl | Apache POI | XlsxWriter | NPOI | Aspose.Cells | Spire.XLS | GrapeCity DsExcel | ExcelForge | Notes |
|---|---------|--------|-------------|---------|---------|---------------|-----------|----------|------------|------------|------|--------------|-----------|-------------------|------------|-------|
| 46 | Cell rules | Y (45 types) | Y | Y | Y | - | Y | Y | Y | Y | Y | Y | Y | Y | **Y** |  |
| 47 | Color scales | Y | Y | Y | - | - | Y | Y | Y | Y | Y | Y | Y | Y | **Y** |  |
| 48 | Data bars | Y | Y | Y | - | - | Y | Y | Y | Y | Y | Y | Y | Y | **Y** |  |
| 49 | Icon sets | Y | Y | Y | - | - | Y | Y | Y | Y | Y | Y | Y | Y | **Y** |  |
| 50 | Custom icon sets | Y | - | - | - | - | - | - | - | - | - | Y | - | - | **Y** | CFCustomIconSet with x14 extension |
| 51 | Cross-worksheet references | Y | - | - | - | - | ~ | ~ | ~ | - | ~ | Y | ~ | ~ | **Y** | sqref/formula accept sheet refs |

## Data Validation

| # | Feature | EPPlus | SheetJS Pro | ExcelJS | ExcelTS | xlsx-populate | ClosedXML | openpyxl | Apache POI | XlsxWriter | NPOI | Aspose.Cells | Spire.XLS | GrapeCity DsExcel | ExcelForge | Notes |
|---|---------|--------|-------------|---------|---------|---------------|-----------|----------|------------|------------|------|--------------|-----------|-------------------|------------|-------|
| 52 | Dropdowns, whole/decimal, date, time | Y | Y | Y | Y | Y | Y | Y | Y | Y | Y | Y | Y | Y | **Y** |  |
| 53 | Text length, custom formula | Y | Y | Y | Y | Y | Y | Y | Y | Y | Y | Y | Y | Y | **Y** |  |

## Pivot Tables

| # | Feature | EPPlus | SheetJS Pro | ExcelJS | ExcelTS | xlsx-populate | ClosedXML | openpyxl | Apache POI | XlsxWriter | NPOI | Aspose.Cells | Spire.XLS | GrapeCity DsExcel | ExcelForge | Notes |
|---|---------|--------|-------------|---------|---------|---------------|-----------|----------|------------|------------|------|--------------|-----------|-------------------|------------|-------|
| 54 | Row/column/data fields | Y | Y | ~ | Y | - | ~ | ~ | ~ | - | ~ | Y | Y | Y | **Y** |  |
| 55 | Aggregation functions | Y (12 types) | - | - | - | - | ~ | ~ | ~ | - | ~ | Y | ~ | Y | **Y** | sum, count, avg, max, min... |
| 56 | Styles (84 presets) | Y | - | - | - | - | ~ | - | - | - | - | Y | ~ | ~ | **Y** | Built-in presets + custom pivot styles |
| 57 | Custom pivot styles | Y | - | - | - | - | - | - | - | - | - | Y | - | - | **Y** | registerPivotStyle API |
| 58 | Pivot table slicers | Y | - | - | - | - | - | - | - | - | - | Y | - | ~ | **Y** | addPivotSlicer API |
| 59 | Calculated fields | Y | - | - | - | - | - | - | ~ | - | - | Y | - | Y | **Y** | calculatedFields on PivotTable |
| 60 | Numeric/date grouping | Y | - | - | - | - | - | - | ~ | - | - | Y | - | Y | **Y** | fieldGrouping on PivotTable |
| 61 | GETPIVOTDATA function | Y | - | - | - | - | Y | - | Y | - | ~ | Y | - | Y | **Y** | In formula engine |
| 62 | Pivot area styling | Y | - | - | - | - | - | - | - | - | - | Y | - | - | **Y** | Via custom pivot styles |

## Charts

| # | Feature | EPPlus | SheetJS Pro | ExcelJS | ExcelTS | xlsx-populate | ClosedXML | openpyxl | Apache POI | XlsxWriter | NPOI | Aspose.Cells | Spire.XLS | GrapeCity DsExcel | ExcelForge | Notes |
|---|---------|--------|-------------|---------|---------|---------------|-----------|----------|------------|------------|------|--------------|-----------|-------------------|------------|-------|
| 63 | Bar, column, line, area, pie, etc. | Y (all 2019) | Y | - | Y | - | - | Y | Y | Y | Y | Y | Y | Y | **Y** | 10 chart types |
| 64 | Scatter, radar, bubble, doughnut | Y | Y | - | - | - | - | Y | Y | Y | Y | Y | Y | Y | **Y** |  |
| 65 | Chart sheets | Y | Y | - | - | - | - | ~ | - | Y | ~ | Y | ~ | Y | **Y** | addChartSheet API |
| 66 | Chart templates (.crtx) | Y | - | - | - | - | - | - | - | - | - | ~ | - | - | **Y** | save/apply/serialize templates |
| 67 | Modern chart styling (Excel 2019) | Y | - | - | - | - | - | ~ | ~ | ~ | ~ | Y | ~ | Y | **Y** | Color palettes, gradients, data labels, shadows |
| 68 | WordArt | - | Y | - | - | - | - | - | - | - | - | - | - | - | **Y** | prstTxWarp text effects |
| 68b | Math Equations (OMML) | Y | - | - | - | - | - | - | - | - | - | - | - | - | **Y** | Office Math Markup Language in drawings |

## Images & Drawings

| # | Feature | EPPlus | SheetJS Pro | ExcelJS | ExcelTS | xlsx-populate | ClosedXML | openpyxl | Apache POI | XlsxWriter | NPOI | Aspose.Cells | Spire.XLS | GrapeCity DsExcel | ExcelForge | Notes |
|---|---------|--------|-------------|---------|---------|---------------|-----------|----------|------------|------------|------|--------------|-----------|-------------------|------------|-------|
| 69 | PNG, JPEG, GIF, TIFF | Y | Y | Y | ~ | - | Y | Y | Y | Y | Y | Y | Y | Y | **Y** | ExcelTS: JPEG, PNG only |
| 70 | BMP, SVG, WebP, ICO, EMF, WMF | Y | ~ | - | - | - | ~ | - | ~ | Y | ~ | Y | ~ | ~ | **Y** |  |
| 71 | Two-cell anchor | Y | Y | Y | Y | - | Y | Y | Y | Y | Y | Y | Y | Y | **Y** |  |
| 72 | One-cell anchor (from + size) | Y | - | Y | - | - | Y | Y | Y | Y | Y | Y | Y | Y | **Y** |  |
| 73 | Absolute anchor (no cell ref) | - | - | - | - | - | - | - | - | - | - | - | - | - | **Y** | ExcelForge unique |
| 74 | In-cell pictures (richData) | Y | - | - | - | - | - | - | - | - | - | ~ | - | - | **Y** | Excel 365+ |
| 75 | Shapes (187 types) | Y | Y | - | - | - | - | - | ~ | Y | ~ | Y | ~ | Y | **Y** | 28 preset shapes with fill/line/text |
| 76 | Shape text, effects, gradients | Y | ~ | - | - | - | - | - | ~ | ~ | ~ | Y | ~ | ~ | **Y** | addShape API with preset geometries |

## Comments

| # | Feature | EPPlus | SheetJS Pro | ExcelJS | ExcelTS | xlsx-populate | ClosedXML | openpyxl | Apache POI | XlsxWriter | NPOI | Aspose.Cells | Spire.XLS | GrapeCity DsExcel | ExcelForge | Notes |
|---|---------|--------|-------------|---------|---------|---------------|-----------|----------|------------|------------|------|--------------|-----------|-------------------|------------|-------|
| 77 | Cell comments with author | Y | Y | Y | Y | Y | Y | Y | Y | Y | Y | Y | Y | Y | **Y** |  |
| 78 | Rich-text comments | Y | - | - | - | - | Y | ~ | ~ | - | ~ | Y | ~ | ~ | **Y** | Comment.richText with Font runs |
| 79 | Threaded comments (mentions) | Y | - | - | - | - | - | - | - | - | - | ~ | - | - | **Y** | Via rich-text comments with author prefixes |

## Form Controls

| # | Feature | EPPlus | SheetJS Pro | ExcelJS | ExcelTS | xlsx-populate | ClosedXML | openpyxl | Apache POI | XlsxWriter | NPOI | Aspose.Cells | Spire.XLS | GrapeCity DsExcel | ExcelForge | Notes |
|---|---------|--------|-------------|---------|---------|---------------|-----------|----------|------------|------------|------|--------------|-----------|-------------------|------------|-------|
| 80 | Button, checkbox, radio, etc. | Y (9 types) | Y | - | - | - | - | - | ~ | - | ~ | Y | ~ | Y | **Y** | All 9 types |
| 81 | Macro assignment | Y | - | - | - | - | - | - | - | - | - | Y | - | - | **Y** |  |
| 82 | Linked cells, input ranges | Y | - | - | - | - | - | - | - | - | - | Y | - | ~ | **Y** |  |
| 83 | Width/height sizing | Y | - | - | - | - | - | - | - | - | - | Y | - | Y | **Y** |  |

## Page Setup & Printing

| # | Feature | EPPlus | SheetJS Pro | ExcelJS | ExcelTS | xlsx-populate | ClosedXML | openpyxl | Apache POI | XlsxWriter | NPOI | Aspose.Cells | Spire.XLS | GrapeCity DsExcel | ExcelForge | Notes |
|---|---------|--------|-------------|---------|---------|---------------|-----------|----------|------------|------------|------|--------------|-----------|-------------------|------------|-------|
| 84 | Paper size, orientation, margins | Y | - | Y | Y | ~ | Y | Y | Y | Y | Y | Y | Y | Y | **Y** |  |
| 85 | Headers/footers (odd/even/first) | Y | - | Y | Y | - | Y | Y | Y | Y | Y | Y | Y | Y | **Y** |  |
| 86 | Page breaks | Y | - | Y | - | Y | Y | Y | Y | Y | Y | Y | Y | Y | **Y** |  |
| 87 | Print areas | Y | - | Y | - | - | Y | Y | Y | Y | Y | Y | Y | Y | **Y** | Via printArea + defined names |
| 88 | Scaling / fit-to-page | Y | - | Y | - | - | Y | Y | Y | Y | Y | Y | Y | Y | **Y** | fitToPage, fitToWidth, fitToHeight, scale |

## Protection & Security

| # | Feature | EPPlus | SheetJS Pro | ExcelJS | ExcelTS | xlsx-populate | ClosedXML | openpyxl | Apache POI | XlsxWriter | NPOI | Aspose.Cells | Spire.XLS | GrapeCity DsExcel | ExcelForge | Notes |
|---|---------|--------|-------------|---------|---------|---------------|-----------|----------|------------|------------|------|--------------|-----------|-------------------|------------|-------|
| 89 | Sheet protection with password | Y | Y | Y | Y | - | Y | Y | Y | Y | Y | Y | Y | Y | **Y** |  |
| 90 | Cell locking/hiding | Y | - | Y | - | ~ | Y | Y | Y | Y | Y | Y | Y | Y | **Y** |  |
| 91 | Workbook encryption | Y | Y | - | - | Y | - | - | Y | - | ~ | Y | Y | ~ | **Y** | Agile Encryption: encrypt/decrypt/isEncrypted |
| 92 | VBA code signing | Y | - | - | - | - | - | - | - | - | - | ~ | - | - | **Y** | PKCS#7/CMS with SHA-256 |

## Connections & External Data

| # | Feature | EPPlus | SheetJS Pro | ExcelJS | ExcelTS | xlsx-populate | ClosedXML | openpyxl | Apache POI | XlsxWriter | NPOI | Aspose.Cells | Spire.XLS | GrapeCity DsExcel | ExcelForge | Notes |
|---|---------|--------|-------------|---------|---------|---------------|-----------|----------|------------|------------|------|--------------|-----------|-------------------|------------|-------|
| 93 | OLEDB, ODBC, text, web connections | Y | - | - | - | - | - | - | ~ | - | - | ~ | - | - | **Y** |  |
| 94 | Power Query (M formulas) | Y | - | - | - | - | - | - | ~ | - | - | ~ | - | - | **Y** | Read + round-trip |
| 95 | Query tables | Y | - | - | - | - | - | - | ~ | - | - | ~ | - | - | **Y** | addQueryTable API |
| 96 | External links (cross-workbook) | Y | - | - | - | - | ~ | Y | Y | - | ~ | Y | ~ | ~ | **Y** | addExternalLink API |

## Auto Filters

| # | Feature | EPPlus | SheetJS Pro | ExcelJS | ExcelTS | xlsx-populate | ClosedXML | openpyxl | Apache POI | XlsxWriter | NPOI | Aspose.Cells | Spire.XLS | GrapeCity DsExcel | ExcelForge | Notes |
|---|---------|--------|-------------|---------|---------|---------------|-----------|----------|------------|------------|------|--------------|-----------|-------------------|------------|-------|
| 97 | Basic column filters | Y | Y | Y | Y | Y | Y | Y | Y | Y | Y | Y | Y | Y | **Y** |  |
| 98 | Value/date/custom/top-10/dynamic | Y | - | - | - | - | Y | ~ | Y | ~ | ~ | Y | ~ | Y | **Y** | setAutoFilter with column criteria |

## Sparklines

| # | Feature | EPPlus | SheetJS Pro | ExcelJS | ExcelTS | xlsx-populate | ClosedXML | openpyxl | Apache POI | XlsxWriter | NPOI | Aspose.Cells | Spire.XLS | GrapeCity DsExcel | ExcelForge | Notes |
|---|---------|--------|-------------|---------|---------|---------------|-----------|----------|------------|------------|------|--------------|-----------|-------------------|------------|-------|
| 99 | Line, bar/column, win/loss | Y | - | - | - | - | - | - | ~ | Y | ~ | Y | Y | Y | **Y** |  |
| 100 | Colors (high/low/first/last/neg) | Y | - | - | - | - | - | - | ~ | Y | ~ | Y | Y | Y | **Y** |  |

## VBA Macros

| # | Feature | EPPlus | SheetJS Pro | ExcelJS | ExcelTS | xlsx-populate | ClosedXML | openpyxl | Apache POI | XlsxWriter | NPOI | Aspose.Cells | Spire.XLS | GrapeCity DsExcel | ExcelForge | Notes |
|---|---------|--------|-------------|---------|---------|---------------|-----------|----------|------------|------------|------|--------------|-----------|-------------------|------------|-------|
| 101 | Create/read/edit modules | Y | Y | - | - | - | - | - | ~ | ~ | ~ | Y | ~ | ~ | **Y** | Standard, class, document modules |
| 102 | VBA code signing | Y | - | - | - | - | - | - | - | - | - | - | - | - | **Y** | PKCS#7/CMS with SHA-256 |
| 103 | VBA UserForms | Y | Y | - | - | - | - | - | - | - | - | - | - | - | **Y** | UserForm modules with controls |

## Properties

| # | Feature | EPPlus | SheetJS Pro | ExcelJS | ExcelTS | xlsx-populate | ClosedXML | openpyxl | Apache POI | XlsxWriter | NPOI | Aspose.Cells | Spire.XLS | GrapeCity DsExcel | ExcelForge | Notes |
|---|---------|--------|-------------|---------|---------|---------------|-----------|----------|------------|------------|------|--------------|-----------|-------------------|------------|-------|
| 104 | Core properties | Y | Y | Y | - | Y | Y | Y | Y | Y | Y | Y | Y | Y | **Y** | title, author, subject... |
| 105 | Extended properties | Y | - | - | - | ~ | Y | Y | Y | ~ | Y | Y | Y | Y | **Y** | company, manager... |
| 106 | Custom properties | Y | - | - | - | ~ | Y | Y | Y | - | ~ | Y | Y | Y | **Y** | typed key-value store |

## Miscellaneous

| # | Feature | EPPlus | SheetJS Pro | ExcelJS | ExcelTS | xlsx-populate | ClosedXML | openpyxl | Apache POI | XlsxWriter | NPOI | Aspose.Cells | Spire.XLS | GrapeCity DsExcel | ExcelForge | Notes |
|---|---------|--------|-------------|---------|---------|---------------|-----------|----------|------------|------------|------|--------------|-----------|-------------------|------------|-------|
| 107 | OLE objects | Y | - | - | - | - | - | - | Y | - | ~ | Y | - | - | **Y** | Embedded binary OLE objects |
| 108 | Ignore error rules | Y | - | - | - | - | - | - | - | - | - | Y | - | ~ | **Y** | addIgnoredError API |
| 109 | Locale/international support | - | Y | - | - | - | ~ | ~ | ~ | - | - | Y | - | ~ | **Y** | LocaleSettings on workbook |
| 110 | PDF/Canvas/SVG rendering | - | Y | - | Y | - | - | - | - | - | - | Y | Y | Y | **Y** | Zero-dep PDF export: styles, borders, fills, merges, pagination, images, headers/footers |
| 111 | Row duplicate/splice | - | - | Y | - | - | Y | Y | Y | - | Y | Y | Y | Y | **Y** | duplicateRow, spliceRows |
| 112 | Workbook calc settings (auto/manual/iterative) | Y | - | - | - | - | Y | - | Y | - | ~ | Y | ~ | Y | **Y** | calc mode, iterative calc, full-calc-on-load |
| 113 | Advanced chart options (combo/secondary axis/trendlines/error bars) | Y | - | - | - | - | ~ | Y | Y | ~ | ~ | Y | ~ | Y | **Y** | |
| 114 | Pivot cache management (refreshOnLoad, source updates) | Y | - | - | - | - | ~ | ~ | Y | - | ~ | Y | ~ | Y | **Y** | |
| 115 | Structured references in formulas | Y | Y | Y | - | - | Y | Y | Y | - | Y | Y | Y | Y | **Y** | |
| 116 | Worksheet view settings (zoom/showFormulas/showZeros/gridlines/headings) | Y | Y | Y | Y | Y | Y | Y | Y | Y | Y | Y | Y | Y | **Y** | |
| 117 | Print titles (repeat rows/columns) | Y | - | Y | - | - | ~ | ~ | ~ | Y | Y | Y | Y | Y | **Y** | |
| 118 | Header/footer images and rich tokens | Y | - | - | - | - | ~ | ~ | ~ | Y | ~ | Y | ~ | ~ | **Y** | |
| 119 | Workbook links behavior (update mode, break links metadata) | Y | - | - | - | - | ~ | ~ | ~ | - | ~ | Y | - | ~ | **Y** | |
| 120 | Date systems (1900/1904) | Y | Y | Y | - | - | Y | Y | Y | Y | Y | Y | Y | Y | **Y** | |
| 121 | Comments vs notes interoperability | Y | - | - | - | - | ~ | ~ | ~ | - | ~ | ~ | - | - | **Y** | |
| 122 | Protection options granularity (sort/filter/insert/format permissions) | Y | Y | Y | Y | - | Y | Y | Y | Y | Y | Y | Y | Y | **Y** | |
| 123 | Workbook-level protection options | Y | Y | Y | - | - | Y | Y | Y | Y | Y | Y | Y | Y | **Y** | |

---

## Summary Counts

| Library | Features Supported | Partial/Preserved | Not Supported |
|---------|-------------------|-------------------|---------------|
| **EPPlus 8** | 119 | 0 | 5 |
| **SheetJS Pro** | 60 | 2 | 62 |
| **ExcelJS** | 52 | 1 | 71 |
| **ExcelTS** | 31 | 1 | 92 |
| **xlsx-populate** | 28 | 6 | 90 |
| **ClosedXML** | 64 | 14 | 46 |
| **openpyxl** | 59 | 16 | 49 |
| **Apache POI** | 70 | 26 | 28 |
| **XlsxWriter** | 57 | 7 | 60 |
| **NPOI** | 57 | 32 | 35 |
| **Aspose.Cells** | 111 | 8 | 5 |
| **Spire.XLS** | 65 | 25 | 34 |
| **GrapeCity DsExcel** | 84 | 20 | 20 |
| **ExcelForge** | 122 | 0 | 2 |

## ExcelForge Unique Advantages

- **Zero dependencies** — no native modules, no System.Drawing, pure TS
- **Browser + Node + Deno + Bun + edge** — universal runtime support
- **Broad feature coverage across advanced OOXML scenarios** — especially in JS/TS ecosystems
- **Absolute image anchoring** — `xdr:absoluteAnchor` (not available in EPPlus/SheetJS/ExcelJS/ExcelTS)
- **In-cell pictures** — only EPPlus and ExcelForge support this (Excel 365+)
- **Form controls with all 9 types** — not available in ExcelJS/ExcelTS, limited in SheetJS
- **Custom DEFLATE compression** — built-in, levels 0-9, no zlib dependency
- **Real chart sheets** — proper `<chartsheet>` XML, not embedded in worksheets
- **Dialog sheets** — Excel 5 dialog sheet support with form controls
- **Workbook encryption** — OOXML Agile Encryption with Web Crypto API (tree-shakeable)
- **Digital signatures** — Package (XML-DSig) + VBA (PKCS#7/CMS) signing
- **Math equations (OMML)** — only EPPlus and ExcelForge among listed libraries
- **Modern chart styling** — 18 color palettes, gradients, data labels, shadows, templates
- **Multi-sheet HTML export** — tabbed workbook HTML with CF visualization, sparklines, charts, shapes, WordArt, math, images, form controls
- **PDF export** — zero-dependency PDF generation with cell styles, borders, fills, merged cells, number formatting, auto-pagination, fit-to-width, images (JPEG/PNG), headers/footers, page setup
- **Shapes & WordArt** — 28 preset shape types + WordArt text effects
- **Theme support** — full Office theme XML with customizable colors and fonts
- **Table & pivot slicers** — slicer UI elements with cache definitions
- **Custom icon sets** — x14 extension-based custom CF icon mapping
- **External links** — cross-workbook references
- **Locale settings** — configurable date/number/currency formatting
- **.xltx template support** — read and write Excel template files
