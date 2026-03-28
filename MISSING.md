EPPlus Features Missing in ExcelForge                                                                                 
                                                                                                                        
  High Impact (Core Functionality)                                                                                      
                                                                                                                        
  ┌─────┬─────────────────────────────────┬──────────────────────────────────────┬──────────────────────────────────┐   
  │  #  │             Feature             │                EPPlus                │            ExcelForge            │
  ├─────┼─────────────────────────────────┼──────────────────────────────────────┼──────────────────────────────────┤
  │ 1   │ Formula Calculation Engine      │ Built-in engine, 463 functions,      │ Stores/preserves formulas only,  │
  │     │                                 │ LAMBDA support                       │ no calculation                   │
  ├─────┼─────────────────────────────────┼──────────────────────────────────────┼──────────────────────────────────┤
  │ 2   │ Data Export (CSV, JSON, HTML)   │ Export to CSV, JSON, HTML/CSS,       │ No export formats besides .xlsx  │
  │     │                                 │ DataTable                            │                                  │
  ├─────┼─────────────────────────────────┼──────────────────────────────────────┼──────────────────────────────────┤
  │ 3   │ Data Import (from CSV, text,    │ Import from enumerables, text files, │ No import from external formats  │
  │     │ DataReader)                     │  DataSets                            │                                  │
  ├─────┼─────────────────────────────────┼──────────────────────────────────────┼──────────────────────────────────┤
  │ 4   │ Workbook Encryption/Decryption  │ Password-protect & open encrypted    │ No encryption support            │
  │     │                                 │ files                                │                                  │
  ├─────┼─────────────────────────────────┼──────────────────────────────────────┼──────────────────────────────────┤
  │ 5   │ Digital Signatures              │ Read/add signatures, 3 signature     │ Not supported                    │
  │     │                                 │ types, 5 hash algos                  │                                  │
  ├─────┼─────────────────────────────────┼──────────────────────────────────────┼──────────────────────────────────┤
  │ 6   │ External Links                  │ Cross-workbook references, link      │ Not supported                    │
  │     │                                 │ breaking, cache update               │                                  │
  └─────┴─────────────────────────────────┴──────────────────────────────────────┴──────────────────────────────────┘

  Medium Impact (Editing & Layout)

  ┌─────┬────────────────────────┬────────────────────────────────────────────────┬────────────────────────────────┐
  │  #  │        Feature         │                     EPPlus                     │           ExcelForge           │
  ├─────┼────────────────────────┼────────────────────────────────────────────────┼────────────────────────────────┤
  │ 7   │ Copy Worksheets        │ Copy with styling + reference shifting         │ Can add/remove, no copy        │
  ├─────┼────────────────────────┼────────────────────────────────────────────────┼────────────────────────────────┤
  │ 8   │ Copy/Move Ranges       │ Copy ranges across sheets/workbooks with       │ Not supported                  │
  │     │                        │ styles                                         │                                │
  ├─────┼────────────────────────┼────────────────────────────────────────────────┼────────────────────────────────┤
  │ 9   │ Insert/Delete Ranges   │ Auto-shift addresses when inserting/deleting   │ Not supported                  │
  ├─────┼────────────────────────┼────────────────────────────────────────────────┼────────────────────────────────┤
  │ 10  │ Sort Ranges            │ Multi-column sort, custom lists, asc/desc      │ Not supported                  │
  ├─────┼────────────────────────┼────────────────────────────────────────────────┼────────────────────────────────┤
  │ 11  │ AutoFit Columns        │ Calculate optimal column width from content    │ Not supported                  │
  ├─────┼────────────────────────┼────────────────────────────────────────────────┼────────────────────────────────┤
  │ 12  │ Fill Operations        │ FillNumber, FillDateTime, FillList             │ Not supported                  │
  ├─────┼────────────────────────┼────────────────────────────────────────────────┼────────────────────────────────┤
  │ 13  │ Shapes (creation)      │ 187 shape types with text, effects, gradients  │ Preserves existing, no         │
  │     │                        │                                                │ creation API                   │
  ├─────┼────────────────────────┼────────────────────────────────────────────────┼────────────────────────────────┤
  │ 14  │ Slicers                │ Table + pivot table slicers, 14 styles, custom │ Preserves existing, no         │
  │     │                        │  styles                                        │ creation API                   │
  ├─────┼────────────────────────┼────────────────────────────────────────────────┼────────────────────────────────┤
  │ 15  │ Custom Table Styles    │ Create from scratch or template                │ Only 27 built-in presets       │
  ├─────┼────────────────────────┼────────────────────────────────────────────────┼────────────────────────────────┤
  │ 16  │ Custom Pivot Table     │ Create from scratch                            │ Only built-in presets          │
  │     │ Styles                 │                                                │                                │
  └─────┴────────────────────────┴────────────────────────────────────────────────┴────────────────────────────────┘

  Lower Impact (Niche / Advanced)

  ┌─────┬────────────────────────┬────────────────────────────────────────────┬─────────────────────────────────────┐
  │  #  │        Feature         │                   EPPlus                   │             ExcelForge              │
  ├─────┼────────────────────────┼────────────────────────────────────────────┼─────────────────────────────────────┤
  │ 17  │ Threaded Comments      │ Mentions, thread resolution                │ Basic comments only                 │
  ├─────┼────────────────────────┼────────────────────────────────────────────┼─────────────────────────────────────┤
  │ 20  │ OLE Objects            │ Add, remove, read, link, embed             │ Not supported                       │
  ├─────┼────────────────────────┼────────────────────────────────────────────┼─────────────────────────────────────┤
  │ 21  │ VBA Code Signing       │ Certificate-based signing                  │ VBA creation/editing, no signing    │
  ├─────┼────────────────────────┼────────────────────────────────────────────┼─────────────────────────────────────┤
  │ 22  │ Theme loading (THMX)   │ Load/modify themes from files              │ Not supported                       │
  ├─────┼────────────────────────┼────────────────────────────────────────────┼─────────────────────────────────────┤
  │ 23  │ Chart templates (CRTX) │ Load chart designs from .crtx              │ Not supported                       │
  ├─────┼────────────────────────┼────────────────────────────────────────────┼─────────────────────────────────────┤
  │ 24  │ R1C1 Reference Style   │ Full support                               │ Not supported                       │
  ├─────┼────────────────────────┼────────────────────────────────────────────┼─────────────────────────────────────┤
  │ 25  │ Calculated Pivot       │ Custom calculated fields                   │ Not supported                       │
  │     │ Fields                 │                                            │                                     │
  ├─────┼────────────────────────┼────────────────────────────────────────────┼─────────────────────────────────────┤
  │ 26  │ Print Areas            │ Configure print areas                      │ Not explicitly supported            │
  ├─────┼────────────────────────┼────────────────────────────────────────────┼─────────────────────────────────────┤
  │ 27  │ XLTX template reading  │ Read template files                        │ Not supported                       │
  ├─────┼────────────────────────┼────────────────────────────────────────────┼─────────────────────────────────────┤
  │ 28  │ Async/Stream writing   │ Async, direct response stream              │ File/bytes output only              │
  └─────┴────────────────────────┴────────────────────────────────────────────┴─────────────────────────────────────┘

  Not Applicable (EPPlus .NET-specific)

  - LINQ queryable cell stores (.NET only)
  - DataReader/DataSet import (.NET types)
  - Strongly typed VBA API style (C# paradigm)

  ---
  Huge File Test Added

  Added example_huge_file() at src/test/examples.ts:1840 — generates 100 columns x 100,000 rows (10M cells) with mixed
  data types (numbers, strings, dates, nulls) plus summary formulas, then round-trips through write and read with
  verification.

  Benchmark results on this machine:
  - Populate: ~13s
  - Write: ~39s
  - Read: ~50s

  Note: 200 columns x 100k rows (20M cells) hits V8's Map maximum size limit (~16.7M entries) since cells are stored in
  a Map<string, Cell>. To reach the EPPlus benchmark of 200 cols, the cell storage would need to switch to a different
  data structure (e.g., a flat array indexed by row * maxCols + col, or a Map<number, Map<number, Cell>>).
