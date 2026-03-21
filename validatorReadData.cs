#!/usr/bin/env dotnet run

#:package DocumentFormat.OpenXml@3.0.2

#pragma warning disable
using System.Text.Json;
using System.Text.Json.Serialization.Metadata;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

// Usage: dotnet run validatorReadData.cs <file.xlsx> <sheetName> <row> <colStart> <colEnd>
// Reads a single row from the given sheet and outputs cell values as JSON array.
// Example: dotnet run validatorReadData.cs output/20_loaded_table.xlsx ErrorsAndWarnings 10000 1 12

if (args.Length < 5)
{
    Console.Error.WriteLine("Usage: dotnet run validatorReadData.cs <file> <sheet> <row> <colStart> <colEnd>");
    Environment.Exit(1);
}

var file = args[0];
var sheetName = args[1];
var targetRow = int.Parse(args[2]);
var colStart = int.Parse(args[3]);
var colEnd = int.Parse(args[4]);

if (!File.Exists(file))
{
    Console.Error.WriteLine("File not found: " + file);
    Environment.Exit(1);
}

using var doc = SpreadsheetDocument.Open(file, false);
var wbPart = doc.WorkbookPart!;

// Find sheet by name
var sheet = wbPart.Workbook.Sheets!.Elements<Sheet>()
    .FirstOrDefault(s => s.Name?.Value == sheetName);

if (sheet == null)
{
    Console.Error.WriteLine("Sheet not found: " + sheetName);
    Environment.Exit(1);
}

var wsPart = (WorksheetPart)wbPart.GetPartById(sheet.Id!.Value!);
var sheetData = wsPart.Worksheet.GetFirstChild<SheetData>()!;

// Get shared strings
var sstPart = wbPart.SharedStringTablePart;
var sst = sstPart?.SharedStringTable;

// Find the target row
var rowEl = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex?.Value == (uint)targetRow);

var result = new List<object?>();

for (int col = colStart; col <= colEnd; col++)
{
    string colLetter = ColToLetter(col);
    string cellRef = colLetter + targetRow;

    Cell? cell = rowEl?.Elements<Cell>().FirstOrDefault(c => c.CellReference?.Value == cellRef);
    if (cell == null || cell.CellValue == null)
    {
        result.Add(null);
        continue;
    }

    string raw = cell.CellValue.Text;
    var dataType = cell.DataType?.Value;

    if (dataType == CellValues.SharedString)
    {
        int idx = int.Parse(raw);
        var item = sst?.Elements<SharedStringItem>().ElementAt(idx);
        result.Add(item?.InnerText ?? raw);
    }
    else if (dataType == CellValues.Boolean)
    {
        result.Add(raw == "1");
    }
    else
    {
        // Try numeric
        if (double.TryParse(raw, System.Globalization.NumberStyles.Any,
            System.Globalization.CultureInfo.InvariantCulture, out var num))
            result.Add(num);
        else
            result.Add(raw);
    }
}

var options = new JsonSerializerOptions
{
    WriteIndented = true,
    TypeInfoResolver = new DefaultJsonTypeInfoResolver()
};
Console.WriteLine(JsonSerializer.Serialize(result, options));

static string ColToLetter(int col)
{
    string s = "";
    while (col > 0)
    {
        int r = (col - 1) % 26;
        s = (char)('A' + r) + s;
        col = (col - 1) / 26;
    }
    return s;
}
