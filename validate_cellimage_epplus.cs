#!/usr/bin/env dotnet run
#:package EPPlus@8.5.0

#pragma warning disable

using OfficeOpenXml;

ExcelPackage.License.SetNonCommercialOrganization("NonCommercialOrganization:@node-projects");

if (args.Length == 0) { Console.WriteLine("Usage: dotnet run validate_cellimage_epplus.cs <file>"); return; }

try
{
    using var pkg = new ExcelPackage(new FileInfo(args[0]));
    var ws = pkg.Workbook.Worksheets[0];
    Console.WriteLine($"Sheet: {ws.Name}");
    Console.WriteLine($"Dimensions: {ws.Dimension}");
    Console.WriteLine($"Drawings: {ws.Drawings.Count}");
    for (int i = 0; i < ws.Drawings.Count; i++)
        Console.WriteLine($"  Drawing {i}: {ws.Drawings[i].Name} ({ws.Drawings[i].GetType().Name})");

    // Check cell values
    if (ws.Dimension != null)
    {
        for (int r = ws.Dimension.Start.Row; r <= Math.Min(ws.Dimension.End.Row, 10); r++)
            for (int c = ws.Dimension.Start.Column; c <= ws.Dimension.End.Column; c++)
            {
                var val = ws.Cells[r, c].Value;
                if (val != null)
                    Console.WriteLine($"  [{r},{c}] = {val}");
            }
    }

    // Try accessing cell pictures
    try
    {
        var picB2 = ws.Cells["B2"].Picture;
        Console.WriteLine($"  B2.Picture: {(picB2 != null ? "EXISTS (" + picB2.GetType().Name + ")" : "null")}");
    }
    catch (Exception ex) { Console.WriteLine($"  B2.Picture: error - {ex.Message}"); }

    Console.WriteLine("EPPlus validation: OK");
}
catch (Exception ex)
{
    Console.WriteLine("EPPlus error: " + ex.ToString());
}
