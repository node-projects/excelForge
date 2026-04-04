#!/usr/bin/env dotnet run
#:package EPPlus@8.5.0
#:package SixLabors.ImageSharp@3.1.7

#pragma warning disable

using OfficeOpenXml;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.PixelFormats;
using SixLabors.ImageSharp.Formats.Png;

ExcelPackage.License.SetNonCommercialOrganization("NonCommercialOrganization:@node-projects");

// Create a real PNG in memory
byte[] pngBytes;
using (var img = new Image<Rgba32>(40, 40))
{
    for (int y = 0; y < 40; y++)
        for (int x = 0; x < 40; x++)
            img[x, y] = new Rgba32(255, 0, 0, 255);
    using var ms = new MemoryStream();
    img.Save(ms, new PngEncoder());
    pngBytes = ms.ToArray();
}

try
{
    using var pkg = new ExcelPackage();
    var ws = pkg.Workbook.Worksheets.Add("CellImages");
    ws.Cells["A1"].Value = "In-cell picture test";
    ws.Cells["A2"].Value = "Red:";
    ws.Row(2).Height = 40;
    ws.Column(2).Width = 12;

    // Use the public Range.Picture.Set API
    ws.Cells["B2"].Picture.Set(pngBytes, "Red square");

    pkg.SaveAs(new FileInfo("output/epplus_cellimage.xlsx"));
    Console.WriteLine("Created output/epplus_cellimage.xlsx");
}
catch (Exception ex)
{
    Console.WriteLine("Error: " + ex);
}
