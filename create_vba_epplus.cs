#!/usr/bin/env dotnet run
#:package EPPlus@8.5.0

#pragma warning disable

using OfficeOpenXml;
using System;
using System.IO;

ExcelPackage.License.SetNonCommercialOrganization("Test");

using (var package = new ExcelPackage())
{
    var worksheet = package.Workbook.Worksheets.Add("Sheet1");
    worksheet.Cells["A1"].Value = "Hello";
    worksheet.Cells["A2"].Value = "Click the button to run the macro!";

    package.Workbook.CreateVBAProject();

    var module = package.Workbook.VbaProject.Modules.AddModule("Module1");
    module.Code = "Sub HelloWorld()\r\n    MsgBox \"Hello from EPPlus VBA!\"\r\nEnd Sub";

    package.SaveAs(new FileInfo("output/epplus_vba.xlsm"));
    Console.WriteLine("EPPlus VBA file created.");
}
