#!/usr/bin/env dotnet run
#:package EPPlus@8.5.0

#pragma warning disable

using OfficeOpenXml;
using System;
using System.IO;

ExcelPackage.License.SetNonCommercialOrganization("Test");

var file = args.Length > 0 ? args[0] : "output/22_vba_macros.xlsm";

try
{
    using (var package = new ExcelPackage(new FileInfo(file)))
    {
        Console.WriteLine($"File: {file}");
        Console.WriteLine($"Sheets: {package.Workbook.Worksheets.Count}");
        foreach (var ws in package.Workbook.Worksheets)
            Console.WriteLine($"  Sheet: {ws.Name}");

        if (package.Workbook.VbaProject != null)
        {
            Console.WriteLine($"VBA Project: YES");
            Console.WriteLine($"  Name: {package.Workbook.VbaProject.Name}");
            Console.WriteLine($"  Modules: {package.Workbook.VbaProject.Modules.Count}");
            foreach (var mod in package.Workbook.VbaProject.Modules)
            {
                Console.WriteLine($"  Module: {mod.Name} Type={mod.Type} Code={mod.Code?.Length ?? 0} chars");
                if (mod.Code?.Length > 0)
                    Console.WriteLine($"    Code: {mod.Code.Substring(0, Math.Min(100, mod.Code.Length))}");
            }
        }
        else
        {
            Console.WriteLine("VBA Project: NO");
        }
    }
}
catch (Exception ex)
{
    Console.WriteLine($"ERROR: {ex.GetType().Name}: {ex.Message}");
    Console.WriteLine($"  StackTrace: {ex.StackTrace}");
    if (ex.InnerException != null)
    {
        Console.WriteLine($"  Inner: {ex.InnerException.GetType().Name}: {ex.InnerException.Message}");
        Console.WriteLine($"  Inner StackTrace: {ex.InnerException.StackTrace}");
    }
}
