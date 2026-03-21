#!/usr/bin/env dotnet run
#:package EPPlus@8.5.0

#pragma warning disable

using OfficeOpenXml;

if (args.Length == 0)
{
    Console.WriteLine("Usage: dotnet run validator-epplus.cs <file.xlsx>");
    return;
}

var file = args[0];

ExcelPackage.License.SetNonCommercialOrganization("NonCommercialOrganization:@node-projects");

try
{
    using var package = new ExcelPackage(new FileInfo(file));
    Console.WriteLine("EPPlus opened the file successfully.");
}
catch (Exception ex)
{
    Console.WriteLine("EPPlus error:");
    Console.WriteLine(ex.ToString());
}
