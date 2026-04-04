#!/usr/bin/env dotnet run

#:package DocumentFormat.OpenXml@3.0.2

#pragma warning disable
using System.IO.Compression;
using System.Text.Json;
using System.Text.Json.Serialization.Metadata;
using System.Xml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;

var result = new List<object>();

if (args.Length == 0)
{
    Console.WriteLine("Usage: dotnet run validator.cs <file.xlsx>");
    Environment.Exit(1);
}

var file = args[0];

if (!File.Exists(file))
{
    Console.WriteLine("File not found.");
    Environment.Exit(1);
}

// ------------------------
// 1. ZIP + XML inspection
// ------------------------
try
{
    using var zip = ZipFile.OpenRead(file);

    foreach (var entry in zip.Entries)
    {
        try
        {
            using var stream = entry.Open();
            using var ms = new MemoryStream();
            stream.CopyTo(ms);
            ms.Position = 0;

            if (entry.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
            {
                try
                {
                    var settings = new XmlReaderSettings
                    {
                        DtdProcessing = DtdProcessing.Prohibit,
                        IgnoreWhitespace = true,
                        IgnoreComments = true
                    };

                    ms.Position = 0;
                    using var reader = XmlReader.Create(ms, settings);
                    while (reader.Read()) { }
                }
                catch (Exception ex)
                {
                    result.Add(new
                    {
                        type = "xml",
                        entry = entry.FullName,
                        error = ex.Message
                    });
                }
            }
        }
        catch (Exception ex)
        {
            result.Add(new
            {
                type = "zip-entry",
                entry = entry.FullName,
                error = ex.Message
            });
        }
    }
}
catch (Exception ex)
{
    result.Add(new
    {
        type = "zip",
        error = ex.Message
    });

    PrintAndExit(result);
}

// ------------------------
// 2. Attempt OpenXML validation if possible
// ------------------------
try
{
    using var doc = SpreadsheetDocument.Open(file, false);
    var validator = new OpenXmlValidator(DocumentFormat.OpenXml.FileFormatVersions.Microsoft365);
    var errors = validator.Validate(doc);

    foreach (var e in errors)
    {
        // Skip false positives for mc:AlternateContent in drawing anchors —
        // the OpenXML SDK doesn't process markup compatibility elements within
        // twoCellAnchor / oneCellAnchor properly. Excel itself handles them fine.
        if (e.Part?.Uri?.ToString()?.Contains("/drawings/") == true &&
            e.Node?.OuterXml?.Contains("AlternateContent") == true)
            continue;

        result.Add(new
        {
            type = "openxml",
            description = e.Description,
            part = e.Part?.Uri?.ToString(),
            path = e.Path?.XPath,
            errorType = e.ErrorType.ToString(),
            node = e.Node?.OuterXml,
            relatedNode = e.RelatedNode?.OuterXml
        });
    }
}
catch (Exception ex)
{
    // Report all nested exceptions for OpenXML
    int level = 0;
    var current = ex;
    while (current != null)
    {
        result.Add(new
        {
            type = "openxml-exception",
            level,
            exception = current.GetType().Name,
            message = current.Message,
            innerMessage = ex.InnerException?.Message
        });

        current = current.InnerException;
        level++;
    }
}

// ------------------------
// 3. Output JSON + exit code
// ------------------------
PrintAndExit(result);

// ------------------------
void PrintAndExit(List<object> result)
{
    var options = new JsonSerializerOptions
    {
        WriteIndented = true,
        TypeInfoResolver = new DefaultJsonTypeInfoResolver()
    };

    string json = JsonSerializer.Serialize(result, options);
    Console.WriteLine(json);

    Environment.Exit(result.Count == 0 ? 0 : 1);
}
