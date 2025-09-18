using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Xml.Linq;

namespace office_tools.WordProcess;

public static class ExtractContentData
{
    public static void Generate(string? sourceDocumentPath = null, string? outputPath = null)
    {
        var projectRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", ".."));
        var documentPath = sourceDocumentPath ?? Path.Combine(projectRoot, "..", "..", "ZData", "test.docx");
        if (!File.Exists(documentPath))
        {
            throw new FileNotFoundException($"Word document not found at {documentPath}");
        }

        var targetPath = outputPath ?? Path.Combine(projectRoot, "Data", "ExtractContentData.json");
        Directory.CreateDirectory(Path.GetDirectoryName(targetPath)!);

        var lines = ExtractLines(documentPath);
        var payload = lines.Select(line => new ContentEntry
        {
            OriginContent = line,
            TranlastedContent = string.Empty
        }).ToList();

        var json = JsonSerializer.Serialize(payload, new JsonSerializerOptions
        {
            WriteIndented = true,
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase
        });

        File.WriteAllText(targetPath, json, Encoding.UTF8);
    }

    private static List<string> ExtractLines(string documentPath)
    {
        var lines = new List<string>();
        using var archive = ZipFile.OpenRead(documentPath);
        var entry = archive.GetEntry("word/document.xml")
                    ?? throw new InvalidOperationException("word/document.xml entry not found in the document.");

        using var stream = entry.Open();
        var xdoc = XDocument.Load(stream);
        XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

        foreach (var paragraph in xdoc.Descendants(w + "p"))
        {
            var builder = new StringBuilder();
            foreach (var node in paragraph.Descendants())
            {
                if (node.Name == w + "t")
                {
                    builder.Append(node.Value);
                }
                else if (node.Name == w + "tab")
                {
                    builder.Append('\t');
                }
                else if (node.Name == w + "br")
                {
                    builder.Append('\n');
                }
            }

            var normalized = builder.ToString().Replace("\r", string.Empty);
            var segments = normalized.Split('\n', StringSplitOptions.RemoveEmptyEntries);
            foreach (var segment in segments)
            {
                var trimmed = segment.Trim();
                if (!string.IsNullOrEmpty(trimmed))
                {
                    lines.Add(trimmed);
                }
            }
        }

        return lines;
    }

    private sealed class ContentEntry
    {
        public string OriginContent { get; init; } = string.Empty;
        public string TranlastedContent { get; init; } = string.Empty;
    }
}
