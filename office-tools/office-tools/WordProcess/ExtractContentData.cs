using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Xml.Linq;

namespace office_tools.WordProcess;

public static class ExtractContentData
{
    private static readonly XNamespace WordNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

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
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping
        });

        File.WriteAllText(targetPath, json, Encoding.UTF8);
    }

    private static List<string> ExtractLines(string documentPath)
    {
        using var fileStream = new FileStream(documentPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        using var archive = new ZipArchive(fileStream, ZipArchiveMode.Read, leaveOpen: false);

        var relevantEntries = archive.Entries
            .Where(IsRelevantWordEntry)
            .OrderBy(entry => entry.FullName, StringComparer.OrdinalIgnoreCase)
            .ToList();

        var lines = new List<string>();

        foreach (var entry in relevantEntries)
        {
            using var stream = entry.Open();
            var document = XDocument.Load(stream);
            var root = document.Root;
            if (root is null)
            {
                continue;
            }

            ExtractFromContainer(root, lines);
        }

        return lines;
    }

    private static void ExtractFromContainer(XElement container, ICollection<string> lines)
    {
        foreach (var element in container.Elements())
        {
            if (element.Name == WordNamespace + "p")
            {
                AppendNormalizedLines(GetParagraphText(element), lines);
                continue;
            }

            if (element.Name == WordNamespace + "tbl")
            {
                ExtractTable(element, lines);
                continue;
            }

            ExtractFromContainer(element, lines);
        }
    }

    private static void ExtractTable(XElement table, ICollection<string> lines)
    {
        foreach (var row in table.Elements(WordNamespace + "tr"))
        {
            foreach (var cell in row.Elements(WordNamespace + "tc"))
            {
                ExtractFromContainer(cell, lines);
            }
        }
    }

    private static string GetParagraphText(XElement paragraph)
    {
        var builder = new StringBuilder();

        foreach (var node in paragraph.Nodes())
        {
            AppendNodeText(node, builder);
        }

        return builder.ToString();
    }

    private static void AppendNodeText(XNode node, StringBuilder builder)
    {
        if (node is not XElement element)
        {
            return;
        }

        if (element.Name == WordNamespace + "t" ||
            element.Name == WordNamespace + "delText" ||
            element.Name == WordNamespace + "instrText")
        {
            builder.Append(element.Value);
            return;
        }

        if (element.Name == WordNamespace + "tab")
        {
            builder.Append('\t');
            return;
        }

        if (element.Name == WordNamespace + "br" || element.Name == WordNamespace + "cr")
        {
            builder.Append('\n');
            return;
        }

        foreach (var child in element.Nodes())
        {
            AppendNodeText(child, builder);
        }
    }

    private static void AppendNormalizedLines(string paragraphText, ICollection<string> lines)
    {
        if (string.IsNullOrWhiteSpace(paragraphText))
        {
            return;
        }

        var normalized = paragraphText.Replace("\r", string.Empty);
        foreach (var segment in normalized.Split('\n', StringSplitOptions.RemoveEmptyEntries))
        {
            var trimmed = segment.Trim();
            if (!string.IsNullOrWhiteSpace(trimmed))
            {
                lines.Add(trimmed);
            }
        }
    }

    private static bool IsRelevantWordEntry(ZipArchiveEntry entry)
    {
        if (!entry.FullName.StartsWith("word/", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        if (entry.FullName.Contains("/_rels/", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        var relativeName = entry.FullName.Substring("word/".Length);

        if (!relativeName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        if (relativeName.Equals("glossary/document.xml", StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        return relativeName.StartsWith("document", StringComparison.OrdinalIgnoreCase) ||
               relativeName.StartsWith("header", StringComparison.OrdinalIgnoreCase) ||
               relativeName.StartsWith("footer", StringComparison.OrdinalIgnoreCase) ||
               relativeName.StartsWith("footnotes", StringComparison.OrdinalIgnoreCase) ||
               relativeName.StartsWith("endnotes", StringComparison.OrdinalIgnoreCase) ||
               relativeName.StartsWith("comments", StringComparison.OrdinalIgnoreCase);
    }

    private sealed class ContentEntry
    {
        public string OriginContent { get; init; } = string.Empty;
        public string TranlastedContent { get; init; } = string.Empty;
    }
}
