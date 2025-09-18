#nullable enable

using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Xml.Linq;

namespace office_tools.WordProcess;

public static class ExtractWordContentData
{
    private const string SourceFileName = "OriginWord.docx";
    private const string OutputFileName = "ExtractWordContentData.json";
    private static readonly string[] LineSeparators = new[] { "\r\n", "\n", "\r" };

    public static void Generate()
    {
        var projectRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", ".."));
        var dataDirectory = Path.Combine(projectRoot, "Data");

        var sourcePath = Path.Combine(dataDirectory, SourceFileName);
        if (!File.Exists(sourcePath))
        {
            throw new FileNotFoundException("Source Word document not found.", sourcePath);
        }

        Directory.CreateDirectory(dataDirectory);

        using var stream = File.OpenRead(sourcePath);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: false);
        var entry = archive.GetEntry("word/document.xml")
                   ?? throw new InvalidDataException("Unable to locate main document part within the Word file.");

        using var entryStream = entry.Open();
        var document = XDocument.Load(entryStream);
        var wordNamespace = (XNamespace)"http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        var body = document.Root?.Element(wordNamespace + "body")
                   ?? throw new InvalidDataException("Word document is missing the body element.");

        var uniqueContents = new HashSet<string>(StringComparer.Ordinal);
        var results = new List<ExtractedEntry>();

        foreach (var paragraph in body.Elements(wordNamespace + "p"))
        {
            var builder = new StringBuilder();

            foreach (var node in paragraph.Descendants())
            {
                if (node.Name == wordNamespace + "t")
                {
                    builder.Append(node.Value);
                }
                else if (node.Name == wordNamespace + "tab")
                {
                    builder.Append('\t');
                }
                else if (node.Name == wordNamespace + "br" || node.Name == wordNamespace + "cr")
                {
                    builder.Append('\n');
                }
            }

            var paragraphText = builder.ToString();
            if (paragraphText.Length == 0)
            {
                continue;
            }

            var segments = paragraphText.Split(LineSeparators, StringSplitOptions.RemoveEmptyEntries);
            foreach (var segment in segments)
            {
                var content = segment.Trim();
                if (content.Length == 0)
                {
                    continue;
                }

                if (!uniqueContents.Add(content))
                {
                    continue;
                }

                results.Add(new ExtractedEntry(content, string.Empty));
            }
        }

        results.Sort(static (left, right) =>
        {
            var lengthComparison = right.OriginContent.Length.CompareTo(left.OriginContent.Length);
            if (lengthComparison != 0)
            {
                return lengthComparison;
            }

            return string.CompareOrdinal(left.OriginContent, right.OriginContent);
        });

        var options = new JsonSerializerOptions
        {
            Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            WriteIndented = true
        };

        var outputPath = Path.Combine(dataDirectory, OutputFileName);
        var json = JsonSerializer.Serialize(results, options);
        File.WriteAllText(outputPath, json);
    }

    private sealed record ExtractedEntry(string OriginContent, string TranlastedContent);
}
