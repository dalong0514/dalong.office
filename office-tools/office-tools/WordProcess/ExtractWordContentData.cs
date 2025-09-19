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
    private static readonly XNamespace WordNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

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
        var body = document.Root?.Element(WordNamespace + "body")
                   ?? throw new InvalidDataException("Word document is missing the body element.");

        var uniqueContents = new HashSet<string>(StringComparer.Ordinal);
        var results = new List<ExtractedEntry>();

        foreach (var paragraph in body.Descendants(WordNamespace + "p"))
        {
            var paragraphText = ExtractParagraphText(paragraph);
            if (paragraphText.Length == 0)
            {
                continue;
            }

            var segments = paragraphText.Split(LineSeparators, StringSplitOptions.RemoveEmptyEntries);
            foreach (var segment in segments)
            {
                var content = segment.Trim();
                if (content.Length == 0 || IsPureEnglishContent(content))
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

    private static string ExtractParagraphText(XElement paragraph)
    {
        var builder = new StringBuilder();

        foreach (var node in paragraph.Descendants())
        {
            if (node.Name == WordNamespace + "t" || node.Name == WordNamespace + "delText" || node.Name == WordNamespace + "instrText")
            {
                builder.Append(node.Value);
            }
            else if (node.Name == WordNamespace + "tab")
            {
                builder.Append('\t');
            }
            else if (node.Name == WordNamespace + "br" || node.Name == WordNamespace + "cr")
            {
                builder.Append('\n');
            }
        }

        return builder.ToString();
    }

    private static bool IsPureEnglishContent(string text)
    {
        var hasLetterOrDigit = false;

        foreach (var ch in text)
        {
            if (ch > 127)
            {
                return false;
            }

            if (char.IsLetterOrDigit(ch))
            {
                hasLetterOrDigit = true;
                continue;
            }

            if (char.IsWhiteSpace(ch) || char.IsPunctuation(ch) || char.IsSymbol(ch))
            {
                continue;
            }

            return false;
        }

        return hasLetterOrDigit;
    }

    private sealed record ExtractedEntry(string OriginContent, string TranlastedContent);
}
