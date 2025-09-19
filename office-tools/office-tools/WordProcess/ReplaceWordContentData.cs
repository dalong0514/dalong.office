#nullable enable

using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Xml.Linq;

namespace office_tools.WordProcess;

public static class ReplaceWordContentData
{
    private const string SourceJsonFileName = "TranslatedWordContentData.json";
    private const string SourceWordFileName = "OriginWord.docx";
    private const string OutputWordFileName = "TranslatedWord.docx";

    /// <summary>
    /// Applies translated content to the source Word document and saves a new file.
    /// </summary>
    public static void Generate()
    {
        var projectRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", ".."));
        var dataDirectory = Path.Combine(projectRoot, "Data");

        var jsonPath = Path.Combine(dataDirectory, SourceJsonFileName);
        if (!File.Exists(jsonPath))
        {
            throw new FileNotFoundException("Translated JSON file not found.", jsonPath);
        }

        var sourceDocPath = Path.Combine(dataDirectory, SourceWordFileName);
        if (!File.Exists(sourceDocPath))
        {
            throw new FileNotFoundException("Source Word document not found.", sourceDocPath);
        }

        Directory.CreateDirectory(dataDirectory);
        var translations = ReadTranslations(jsonPath);
        if (translations.Count == 0)
        {
            var outputPathEmpty = Path.Combine(dataDirectory, OutputWordFileName);
            File.Copy(sourceDocPath, outputPathEmpty, overwrite: true);
            return;
        }

        var outputPath = Path.Combine(dataDirectory, OutputWordFileName);
        File.Copy(sourceDocPath, outputPath, overwrite: true);

        using var wordStream = new FileStream(outputPath, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
        using var archive = new ZipArchive(wordStream, ZipArchiveMode.Update, leaveOpen: false);
        var documentEntry = archive.GetEntry("word/document.xml")
                           ?? throw new InvalidDataException("Unable to locate main document part within the Word file.");

        XDocument document;
        using (var entryStream = documentEntry.Open())
        {
            document = XDocument.Load(entryStream);
        }

        var wordNamespace = (XNamespace)"http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        var body = document.Root?.Element(wordNamespace + "body")
                   ?? throw new InvalidDataException("Word document is missing the body element.");

        foreach (var paragraph in body.Descendants(wordNamespace + "p"))
        {
            var textElements = paragraph.Descendants(wordNamespace + "t").ToList();
            if (textElements.Count == 0)
            {
                continue;
            }

            var segments = textElements.Select(e => new TextSegment(e)).ToList();
            var combined = CombineSegments(segments);
            if (combined.Length == 0)
            {
                continue;
            }

            var updated = combined;
            foreach (var translation in translations)
            {
                if (string.IsNullOrEmpty(translation.OriginContent))
                {
                    continue;
                }

                if (!updated.Contains(translation.OriginContent, StringComparison.Ordinal))
                {
                    continue;
                }

                var replacement = translation.TranlastedContent ?? string.Empty;
                updated = updated.Replace(translation.OriginContent, replacement, StringComparison.Ordinal);
            }

            DistributeTextAcrossSegments(segments, updated);

            foreach (var segment in segments)
            {
                segment.Element.Value = segment.Text;
            }
        }

        documentEntry.Delete();
        var newEntry = archive.CreateEntry("word/document.xml", CompressionLevel.Optimal);
        document.Declaration ??= new XDeclaration("1.0", "UTF-8", "yes");
        using (var outputStream = newEntry.Open())
        {
            document.Save(outputStream);
        }
    }

    /// <summary>
    /// Loads translation entries from JSON and sorts them by original length.
    /// </summary>
    private static List<TranslationEntry> ReadTranslations(string jsonPath)
    {
        var json = File.ReadAllText(jsonPath);
        var options = new JsonSerializerOptions
        {
            PropertyNameCaseInsensitive = true,
            ReadCommentHandling = JsonCommentHandling.Skip,
            AllowTrailingCommas = true,
            Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping
        };

        var entries = JsonSerializer.Deserialize<List<TranslationEntry>>(json, options) ?? new List<TranslationEntry>();
        entries.RemoveAll(static e => string.IsNullOrWhiteSpace(e.OriginContent));
        entries.Sort(static (left, right) =>
        {
            var lengthComparison = right.OriginContent.Length.CompareTo(left.OriginContent.Length);
            if (lengthComparison != 0)
            {
                return lengthComparison;
            }

            return string.CompareOrdinal(left.OriginContent, right.OriginContent);
        });

        return entries;
    }

    /// <summary>
    /// Creates a contiguous string from the ordered text segments of a paragraph.
    /// </summary>
    private static string CombineSegments(List<TextSegment> segments)
    {
        if (segments.Count == 1)
        {
            return segments[0].Text;
        }

        var builder = new StringBuilder();
        foreach (var segment in segments)
        {
            builder.Append(segment.Text);
        }

        return builder.ToString();
    }

    /// <summary>
    /// Writes the updated text back into the original segment boundaries.
    /// </summary>
    private static void DistributeTextAcrossSegments(List<TextSegment> segments, string text)
    {
        var offset = 0;
        for (var i = 0; i < segments.Count; i++)
        {
            var segment = segments[i];
            var remaining = text.Length - offset;
            if (remaining <= 0)
            {
                segment.Text = string.Empty;
                continue;
            }

            var take = i == segments.Count - 1
                ? remaining
                : Math.Min(remaining, segment.OriginalLength);

            segment.Text = take > 0 ? text.Substring(offset, take) : string.Empty;
            offset += take;
        }
    }

    /// <summary>
    /// Represents a translation entry with original and translated content.
    /// </summary>
    /// <param name="OriginContent"></param>
    /// <param name="TranlastedContent"></param>
    private sealed record TranslationEntry(
        [property: JsonPropertyName("originContent")] string OriginContent,
        [property: JsonPropertyName("tranlastedContent")] string? TranlastedContent);

    /// <summary>
    /// Represents a text segment within a paragraph, tracking its original length.
    /// </summary>
    private sealed class TextSegment
    {
        public TextSegment(XElement element)
        {
            Element = element;
            Text = element.Value;
            OriginalLength = Text.Length;
        }

        public XElement Element { get; }
        public string Text { get; set; }
        public int OriginalLength { get; }
    }
}
