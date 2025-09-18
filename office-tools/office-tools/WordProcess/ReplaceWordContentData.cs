#nullable enable

using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
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
            foreach (var translation in translations)
            {
                ReplaceInSegments(segments, translation);
            }

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

    private static void ReplaceInSegments(List<TextSegment> segments, TranslationEntry translation)
    {
        var target = translation.OriginContent;
        if (target.Length == 0)
        {
            return;
        }

        var replacement = translation.TranlastedContent ?? string.Empty;

        while (true)
        {
            RecalculateStarts(segments);
            var combined = CombineSegments(segments);
            var index = combined.IndexOf(target, StringComparison.Ordinal);
            if (index < 0)
            {
                break;
            }

            ApplyReplacement(segments, index, target.Length, replacement);
        }
    }

    private static void ApplyReplacement(List<TextSegment> segments, int matchStart, int matchLength, string replacement)
    {
        RecalculateStarts(segments);
        var matchEnd = matchStart + matchLength;

        var startSegmentIndex = FindSegmentIndex(segments, matchStart);
        var endSegmentIndex = FindSegmentIndex(segments, matchEnd - 1);
        if (startSegmentIndex < 0 || endSegmentIndex < 0)
        {
            return;
        }

        var startSegment = segments[startSegmentIndex];
        var endSegment = segments[endSegmentIndex];

        var startOffset = matchStart - startSegment.Start;
        var endOffset = matchEnd - endSegment.Start;

        var prefix = startOffset > 0 ? startSegment.Text[..startOffset] : string.Empty;
        var suffix = endOffset < endSegment.Text.Length ? endSegment.Text[endOffset..] : string.Empty;

        startSegment.Text = prefix + replacement + suffix;

        for (var i = startSegmentIndex + 1; i <= endSegmentIndex; i++)
        {
            segments[i].Text = string.Empty;
        }
    }

    private static void RecalculateStarts(List<TextSegment> segments)
    {
        var offset = 0;
        foreach (var segment in segments)
        {
            segment.Start = offset;
            offset += segment.Text.Length;
        }
    }

    private static int FindSegmentIndex(List<TextSegment> segments, int position)
    {
        for (var i = 0; i < segments.Count; i++)
        {
            var segment = segments[i];
            var endExclusive = segment.Start + segment.Text.Length;
            if (position < endExclusive)
            {
                return i;
            }
        }

        return -1;
    }

    private static string CombineSegments(List<TextSegment> segments)
    {
        if (segments.Count == 1)
        {
            return segments[0].Text;
        }

        var builder = new System.Text.StringBuilder();
        foreach (var segment in segments)
        {
            builder.Append(segment.Text);
        }

        return builder.ToString();
    }

    private sealed record TranslationEntry(
        [property: JsonPropertyName("originContent")] string OriginContent,
        [property: JsonPropertyName("tranlastedContent")] string? TranlastedContent);

    private sealed class TextSegment
    {
        public TextSegment(XElement element)
        {
            Element = element;
            Text = element.Value;
        }

        public XElement Element { get; }
        public string Text { get; set; }
        public int Start { get; set; }
    }
}
