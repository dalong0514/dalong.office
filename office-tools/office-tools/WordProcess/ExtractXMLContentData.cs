#nullable enable

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace office_tools.WordProcess;

public static class ExtractXMLContentData
{
    private const string SourceFileName = "test.xml";
    private const string OutputFileName = "ExtractXMLContentData.json";
    private static readonly Regex TextPattern = new("<w:t(?:\\s[^>]*)?>(.*?)</w:t>", RegexOptions.Singleline | RegexOptions.Compiled);

    public static void Generate()
    {
        var projectRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", ".."));
        var dataDirectory = Path.Combine(projectRoot, "Data");

        var sourcePath = Path.Combine(dataDirectory, SourceFileName);
        if (!File.Exists(sourcePath))
        {
            throw new FileNotFoundException("Source XML file not found.", sourcePath);
        }

        Directory.CreateDirectory(dataDirectory);

        var xmlContent = File.ReadAllText(sourcePath);
        var uniqueContents = new HashSet<string>(StringComparer.Ordinal);
        var results = new List<ExtractedEntry>();

        foreach (Match match in TextPattern.Matches(xmlContent))
        {
            var inner = match.Groups[1].Value;
            if (inner.TrimStart().StartsWith("<w:titlePg/>", StringComparison.Ordinal))
            {
                continue;
            }

            var content = inner.Trim();
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
