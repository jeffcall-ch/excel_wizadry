// Copyright (c) FileExplorer contributors. All rights reserved.

using System.Collections.Concurrent;
using System.Diagnostics;
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using FileExplorer.Core.Interfaces;
using FileExplorer.Core.Models;
using Microsoft.Extensions.Logging;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;

namespace FileExplorer.Infrastructure.Services;

/// <summary>
/// Extracts identity fields from Word and Excel families using anchor-based parsing.
/// </summary>
public sealed class IdentityExtractionService : IIdentityExtractionService
{
    private static readonly Regex ProjectAnchorRegex = new(
        @"\bproject\b",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private static readonly Regex ProjectAnchorValueRegex = new(
        @"\bproject(?:\s*name)?\b\s*[:\-]\s*(?<v>.+)$",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private static readonly Regex ProjectRegex = new(
        @"\bproject(?!\s*number)\s*(?:name)?\s*[:\-]\s*(?<v>[^\r\n\t\|;]+)",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private static readonly Regex RevisionRegex = new(
        @"\b(revision|rev\.?)\s*[:\-]\s*(?<v>[^\r\n\t\|;]+)",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private static readonly Regex PdfInternalRevisionRegex = new(
        @"(?:revision|rev)\s*[:\-]?\s*(?<v>[A-Z0-9.\-]+)",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private static readonly Regex PdfRevisionTokenRegex = new(
        @"^(?=.*\d)[A-Z0-9.\-]+$",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private static readonly Regex DocumentNameRegex = new(
        @"\b(doc(?:ument)?\s*name|title)\s*[:\-]\s*(?<v>[^\r\n\t\|;]+)",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private static readonly Regex RevisionFromFileNameRegex = new(
        @"(?<!\d)(?<rev>\d+\.\d+)(?!\d)",
        RegexOptions.Compiled);

    private static readonly HashSet<string> SupportedExtensions = new(StringComparer.OrdinalIgnoreCase)
    {
        ".doc", ".docx", ".docm", ".dot", ".dotx", ".dotm",
        ".xls", ".xlsx", ".xlsm", ".xlt", ".xltx", ".xltm", ".xlsb",
        ".pdf"
    };

    private readonly ConcurrentDictionary<string, CachedIdentity> _cache = new(StringComparer.OrdinalIgnoreCase);
    private readonly ILogger<IdentityExtractionService> _logger;

    private sealed record CachedIdentity(long LastWriteTicksUtc, long Length, DocumentIdentity Identity);

    /// <summary>
    /// Initializes a new instance of the <see cref="IdentityExtractionService"/> class.
    /// </summary>
    public IdentityExtractionService(ILogger<IdentityExtractionService> logger)
    {
        _logger = logger;
    }

    /// <inheritdoc/>
    public bool SupportsExtension(string extension)
    {
        if (string.IsNullOrWhiteSpace(extension))
            return false;

        if (!extension.StartsWith('.'))
            extension = "." + extension;

        return SupportedExtensions.Contains(extension);
    }

    /// <inheritdoc/>
    public async Task<DocumentIdentity> ExtractIdentityAsync(string filePath, CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
            return DocumentIdentity.Empty;

        var extension = Path.GetExtension(filePath);
        if (!SupportsExtension(extension))
            return DocumentIdentity.Empty;

        var fileInfo = new FileInfo(filePath);
        if (_cache.TryGetValue(filePath, out var cached) &&
            cached.LastWriteTicksUtc == fileInfo.LastWriteTimeUtc.Ticks &&
            cached.Length == fileInfo.Length)
        {
            return cached.Identity;
        }

        var stopwatch = Stopwatch.StartNew();

        try
        {
            cancellationToken.ThrowIfCancellationRequested();

            DocumentIdentity identity = extension.ToLowerInvariant() switch
            {
                ".docx" or ".docm" or ".dotx" or ".dotm" => await ExtractFromWordOpenXmlAsync(filePath, cancellationToken),
                ".xlsx" or ".xlsm" or ".xltx" or ".xltm" => await ExtractFromExcelOpenXmlAsync(filePath, cancellationToken),
                ".pdf" => await ExtractFromPdfFirstPageAsync(filePath, cancellationToken),
                ".doc" or ".dot" or ".xls" or ".xlt" or ".xlsb" => await ExtractFromLegacyBinaryScanAsync(filePath, cancellationToken),
                _ => DocumentIdentity.Empty
            };

            var revisionFromFileName = ExtractRevisionFromFileName(filePath);
            if (!string.IsNullOrWhiteSpace(revisionFromFileName))
            {
                // For PDFs we prefer in-document revision and keep filename as fallback only.
                if (!string.Equals(extension, ".pdf", StringComparison.OrdinalIgnoreCase) ||
                    string.IsNullOrWhiteSpace(identity.Revision))
                {
                    identity = identity with { Revision = revisionFromFileName };
                }
            }

            if (string.IsNullOrWhiteSpace(identity.DocumentName))
            {
                var documentNameFromFileName = ExtractDocumentNameFromFileName(filePath);
                if (!string.IsNullOrWhiteSpace(documentNameFromFileName))
                {
                    identity = identity with { DocumentName = documentNameFromFileName };
                }
            }

            _cache[filePath] = new CachedIdentity(fileInfo.LastWriteTimeUtc.Ticks, fileInfo.Length, identity);
            return identity;
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (Exception ex)
        {
            _logger.LogDebug(ex, "Identity extraction failed for {Path}", filePath);

            if (IsWordOpenXmlExtension(extension))
            {
                var fallbackIdentity = await TryExtractWordProjectFallbackAsync(filePath, cancellationToken);
                if (fallbackIdentity.HasAnyValue)
                    return fallbackIdentity;
            }

            return DocumentIdentity.Empty;
        }
        finally
        {
            stopwatch.Stop();
            _logger.LogDebug("Identity extraction elapsed: {ElapsedMs} ms for {Path}", stopwatch.ElapsedMilliseconds, filePath);
        }
    }

    /// <inheritdoc/>
    public async Task<IReadOnlyDictionary<string, DocumentIdentity>> ExtractIdentityBatchAsync(
        IReadOnlyList<string> filePaths,
        CancellationToken cancellationToken = default)
    {
        if (filePaths.Count == 0)
            return new Dictionary<string, DocumentIdentity>(StringComparer.OrdinalIgnoreCase);

        var filteredPaths = filePaths
            .Where(path => !string.IsNullOrWhiteSpace(path) && SupportsExtension(Path.GetExtension(path)))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        if (filteredPaths.Count == 0)
            return new Dictionary<string, DocumentIdentity>(StringComparer.OrdinalIgnoreCase);

        var stopwatch = Stopwatch.StartNew();
        var results = new ConcurrentDictionary<string, DocumentIdentity>(StringComparer.OrdinalIgnoreCase);

        var parallelOptions = new ParallelOptions
        {
            CancellationToken = cancellationToken,
            MaxDegreeOfParallelism = Math.Max(2, Environment.ProcessorCount)
        };

        await Parallel.ForEachAsync(filteredPaths, parallelOptions, async (path, ct) =>
        {
            var identity = await Task.Run(
                async () => await ExtractIdentityAsync(path, ct).ConfigureAwait(false),
                ct).ConfigureAwait(false);

            if (identity.HasAnyValue)
            {
                results[path] = identity;
            }
        });

        stopwatch.Stop();
        _logger.LogInformation(
            "Identity batch extraction finished in {ElapsedMs} ms. Input={InputCount} Populated={PopulatedCount}",
            stopwatch.ElapsedMilliseconds,
            filteredPaths.Count,
            results.Count);

        return new Dictionary<string, DocumentIdentity>(results, StringComparer.OrdinalIgnoreCase);
    }

    private static async Task<DocumentIdentity> ExtractFromWordOpenXmlAsync(string filePath, CancellationToken cancellationToken)
    {
        using var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: false);

        var documentEntry = archive.GetEntry("word/document.xml");
        if (documentEntry is null)
            return DocumentIdentity.Empty;

        await using var entryStream = documentEntry.Open();
        var documentXml = await XDocument.LoadAsync(entryStream, LoadOptions.None, cancellationToken);

        var firstPageBodyTokens = ExtractWordBodyFirstPageTokens(documentXml);
        var headerFooterTokens = await LoadWordHeaderFooterTokensAsync(archive, cancellationToken);

        var allTokens = new List<string>(headerFooterTokens.Count + firstPageBodyTokens.Count);
        allTokens.AddRange(headerFooterTokens);
        allTokens.AddRange(firstPageBodyTokens);

        if (allTokens.Count == 0)
            return DocumentIdentity.Empty;

        var firstPageText = string.Join('\n', allTokens);
        var parsedIdentity = ParseIdentityFromText(firstPageText);
        var projectFromAnchors = ExtractProjectFromWordTokens(allTokens);

        if (string.IsNullOrWhiteSpace(projectFromAnchors))
        {
            projectFromAnchors = await ExtractProjectFromWordHeaderFooterRawXmlAsync(archive, cancellationToken);
        }

        return new DocumentIdentity
        {
            ProjectTitle = FirstNonEmpty(projectFromAnchors, parsedIdentity.ProjectTitle),
            Revision = parsedIdentity.Revision,
            DocumentName = parsedIdentity.DocumentName
        };
    }

    private static async Task<DocumentIdentity> ExtractFromExcelOpenXmlAsync(string filePath, CancellationToken cancellationToken)
    {
        using var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: false);

        var workbookEntry = archive.GetEntry("xl/workbook.xml");
        if (workbookEntry is null)
            return DocumentIdentity.Empty;

        await using var workbookStream = workbookEntry.Open();
        var workbook = await XDocument.LoadAsync(workbookStream, LoadOptions.None, cancellationToken);
        var relationshipTargets = await LoadWorkbookRelationshipsAsync(archive, cancellationToken);
        var sharedStrings = await LoadSharedStringsAsync(archive, cancellationToken);

        var sheets = workbook
            .Descendants()
            .Where(node => node.Name.LocalName == "sheet")
            .Select(node =>
            {
                var relId = node.Attributes().FirstOrDefault(attr => attr.Name.LocalName == "id")?.Value;
                var name = node.Attribute("name")?.Value ?? string.Empty;
                return new SheetCandidate(name, relId);
            })
            .Where(candidate => !string.IsNullOrWhiteSpace(candidate.RelationshipId))
            .OrderByDescending(candidate => ScoreSheetName(candidate.Name))
            .ToList();

        if (sheets.Count == 0)
            return DocumentIdentity.Empty;

        var prioritizedSheets = sheets
            .Where(sheet => IsCoverSheetName(sheet.Name))
            .Concat(sheets.Where(sheet => !IsCoverSheetName(sheet.Name)))
            .ToList();

        IReadOnlyDictionary<string, string>? fallbackCells = null;

        foreach (var sheet in prioritizedSheets)
        {
            cancellationToken.ThrowIfCancellationRequested();

            if (!relationshipTargets.TryGetValue(sheet.RelationshipId!, out var targetPath))
                continue;

            var worksheetPath = NormalizeWorksheetPath(targetPath);
            var worksheetEntry = archive.GetEntry(worksheetPath);
            if (worksheetEntry is null)
                continue;

            await using var worksheetStream = worksheetEntry.Open();
            var worksheet = await XDocument.LoadAsync(worksheetStream, LoadOptions.None, cancellationToken);

            var cells = worksheet
                .Descendants()
                .Where(node => node.Name.LocalName == "c")
                .Select(node =>
                {
                    var reference = node.Attribute("r")?.Value ?? string.Empty;
                    var value = ResolveWorksheetCellValue(node, sharedStrings);
                    return (reference, value);
                })
                .Where(item => !string.IsNullOrWhiteSpace(item.reference) && !string.IsNullOrWhiteSpace(item.value))
                .ToDictionary(item => item.reference, item => item.value, StringComparer.OrdinalIgnoreCase);

            if (cells.Count == 0)
                continue;

            fallbackCells ??= cells;

            if (!string.IsNullOrWhiteSpace(FindProjectNameCellBelowAnchor(cells)))
                return ParseIdentityFromWorksheet(cells);
        }

        return fallbackCells is null
            ? DocumentIdentity.Empty
            : ParseIdentityFromWorksheet(fallbackCells);
    }

    private static async Task<DocumentIdentity> ExtractFromLegacyBinaryScanAsync(string filePath, CancellationToken cancellationToken)
    {
        const int maxBytesToScan = 4 * 1024 * 1024;

        using var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        var length = (int)Math.Min(stream.Length, maxBytesToScan);
        if (length <= 0)
            return DocumentIdentity.Empty;

        var buffer = new byte[length];
        var read = await stream.ReadAsync(buffer.AsMemory(0, length), cancellationToken);
        if (read <= 0)
            return DocumentIdentity.Empty;

        // Latin-1 preserves byte values for a best-effort anchor scan in binary legacy formats.
        var text = Encoding.Latin1.GetString(buffer, 0, read);
        return ParseIdentityFromText(text);
    }

    private static Task<DocumentIdentity> ExtractFromPdfFirstPageAsync(string filePath, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();

        return Task.Run(() =>
        {
            cancellationToken.ThrowIfCancellationRequested();

            var project = string.Empty;
            var revision = string.Empty;
            var documentName = string.Empty;

            try
            {
                using var document = PdfDocument.Open(filePath);
                if (document.NumberOfPages > 0)
                {
                    var page = document.GetPage(1);
                    var words = GetOrderedPdfWords(page);
                    var sortedFlowText = BuildPdfSortedFlowText(words);

                    project = PickLongerNonEmpty(
                        ExtractPdfProjectFromAnchors(words),
                        ExtractProjectFromPdfFirstPageLines(sortedFlowText));

                    revision = FirstNonEmpty(
                        ExtractPdfRevisionFromAnchors(words),
                        ExtractUsingRegex(sortedFlowText, PdfInternalRevisionRegex),
                        ExtractUsingRegex(sortedFlowText, RevisionRegex));

                    documentName = FirstNonEmpty(
                        ExtractPdfDocumentNameFromAnchors(words),
                        ExtractUsingRegex(sortedFlowText, DocumentNameRegex));
                }
            }
            catch
            {
                // Fall back to raw-content extraction below.
            }

            if (string.IsNullOrWhiteSpace(project) ||
                string.IsNullOrWhiteSpace(revision) ||
                string.IsNullOrWhiteSpace(documentName))
            {
                var rawFallbackText = ExtractPdfFirstPageTextRawFallback(filePath);
                if (!string.IsNullOrWhiteSpace(rawFallbackText))
                {
                    project = FirstNonEmpty(
                        project,
                        ExtractProjectFromPdfFirstPageLines(rawFallbackText));

                    revision = FirstNonEmpty(
                        revision,
                        ExtractUsingRegex(rawFallbackText, PdfInternalRevisionRegex),
                        ExtractUsingRegex(rawFallbackText, RevisionRegex));

                    documentName = FirstNonEmpty(
                        documentName,
                        ExtractUsingRegex(rawFallbackText, DocumentNameRegex));
                }
            }

            return new DocumentIdentity
            {
                ProjectTitle = project,
                Revision = revision,
                DocumentName = documentName
            };
        }, cancellationToken);
    }

    private static List<PdfWordToken> GetOrderedPdfWords(Page page)
    {
        return page
            .GetWords()
            .Select(word => new PdfWordToken(
                NormalizeWhitespace(word.Text),
                word.BoundingBox.Left,
                word.BoundingBox.Right,
                word.BoundingBox.Top,
                word.BoundingBox.Bottom))
            .Where(word => !string.IsNullOrWhiteSpace(word.Text))
            .OrderByDescending(word => word.Top)
            .ThenBy(word => word.Left)
            .ToList();
    }

    private static string BuildPdfSortedFlowText(IReadOnlyList<PdfWordToken> words)
    {
        if (words.Count == 0)
            return string.Empty;

        const double sameLineTolerance = 5d;
        var lines = new List<string>();
        var lineBuilder = new StringBuilder();
        var currentTop = words[0].Top;

        foreach (var word in words)
        {
            if (Math.Abs(word.Top - currentTop) > sameLineTolerance)
            {
                if (lineBuilder.Length > 0)
                {
                    lines.Add(NormalizeWhitespace(lineBuilder.ToString()));
                    lineBuilder.Clear();
                }

                currentTop = word.Top;
            }

            if (lineBuilder.Length > 0)
                lineBuilder.Append(' ');

            lineBuilder.Append(word.Text);
        }

        if (lineBuilder.Length > 0)
            lines.Add(NormalizeWhitespace(lineBuilder.ToString()));

        return string.Join('\n', lines.Where(line => !string.IsNullOrWhiteSpace(line)));
    }

    private static string ExtractPdfProjectFromAnchors(IReadOnlyList<PdfWordToken> words)
    {
        var anchors = FindPdfAnchors(words, "project name", "project");

        foreach (var anchor in anchors)
        {
            var belowLine = ExtractPdfLineBelowAnchor(words, anchor);
            if (IsValidProjectValueCandidate(belowLine) &&
                !ContainsAnyToken(belowLine, "revision", "rev", "document", "doc", "title"))
            {
                return belowLine.Length > 200 ? belowLine[..200] : belowLine;
            }

            var zoneWords = GetPdfTargetZoneWords(words, anchor);
            if (zoneWords.Count == 0)
                continue;

            var zoneText = BuildPdfSortedFlowText(zoneWords);
            var inline = TryExtractProjectValueFromAnchorToken($"{anchor.AnchorText}: {zoneText}");
            if (!string.IsNullOrWhiteSpace(inline))
                return inline;

            var candidate = zoneText
                .Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(NormalizeWhitespace)
                .Where(value =>
                    IsValidProjectValueCandidate(value) &&
                    !ContainsAnyToken(value, "project", "revision", "rev", "document", "doc", "title"))
                .OrderByDescending(value => value.Length)
                .FirstOrDefault();

            if (!string.IsNullOrWhiteSpace(candidate))
                return candidate.Length > 200 ? candidate[..200] : candidate;
        }

        return string.Empty;
    }

    private static string ExtractPdfRevisionFromAnchors(IReadOnlyList<PdfWordToken> words)
    {
        var anchors = FindPdfAnchors(words, "revision", "rev");

        foreach (var anchor in anchors)
        {
            var rightLine = ExtractPdfLineRightOfAnchor(words, anchor);
            var rightToken = rightLine
                .Split(' ', StringSplitOptions.RemoveEmptyEntries)
                .Select(NormalizeWhitespace)
                .FirstOrDefault(token => PdfRevisionTokenRegex.IsMatch(token));
            if (!string.IsNullOrWhiteSpace(rightToken))
                return rightToken;

            var belowLine = ExtractPdfLineBelowAnchor(words, anchor);
            var belowToken = belowLine
                .Split(' ', StringSplitOptions.RemoveEmptyEntries)
                .Select(NormalizeWhitespace)
                .FirstOrDefault(token => PdfRevisionTokenRegex.IsMatch(token));
            if (!string.IsNullOrWhiteSpace(belowToken))
                return belowToken;

            var zoneWords = GetPdfTargetZoneWords(words, anchor);
            if (zoneWords.Count == 0)
                continue;

            var zoneText = BuildPdfSortedFlowText(zoneWords);
            var revision = ExtractUsingRegex(zoneText, PdfInternalRevisionRegex);
            if (!string.IsNullOrWhiteSpace(revision))
                return revision;

            var revisionToken = zoneWords
                .Select(word => NormalizeWhitespace(word.Text))
                .FirstOrDefault(value =>
                    PdfRevisionTokenRegex.IsMatch(value) &&
                    !ContainsAnyToken(value, "project", "document", "title") &&
                    !string.Equals(value, "pdf", StringComparison.OrdinalIgnoreCase));

            if (!string.IsNullOrWhiteSpace(revisionToken))
                return revisionToken;
        }

        return string.Empty;
    }

    private static string ExtractPdfDocumentNameFromAnchors(IReadOnlyList<PdfWordToken> words)
    {
        var anchors = FindPdfAnchors(words, "document name", "doc name", "document", "title");

        foreach (var anchor in anchors)
        {
            var zoneWords = GetPdfTargetZoneWords(words, anchor);
            if (zoneWords.Count == 0)
                continue;

            var zoneText = BuildPdfSortedFlowText(zoneWords);
            var documentName = ExtractUsingRegex($"{anchor.AnchorText}: {zoneText}", DocumentNameRegex);
            if (!string.IsNullOrWhiteSpace(documentName))
                return documentName;

            var lineCandidate = zoneText
                .Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(NormalizeWhitespace)
                .FirstOrDefault(value =>
                    !string.IsNullOrWhiteSpace(value) &&
                    !ContainsAnyToken(value, "project", "revision", "rev") &&
                    value.Length >= 4);

            if (!string.IsNullOrWhiteSpace(lineCandidate))
                return lineCandidate.Length > 200 ? lineCandidate[..200] : lineCandidate;
        }

        return string.Empty;
    }

    private static List<PdfAnchorToken> FindPdfAnchors(IReadOnlyList<PdfWordToken> words, params string[] anchorPhrases)
    {
        var patterns = anchorPhrases
            .Where(phrase => !string.IsNullOrWhiteSpace(phrase))
            .Select(phrase => phrase
                .Split(' ', StringSplitOptions.RemoveEmptyEntries)
                .Select(token => token.Trim().ToLowerInvariant())
                .ToArray())
            .Where(tokens => tokens.Length > 0)
            .OrderByDescending(tokens => tokens.Length)
            .ToList();

        if (patterns.Count == 0 || words.Count == 0)
            return [];

        var anchors = new List<PdfAnchorToken>();
        for (var index = 0; index < words.Count; index++)
        {
            foreach (var pattern in patterns)
            {
                if (!TryMatchAnchorPattern(words, index, pattern, out var endIndex))
                    continue;

                if (pattern.Length == 1 &&
                    pattern[0] == "project" &&
                    index + 1 < words.Count &&
                    NormalizePdfToken(words[index + 1].Text) == "number")
                {
                    continue;
                }

                var matchedWords = words.Skip(index).Take(endIndex - index + 1).ToList();
                anchors.Add(new PdfAnchorToken(
                    index,
                    endIndex,
                    pattern[0],
                    matchedWords.Min(word => word.Left),
                    matchedWords.Max(word => word.Right),
                    matchedWords.Max(word => word.Top),
                    matchedWords.Min(word => word.Bottom),
                    string.Join(' ', pattern)));
                break;
            }
        }

        return anchors;
    }

    private static List<PdfWordToken> GetPdfTargetZoneWords(IReadOnlyList<PdfWordToken> words, PdfAnchorToken anchor)
    {
        const double horizontalBottomTolerance = 5d;
        const double verticalLeftTolerance = 20d;
        const double radius = 50d;

        return words
            .Select((word, index) => (word, index))
            .Where(item => item.index < anchor.StartIndex || item.index > anchor.EndIndex)
            .Where(item =>
            {
                var isHorizontal =
                    item.word.Left > anchor.Left &&
                    Math.Abs(item.word.Bottom - anchor.Bottom) <= horizontalBottomTolerance;

                var isVertical =
                    Math.Abs(item.word.Left - anchor.Left) <= verticalLeftTolerance &&
                    item.word.Bottom < anchor.Bottom;

                var inRadius = DistanceToAnchor(item.word, anchor) <= radius;

                return isHorizontal || isVertical || inRadius;
            })
            .Select(item => item.word)
            .OrderByDescending(word => word.Top)
            .ThenBy(word => word.Left)
            .ToList();
    }

    private static string ExtractPdfLineRightOfAnchor(IReadOnlyList<PdfWordToken> words, PdfAnchorToken anchor)
    {
        const double bottomTolerance = 5d;

        var rightWords = words
            .Where(word =>
                word.Left > anchor.Right &&
                Math.Abs(word.Bottom - anchor.Bottom) <= bottomTolerance)
            .OrderBy(word => word.Left)
            .ToList();

        if (rightWords.Count == 0)
            return string.Empty;

        return NormalizeWhitespace(string.Join(' ', rightWords.Select(word => word.Text)));
    }

    private static string ExtractPdfLineBelowAnchor(IReadOnlyList<PdfWordToken> words, PdfAnchorToken anchor)
    {
        const double leftTolerance = 20d;
        const double bottomTolerance = 5d;

        var firstBelowWord = words
            .Where(word =>
                word.Bottom < anchor.Bottom &&
                Math.Abs(word.Left - anchor.Left) <= leftTolerance)
            .OrderByDescending(word => word.Top)
            .ThenBy(word => word.Left)
            .FirstOrDefault();

        if (firstBelowWord is null)
            return string.Empty;

        var lineWords = words
            .Where(word =>
                Math.Abs(word.Bottom - firstBelowWord.Bottom) <= bottomTolerance &&
                word.Left >= firstBelowWord.Left)
            .OrderBy(word => word.Left)
            .ToList();

        if (lineWords.Count == 0)
            return string.Empty;

        return NormalizeWhitespace(string.Join(' ', lineWords.Select(word => word.Text)));
    }

    private static bool TryMatchAnchorPattern(
        IReadOnlyList<PdfWordToken> words,
        int startIndex,
        IReadOnlyList<string> pattern,
        out int endIndex)
    {
        endIndex = startIndex;
        if (startIndex + pattern.Count > words.Count)
            return false;

        for (var index = 0; index < pattern.Count; index++)
        {
            var token = NormalizePdfToken(words[startIndex + index].Text);
            if (!string.Equals(token, pattern[index], StringComparison.OrdinalIgnoreCase))
                return false;
        }

        endIndex = startIndex + pattern.Count - 1;
        return true;
    }

    private static string NormalizePdfToken(string token)
    {
        return NormalizeWhitespace(token)
            .Trim(':', ';', '-', '_', ',', '.', '(', ')', '[', ']', '{', '}', '"', '\'')
            .ToLowerInvariant();
    }

    private static double DistanceToAnchor(PdfWordToken word, PdfAnchorToken anchor)
    {
        var wordCenterX = (word.Left + word.Right) / 2d;
        var wordCenterY = (word.Top + word.Bottom) / 2d;

        var anchorCenterX = (anchor.Left + anchor.Right) / 2d;
        var anchorCenterY = (anchor.Top + anchor.Bottom) / 2d;

        var deltaX = wordCenterX - anchorCenterX;
        var deltaY = wordCenterY - anchorCenterY;
        return Math.Sqrt((deltaX * deltaX) + (deltaY * deltaY));
    }

    private static string ExtractProjectFromPdfFirstPageLines(string firstPageText)
    {
        var lines = firstPageText
            .Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
            .Select(NormalizeWhitespace)
            .Where(line => !string.IsNullOrWhiteSpace(line))
            .ToList();

        for (var index = 0; index < lines.Count; index++)
        {
            var line = lines[index];
            if (!line.Contains("project", StringComparison.OrdinalIgnoreCase))
                continue;

            if (IsDisallowedProjectAnchor(line))
                continue;

            var inlineValue = TryExtractProjectValueFromAnchorToken(line);
            if (!string.IsNullOrWhiteSpace(inlineValue))
                return inlineValue;

            if (!line.Contains("project name", StringComparison.OrdinalIgnoreCase))
                continue;

            for (var lookAhead = index + 1; lookAhead < lines.Count && lookAhead <= index + 3; lookAhead++)
            {
                var candidate = lines[lookAhead];
                if (string.IsNullOrWhiteSpace(candidate))
                    continue;

                if (!IsValidProjectValueCandidate(candidate))
                    continue;

                return candidate.Length > 200 ? candidate[..200] : candidate;
            }
        }

        return string.Empty;
    }

    private static string ExtractPdfFirstPageTextRawFallback(string filePath)
    {
        try
        {
            var rawBytes = File.ReadAllBytes(filePath);
            var rawText = Encoding.Latin1.GetString(rawBytes);

            var streamStart = rawText.IndexOf("stream", StringComparison.OrdinalIgnoreCase);
            if (streamStart < 0)
                return rawText;

            streamStart += "stream".Length;
            var streamEnd = rawText.IndexOf("endstream", streamStart, StringComparison.OrdinalIgnoreCase);
            if (streamEnd < 0)
                return rawText;

            var streamContent = rawText[streamStart..streamEnd];
            var matches = Regex.Matches(streamContent, @"\((?<t>(?:\\.|[^\\)])*)\)\s*T[Jj]", RegexOptions.Singleline);
            if (matches.Count == 0)
                return rawText;

            var lines = matches
                .Select(match => DecodePdfLiteral(match.Groups["t"].Value))
                .Select(NormalizeWhitespace)
                .Where(line => !string.IsNullOrWhiteSpace(line))
                .ToList();

            return lines.Count == 0 ? rawText : string.Join('\n', lines);
        }
        catch
        {
            return string.Empty;
        }
    }

    private static string DecodePdfLiteral(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
            return string.Empty;

        var builder = new StringBuilder(value.Length);
        for (var index = 0; index < value.Length; index++)
        {
            var ch = value[index];
            if (ch == '\\' && index + 1 < value.Length)
            {
                var next = value[index + 1];
                builder.Append(next switch
                {
                    'n' => '\n',
                    'r' => '\r',
                    't' => '\t',
                    '(' => '(',
                    ')' => ')',
                    '\\' => '\\',
                    _ => next
                });
                index++;
                continue;
            }

            builder.Append(ch);
        }

        return builder.ToString();
    }

    private static DocumentIdentity ParseIdentityFromWorksheet(IReadOnlyDictionary<string, string> cells)
    {
        var project = string.Empty;
        var revision = string.Empty;
        var documentName = string.Empty;

        foreach (var (cellRef, rawValue) in cells)
        {
            var value = NormalizeWhitespace(rawValue);
            if (string.IsNullOrWhiteSpace(value))
                continue;

            project = FirstNonEmpty(project,
                FindProjectNameCellBelowAnchor(cells, cellRef, value),
                ExtractUsingRegex(value, ProjectRegex),
                TryGetAdjacentAnchorValue(cells, cellRef, value, "project"));

            revision = FirstNonEmpty(revision,
                ExtractUsingRegex(value, RevisionRegex),
                TryGetAdjacentAnchorValue(cells, cellRef, value, "revision", "rev"));

            documentName = FirstNonEmpty(documentName,
                ExtractUsingRegex(value, DocumentNameRegex),
                TryGetAdjacentAnchorValue(cells, cellRef, value, "doc name", "document name", "title"));

            if (!string.IsNullOrWhiteSpace(project) &&
                !string.IsNullOrWhiteSpace(revision) &&
                !string.IsNullOrWhiteSpace(documentName))
            {
                break;
            }
        }

        if (string.IsNullOrWhiteSpace(project) || string.IsNullOrWhiteSpace(revision) || string.IsNullOrWhiteSpace(documentName))
        {
            var aggregateText = string.Join('\n', cells.Values);
            var fallback = ParseIdentityFromText(aggregateText);
            project = FirstNonEmpty(project, fallback.ProjectTitle);
            revision = FirstNonEmpty(revision, fallback.Revision);
            documentName = FirstNonEmpty(documentName, fallback.DocumentName);
        }

        return new DocumentIdentity
        {
            ProjectTitle = project,
            Revision = revision,
            DocumentName = documentName
        };
    }

    private static string FindProjectNameCellBelowAnchor(IReadOnlyDictionary<string, string> cells)
    {
        foreach (var (cellRef, rawValue) in cells)
        {
            var candidate = FindProjectNameCellBelowAnchor(cells, cellRef, rawValue);
            if (!string.IsNullOrWhiteSpace(candidate))
                return candidate;
        }

        return string.Empty;
    }

    private static string FindProjectNameCellBelowAnchor(IReadOnlyDictionary<string, string> cells, string cellReference, string value)
    {
        var anchorText = NormalizeWhitespace(value);
        if (!anchorText.Contains("project name", StringComparison.OrdinalIgnoreCase))
            return string.Empty;

        if (IsDisallowedProjectAnchor(anchorText))
            return string.Empty;

        var belowCellReference = TryGetAdjacentCellReference(cellReference, columnDelta: 0, rowDelta: 1);
        if (belowCellReference is not null &&
            cells.TryGetValue(belowCellReference, out var belowValue) &&
            IsValidProjectValueCandidate(belowValue))
        {
            return NormalizeWhitespace(belowValue);
        }

        var rightCellReference = TryGetAdjacentCellReference(cellReference, columnDelta: 1, rowDelta: 0);
        if (rightCellReference is not null &&
            cells.TryGetValue(rightCellReference, out var rightValue) &&
            IsValidProjectValueCandidate(rightValue))
        {
            return NormalizeWhitespace(rightValue);
        }

        return string.Empty;
    }

    private static bool IsValidProjectValueCandidate(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
            return false;

        var normalized = NormalizeWhitespace(value);
        if (IsDisallowedProjectAnchor(normalized))
            return false;

        if (string.Equals(normalized, "name", StringComparison.OrdinalIgnoreCase))
            return false;

        return !Regex.IsMatch(
            normalized,
            @"^project\s*(name|number)?\s*[:\-]?$",
            RegexOptions.IgnoreCase);
    }

    private static DocumentIdentity ParseIdentityFromText(string text)
    {
        if (string.IsNullOrWhiteSpace(text))
            return DocumentIdentity.Empty;

        return new DocumentIdentity
        {
            ProjectTitle = ExtractUsingRegex(text, ProjectRegex),
            Revision = ExtractUsingRegex(text, RevisionRegex),
            DocumentName = ExtractUsingRegex(text, DocumentNameRegex)
        };
    }

    private static string ExtractUsingRegex(string text, Regex regex)
    {
        var match = regex.Match(text);
        if (!match.Success)
            return string.Empty;

        var value = match.Groups["v"].Value;
        value = NormalizeWhitespace(value);

        return value.Length > 200 ? value[..200] : value;
    }

    private static string TryGetAdjacentAnchorValue(
        IReadOnlyDictionary<string, string> cells,
        string cellReference,
        string value,
        params string[] anchorTokens)
    {
        if (!ContainsAnyToken(value, anchorTokens))
            return string.Empty;

        if (anchorTokens.Any(token => token.Equals("project", StringComparison.OrdinalIgnoreCase)) &&
            IsDisallowedProjectAnchor(value))
        {
            return string.Empty;
        }

        var rightCellReference = TryGetAdjacentCellReference(cellReference, 1);
        if (rightCellReference is null)
            return string.Empty;

        if (!cells.TryGetValue(rightCellReference, out var adjacentValue))
            adjacentValue = string.Empty;

        var normalizedAdjacent = NormalizeWhitespace(adjacentValue);

        if (string.IsNullOrWhiteSpace(normalizedAdjacent))
        {
            var belowCellReference = TryGetAdjacentCellReference(cellReference, columnDelta: 0, rowDelta: 1);
            if (belowCellReference is not null && cells.TryGetValue(belowCellReference, out var belowValue))
            {
                normalizedAdjacent = NormalizeWhitespace(belowValue);
            }
        }

        if (string.IsNullOrWhiteSpace(normalizedAdjacent))
            return string.Empty;

        if (anchorTokens.Any(token => token.Equals("project", StringComparison.OrdinalIgnoreCase)) &&
            !IsValidProjectValueCandidate(normalizedAdjacent))
        {
            return string.Empty;
        }

        return normalizedAdjacent;
    }

    private static bool ContainsAnyToken(string value, params string[] tokens)
    {
        return tokens.Any(token => value.Contains(token, StringComparison.OrdinalIgnoreCase));
    }

    private static string NormalizeWhitespace(string value)
    {
        return Regex.Replace(value, @"\s+", " ").Trim();
    }

    private static string FirstNonEmpty(params string[] values)
    {
        return values.FirstOrDefault(value => !string.IsNullOrWhiteSpace(value)) ?? string.Empty;
    }

    private static string PickLongerNonEmpty(string primary, string secondary)
    {
        if (string.IsNullOrWhiteSpace(primary))
            return secondary;

        if (string.IsNullOrWhiteSpace(secondary))
            return primary;

        return secondary.Length > primary.Length ? secondary : primary;
    }

    private static string? TryGetAdjacentCellReference(string cellReference, int columnDelta)
    {
        return TryGetAdjacentCellReference(cellReference, columnDelta, rowDelta: 0);
    }

    private static string? TryGetAdjacentCellReference(string cellReference, int columnDelta, int rowDelta)
    {
        if (string.IsNullOrWhiteSpace(cellReference))
            return null;

        var letters = new string(cellReference.TakeWhile(char.IsLetter).ToArray());
        var digits = new string(cellReference.SkipWhile(char.IsLetter).ToArray());

        if (string.IsNullOrWhiteSpace(letters) || string.IsNullOrWhiteSpace(digits) || !int.TryParse(digits, out var row))
            return null;

        var column = ColumnLettersToNumber(letters) + columnDelta;
        var targetRow = row + rowDelta;
        if (column <= 0 || targetRow <= 0)
            return null;

        return NumberToColumnLetters(column) + targetRow;
    }

    private static int ColumnLettersToNumber(string letters)
    {
        var result = 0;
        foreach (var ch in letters.ToUpperInvariant())
        {
            result = (result * 26) + (ch - 'A' + 1);
        }

        return result;
    }

    private static string NumberToColumnLetters(int column)
    {
        var result = new StringBuilder();

        while (column > 0)
        {
            column--;
            result.Insert(0, (char)('A' + (column % 26)));
            column /= 26;
        }

        return result.ToString();
    }

    private static async Task<Dictionary<string, string>> LoadWorkbookRelationshipsAsync(ZipArchive archive, CancellationToken cancellationToken)
    {
        var entry = archive.GetEntry("xl/_rels/workbook.xml.rels");
        if (entry is null)
            return new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        await using var stream = entry.Open();
        var relationships = await XDocument.LoadAsync(stream, LoadOptions.None, cancellationToken);

        return relationships
            .Descendants()
            .Where(node => node.Name.LocalName == "Relationship")
            .Select(node =>
            {
                var id = node.Attribute("Id")?.Value;
                var target = node.Attribute("Target")?.Value;
                return (id, target);
            })
            .Where(item => !string.IsNullOrWhiteSpace(item.id) && !string.IsNullOrWhiteSpace(item.target))
            .ToDictionary(item => item.id!, item => item.target!, StringComparer.OrdinalIgnoreCase);
    }

    private static async Task<List<string>> LoadWordHeaderFooterTokensAsync(
        ZipArchive archive,
        CancellationToken cancellationToken)
    {
        var tokens = new List<string>();
        var partPaths = archive.Entries
            .Select(entry => entry.FullName)
            .Where(path =>
                path.StartsWith("word/header", StringComparison.OrdinalIgnoreCase) ||
                path.StartsWith("word/footer", StringComparison.OrdinalIgnoreCase))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        foreach (var partPath in partPaths)
        {
            cancellationToken.ThrowIfCancellationRequested();

            var partEntry = archive.GetEntry(partPath);
            if (partEntry is null)
                continue;

            try
            {
                await using var partStream = partEntry.Open();
                var partDocument = await XDocument.LoadAsync(partStream, LoadOptions.None, cancellationToken);
                tokens.AddRange(ExtractWordTextTokens(partDocument));
            }
            catch
            {
                // Skip malformed header/footer parts and continue best-effort extraction.
            }
        }

        return tokens;
    }

    private static List<string> ExtractWordBodyFirstPageTokens(XDocument document)
    {
        var tokens = new List<string>();
        var body = document.Descendants().FirstOrDefault(node => node.Name.LocalName == "body");
        if (body is null)
            return tokens;

        foreach (var element in body.Descendants())
        {
            if (IsWordPageBreakElement(element))
                break;

            if (element.Name.LocalName != "t")
                continue;

            var text = NormalizeWhitespace(element.Value);
            if (!string.IsNullOrWhiteSpace(text))
                tokens.Add(text);
        }

        return tokens;
    }

    private static List<string> ExtractWordTextTokens(XDocument document)
    {
        return document
            .Descendants()
            .Where(node => node.Name.LocalName == "t")
            .Select(node => NormalizeWhitespace(node.Value))
            .Where(value => !string.IsNullOrWhiteSpace(value))
            .ToList();
    }

    private static bool IsWordPageBreakElement(XElement element)
    {
        if (element.Name.LocalName == "lastRenderedPageBreak")
            return true;

        if (element.Name.LocalName != "br")
            return false;

        var breakType = element.Attributes().FirstOrDefault(attribute => attribute.Name.LocalName == "type")?.Value;
        return string.Equals(breakType, "page", StringComparison.OrdinalIgnoreCase);
    }

    private static string ExtractProjectFromWordTokens(IReadOnlyList<string> tokens)
    {
        for (var index = 0; index < tokens.Count; index++)
        {
            var token = NormalizeWhitespace(tokens[index]);
            if (string.IsNullOrWhiteSpace(token))
                continue;

            if (IsDisallowedProjectAnchor(token))
                continue;

            var inlineValue = TryExtractProjectValueFromAnchorToken(token);
            if (!string.IsNullOrWhiteSpace(inlineValue))
                return inlineValue;

            if (!ProjectAnchorRegex.IsMatch(token))
                continue;

            for (var lookAhead = index + 1; lookAhead < tokens.Count && lookAhead <= index + 3; lookAhead++)
            {
                var candidate = NormalizeWhitespace(tokens[lookAhead]);
                if (string.IsNullOrWhiteSpace(candidate))
                    continue;

                if (ContainsAnyToken(candidate, "project", "revision", "rev", "document", "doc name", "title"))
                    continue;

                return candidate.Length > 200 ? candidate[..200] : candidate;
            }
        }

        return string.Empty;
    }

    private static string TryExtractProjectValueFromAnchorToken(string token)
    {
        if (IsDisallowedProjectAnchor(token))
            return string.Empty;

        var match = ProjectAnchorValueRegex.Match(token);
        if (!match.Success)
            return string.Empty;

        var extracted = NormalizeWhitespace(match.Groups["v"].Value.Trim(' ', ':', '-'));
        if (string.IsNullOrWhiteSpace(extracted))
            return string.Empty;

        if (string.Equals(extracted, "name", StringComparison.OrdinalIgnoreCase))
            return string.Empty;

        if (ContainsAnyToken(extracted, "revision", "rev", "document", "doc name", "title"))
            return string.Empty;

        if (!IsValidProjectValueCandidate(extracted))
            return string.Empty;

        return extracted.Length > 200 ? extracted[..200] : extracted;
    }

    private static bool IsDisallowedProjectAnchor(string text)
    {
        if (string.IsNullOrWhiteSpace(text))
            return false;

        var normalized = NormalizeWhitespace(text);
        return normalized.Contains("project number", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsWordOpenXmlExtension(string extension)
    {
        return extension.Equals(".docx", StringComparison.OrdinalIgnoreCase) ||
            extension.Equals(".docm", StringComparison.OrdinalIgnoreCase) ||
            extension.Equals(".dotx", StringComparison.OrdinalIgnoreCase) ||
            extension.Equals(".dotm", StringComparison.OrdinalIgnoreCase);
    }

    private static async Task<DocumentIdentity> TryExtractWordProjectFallbackAsync(string filePath, CancellationToken cancellationToken)
    {
        try
        {
            using var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
            using var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: false);

            var headerFooterEntries = archive.Entries
                .Where(entry =>
                    entry.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase) &&
                    (entry.FullName.StartsWith("word/header", StringComparison.OrdinalIgnoreCase) ||
                     entry.FullName.StartsWith("word/footer", StringComparison.OrdinalIgnoreCase)))
                .ToList();

            foreach (var entry in headerFooterEntries)
            {
                cancellationToken.ThrowIfCancellationRequested();

                await using var entryStream = entry.Open();
                using var reader = new StreamReader(entryStream);
                var xml = await reader.ReadToEndAsync(cancellationToken);
                if (string.IsNullOrWhiteSpace(xml))
                    continue;

                var flattened = NormalizeWhitespace(Regex.Replace(xml, "<[^>]+>", " "));
                if (string.IsNullOrWhiteSpace(flattened))
                    continue;

                var headerFooterProject = FirstNonEmpty(
                    ExtractUsingRegex(flattened, ProjectRegex),
                    TryExtractProjectValueFromAnchorToken(flattened));

                if (!string.IsNullOrWhiteSpace(headerFooterProject))
                {
                    return new DocumentIdentity
                    {
                        ProjectTitle = headerFooterProject
                    };
                }
            }

            var documentEntry = archive.GetEntry("word/document.xml");
            if (documentEntry is null)
                return DocumentIdentity.Empty;

            cancellationToken.ThrowIfCancellationRequested();
            await using var documentStream = documentEntry.Open();
            using var documentReader = new StreamReader(documentStream);
            var documentXml = await documentReader.ReadToEndAsync(cancellationToken);
            if (string.IsNullOrWhiteSpace(documentXml))
                return DocumentIdentity.Empty;

            var firstPageXml = TakeFirstPageWordXml(documentXml);
            var mergedText = NormalizeWhitespace(Regex.Replace(firstPageXml, "<[^>]+>", " "));
            if (string.IsNullOrWhiteSpace(mergedText))
                return DocumentIdentity.Empty;

            var project = FirstNonEmpty(
                ExtractUsingRegex(mergedText, ProjectRegex),
                TryExtractProjectValueFromAnchorToken(mergedText));

            if (string.IsNullOrWhiteSpace(project))
                return DocumentIdentity.Empty;

            return new DocumentIdentity
            {
                ProjectTitle = project
            };
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch
        {
            return DocumentIdentity.Empty;
        }
    }

    private static string TakeFirstPageWordXml(string documentXml)
    {
        if (string.IsNullOrWhiteSpace(documentXml))
            return string.Empty;

        var pageBreakPattern = "<[^>]*lastRenderedPageBreak[^>]*>|<[^>]*:br[^>]*:type\\s*=\\s*['\\\"]page['\\\"][^>]*>";
        var split = Regex.Split(documentXml, pageBreakPattern, RegexOptions.IgnoreCase);
        return split.Length == 0 ? documentXml : split[0];
    }

    private static async Task<string> ExtractProjectFromWordHeaderFooterRawXmlAsync(ZipArchive archive, CancellationToken cancellationToken)
    {
        foreach (var entry in archive.Entries)
        {
            if (!entry.FullName.StartsWith("word/header", StringComparison.OrdinalIgnoreCase) &&
                !entry.FullName.StartsWith("word/footer", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            cancellationToken.ThrowIfCancellationRequested();

            try
            {
                await using var stream = entry.Open();
                using var reader = new StreamReader(stream);
                var xml = await reader.ReadToEndAsync(cancellationToken);
                if (string.IsNullOrWhiteSpace(xml))
                    continue;

                var flattened = NormalizeWhitespace(Regex.Replace(xml, "<[^>]+>", " "));
                if (string.IsNullOrWhiteSpace(flattened))
                    continue;

                var parsed = ExtractUsingRegex(flattened, ProjectRegex);
                if (!string.IsNullOrWhiteSpace(parsed))
                    return parsed;

                var anchorParsed = TryExtractProjectValueFromAnchorToken(flattened);
                if (!string.IsNullOrWhiteSpace(anchorParsed))
                    return anchorParsed;
            }
            catch
            {
                // Best-effort fallback: ignore malformed or inaccessible header/footer content.
            }
        }

        return string.Empty;
    }

    private static async Task<Dictionary<int, string>> LoadSharedStringsAsync(ZipArchive archive, CancellationToken cancellationToken)
    {
        var entry = archive.GetEntry("xl/sharedStrings.xml");
        if (entry is null)
            return new Dictionary<int, string>();

        await using var stream = entry.Open();
        var sharedStringsDocument = await XDocument.LoadAsync(stream, LoadOptions.None, cancellationToken);

        var sharedStrings = new Dictionary<int, string>();
        var index = 0;

        foreach (var sharedStringNode in sharedStringsDocument.Descendants().Where(node => node.Name.LocalName == "si"))
        {
            var text = string.Join(
                string.Empty,
                sharedStringNode.Descendants().Where(node => node.Name.LocalName == "t").Select(node => node.Value));

            sharedStrings[index++] = text;
        }

        return sharedStrings;
    }

    private static string ResolveWorksheetCellValue(XElement cell, IReadOnlyDictionary<int, string> sharedStrings)
    {
        var cellType = cell.Attribute("t")?.Value;

        if (string.Equals(cellType, "inlineStr", StringComparison.OrdinalIgnoreCase))
        {
            var inlineText = string.Join(
                string.Empty,
                cell.Descendants().Where(node => node.Name.LocalName == "t").Select(node => node.Value));

            return NormalizeWhitespace(inlineText);
        }

        var valueNode = cell.Elements().FirstOrDefault(node => node.Name.LocalName == "v");
        if (valueNode is null)
            return string.Empty;

        var rawValue = valueNode.Value;

        if (string.Equals(cellType, "s", StringComparison.OrdinalIgnoreCase) && int.TryParse(rawValue, out var sharedStringIndex))
        {
            return sharedStrings.TryGetValue(sharedStringIndex, out var sharedValue)
                ? NormalizeWhitespace(sharedValue)
                : string.Empty;
        }

        return NormalizeWhitespace(rawValue);
    }

    private static string NormalizeWorksheetPath(string relationshipTarget)
    {
        var target = relationshipTarget.Replace('\\', '/');
        if (target.StartsWith('/'))
            target = target.TrimStart('/');

        if (!target.StartsWith("xl/", StringComparison.OrdinalIgnoreCase))
            target = "xl/" + target.TrimStart('/');

        return target;
    }

    private static int ScoreSheetName(string? name)
    {
        if (string.IsNullOrWhiteSpace(name))
            return 0;

        var normalized = name.Trim().ToLowerInvariant();

        if (normalized.Contains("coversheet") || normalized.Contains("cover sheet"))
            return 120;

        if (normalized.Contains("cover") || normalized.Contains("deckblatt") || normalized.Contains("titelblatt"))
            return 100;

        if (normalized.Contains("title") || normalized.Contains("summary"))
            return 80;

        return normalized switch
        {
            "sheet1" => 30,
            "blatt1" => 30,
            _ => 0
        };
    }

    private static bool IsCoverSheetName(string? name)
    {
        if (string.IsNullOrWhiteSpace(name))
            return false;

        var normalized = name.Trim().ToLowerInvariant();
        return normalized.Contains("coversheet") ||
               normalized.Contains("cover sheet") ||
               normalized.Contains("cover") ||
               normalized.Contains("deckblatt") ||
               normalized.Contains("titelblatt");
    }

    private static string ExtractRevisionFromFileName(string filePath)
    {
        var name = Path.GetFileNameWithoutExtension(filePath);
        if (string.IsNullOrWhiteSpace(name))
            return string.Empty;

        var matches = RevisionFromFileNameRegex.Matches(name);
        if (matches.Count == 0)
            return string.Empty;

        var value = matches[^1].Groups["rev"].Value;
        return string.IsNullOrWhiteSpace(value) ? string.Empty : value;
    }

    private static string ExtractDocumentNameFromFileName(string filePath)
    {
        var name = Path.GetFileNameWithoutExtension(filePath);
        if (string.IsNullOrWhiteSpace(name))
            return string.Empty;

        var normalized = NormalizeWhitespace(name.Replace('_', ' '));
        return normalized.Length > 200 ? normalized[..200] : normalized;
    }

    private sealed record SheetCandidate(string Name, string? RelationshipId);

    private sealed record PdfWordToken(string Text, double Left, double Right, double Top, double Bottom);

    private sealed record PdfAnchorToken(
        int StartIndex,
        int EndIndex,
        string AnchorKey,
        double Left,
        double Right,
        double Top,
        double Bottom,
        string AnchorText);
}
