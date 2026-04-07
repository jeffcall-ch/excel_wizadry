// Copyright (c) FileExplorer contributors. All rights reserved.

using System.Diagnostics;
using System.IO.Compression;
using System.Text;
using FileExplorer.Infrastructure.Services;
using FluentAssertions;
using Microsoft.Extensions.Logging.Abstractions;
using Xunit;

namespace FileExplorer.Tests;

public sealed class IdentityExtractionServiceTests
{
    [Fact]
    public async Task ExtractIdentityAsync_Docx_ReturnsAnchorValues()
    {
        var service = new IdentityExtractionService(NullLogger<IdentityExtractionService>.Instance);
        var tempPath = Path.Combine(Path.GetTempPath(), $"id-docx-{Guid.NewGuid():N}.docx");

        try
        {
            CreateDocxSample(tempPath, "Project: CA100", "Revision: 04", "Doc Name: Valve List");

            var stopwatch = Stopwatch.StartNew();
            var identity = await service.ExtractIdentityAsync(tempPath);
            stopwatch.Stop();

            identity.ProjectTitle.Should().Be("CA100");
            identity.Revision.Should().Be("04");
            identity.DocumentName.Should().Be("Valve List");
            stopwatch.ElapsedMilliseconds.Should().BeLessThan(2000);
        }
        finally
        {
            TryDelete(tempPath);
        }
    }

    [Fact]
    public async Task ExtractIdentityAsync_Xlsx_UsesAnchorAdjacentCells()
    {
        var service = new IdentityExtractionService(NullLogger<IdentityExtractionService>.Instance);
        var tempPath = Path.Combine(Path.GetTempPath(), $"id-xlsx-{Guid.NewGuid():N}.xlsx");

        try
        {
            CreateXlsxSample(tempPath, "Project X", "B-17", "Pump Data");

            var stopwatch = Stopwatch.StartNew();
            var identity = await service.ExtractIdentityAsync(tempPath);
            stopwatch.Stop();

            identity.ProjectTitle.Should().Be("Project X");
            identity.Revision.Should().Be("B-17");
            identity.DocumentName.Should().Be("Pump Data");
            stopwatch.ElapsedMilliseconds.Should().BeLessThan(2000);
        }
        finally
        {
            TryDelete(tempPath);
        }
    }

    [Fact]
    public async Task ExtractIdentityAsync_Xlsx_PrefersCoverSheet_ProjectNameFromCellBelow()
    {
        var service = new IdentityExtractionService(NullLogger<IdentityExtractionService>.Instance);
        var tempPath = Path.Combine(Path.GetTempPath(), $"id-xlsx-coversheet-{Guid.NewGuid():N}.xlsx");

        try
        {
            CreateXlsxCoverSheetProjectNameBelowSample(tempPath, "CA100 Cover Project", "WRONG SUMMARY PROJECT");

            var identity = await service.ExtractIdentityAsync(tempPath);

            identity.ProjectTitle.Should().Be("CA100 Cover Project");
        }
        finally
        {
            TryDelete(tempPath);
        }
    }

    [Fact]
    public async Task ExtractIdentityAsync_Xlsx_DocumentNameFallsBackToFileName_WhenMissingInSheet()
    {
        var service = new IdentityExtractionService(NullLogger<IdentityExtractionService>.Instance);
        var tempFolder = Path.Combine(Path.GetTempPath(), $"id-xlsx-docname-fallback-{Guid.NewGuid():N}");
        Directory.CreateDirectory(tempFolder);
        var tempPath = Path.Combine(tempFolder, "price list lot general piping pipe material.xlsx");

        try
        {
            CreateXlsxSample(tempPath, "Project X", "B-17", string.Empty);

            var identity = await service.ExtractIdentityAsync(tempPath);

            identity.ProjectTitle.Should().Be("Project X");
            identity.Revision.Should().Be("B-17");
            identity.DocumentName.Should().Be("price list lot general piping pipe material");
        }
        finally
        {
            TryDelete(tempFolder);
        }
    }

    [Fact]
    public async Task ExtractIdentityAsync_Pdf_FirstPageProjectNameBelowAnchor_ReturnsProject()
    {
        var service = new IdentityExtractionService(NullLogger<IdentityExtractionService>.Instance);
        var tempPath = Path.Combine(Path.GetTempPath(), $"id-pdf-{Guid.NewGuid():N}.pdf");

        try
        {
            CreatePdfSample(tempPath,
                "Project Name",
                "CA100 PDF Project",
                "Revision: 3.0",
                "Doc Name: PDF Sheet");

            var identity = await service.ExtractIdentityAsync(tempPath);

            identity.ProjectTitle.Should().Be("CA100 PDF Project");
        }
        finally
        {
            TryDelete(tempPath);
        }
    }

    [Fact]
    public async Task ExtractIdentityAsync_Pdf_InternalRevisionPreferredOverFileName()
    {
        var service = new IdentityExtractionService(NullLogger<IdentityExtractionService>.Instance);
        var tempPath = Path.Combine(Path.GetTempPath(), $"id-pdf-rev-9.9-{Guid.NewGuid():N}.pdf");

        try
        {
            CreatePdfSample(tempPath,
                "Project Name",
                "CA100 PDF Project",
                "Revision: 3.0",
                "Doc Name: PDF Sheet");

            var identity = await service.ExtractIdentityAsync(tempPath);

            identity.Revision.Should().Be("3.0");
        }
        finally
        {
            TryDelete(tempPath);
        }
    }

    [Fact]
    public async Task ExtractIdentityAsync_Pdf_FileNameRevisionUsedWhenInternalMissing()
    {
        var service = new IdentityExtractionService(NullLogger<IdentityExtractionService>.Instance);
        var tempPath = Path.Combine(Path.GetTempPath(), $"id-pdf-rev-5.6-{Guid.NewGuid():N}.pdf");

        try
        {
            CreatePdfSample(tempPath,
                "Project Name",
                "CA100 PDF Project",
                "Doc Name: PDF Sheet");

            var identity = await service.ExtractIdentityAsync(tempPath);

            identity.Revision.Should().Be("5.6");
        }
        finally
        {
            TryDelete(tempPath);
        }
    }

    [Fact]
    public async Task ExtractIdentityAsync_Docx_ProjectUsesFirstPageHeaderAndIgnoresSecondPage()
    {
        var service = new IdentityExtractionService(NullLogger<IdentityExtractionService>.Instance);
        var tempPath = Path.Combine(Path.GetTempPath(), $"id-docx-header-{Guid.NewGuid():N}.docx");

        try
        {
            CreateDocxWithHeaderProjectAndSecondPageProject(tempPath, "CA100 HEADER", "WRONG SECOND PAGE");

            var identity = await service.ExtractIdentityAsync(tempPath);

            identity.ProjectTitle.Should().Be("CA100 HEADER");
        }
        finally
        {
            TryDelete(tempPath);
        }
    }

    [Fact]
    public async Task ExtractIdentityAsync_Docx_ProjectNumberIsIgnored_ProjectNameIsUsed()
    {
        var service = new IdentityExtractionService(NullLogger<IdentityExtractionService>.Instance);
        var tempPath = Path.Combine(Path.GetTempPath(), $"id-docx-projectname-{Guid.NewGuid():N}.docx");

        try
        {
            CreateDocxWithTextLines(tempPath,
                "Project Number: 14571",
                "Project Name: CA100 Boiler Hall",
                "Revision: 04",
                "Doc Name: Valve List");

            var identity = await service.ExtractIdentityAsync(tempPath);

            identity.ProjectTitle.Should().Be("CA100 Boiler Hall");
        }
        finally
        {
            TryDelete(tempPath);
        }
    }

    [Fact]
    public async Task ExtractIdentityAsync_Docx_RevisionFromFileName_IsPreferred()
    {
        var service = new IdentityExtractionService(NullLogger<IdentityExtractionService>.Instance);
        var tempPath = Path.Combine(Path.GetTempPath(), $"id-docx-rev-2.4-{Guid.NewGuid():N}.docx");

        try
        {
            CreateDocxSample(tempPath, "Project: CA100", "Revision: 9.9", "Doc Name: Valve List");

            var identity = await service.ExtractIdentityAsync(tempPath);

            identity.Revision.Should().Be("2.4");
        }
        finally
        {
            TryDelete(tempPath);
        }
    }

    [Fact]
    public async Task ExtractIdentityBatchAsync_MultipleFiles_CompletesWithinBudget()
    {
        var service = new IdentityExtractionService(NullLogger<IdentityExtractionService>.Instance);
        var tempFolder = Path.Combine(Path.GetTempPath(), $"id-batch-{Guid.NewGuid():N}");
        Directory.CreateDirectory(tempFolder);

        try
        {
            var files = new List<string>();
            for (var i = 0; i < 18; i++)
            {
                var path = Path.Combine(tempFolder, $"sample-{i:00}.docx");
                CreateDocxSample(path, $"Project: P-{i:00}", $"Revision: R-{i:00}", $"Doc Name: D-{i:00}");
                files.Add(path);
            }

            var stopwatch = Stopwatch.StartNew();
            var results = await service.ExtractIdentityBatchAsync(files);
            stopwatch.Stop();

            results.Count.Should().Be(files.Count);
            stopwatch.ElapsedMilliseconds.Should().BeLessThan(3000);
        }
        finally
        {
            TryDelete(tempFolder);
        }
    }

    private static void CreateDocxSample(string path, string projectLine, string revisionLine, string docNameLine)
    {
        using var zip = ZipFile.Open(path, ZipArchiveMode.Create);
        var docEntry = zip.CreateEntry("word/document.xml");

        using var writer = new StreamWriter(docEntry.Open());
        writer.Write("""
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:body>
                <w:p><w:r><w:t>
            """);
        writer.Write(projectLine);
        writer.Write("""
                </w:t></w:r></w:p>
                <w:p><w:r><w:t>
            """);
        writer.Write(revisionLine);
        writer.Write("""
                </w:t></w:r></w:p>
                <w:p><w:r><w:t>
            """);
        writer.Write(docNameLine);
        writer.Write("""
                </w:t></w:r></w:p>
              </w:body>
            </w:document>
            """);
    }

        private static void CreateDocxWithTextLines(string path, params string[] lines)
        {
                using var zip = ZipFile.Open(path, ZipArchiveMode.Create);
                var docEntry = zip.CreateEntry("word/document.xml");

                using var writer = new StreamWriter(docEntry.Open());
                writer.Write("""
                        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                            <w:body>
                        """);

                foreach (var line in lines)
                {
                        writer.Write("<w:p><w:r><w:t>");
                        writer.Write(line);
                        writer.Write("</w:t></w:r></w:p>");
                }

                writer.Write("""
                            </w:body>
                        </w:document>
                        """);
        }

        private static void CreateDocxWithHeaderProjectAndSecondPageProject(string path, string headerProject, string secondPageProject)
        {
                using var zip = ZipFile.Open(path, ZipArchiveMode.Create);

                var docEntry = zip.CreateEntry("word/document.xml");
                using (var writer = new StreamWriter(docEntry.Open()))
                {
                        writer.Write("""
                                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                                                        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                                    <w:body>
                                        <w:p><w:r><w:t>Cover page</w:t></w:r></w:p>
                                        <w:p><w:r><w:br w:type="page"/></w:r></w:p>
                                        <w:p><w:r><w:t>Project: 
                        """);
                        writer.Write(secondPageProject);
                        writer.Write("""
                                        </w:t></w:r></w:p>
                                        <w:sectPr>
                                            <w:headerReference w:type="default" r:id="rId1" />
                                        </w:sectPr>
                                    </w:body>
                                </w:document>
                                """);
                }

                var relsEntry = zip.CreateEntry("word/_rels/document.xml.rels");
                using (var writer = new StreamWriter(relsEntry.Open()))
                {
                        writer.Write("""
                                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                                    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml" />
                                </Relationships>
                                """);
                }

                var headerEntry = zip.CreateEntry("word/header1.xml");
                using (var writer = new StreamWriter(headerEntry.Open()))
                {
                        writer.Write("""
                                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                                <w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                                    <w:p><w:r><w:t>Project: 
                        """);
                        writer.Write(headerProject);
                        writer.Write("""
                                    </w:t></w:r></w:p>
                                </w:hdr>
                                """);
                }
        }

    private static void CreateXlsxSample(string path, string project, string revision, string docName)
    {
        using var zip = ZipFile.Open(path, ZipArchiveMode.Create);

        var workbook = zip.CreateEntry("xl/workbook.xml");
        using (var writer = new StreamWriter(workbook.Open()))
        {
            writer.Write("""
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <sheets>
                    <sheet name="CoverSheet" sheetId="1" r:id="rId1" />
                  </sheets>
                </workbook>
                """);
        }

        var rels = zip.CreateEntry("xl/_rels/workbook.xml.rels");
        using (var writer = new StreamWriter(rels.Open()))
        {
            writer.Write("""
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml" />
                </Relationships>
                """);
        }

        var sheet = zip.CreateEntry("xl/worksheets/sheet1.xml");
        using (var writer = new StreamWriter(sheet.Open()))
        {
            writer.Write("""
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                  <sheetData>
                    <row r="1">
                      <c r="A1" t="inlineStr"><is><t>Project:</t></is></c>
                      <c r="B1" t="inlineStr"><is><t>
                """);
            writer.Write(project);
            writer.Write("""
                      </t></is></c>
                    </row>
                    <row r="2">
                      <c r="A2" t="inlineStr"><is><t>Revision:</t></is></c>
                      <c r="B2" t="inlineStr"><is><t>
                """);
            writer.Write(revision);
            writer.Write("""
                      </t></is></c>
                    </row>
                    <row r="3">
                      <c r="A3" t="inlineStr"><is><t>Doc Name:</t></is></c>
                      <c r="B3" t="inlineStr"><is><t>
                """);
            writer.Write(docName);
            writer.Write("""
                      </t></is></c>
                    </row>
                  </sheetData>
                </worksheet>
                """);
        }
    }

        private static void CreateXlsxCoverSheetProjectNameBelowSample(string path, string coverProjectName, string summaryProjectName)
        {
                using var zip = ZipFile.Open(path, ZipArchiveMode.Create);

                var workbook = zip.CreateEntry("xl/workbook.xml");
                using (var writer = new StreamWriter(workbook.Open()))
                {
                        writer.Write("""
                                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                                <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                                                    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                                    <sheets>
                                        <sheet name="Summary" sheetId="1" r:id="rId1" />
                                        <sheet name="CoverSheet" sheetId="2" r:id="rId2" />
                                    </sheets>
                                </workbook>
                                """);
                }

                var rels = zip.CreateEntry("xl/_rels/workbook.xml.rels");
                using (var writer = new StreamWriter(rels.Open()))
                {
                        writer.Write("""
                                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                                    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml" />
                                    <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml" />
                                </Relationships>
                                """);
                }

                var summarySheet = zip.CreateEntry("xl/worksheets/sheet1.xml");
                using (var writer = new StreamWriter(summarySheet.Open()))
                {
                        writer.Write("""
                                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                                <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                                    <sheetData>
                                        <row r="1">
                                            <c r="A1" t="inlineStr"><is><t>Project:</t></is></c>
                                            <c r="B1" t="inlineStr"><is><t>
                                """);
                        writer.Write(summaryProjectName);
                        writer.Write("""
                                            </t></is></c>
                                        </row>
                                    </sheetData>
                                </worksheet>
                                """);
                }

                var coverSheet = zip.CreateEntry("xl/worksheets/sheet2.xml");
                using (var writer = new StreamWriter(coverSheet.Open()))
                {
                        writer.Write("""
                                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                                <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                                    <sheetData>
                                        <row r="5">
                                            <c r="B5" t="inlineStr"><is><t>Project Name</t></is></c>
                                        </row>
                                        <row r="6">
                                            <c r="B6" t="inlineStr"><is><t>
                                """);
                        writer.Write(coverProjectName);
                        writer.Write("""
                                            </t></is></c>
                                        </row>
                                    </sheetData>
                                </worksheet>
                                """);
                }
        }

    private static void TryDelete(string path)
    {
        try
        {
            if (File.Exists(path))
                File.Delete(path);
            else if (Directory.Exists(path))
                Directory.Delete(path, recursive: true);
        }
        catch
        {
            // Best effort cleanup for temp artifacts.
        }
    }

    private static void CreatePdfSample(string path, params string[] lines)
    {
        var contentBuilder = new StringBuilder();
        contentBuilder.AppendLine("BT");
        contentBuilder.AppendLine("/F1 12 Tf");
        contentBuilder.AppendLine("72 740 Td");

        foreach (var line in lines)
        {
            contentBuilder.Append('(');
            contentBuilder.Append(EscapePdfText(line));
            contentBuilder.AppendLine(") Tj");
            contentBuilder.AppendLine("0 -16 Td");
        }

        contentBuilder.AppendLine("ET");
        var contentStream = contentBuilder.ToString();
        var contentLength = Encoding.ASCII.GetByteCount(contentStream);

        var objects = new List<string>
        {
            "<< /Type /Catalog /Pages 2 0 R >>",
            "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Contents 4 0 R /Resources << /Font << /F1 5 0 R >> >> >>",
            $"<< /Length {contentLength} >>\nstream\n{contentStream}endstream",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>"
        };

        WriteMinimalPdf(path, objects);
    }

    private static void WriteMinimalPdf(string path, IReadOnlyList<string> objects)
    {
        using var memory = new MemoryStream();
        void WriteAscii(string text)
        {
            var bytes = Encoding.ASCII.GetBytes(text);
            memory.Write(bytes, 0, bytes.Length);
        }

        WriteAscii("%PDF-1.4\n");

        var offsets = new List<long>(objects.Count);
        for (var i = 0; i < objects.Count; i++)
        {
            offsets.Add(memory.Position);
            WriteAscii($"{i + 1} 0 obj\n");
            WriteAscii(objects[i]);
            WriteAscii("\nendobj\n");
        }

        var xrefStart = memory.Position;
        WriteAscii($"xref\n0 {objects.Count + 1}\n");
        WriteAscii("0000000000 65535 f \n");

        foreach (var offset in offsets)
        {
            WriteAscii($"{offset:0000000000} 00000 n \n");
        }

        WriteAscii($"trailer\n<< /Size {objects.Count + 1} /Root 1 0 R >>\n");
        WriteAscii($"startxref\n{xrefStart}\n%%EOF");

        File.WriteAllBytes(path, memory.ToArray());
    }

    private static string EscapePdfText(string value)
    {
        return value
            .Replace("\\", "\\\\", StringComparison.Ordinal)
            .Replace("(", "\\(", StringComparison.Ordinal)
            .Replace(")", "\\)", StringComparison.Ordinal);
    }
}
