using System;
using System.IO;
using System.IO.Compression;
using System.Text;
using MSOfficeAIAssistant.Attachments;
using MSOfficeAIAssistant.Hosts;
using System.Collections.Generic;

namespace MSOfficeAIAssistant.Tests
{
    public static class AttachmentExtractorProvenanceTests
    {
        public static void RunAll()
        {
            TestDocxParagraphIndexing();
            TestXlsxRealSheetNamesAndCellAddresses();
            TestWordControllerExcerptLineOffsets();
        }

        private static void TestDocxParagraphIndexing()
        {
            // Build a minimal in-memory .docx with 3 paragraphs (including one empty)
            string tempDir = Path.Combine(Path.GetTempPath(), "docx_test_" + Guid.NewGuid().ToString());
            Directory.CreateDirectory(tempDir);
            string docxPath = Path.Combine(tempDir, "test.docx");

            try
            {
                CreateTestDocx(docxPath);

                // Call the public async method and wait for result
                var task = AttachmentExtractor.ExtractAsync(docxPath);
                var block = task.GetAwaiter().GetResult();

                string text = block.ExtractedText;
                Assert(text.Contains("[¶1]"), "Expected paragraph 1 tag");
                Assert(text.Contains("[¶3]"), "Expected paragraph 3 tag (paragraph 2 is empty so counter advanced but nothing output)");
                Assert(text.Contains("[¶1] First paragraph"), "Expected paragraph 1 content");
                Assert(text.Contains("[¶3] Third paragraph"), "Expected paragraph 3 content");
                Assert(!text.Contains("[¶2]"), "Empty paragraph 2 should not appear in output");

                // Verify that non-empty paragraphs have the index prefix
                int para1Index = text.IndexOf("[¶1]");
                int para3Index = text.IndexOf("[¶3]");
                Assert(para1Index >= 0, "Paragraph 1 index should be found");
                Assert(para3Index > para1Index, "Paragraph 3 should come after paragraph 1");
            }
            finally
            {
                try { File.Delete(docxPath); } catch { }
                try { Directory.Delete(tempDir); } catch { }
            }
        }

        private static void TestXlsxRealSheetNamesAndCellAddresses()
        {
            // Build a minimal in-memory .xlsx with a real sheet name "Budget" and cell addresses
            string tempDir = Path.Combine(Path.GetTempPath(), "xlsx_test_" + Guid.NewGuid().ToString());
            Directory.CreateDirectory(tempDir);
            string xlsxPath = Path.Combine(tempDir, "test.xlsx");

            try
            {
                CreateTestXlsx(xlsxPath);

                // Call the public async method and wait for result
                var task = AttachmentExtractor.ExtractAsync(xlsxPath);
                var block = task.GetAwaiter().GetResult();

                string text = block.ExtractedText;
                Assert(text.Contains("Budget"), "Expected real sheet name 'Budget'");
                Assert(text.Contains("B2="), "Expected cell address B2");
                Assert(text.Contains("100"), "Expected cell value");
                Assert(!text.Contains("Sheet 1"), "Should not use fallback ordinal when real name is available");
            }
            finally
            {
                try { File.Delete(xlsxPath); } catch { }
                try { Directory.Delete(tempDir); } catch { }
            }
        }

        private static void TestWordControllerExcerptLineOffsets()
        {
            // Build a multi-paragraph document text that will be chunked
            var sb = new StringBuilder();
            for (int i = 0; i < 10; i++)
            {
                sb.AppendLine("Line " + (i + 1) + ": This is a paragraph with some content to make it long enough.");
            }
            string documentText = sb.ToString();

            // Call the public static method with a limited character budget to force chunking
            string context = WordController.BuildRelevantDocumentContext(documentText, "content", 500);

            // Verify that the excerpt labels include paragraph line offsets
            Assert(context.Contains("~Paragraph"), "Expected excerpt label with paragraph offset indicator (~Paragraph)");
            Assert(context.Contains("[Excerpt"), "Expected excerpt label format");
            Assert(context.Contains("of"), "Expected 'of' in excerpt count");

            // Extract and verify paragraph numbers are present and reasonable
            int excerptCount = CountOccurrences(context, "[Excerpt");
            Assert(excerptCount >= 1, string.Format("Expected at least 1 excerpt, found {0}", excerptCount));

            // Verify that paragraph numbers increase (or at least are present)
            int paraPos = context.IndexOf("~Paragraph");
            Assert(paraPos >= 0, "Expected ~Paragraph marker");
        }

        private static void CreateTestDocx(string path)
        {
            using (var zip = ZipFile.Open(path, ZipArchiveMode.Create))
            {
                // Add [Content_Types].xml
                var contentTypesEntry = zip.CreateEntry("[Content_Types].xml");
                using (var stream = contentTypesEntry.Open())
                using (var writer = new StreamWriter(stream, Encoding.UTF8))
                {
                    writer.Write("<?xml version=\"1.0\"?>");
                    writer.Write("<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">");
                    writer.Write("<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>");
                    writer.Write("<Default Extension=\"xml\" ContentType=\"application/xml\"/>");
                    writer.Write("<Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>");
                    writer.Write("</Types>");
                }

                // Add _rels/.rels
                var relsDir = zip.CreateEntry("_rels/");
                var relsEntry = zip.CreateEntry("_rels/.rels");
                using (var stream = relsEntry.Open())
                using (var writer = new StreamWriter(stream, Encoding.UTF8))
                {
                    writer.Write("<?xml version=\"1.0\"?>");
                    writer.Write("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
                    writer.Write("<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/>");
                    writer.Write("</Relationships>");
                }

                // Add word/document.xml with 3 paragraphs (1st and 3rd non-empty, 2nd empty)
                var wordDir = zip.CreateEntry("word/");
                var docEntry = zip.CreateEntry("word/document.xml");
                using (var stream = docEntry.Open())
                using (var writer = new StreamWriter(stream, Encoding.UTF8))
                {
                    writer.Write("<?xml version=\"1.0\"?>");
                    writer.Write("<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">");
                    writer.Write("<w:body>");
                    writer.Write("<w:p><w:r><w:t>First paragraph</w:t></w:r></w:p>");
                    writer.Write("<w:p></w:p>"); // Empty paragraph
                    writer.Write("<w:p><w:r><w:t>Third paragraph</w:t></w:r></w:p>");
                    writer.Write("</w:body>");
                    writer.Write("</w:document>");
                }
            }
        }

        private static void CreateTestXlsx(string path)
        {
            using (var zip = ZipFile.Open(path, ZipArchiveMode.Create))
            {
                // Add [Content_Types].xml
                var contentTypesEntry = zip.CreateEntry("[Content_Types].xml");
                using (var stream = contentTypesEntry.Open())
                using (var writer = new StreamWriter(stream, Encoding.UTF8))
                {
                    writer.Write("<?xml version=\"1.0\"?>");
                    writer.Write("<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">");
                    writer.Write("<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>");
                    writer.Write("<Default Extension=\"xml\" ContentType=\"application/xml\"/>");
                    writer.Write("<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>");
                    writer.Write("<Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>");
                    writer.Write("<Override PartName=\"/xl/sharedStrings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/>");
                    writer.Write("</Types>");
                }

                // Add _rels/.rels
                var relsEntry = zip.CreateEntry("_rels/.rels");
                using (var stream = relsEntry.Open())
                using (var writer = new StreamWriter(stream, Encoding.UTF8))
                {
                    writer.Write("<?xml version=\"1.0\"?>");
                    writer.Write("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
                    writer.Write("<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>");
                    writer.Write("</Relationships>");
                }

                // Add xl/_rels/workbook.xml.rels
                var xlRelsEntry = zip.CreateEntry("xl/_rels/workbook.xml.rels");
                using (var stream = xlRelsEntry.Open())
                using (var writer = new StreamWriter(stream, Encoding.UTF8))
                {
                    writer.Write("<?xml version=\"1.0\"?>");
                    writer.Write("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
                    writer.Write("<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>");
                    writer.Write("<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\" Target=\"sharedStrings.xml\"/>");
                    writer.Write("</Relationships>");
                }

                // Add xl/workbook.xml with sheet name "Budget"
                var xlDir = zip.CreateEntry("xl/");
                var workbookEntry = zip.CreateEntry("xl/workbook.xml");
                using (var stream = workbookEntry.Open())
                using (var writer = new StreamWriter(stream, Encoding.UTF8))
                {
                    writer.Write("<?xml version=\"1.0\"?>");
                    writer.Write("<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
                    writer.Write("<sheets>");
                    writer.Write("<sheet name=\"Budget\" sheetId=\"1\" r:id=\"rId1\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"/>");
                    writer.Write("</sheets>");
                    writer.Write("</workbook>");
                }

                // Add xl/sharedStrings.xml (empty for this test, we'll use numeric literals)
                var sstEntry = zip.CreateEntry("xl/sharedStrings.xml");
                using (var stream = sstEntry.Open())
                using (var writer = new StreamWriter(stream, Encoding.UTF8))
                {
                    writer.Write("<?xml version=\"1.0\"?>");
                    writer.Write("<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"></sst>");
                }

                // Add xl/worksheets/sheet1.xml with cell B2=100
                var worksheetsDir = zip.CreateEntry("xl/worksheets/");
                var sheet1Entry = zip.CreateEntry("xl/worksheets/sheet1.xml");
                using (var stream = sheet1Entry.Open())
                using (var writer = new StreamWriter(stream, Encoding.UTF8))
                {
                    writer.Write("<?xml version=\"1.0\"?>");
                    writer.Write("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
                    writer.Write("<sheetData>");
                    writer.Write("<row r=\"1\"><c r=\"A1\"><v>Header1</v></c><c r=\"B1\"><v>Header2</v></c></row>");
                    writer.Write("<row r=\"2\"><c r=\"A2\"><v>Data</v></c><c r=\"B2\"><v>100</v></c></row>");
                    writer.Write("</sheetData>");
                    writer.Write("</worksheet>");
                }
            }
        }

        private static int CountOccurrences(string text, string pattern)
        {
            if (string.IsNullOrEmpty(text) || string.IsNullOrEmpty(pattern)) return 0;
            int count = 0;
            int index = 0;
            while ((index = text.IndexOf(pattern, index, StringComparison.Ordinal)) >= 0)
            {
                count++;
                index += pattern.Length;
            }
            return count;
        }

        private static void Assert(bool condition, string message)
        {
            if (!condition)
            {
                throw new Exception("Assertion failed: " + message);
            }
        }
    }
}
