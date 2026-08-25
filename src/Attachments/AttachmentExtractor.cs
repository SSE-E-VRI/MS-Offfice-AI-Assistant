using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml.Linq;
using MSOfficeAIAssistant.Core;
using MSOfficeAIAssistant.Providers;

namespace MSOfficeAIAssistant.Attachments
{
    public static class AttachmentExtractor
    {
        public const long MaxPerFileSizeBytes = 20 * 1024 * 1024; // 20 MB
        public const long MaxTotalSizeBytes = 30 * 1024 * 1024;   // 30 MB
        public const int MaxFileCount = 10;
        public const int MaxExtractedCharacters = 50000;
        private const long MaxDecompressedBytes = 100 * 1024 * 1024; // 100 MB decompressed limit

        private static readonly HashSet<string> ImageExtensions = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            ".png", ".jpg", ".jpeg", ".webp", ".gif"
        };

        private static readonly HashSet<string> TextExtensions = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            ".txt", ".csv", ".json", ".xml", ".md", ".cs", ".py", ".js", ".html", ".css", ".sql", ".log", ".ini", ".yaml", ".yml"
        };

        private static readonly HashSet<string> LegacyExtensions = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            ".doc", ".xls", ".ppt", ".rtf"
        };

        public static async Task<AttachmentBlock> ExtractAsync(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath))
                throw new ArgumentNullException("filePath");

            var fileInfo = new FileInfo(filePath);
            if (!fileInfo.Exists)
                throw new FileNotFoundException("Attachment file not found: " + filePath);

            if (fileInfo.Length > MaxPerFileSizeBytes)
            {
                throw new InvalidOperationException(string.Format(
                    "File '{0}' exceeds the maximum allowed single file size limit of 20 MB ({1:F2} MB).",
                    fileInfo.Name, fileInfo.Length / (1024.0 * 1024.0)));
            }

            string ext = fileInfo.Extension.ToLowerInvariant();

            // Explicit rejection of legacy formats
            if (LegacyExtensions.Contains(ext))
            {
                throw new NotSupportedException(string.Format(
                    "Legacy binary format '{0}' is not supported. Please save as modern OpenXML ({1}x) or export as PDF.",
                    fileInfo.Name, ext));
            }

            var block = new AttachmentBlock
            {
                FileName = fileInfo.Name,
                FileSizeBytes = fileInfo.Length
            };

            if (ImageExtensions.Contains(ext))
            {
                block.IsImage = true;
                block.ContentType = GetMimeType(ext);
                block.RawBytes = File.ReadAllBytes(filePath);
                block.ExtractedText = string.Format("[Image Attachment: {0}]", fileInfo.Name);
                return block;
            }

            block.IsImage = false;
            string text = string.Empty;

            if (ext == ".docx")
            {
                text = await Task.Run(() => ExtractDocx(filePath)).ConfigureAwait(false);
                block.ContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
            }
            else if (ext == ".xlsx")
            {
                text = await Task.Run(() => ExtractXlsx(filePath)).ConfigureAwait(false);
                block.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            }
            else if (ext == ".pptx")
            {
                text = await Task.Run(() => ExtractPptx(filePath)).ConfigureAwait(false);
                block.ContentType = "application/vnd.openxmlformats-officedocument.presentationml.presentation";
            }
            else if (ext == ".pdf")
            {
                text = await Task.Run(() => ExtractPdf(filePath)).ConfigureAwait(false);
                block.ContentType = "application/pdf";
            }
            else if (TextExtensions.Contains(ext))
            {
                using (var reader = new StreamReader(filePath, Encoding.UTF8, true))
                {
                    text = await reader.ReadToEndAsync().ConfigureAwait(false);
                }
                block.ContentType = "text/plain";
            }
            else
            {
                throw new NotSupportedException(string.Format("Unsupported file extension: '{0}'. Supported: .docx, .xlsx, .pptx, .pdf, images, and text files.", ext));
            }

            if (text.Length > MaxExtractedCharacters)
            {
                text = text.Substring(0, MaxExtractedCharacters) + "\n\n[... Truncated to fit context window ...]";
            }

            block.ExtractedText = text;
            return block;
        }

        private static string ExtractDocx(string path)
        {
            var sb = new StringBuilder();
            using (var zip = ZipFile.OpenRead(path))
            {
                if (ArchiveDecompressedSizeExceedsLimit(zip))
                {
                    return "[Attachment too large when decompressed to process safely.]";
                }

                var entry = zip.GetEntry("word/document.xml");
                if (entry != null)
                {
                    using (var stream = entry.Open())
                    {
                        var doc = XDocument.Load(stream);
                        XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
                        int paragraphIndex = 1;
                        foreach (var p in doc.Descendants(w + "p"))
                        {
                            string paragraphText = string.Concat(p.Descendants(w + "t").Select(t => t.Value));
                            if (!string.IsNullOrWhiteSpace(paragraphText))
                            {
                                sb.AppendLine(string.Format("[¶{0}] {1}", paragraphIndex, paragraphText));
                            }
                            paragraphIndex++;
                        }
                    }
                }
            }
            return sb.ToString();
        }

        private static string ExtractXlsx(string path)
        {
            var sb = new StringBuilder();
            var sharedStrings = new List<string>();
            var sheetNames = new List<string>();

            using (var zip = ZipFile.OpenRead(path))
            {
                if (ArchiveDecompressedSizeExceedsLimit(zip))
                {
                    return "[Attachment too large when decompressed to process safely.]";
                }

                // Read Sheet Names from workbook.xml
                try
                {
                    var workbookEntry = zip.GetEntry("xl/workbook.xml");
                    if (workbookEntry != null)
                    {
                        using (var stream = workbookEntry.Open())
                        {
                            var doc = XDocument.Load(stream);
                            XNamespace s = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
                            foreach (var sheet in doc.Descendants(s + "sheet"))
                            {
                                string name = (string)sheet.Attribute("name");
                                if (!string.IsNullOrEmpty(name))
                                {
                                    sheetNames.Add(name);
                                }
                            }
                        }
                    }
                }
                catch
                {
                    // If sheet name resolution fails, sheetNames remains empty and we'll fall back to ordinals
                }

                // Read Shared Strings
                var sstEntry = zip.GetEntry("xl/sharedStrings.xml");
                if (sstEntry != null)
                {
                    using (var stream = sstEntry.Open())
                    {
                        var doc = XDocument.Load(stream);
                        XNamespace s = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
                        foreach (var si in doc.Descendants(s + "si"))
                        {
                            string str = string.Concat(si.Descendants(s + "t").Select(t => t.Value));
                            sharedStrings.Add(str);
                        }
                    }
                }

                // Read Worksheets
                var sheetEntries = zip.Entries.Where(e => e.FullName.StartsWith("xl/worksheets/sheet", StringComparison.OrdinalIgnoreCase) && e.FullName.EndsWith(".xml")).ToList();
                int sheetNum = 1;
                foreach (var sheetEntry in sheetEntries)
                {
                    // Resolve sheet name: use real name if available, fall back to ordinal
                    string sheetName;
                    if (sheetNum - 1 < sheetNames.Count && !string.IsNullOrEmpty(sheetNames[sheetNum - 1]))
                    {
                        sheetName = sheetNames[sheetNum - 1];
                    }
                    else
                    {
                        sheetName = string.Format("Sheet {0}", sheetNum);
                    }

                    sb.AppendLine(string.Format("--- {0} ---", sheetName));
                    using (var stream = sheetEntry.Open())
                    {
                        var doc = XDocument.Load(stream);
                        XNamespace s = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
                        foreach (var row in doc.Descendants(s + "row"))
                        {
                            var cellValues = new List<string>();
                            foreach (var c in row.Descendants(s + "c"))
                            {
                                string cellAddr = (string)c.Attribute("r");
                                if (string.IsNullOrEmpty(cellAddr))
                                    continue;

                                string type = (string)c.Attribute("t");
                                string val = c.Element(s + "v") != null ? c.Element(s + "v").Value : string.Empty;

                                int sstIdx;
                                if (type == "s" && int.TryParse(val, out sstIdx) && sstIdx >= 0 && sstIdx < sharedStrings.Count)
                                {
                                    val = sharedStrings[sstIdx];
                                }
                                else if (type == "inlineStr")
                                {
                                    val = string.Concat(c.Descendants(s + "t").Select(t => t.Value));
                                }

                                if (!string.IsNullOrEmpty(val))
                                {
                                    cellValues.Add(string.Format("{0}={1}", cellAddr, val));
                                }
                            }
                            if (cellValues.Count > 0)
                            {
                                sb.AppendLine(string.Join("\t", cellValues));
                            }
                        }
                    }
                    sheetNum++;
                }
            }
            return sb.ToString();
        }

        private static string ExtractPptx(string path)
        {
            var sb = new StringBuilder();
            using (var zip = ZipFile.OpenRead(path))
            {
                if (ArchiveDecompressedSizeExceedsLimit(zip))
                {
                    return "[Attachment too large when decompressed to process safely.]";
                }

                var slideEntries = zip.Entries
                    .Where(e => e.FullName.StartsWith("ppt/slides/slide", StringComparison.OrdinalIgnoreCase) && e.FullName.EndsWith(".xml"))
                    .OrderBy(e => e.FullName)
                    .ToList();

                int slideIndex = 1;
                foreach (var entry in slideEntries)
                {
                    sb.AppendLine(string.Format("--- Slide {0} ---", slideIndex++));
                    using (var stream = entry.Open())
                    {
                        var doc = XDocument.Load(stream);
                        XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
                        foreach (var p in doc.Descendants(a + "p"))
                        {
                            string line = string.Concat(p.Descendants(a + "t").Select(t => t.Value));
                            if (!string.IsNullOrWhiteSpace(line))
                            {
                                sb.AppendLine(line);
                            }
                        }
                    }
                }
            }
            return sb.ToString();
        }

        private static bool ArchiveDecompressedSizeExceedsLimit(ZipArchive archive)
        {
            long totalDecompressed = 0;
            foreach (var entry in archive.Entries)
            {
                totalDecompressed += entry.Length; // entry.Length is the uncompressed size
                if (totalDecompressed > MaxDecompressedBytes)
                {
                    Logger.Warn(string.Format("AttachmentExtractor: Archive decompressed size exceeds {0}MB limit. Extraction stopped.", MaxDecompressedBytes / (1024 * 1024)));
                    return true;
                }
            }
            return false;
        }

        private static string ExtractPdf(string path)
        {
            var sb = new StringBuilder();
            try
            {
                using (var document = UglyToad.PdfPig.PdfDocument.Open(path))
                {
                    foreach (var page in document.GetPages())
                    {
                        string pageText = page.Text;
                        if (!string.IsNullOrWhiteSpace(pageText))
                        {
                            sb.AppendLine(string.Format("--- Page {0} ---", page.Number));
                            sb.AppendLine(pageText);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Error("AttachmentExtractor: PDF extraction failed", ex);
                throw new InvalidOperationException("Failed to read PDF document text: " + ex.Message, ex);
            }
            return sb.ToString();
        }

        private static string GetMimeType(string extension)
        {
            switch (extension)
            {
                case ".png": return "image/png";
                case ".jpg":
                case ".jpeg": return "image/jpeg";
                case ".webp": return "image/webp";
                case ".gif": return "image/gif";
                case ".bmp": return "image/bmp";
                default: return "application/octet-stream";
            }
        }
    }
}
