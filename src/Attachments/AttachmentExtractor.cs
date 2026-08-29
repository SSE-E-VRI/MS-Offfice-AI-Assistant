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
                        XNamespace wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
                        int paragraphIndex = 1;
                        int tableIndex = 1;
                        var body = doc.Root != null ? doc.Root.Element(w + "body") : null;
                        if (body != null)
                        {
                            foreach (var elem in body.Elements())
                            {
                                if (elem.Name == w + "p")
                                {
                                    string paraText = ExtractParagraphText(elem, w);
                                    // Preserve hyperlink display with marker
                                    // Check for w:hyperlink children to add [link] hint
                                    bool hasHyperlink = elem.Descendants(w + "hyperlink").Any();
                                    if (!string.IsNullOrWhiteSpace(paraText))
                                    {
                                        if (hasHyperlink) paraText = paraText + " [hyperlink]";
                                        sb.AppendLine(string.Format("[¶{0}] {1}", paragraphIndex, paraText));
                                        // Preserve text boxes inside paragraph
                                        string tbText = ExtractTextBoxText(elem, w, wp);
                                        if (!string.IsNullOrWhiteSpace(tbText))
                                            sb.AppendLine(string.Format("[TextBox ¶{0}] {1}", paragraphIndex, tbText));
                                        if (elem.Descendants(w + "drawing").Any() || elem.Descendants(w + "pict").Any())
                                            sb.AppendLine(string.Format("[Image ¶{0}]", paragraphIndex));
                                    }
                                    // Preserve tracked changes markers
                                    string insDel = ExtractInsDelText(elem, w);
                                    if (!string.IsNullOrWhiteSpace(insDel))
                                        sb.AppendLine(string.Format("[Tracked ¶{0}] {1}", paragraphIndex, insDel));
                                    paragraphIndex++;
                                }
                                else if (elem.Name == w + "tbl")
                                {
                                    sb.AppendLine(string.Format("[Table {0}]", tableIndex));
                                    var rows = elem.Elements(w + "tr").ToList();
                                    bool firstRow = true;
                                    foreach (var row in rows)
                                    {
                                        var cells = row.Elements(w + "tc").ToList();
                                        var cellTexts = new List<string>();
                                        foreach (var cell in cells)
                                        {
                                            string cellText = string.Concat(cell.Descendants(w + "t").Select(t => t.Value)).Trim();
                                            // Hyperlink inside cell
                                            if (cell.Descendants(w + "hyperlink").Any()) cellText = cellText + " [link]";
                                            cellTexts.Add(cellText.Replace("|", "/"));
                                        }
                                        if (cellTexts.Count > 0)
                                        {
                                            sb.AppendLine("| " + string.Join(" | ", cellTexts) + " |");
                                            if (firstRow)
                                            {
                                                sb.AppendLine("| " + string.Join(" | ", cellTexts.Select(c => "---")) + " |");
                                                firstRow = false;
                                            }
                                        }
                                    }
                                    tableIndex++;
                                }
                            }
                        }
                        else
                        {
                            // Fallback to old descendant scan
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

                // Headers
                foreach (var headerEntry in zip.Entries.Where(e => e.FullName.StartsWith("word/header", StringComparison.OrdinalIgnoreCase)).OrderBy(e => e.FullName))
                {
                    try
                    {
                        using (var stream = headerEntry.Open())
                        {
                            var doc = XDocument.Load(stream);
                            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
                            foreach (var p in doc.Descendants(w + "p"))
                            {
                                string t = string.Concat(p.Descendants(w + "t").Select(x => x.Value)).Trim();
                                if (!string.IsNullOrWhiteSpace(t))
                                    sb.AppendLine(string.Format("[Header {0}] {1}", Path.GetFileName(headerEntry.FullName), t));
                            }
                        }
                    }
                    catch { }
                }
                // Footers
                foreach (var footerEntry in zip.Entries.Where(e => e.FullName.StartsWith("word/footer", StringComparison.OrdinalIgnoreCase)).OrderBy(e => e.FullName))
                {
                    try
                    {
                        using (var stream = footerEntry.Open())
                        {
                            var doc = XDocument.Load(stream);
                            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
                            foreach (var p in doc.Descendants(w + "p"))
                            {
                                string t = string.Concat(p.Descendants(w + "t").Select(x => x.Value)).Trim();
                                if (!string.IsNullOrWhiteSpace(t))
                                    sb.AppendLine(string.Format("[Footer {0}] {1}", Path.GetFileName(footerEntry.FullName), t));
                            }
                        }
                    }
                    catch { }
                }
                // Footnotes / Endnotes
                foreach (var noteEntry in zip.Entries.Where(e => e.FullName == "word/footnotes.xml" || e.FullName == "word/endnotes.xml"))
                {
                    try
                    {
                        using (var stream = noteEntry.Open())
                        {
                            var doc = XDocument.Load(stream);
                            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
                            int idx = 1;
                            foreach (var p in doc.Descendants(w + "p"))
                            {
                                string t = string.Concat(p.Descendants(w + "t").Select(x => x.Value)).Trim();
                                if (!string.IsNullOrWhiteSpace(t))
                                    sb.AppendLine(string.Format("[{0} {1}] {2}", noteEntry.FullName.Contains("footnote") ? "Footnote" : "Endnote", idx, t));
                                idx++;
                            }
                        }
                    }
                    catch { }
                }
                // Comments
                var commentsEntry = zip.GetEntry("word/comments.xml");
                if (commentsEntry != null)
                {
                    try
                    {
                        using (var stream = commentsEntry.Open())
                        {
                            var doc = XDocument.Load(stream);
                            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
                            foreach (var c in doc.Descendants(w + "comment"))
                            {
                                string author = (string)c.Attribute(w + "author") ?? (string)c.Attribute("author") ?? "Unknown";
                                string txt = string.Concat(c.Descendants(w + "t").Select(x => x.Value)).Trim();
                                if (!string.IsNullOrWhiteSpace(txt))
                                    sb.AppendLine(string.Format("[Comment by {0}] {1}", author, txt));
                            }
                        }
                    }
                    catch { }
                }
            }
            return sb.ToString();
        }

        private static string ExtractParagraphText(System.Xml.Linq.XElement p, System.Xml.Linq.XNamespace w)
        {
            try
            {
                // Preserve style name if heading
                string style = string.Empty;
                var pPr = p.Element(w + "pPr");
                if (pPr != null)
                {
                    var pStyle = pPr.Element(w + "pStyle");
                    if (pStyle != null) style = (string)pStyle.Attribute(w + "val") ?? string.Empty;
                }
                string text = string.Concat(p.Descendants(w + "t").Select(t => t.Value));
                if (!string.IsNullOrWhiteSpace(style)) text = string.Format("[{0}] {1}", style, text);
                return text.Trim();
            }
            catch { return string.Concat(p.Descendants(w + "t").Select(t => t.Value)).Trim(); }
        }

        private static string ExtractTextBoxText(System.Xml.Linq.XElement p, System.Xml.Linq.XNamespace w, System.Xml.Linq.XNamespace wp)
        {
            try
            {
                var txbx = p.Descendants(w + "txbxContent").FirstOrDefault();
                if (txbx != null)
                {
                    string t = string.Concat(txbx.Descendants(w + "t").Select(x => x.Value)).Trim();
                    return t;
                }
            }
            catch { }
            return string.Empty;
        }

        private static string ExtractInsDelText(System.Xml.Linq.XElement p, System.Xml.Linq.XNamespace w)
        {
            try
            {
                var ins = p.Descendants(w + "ins").Select(e => string.Concat(e.Descendants(w + "t").Select(t => t.Value)).Trim()).Where(s => !string.IsNullOrWhiteSpace(s)).ToList();
                var del = p.Descendants(w + "del").Select(e => string.Concat(e.Descendants(w + "t").Select(t => t.Value)).Trim()).Where(s => !string.IsNullOrWhiteSpace(s)).ToList();
                var parts = new List<string>();
                if (ins.Count > 0) parts.Add("Inserted: " + string.Join("; ", ins));
                if (del.Count > 0) parts.Add("Deleted: " + string.Join("; ", del));
                return string.Join(" | ", parts);
            }
            catch { return string.Empty; }
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
                                    // Sheet-qualified so NavigateToCitation resolves the correct sheet
                                    // (adversarial-review fix — a bare "A1=value" tag under a multi-sheet
                                    // workbook resolved to whatever sheet happened to be active at click
                                    // time, not the sheet the value actually came from). Matches the
                                    // existing "SheetName!Address" citation pattern (MarkdownHelper Pattern 3,
                                    // EvidenceLevel, ChatSidebar.NavigateToCitation's excelSheetMatch) — no
                                    // matcher/navigation changes needed, only emission.
                                    cellValues.Add(string.Format("{0}!{1}={2}", sheetName, cellAddr, val));
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

                // Resolve correct slide order from presentation.xml (avoids lexical slide10 < slide2 issue)
                List<string> orderedSlideFiles = GetOrderedSlideFiles(zip);
                if (orderedSlideFiles == null || orderedSlideFiles.Count == 0)
                {
                    orderedSlideFiles = zip.Entries
                        .Where(e => e.FullName.StartsWith("ppt/slides/slide", StringComparison.OrdinalIgnoreCase) && e.FullName.EndsWith(".xml"))
                        .OrderBy(e => ExtractSlideNumber(e.FullName))
                        .ThenBy(e => e.FullName, StringComparer.OrdinalIgnoreCase)
                        .Select(e => e.FullName)
                        .ToList();
                }

                Dictionary<string, string> slideToSection = BuildSectionMap(zip);
                Dictionary<string, string> slideToNotesFile = BuildSlideNotesMap(zip);
                Dictionary<string, string> slideToLayout = BuildSlideLayoutMap(zip);

                int slideIndex = 1;
                foreach (var slideFile in orderedSlideFiles)
                {
                    var entry = zip.GetEntry(slideFile);
                    if (entry == null) continue;

                    sb.AppendLine(string.Format("--- Slide {0} ---", slideIndex));

                    string sectionName = null;
                    if (slideToSection != null && slideToSection.TryGetValue(slideFile, out sectionName) && !string.IsNullOrWhiteSpace(sectionName))
                        sb.AppendLine(string.Format("[Section: {0}]", sectionName));

                    string layoutName = null;
                    if (slideToLayout != null && slideToLayout.TryGetValue(slideFile, out layoutName) && !string.IsNullOrWhiteSpace(layoutName))
                        sb.AppendLine(string.Format("[Layout: {0}]", layoutName));

                    using (var stream = entry.Open())
                    {
                        var doc = XDocument.Load(stream);
                        XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
                        XNamespace p = "http://schemas.openxmlformats.org/presentationml/2006/main";
                        XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
                        XNamespace p14 = "http://schemas.microsoft.com/office/powerpoint/2010/main";

                        string title = ExtractPptxSlideTitle(doc, p, a);
                        if (!string.IsNullOrWhiteSpace(title))
                            sb.AppendLine(string.Format("[Title] {0}", title));

                        // Slide metadata: name, hidden flag
                        try
                        {
                            var cSld = doc.Descendants(p + "cSld").FirstOrDefault();
                            if (cSld != null)
                            {
                                string sldName = (string)cSld.Attribute("name");
                                string titleTrimmed = title != null ? title.Trim() : null;
                                if (!string.IsNullOrWhiteSpace(sldName) && !string.Equals(sldName.Trim(), titleTrimmed, StringComparison.OrdinalIgnoreCase))
                                    sb.AppendLine(string.Format("[Slide name: {0}]", sldName.Trim()));
                            }
                            var show = doc.Descendants(p + "show").FirstOrDefault();
                            if (show != null)
                            {
                                string val = (string)show.Attribute("val");
                                if (!string.IsNullOrWhiteSpace(val) && val.Trim() == "0")
                                    sb.AppendLine("[Hidden slide]");
                            }
                            // also check transition hidden? Not needed for extraction; hidden via show attribute
                            // Check for timing? not needed
                        }
                        catch { }

                        foreach (var para in doc.Descendants(a + "p"))
                        {
                            string line = string.Concat(para.Descendants(a + "t").Select(t => t.Value));
                            if (!string.IsNullOrWhiteSpace(line))
                            {
                                // Avoid duplicating title if already emitted
                                if (!string.IsNullOrWhiteSpace(title) && line.Trim().Equals(title.Trim(), StringComparison.OrdinalIgnoreCase))
                                    continue;
                                sb.AppendLine(line.Trim());
                            }
                        }

                        // Detailed image extraction with alt text / name
                        try
                        {
                            var pics = doc.Descendants(p + "pic").ToList();
                            // Also cNvPr with descr is inside p:cNvPr
                            int picIdx = 1;
                            foreach (var pic in pics)
                            {
                                string picName = string.Empty;
                                string picDescr = string.Empty;
                                var cNvPr = pic.Descendants(p + "cNvPr").FirstOrDefault() ?? pic.Descendants().FirstOrDefault(e => e.Name.LocalName == "cNvPr");
                                if (cNvPr != null)
                                {
                                    picName = (string)cNvPr.Attribute("name") ?? string.Empty;
                                    picDescr = (string)cNvPr.Attribute("descr") ?? string.Empty;
                                }
                                // blip embed target for file name hint
                                string embed = string.Empty;
                                var blip = pic.Descendants().FirstOrDefault(e => e.Name.LocalName == "blip");
                                if (blip != null)
                                {
                                    XNamespace r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
                                    embed = (string)blip.Attribute(r + "embed") ?? string.Empty;
                                }
                                if (!string.IsNullOrWhiteSpace(picDescr))
                                    sb.AppendLine(string.Format("[Image {0}: {1} - {2}]", picIdx, picName.Trim(), picDescr.Trim()));
                                else if (!string.IsNullOrWhiteSpace(picName))
                                    sb.AppendLine(string.Format("[Image {0}: {1}]", picIdx, picName.Trim()));
                                else
                                    sb.AppendLine(string.Format("[Image {0}]", picIdx));
                                picIdx++;
                            }
                            if (pics.Count == 0)
                            {
                                int blipCount = doc.Descendants().Count(e => e.Name.LocalName == "blip");
                                if (blipCount > 0)
                                    sb.AppendLine(string.Format("[Images: {0} on this slide]", blipCount));
                            }
                        }
                        catch
                        {
                            int imageCount = doc.Descendants().Count(e => e.Name.LocalName == "blip");
                            if (imageCount > 0)
                                sb.AppendLine(string.Format("[Images: {0} on this slide]", imageCount));
                        }

                        // Detailed table extraction
                        try
                        {
                            var tables = doc.Descendants(a + "tbl").ToList();
                            if (tables.Count == 0) tables = doc.Descendants().Where(e => e.Name.LocalName == "tbl").ToList();
                            int tblIdx = 1;
                            foreach (var tbl in tables)
                            {
                                sb.AppendLine(string.Format("[Table {0} on slide {1}]", tblIdx, slideIndex));
                                var rows = tbl.Descendants(a + "tr").ToList();
                                if (rows.Count == 0) rows = tbl.Descendants().Where(e => e.Name.LocalName == "tr").ToList();
                                bool firstRow = true;
                                foreach (var row in rows)
                                {
                                    var cells = row.Descendants(a + "tc").ToList();
                                    if (cells.Count == 0) cells = row.Descendants().Where(e => e.Name.LocalName == "tc").ToList();
                                    var cellTexts = new List<string>();
                                    foreach (var cell in cells)
                                    {
                                        string cellText = string.Concat(cell.Descendants(a + "t").Select(t => t.Value)).Trim();
                                        if (string.IsNullOrWhiteSpace(cellText))
                                            cellText = string.Concat(cell.Descendants().Where(e => e.Name.LocalName == "t").Select(t => t.Value)).Trim();
                                        cellTexts.Add(string.IsNullOrWhiteSpace(cellText) ? "" : cellText.Replace("|", "/"));
                                    }
                                    if (cellTexts.Count > 0)
                                    {
                                        sb.AppendLine("| " + string.Join(" | ", cellTexts) + " |");
                                        if (firstRow)
                                        {
                                            sb.AppendLine("| " + string.Join(" | ", cellTexts.Select(x => "---")) + " |");
                                            firstRow = false;
                                        }
                                    }
                                }
                                tblIdx++;
                            }
                        }
                        catch { }

                        // Detailed chart extraction (follow relationship to chart file)
                        try
                        {
                            var chartRefs = doc.Descendants(c + "chart").ToList();
                            if (chartRefs.Count == 0) chartRefs = doc.Descendants().Where(e => e.Name.LocalName == "chart").ToList();
                            int chartIdx = 1;
                            foreach (var chartRef in chartRefs)
                            {
                                XNamespace r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
                                string rId = (string)chartRef.Attribute(r + "id");
                                string chartTitle = string.Empty;
                                string chartTypeHint = string.Empty;
                                // Try to resolve chart file via slide rels
                                if (!string.IsNullOrWhiteSpace(rId))
                                {
                                    try
                                    {
                                        string relsPath = "ppt/slides/_rels/" + System.IO.Path.GetFileName(slideFile) + ".rels";
                                        var relsEntry = zip.GetEntry(relsPath);
                                        if (relsEntry != null)
                                        {
                                            using (var rs = relsEntry.Open())
                                            {
                                                var rdoc = XDocument.Load(rs);
                                                XNamespace relNs = "http://schemas.openxmlformats.org/package/2006/relationships";
                                                foreach (var rel in rdoc.Descendants(relNs + "Relationship"))
                                                {
                                                    string id = (string)rel.Attribute("Id");
                                                    string target = (string)rel.Attribute("Target");
                                                    if (string.Equals(id, rId, StringComparison.OrdinalIgnoreCase) && !string.IsNullOrWhiteSpace(target))
                                                    {
                                                        string chartFile = target.StartsWith("..") ? "ppt" + target.Substring(2).Replace("../", "/") : ("ppt/slides/" + target.TrimStart('/'));
                                                        // Normalize: charts/chart1.xml is relative to ppt/slides -> ppt/charts/chart1.xml
                                                        if (target.Contains("charts/"))
                                                        {
                                                            chartFile = "ppt/charts/" + System.IO.Path.GetFileName(target);
                                                        }
                                                        var chartEntry = zip.GetEntry(chartFile);
                                                        if (chartEntry == null) chartEntry = zip.GetEntry("ppt/charts/" + System.IO.Path.GetFileName(target));
                                                        if (chartEntry != null)
                                                        {
                                                            using (var cs = chartEntry.Open())
                                                            {
                                                                var cdoc = XDocument.Load(cs);
                                                                XNamespace ca = "http://schemas.openxmlformats.org/drawingml/2006/chart";
                                                                XNamespace ca2 = "http://schemas.openxmlformats.org/drawingml/2006/main";
                                                                // Title
                                                                var titleEl = cdoc.Descendants(ca + "title").FirstOrDefault();
                                                                if (titleEl != null)
                                                                {
                                                                    chartTitle = string.Concat(titleEl.Descendants(ca2 + "t").Select(t => t.Value)).Trim();
                                                                    if (string.IsNullOrWhiteSpace(chartTitle))
                                                                        chartTitle = string.Concat(titleEl.Descendants().Where(e => e.Name.LocalName == "t").Select(t => t.Value)).Trim();
                                                                }
                                                                // Chart type detection
                                                                if (cdoc.Descendants(ca + "barChart").Any()) chartTypeHint = "bar";
                                                                else if (cdoc.Descendants(ca + "lineChart").Any()) chartTypeHint = "line";
                                                                else if (cdoc.Descendants(ca + "pieChart").Any()) chartTypeHint = "pie";
                                                                else if (cdoc.Descendants(ca + "areaChart").Any()) chartTypeHint = "area";
                                                                else if (cdoc.Descendants(ca + "scatterChart").Any()) chartTypeHint = "scatter";
                                                                else if (cdoc.Descendants(ca + "doughnutChart").Any()) chartTypeHint = "doughnut";
                                                                else chartTypeHint = "column";
                                                                // Categories / values snippet
                                                                var cats = cdoc.Descendants(ca + "cat").Take(5).Select(e => string.Concat(e.Descendants(ca2 + "v").Select(v => v.Value)).Trim()).Where(s => !string.IsNullOrWhiteSpace(s)).ToList();
                                                                if (cats.Count == 0) cats = cdoc.Descendants().Where(e => e.Name.LocalName == "cat").Take(5).Select(e => string.Concat(e.Descendants().Where(x => x.Name.LocalName == "v").Select(v => v.Value)).Trim()).Where(s => !string.IsNullOrWhiteSpace(s)).ToList();
                                                                if (cats.Count > 0)
                                                                    chartTitle = (string.IsNullOrWhiteSpace(chartTitle) ? "" : chartTitle + " ") + string.Format("[Categories: {0}]", string.Join(", ", cats));
                                                            }
                                                        }
                                                        break;
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    catch { }
                                }
                                if (!string.IsNullOrWhiteSpace(chartTitle))
                                    sb.AppendLine(string.Format("[Chart {0}: {1} ({2})]", chartIdx, chartTitle, string.IsNullOrWhiteSpace(chartTypeHint) ? "chart" : chartTypeHint));
                                else if (!string.IsNullOrWhiteSpace(chartTypeHint))
                                    sb.AppendLine(string.Format("[Chart {0}: {1} chart]", chartIdx, chartTypeHint));
                                else
                                    sb.AppendLine(string.Format("[Chart {0} on slide]", chartIdx));
                                chartIdx++;
                            }
                            if (chartRefs.Count == 0)
                            {
                                // Fallback legacy detection
                                bool hasChart = doc.Descendants(c + "chart").Any() || doc.Descendants().Any(e => e.Name.LocalName == "chart");
                                if (hasChart) sb.AppendLine("[Chart present on slide]");
                            }
                        }
                        catch { }

                        bool hasSmartArt = doc.Descendants().Any(e => e.Name.LocalName == "dgm");
                        if (hasSmartArt)
                            sb.AppendLine("[Diagram/SmartArt present on slide]");

                        // Theme/layout extra: try to read master theme color hint
                        try
                        {
                            // Look for theme reference in slide master rels? Simplified: if slide contains theme element, note it
                            var themeEl = doc.Descendants().FirstOrDefault(e => e.Name.LocalName == "clrScheme");
                            if (themeEl != null)
                                sb.AppendLine("[Theme colors referenced]");
                        }
                        catch { }
                    }

                    // Speaker notes via slides/_rels/slideN.xml.rels -> notesSlide
                    string notesFile = null;
                    if (slideToNotesFile != null && slideToNotesFile.TryGetValue(slideFile, out notesFile) && !string.IsNullOrWhiteSpace(notesFile))
                    {
                        var notesEntry = zip.GetEntry(notesFile);
                        if (notesEntry != null)
                        {
                            try
                            {
                                using (var ns = notesEntry.Open())
                                {
                                    var ndoc = XDocument.Load(ns);
                                    XNamespace a2 = "http://schemas.openxmlformats.org/drawingml/2006/main";
                                    var noteLines = new List<string>();
                                    foreach (var para in ndoc.Descendants(a2 + "p"))
                                    {
                                        string line = string.Concat(para.Descendants(a2 + "t").Select(t => t.Value)).Trim();
                                        if (!string.IsNullOrWhiteSpace(line))
                                            noteLines.Add(line);
                                    }
                                    // Filter out placeholder "Click to edit" leftovers
                                    noteLines = noteLines.Where(l => !l.Equals("Click to edit Master text styles", StringComparison.OrdinalIgnoreCase) && !l.StartsWith("Click to", StringComparison.OrdinalIgnoreCase)).ToList();
                                    if (noteLines.Count > 0)
                                    {
                                        sb.AppendLine("[Speaker Notes]");
                                        foreach (var nl in noteLines) sb.AppendLine(nl);
                                    }
                                }
                            }
                            catch { }
                        }
                    }

                    sb.AppendLine();
                    slideIndex++;
                }

                // Presentation-level fallback: if orderedSlideFiles empty but we had entries, ensure we captured something
                if (orderedSlideFiles.Count == 0)
                    sb.AppendLine("[No slides found in presentation]");

                // Also include top-level presentation-wide sections list as provenance
                try
                {
                    var presEntry = zip.GetEntry("ppt/presentation.xml");
                    if (presEntry != null)
                    {
                        using (var s = presEntry.Open())
                        {
                            var pdoc = XDocument.Load(s);
                            XNamespace p14 = "http://schemas.microsoft.com/office/powerpoint/2010/main";
                            var sections = pdoc.Descendants(p14 + "section").ToList();
                            if (sections.Count > 0)
                            {
                                sb.AppendLine("--- Sections ---");
                                foreach (var sec in sections)
                                {
                                    string sname = (string)sec.Attribute("name") ?? "(unnamed)";
                                    sb.AppendLine(string.Format("[Section] {0}", sname));
                                }
                            }
                        }
                    }
                }
                catch { }
            }
            return sb.ToString();
        }

        private static int ExtractSlideNumber(string fullName)
        {
            try
            {
                string file = Path.GetFileNameWithoutExtension(fullName);
                string numPart = new string(file.Where(char.IsDigit).ToArray());
                int n;
                if (int.TryParse(numPart, out n)) return n;
            }
            catch { }
            return int.MaxValue;
        }

        private static List<string> GetOrderedSlideFiles(ZipArchive zip)
        {
            try
            {
                var presEntry = zip.GetEntry("ppt/presentation.xml");
                if (presEntry == null) return null;
                XDocument presDoc;
                using (var s = presEntry.Open()) presDoc = XDocument.Load(s);
                XNamespace p = "http://schemas.openxmlformats.org/presentationml/2006/main";
                XNamespace r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

                // Build rId -> Target map from ppt/_rels/presentation.xml.rels
                var relsEntry = zip.GetEntry("ppt/_rels/presentation.xml.rels");
                var rIdToTarget = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                if (relsEntry != null)
                {
                    using (var rs = relsEntry.Open())
                    {
                        var rdoc = XDocument.Load(rs);
                        XNamespace relNs = "http://schemas.openxmlformats.org/package/2006/relationships";
                        foreach (var rel in rdoc.Descendants(relNs + "Relationship"))
                        {
                            string id = (string)rel.Attribute("Id");
                            string target = (string)rel.Attribute("Target");
                            if (!string.IsNullOrWhiteSpace(id) && !string.IsNullOrWhiteSpace(target))
                            {
                                // Target is like "slides/slide1.xml"
                                string normalized = target.StartsWith("ppt/") ? target : "ppt/" + target.TrimStart('/');
                                rIdToTarget[id] = normalized;
                            }
                        }
                    }
                }

                var sldIds = presDoc.Descendants(p + "sldId").ToList();
                if (sldIds.Count == 0) return null;
                var ordered = new List<string>();
                foreach (var sldId in sldIds)
                {
                    string rId = (string)sldId.Attribute(r + "id");
                    string target = null;
                    if (!string.IsNullOrWhiteSpace(rId) && rIdToTarget.TryGetValue(rId, out target))
                    {
                        ordered.Add(target);
                    }
                }
                // Validate that files exist; if mapping failed, fall back
                if (ordered.Count > 0 && ordered.All(f => zip.GetEntry(f) != null))
                    return ordered;
            }
            catch { }
            return null;
        }

        private static Dictionary<string, string> BuildSectionMap(ZipArchive zip)
        {
            var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            try
            {
                var presEntry = zip.GetEntry("ppt/presentation.xml");
                if (presEntry == null) return map;
                XDocument presDoc;
                using (var s = presEntry.Open()) presDoc = XDocument.Load(s);
                XNamespace p14 = "http://schemas.microsoft.com/office/powerpoint/2010/main";

                // Need slide file ordering to map section sldId -> slide file
                // We have sldId elements with @id attribute referencing slide index; p14:section has p14:sldIdLst/p14:sldId @id
                // That @id matches presentation's p:sldId @id. So build map from sldId id -> slideFile.
                Dictionary<string, string> idToSlideFile = new Dictionary<string, string>();
                try
                {
                    XNamespace p = "http://schemas.openxmlformats.org/presentationml/2006/main";
                    XNamespace r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
                    var presDoc2 = presDoc;
                    var relsEntry = zip.GetEntry("ppt/_rels/presentation.xml.rels");
                    var rIdToTarget = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                    if (relsEntry != null)
                    {
                        using (var rs = relsEntry.Open())
                        {
                            var rdoc = XDocument.Load(rs);
                            XNamespace relNs = "http://schemas.openxmlformats.org/package/2006/relationships";
                            foreach (var rel in rdoc.Descendants(relNs + "Relationship"))
                            {
                                string rid = (string)rel.Attribute("Id");
                                string target = (string)rel.Attribute("Target");
                                if (!string.IsNullOrWhiteSpace(rid) && !string.IsNullOrWhiteSpace(target))
                                {
                                    string normalized = target.StartsWith("ppt/") ? target : "ppt/" + target.TrimStart('/');
                                    rIdToTarget[rid] = normalized;
                                }
                            }
                        }
                    }
                    var sldIds = presDoc2.Descendants(p + "sldId").ToList();
                    foreach (var sldId in sldIds)
                    {
                        string idVal = (string)sldId.Attribute("id");
                        string rId = (string)sldId.Attribute(r + "id");
                        string target = null;
                        if (!string.IsNullOrWhiteSpace(rId) && rIdToTarget.TryGetValue(rId, out target) && !string.IsNullOrWhiteSpace(idVal))
                        {
                            idToSlideFile[idVal] = target;
                        }
                    }
                }
                catch { }

                var sections = presDoc.Descendants(p14 + "section").ToList();
                foreach (var sec in sections)
                {
                    string secName = (string)sec.Attribute("name") ?? "(unnamed)";
                    var sldIdLst = sec.Element(p14 + "sldIdLst");
                    if (sldIdLst == null) continue;
                    foreach (var sid in sldIdLst.Elements(p14 + "sldId"))
                    {
                        string id = (string)sid.Attribute("id");
                        string slideFile = null;
                        if (!string.IsNullOrWhiteSpace(id) && idToSlideFile.TryGetValue(id, out slideFile))
                        {
                            if (!map.ContainsKey(slideFile))
                                map[slideFile] = secName;
                        }
                        if (!string.IsNullOrWhiteSpace(id) && string.IsNullOrWhiteSpace(slideFile))
                        {
                            // Fallback: try attribute r:id?
                            XNamespace r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
                            string rId = (string)sid.Attribute(r + "id");
                            if (!string.IsNullOrWhiteSpace(rId))
                            {
                                // Already handled via idToSlideFile; if not found, try direct slide number inference
                            }
                        }
                    }
                }
            }
            catch { }
            return map;
        }

        private static Dictionary<string, string> BuildSlideNotesMap(ZipArchive zip)
        {
            var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            try
            {
                var slideEntries = zip.Entries.Where(e => e.FullName.StartsWith("ppt/slides/slide", StringComparison.OrdinalIgnoreCase) && e.FullName.EndsWith(".xml")).ToList();
                foreach (var slideEntry in slideEntries)
                {
                    string relsPath = "ppt/slides/_rels/" + Path.GetFileName(slideEntry.FullName) + ".rels";
                    var relsEntry = zip.GetEntry(relsPath);
                    if (relsEntry == null) continue;
                    try
                    {
                        using (var rs = relsEntry.Open())
                        {
                            var rdoc = XDocument.Load(rs);
                            XNamespace relNs = "http://schemas.openxmlformats.org/package/2006/relationships";
                            foreach (var rel in rdoc.Descendants(relNs + "Relationship"))
                            {
                                string type = (string)rel.Attribute("Type");
                                if (!string.IsNullOrWhiteSpace(type) && type.EndsWith("/notesSlide", StringComparison.OrdinalIgnoreCase))
                                {
                                    string target = (string)rel.Attribute("Target");
                                    if (!string.IsNullOrWhiteSpace(target))
                                    {
                                        // Target is like "../notesSlides/notesSlide1.xml"
                                        string dir = "ppt/slides/";
                                        string combined = target.Replace("../", "ppt/");
                                        if (combined.StartsWith("ppt/notesSlides/"))
                                        {
                                            // Already normalized
                                            if (zip.GetEntry(combined) != null)
                                                map[slideEntry.FullName] = combined;
                                        }
                                        else
                                        {
                                            // Resolve relative to ppt/slides/
                                            string normalized = "ppt/notesSlides/" + Path.GetFileName(target);
                                            if (zip.GetEntry(normalized) != null)
                                                map[slideEntry.FullName] = normalized;
                                        }
                                        break;
                                    }
                                }
                            }
                        }
                    }
                    catch { }
                }
            }
            catch { }
            return map;
        }

        private static Dictionary<string, string> BuildSlideLayoutMap(ZipArchive zip)
        {
            var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            try
            {
                var slideEntries = zip.Entries.Where(e => e.FullName.StartsWith("ppt/slides/slide", StringComparison.OrdinalIgnoreCase) && e.FullName.EndsWith(".xml")).ToList();
                foreach (var slideEntry in slideEntries)
                {
                    try
                    {
                        using (var s = slideEntry.Open())
                        {
                            var doc = XDocument.Load(s);
                            XNamespace r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
                            // In slide xml, layout is referenced via rId in p:sld -> p:cSld ->? Actually <p:sldLayoutId> not per slide; slide's relationship to layout is in its rels.
                            // So check slide rels for slideLayout
                            string relsPath = "ppt/slides/_rels/" + Path.GetFileName(slideEntry.FullName) + ".rels";
                            var relsEntry = zip.GetEntry(relsPath);
                            if (relsEntry != null)
                            {
                                using (var rs = relsEntry.Open())
                                {
                                    var rdoc = XDocument.Load(rs);
                                    XNamespace relNs = "http://schemas.openxmlformats.org/package/2006/relationships";
                                    foreach (var rel in rdoc.Descendants(relNs + "Relationship"))
                                    {
                                        string type = (string)rel.Attribute("Type");
                                        if (!string.IsNullOrWhiteSpace(type) && type.Contains("slideLayout"))
                                        {
                                            string target = (string)rel.Attribute("Target");
                                            if (!string.IsNullOrWhiteSpace(target))
                                            {
                                                string layoutFile = target.Contains("/") ? Path.GetFileName(target) : target;
                                                map[slideEntry.FullName] = layoutFile.Replace(".xml", "");
                                            }
                                            break;
                                        }
                                    }
                                }
                            }
                        }
                    }
                    catch { }
                }
            }
            catch { }
            return map;
        }

        private static string ExtractPptxSlideTitle(XDocument slideDoc, XNamespace pNs, XNamespace aNs)
        {
            try
            {
                // Prefer placeholder title shapes
                var shapes = slideDoc.Descendants(pNs + "sp").ToList();
                foreach (var sp in shapes)
                {
                    var nvPr = sp.Element(pNs + "nvSpPr");
                    if (nvPr == null) continue;
                    var cNvPr = nvPr.Element(pNs + "cNvPr");
                    var nvPr2 = nvPr.Element(pNs + "nvPr");
                    if (nvPr2 != null)
                    {
                        var ph = nvPr2.Element(pNs + "ph");
                        if (ph != null)
                        {
                            string phType = (string)ph.Attribute("type");
                            if (string.Equals(phType, "title", StringComparison.OrdinalIgnoreCase) || string.Equals(phType, "ctrTitle", StringComparison.OrdinalIgnoreCase))
                            {
                                var txBody = sp.Element(pNs + "txBody");
                                if (txBody != null)
                                {
                                    string t = string.Concat(txBody.Descendants(aNs + "t").Select(x => x.Value)).Trim();
                                    if (!string.IsNullOrWhiteSpace(t)) return t;
                                }
                            }
                        }
                    }
                }
                // Fallback: first paragraph that looks like title (largest or first)
                var firstPara = slideDoc.Descendants(aNs + "t").FirstOrDefault();
                if (firstPara != null)
                {
                    // Use shape title heuristic: if doc has cSld name?
                    return null; // avoid misattributing body as title; let caller dedupe
                }
            }
            catch { }
            return null;
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
