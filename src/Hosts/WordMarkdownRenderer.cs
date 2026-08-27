using System;
using System.Collections.Generic;
using Markdig;
using Markdig.Extensions.Tables;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;
using MSOfficeAIAssistant.Core;
using Word = NetOffice.WordApi;

namespace MSOfficeAIAssistant.Hosts
{
    /// <summary>
    /// Converts Markdown text into native Word document formatting using Markdig's AST.
    /// Preserves and aligns with the active document's font family, base font size, and paragraph styles.
    /// </summary>
    public static class WordMarkdownRenderer
    {
        public static void Render(Word.Application app, string markdown)
        {
            if (app == null || string.IsNullOrWhiteSpace(markdown)) return;

            bool oldScreenUpdating = true;
            try
            {
                oldScreenUpdating = app.ScreenUpdating;
                app.ScreenUpdating = false;
            }
            catch { }

            try
            {
                string docFont;
                float docSize;
                GetDocumentFont(app, out docFont, out docSize);

                // 2. Parse markdown into AST using Markdig with advanced extensions (tables, etc.)
                var pipeline = new MarkdownPipelineBuilder()
                    .UseAdvancedExtensions()
                    .Build();
                var document = Markdown.Parse(markdown, pipeline);

                // 3. Prepare insertion range
                var range = app.Selection != null ? app.Selection.Range : null;
                if (range == null) return;

                // If user has a non-collapsed selection, clear it first
                if (app.Selection.Type != Word.Enums.WdSelectionType.wdSelectionIP)
                {
                    range.Text = string.Empty;
                }
                range.Collapse(Word.Enums.WdCollapseDirection.wdCollapseEnd);

                // 4. Walk AST blocks and emit formatted Word objects
                foreach (var block in document)
                {
                    RenderBlock(app, range, block, docFont, docSize);
                }

                // Move selection to end of inserted content
                try
                {
                    app.Selection.SetRange(range.End, range.End);
                }
                catch { }
            }
            catch (Exception ex)
            {
                Logger.Error("WordMarkdownRenderer AST insertion failed, falling back to plain text", ex);
                try
                {
                    if (app.Selection != null)
                        app.Selection.TypeText(markdown);
                }
                catch { }
            }
            finally
            {
                try
                {
                    app.ScreenUpdating = oldScreenUpdating;
                }
                catch { }
            }
        }

        private static void GetDocumentFont(Word.Application app, out string docFont, out float docSize)
        {
            docFont = "Calibri";
            docSize = 11.0f;
            try
            {
                if (app != null && app.Selection != null && app.Selection.Font != null)
                {
                    string selFont = app.Selection.Font.Name;
                    if (!string.IsNullOrWhiteSpace(selFont)) docFont = selFont;
                    float selSize = app.Selection.Font.Size;
                    if (selSize >= 4.0f && selSize <= 100.0f) docSize = selSize;
                }
                else if (app != null && app.ActiveDocument != null && app.ActiveDocument.Styles != null)
                {
                    var normalStyle = app.ActiveDocument.Styles[Word.Enums.WdBuiltinStyle.wdStyleNormal];
                    if (normalStyle != null && normalStyle.Font != null)
                    {
                        if (!string.IsNullOrWhiteSpace(normalStyle.Font.Name)) docFont = normalStyle.Font.Name;
                        if (normalStyle.Font.Size >= 4.0f && normalStyle.Font.Size <= 100.0f)
                            docSize = normalStyle.Font.Size;
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("Could not detect active font properties: {0}", ex.Message));
            }
        }

        private static void RenderBlock(Word.Application app, Word.Range range, Block block, string docFont, float docSize)
        {
            var heading = block as HeadingBlock;
            if (heading != null)
            {
                RenderHeading(app, range, heading, docFont, docSize);
                return;
            }

            var para = block as ParagraphBlock;
            if (para != null)
            {
                RenderParagraph(app, range, para, docFont, docSize);
                return;
            }

            var list = block as ListBlock;
            if (list != null)
            {
                RenderList(app, range, list, docFont, docSize);
                return;
            }

            var table = block as Table;
            if (table != null)
            {
                // A table Word refuses to create must not take the data down with it --
                // fall back to tab-separated rows rather than silently dropping the content.
                if (!RenderTable(app, range, table, docFont, docSize))
                    RenderTableAsText(range, table, docFont, docSize);
                return;
            }

            var fencedCode = block as FencedCodeBlock;
            if (fencedCode != null)
            {
                RenderCodeBlock(app, range, fencedCode, docSize);
                return;
            }

            var code = block as CodeBlock;
            if (code != null)
            {
                RenderCodeBlock(app, range, code, docSize);
                return;
            }

            var quote = block as QuoteBlock;
            if (quote != null)
            {
                RenderQuote(app, range, quote, docFont, docSize);
                return;
            }

            var breakBlock = block as ThematicBreakBlock;
            if (breakBlock != null)
            {
                RenderThematicBreak(app, range);
                return;
            }

            var container = block as ContainerBlock;
            if (container != null)
            {
                foreach (var child in container)
                {
                    RenderBlock(app, range, child, docFont, docSize);
                }
            }
        }

        private static void RenderHeading(Word.Application app, Word.Range range, HeadingBlock heading, string docFont, float docSize)
        {
            float scale = 1.0f;
            switch (heading.Level)
            {
                case 1: scale = 1.50f; break;
                case 2: scale = 1.30f; break;
                case 3: scale = 1.15f; break;
                case 4: scale = 1.05f; break;
                default: scale = 1.00f; break;
            }
            float headingSize = docSize * scale;

            try
            {
                range.ParagraphFormat.SpaceBefore = (float)Math.Max(3.0, 10.0 - heading.Level * 1.5);
                range.ParagraphFormat.SpaceAfter = 3.0f;
                range.ParagraphFormat.LeftIndent = 0.0f;
                range.ParagraphFormat.KeepWithNext = -1;
            }
            catch { }

            if (heading.Inline != null)
            {
                RenderInlines(range, heading.Inline, docFont, headingSize, true, false, false, docFont, docSize);
            }

            range.InsertParagraphAfter();
            range.Collapse(Word.Enums.WdCollapseDirection.wdCollapseEnd);
        }

        private static void RenderParagraph(Word.Application app, Word.Range range, ParagraphBlock para, string docFont, float docSize)
        {
            try
            {
                range.ParagraphFormat.SpaceBefore = 0.0f;
                range.ParagraphFormat.SpaceAfter = 4.0f;
                range.ParagraphFormat.LeftIndent = 0.0f;
            }
            catch { }

            if (para.Inline != null)
            {
                RenderInlines(range, para.Inline, docFont, docSize, false, false, false, docFont, docSize);
            }

            range.InsertParagraphAfter();
            range.Collapse(Word.Enums.WdCollapseDirection.wdCollapseEnd);
        }

        private static void RenderList(Word.Application app, Word.Range range, ListBlock list, string docFont, float docSize)
        {
            // Numbering is applied once the whole list tree is written, not per item as it is
            // rendered. ApplyNumberDefault() on each freshly inserted paragraph makes Word start a
            // brand new single-item list every time, so each point restarted at "1.", and the
            // RemoveNumbers that ended the walk stripped the numbering off whichever list the
            // trailing paragraph still belonged to -- taking that item's number with it.
            var items = new List<ListItemPlacement>();
            RenderListItems(app, range, list, docFont, docSize, 1, items);
            ApplyListFormatting(app, items);
        }

        /// <summary>
        /// One rendered list item, held until the whole list tree is written so numbering can be
        /// applied in a single coordinated pass.
        /// </summary>
        private sealed class ListItemPlacement
        {
            public Word.Range Range;
            public int Level;
            public bool Ordered;
        }

        private static void RenderListItems(Word.Application app, Word.Range range, ListBlock list,
            string docFont, float docSize, int level, List<ListItemPlacement> items)
        {
            foreach (var child in PlanListBlocks(list))
            {
                var para = child as ParagraphBlock;
                if (para != null)
                {
                    int itemStart = range.Start;

                    try
                    {
                        range.ParagraphFormat.SpaceBefore = 0.0f;
                        range.ParagraphFormat.SpaceAfter = 2.0f;
                    }
                    catch { }

                    if (para.Inline != null)
                    {
                        RenderInlines(range, para.Inline, docFont, docSize, false, false, false, docFont, docSize);
                    }

                    CollectItemRange(app, items, itemStart, range.End, level, list.IsOrdered);

                    range.InsertParagraphAfter();
                    range.Collapse(Word.Enums.WdCollapseDirection.wdCollapseEnd);
                    continue;
                }

                var nested = child as ListBlock;
                if (nested != null)
                {
                    // Nested lists join the same numbering pass rather than going through
                    // RenderBlock. A separate pass would leave Word unable to tell where the outer
                    // list resumes, so the outer item after a nested list continued the inner
                    // list's sequence instead of its own.
                    ClearListFormat(range);
                    RenderListItems(app, range, nested, docFont, docSize, level + 1, items);
                    continue;
                }

                // Everything else a list item holds -- a table, a fenced code block, a quote -- is
                // rendered as a block in its own right. These used to fall through an if/else that
                // only knew about paragraphs and nested lists, so a table under a numbered point
                // was silently discarded. Numbering has to come off the insertion point first, or
                // Word carries the list format into the block.
                ClearListFormat(range);
                RenderBlock(app, range, child, docFont, docSize);
            }

            ClearListFormat(range);
        }

        // Outline gallery template 2 numbers "1." at the top level -- matching both a plain
        // one-level list and what the chat pane's preview shows -- and "1.1." below it, and Word
        // restarts and resumes each level on its own. Template 1 ("1)" / "a)") would change the
        // look of every existing flat list, which is the overwhelmingly common case.
        private const int OutlineNumberTemplateIndex = 2;
        private const int MaxWordListLevel = 9;
        private const float ListLevelIndentPoints = 18.0f;

        /// <summary>
        /// Records the span a list item's text occupies so numbering can be applied to it once the
        /// whole tree is written. A failure here costs the item its number, never its text.
        /// </summary>
        private static void CollectItemRange(Word.Application app, List<ListItemPlacement> items,
            int start, int end, int level, bool ordered)
        {
            try
            {
                if (app == null || app.ActiveDocument == null || end <= start) return;

                var placement = new ListItemPlacement();
                placement.Range = app.ActiveDocument.Range(start, end);
                placement.Level = level;
                placement.Ordered = ordered;
                items.Add(placement);
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("Could not capture list item range for numbering: {0}", ex.Message));
            }
        }

        /// <summary>
        /// Numbers or bullets a whole list tree in one pass. Each item is applied at its own outline
        /// level and told whether it continues the list it belongs to, which is what keeps a list
        /// running 1., 2., 3. across items -- including across an item that contains a table or a
        /// nested list, whose paragraphs sit between the item ranges and must stay out of the
        /// numbering.
        /// </summary>
        private static void ApplyListFormatting(Word.Application app, List<ListItemPlacement> items)
        {
            if (app == null || items == null || items.Count == 0) return;

            try
            {
                Word.ListTemplate numbered = null;
                Word.ListTemplate bulleted = null;

                // Only the first item of each template starts a new list; every later item
                // continues it, and its outline level alone decides where it sits. Word restarts
                // and resumes each level on its own from there. Marking the first item of a NESTED
                // list as a fresh start instead put that item back at level 1.
                bool numberedStarted = false;
                bool bulletedStarted = false;

                foreach (var item in items)
                {
                    Word.ListTemplate template;
                    bool continuesList;
                    if (item.Ordered)
                    {
                        if (numbered == null)
                            numbered = app.ListGalleries[Word.Enums.WdListGalleryType.wdOutlineNumberGallery]
                                .ListTemplates[OutlineNumberTemplateIndex];
                        template = numbered;
                        continuesList = numberedStarted;
                    }
                    else
                    {
                        if (bulleted == null)
                            bulleted = app.ListGalleries[Word.Enums.WdListGalleryType.wdBulletGallery]
                                .ListTemplates[1];
                        template = bulleted;
                        continuesList = bulletedStarted;
                    }
                    if (template == null) continue;

                    int level = Math.Min(Math.Max(item.Level, 1), MaxWordListLevel);
                    try
                    {
                        item.Range.ListFormat.ApplyListTemplateWithLevel(
                            template,
                            continuesList,
                            Word.Enums.WdListApplyTo.wdListApplyToSelection,
                            Word.Enums.WdDefaultListBehavior.wdWord10ListBehavior,
                            level);

                        if (item.Ordered) numberedStarted = true;
                        else bulletedStarted = true;

                        // Starting a list is always a level-1 act as far as Word is concerned, so
                        // the first item of a template that happens to be nested (an ordered list
                        // opening underneath a bullet, say) comes back at the top level. Put it
                        // back where it belongs.
                        try
                        {
                            if (item.Range.ListFormat.ListLevelNumber != level)
                                item.Range.ListFormat.ListLevelNumber = level;
                        }
                        catch { }

                        // The bullet gallery is single-level, so nesting has to be shown by indent.
                        // Set it absolutely -- adding to whatever Word left on the paragraph
                        // compounded the inherited indent of the item before it.
                        if (!item.Ordered)
                            item.Range.ParagraphFormat.LeftIndent = ListLevelIndentPoints * (level + 1);
                    }
                    catch (Exception ex)
                    {
                        Logger.Warn(string.Format("Could not apply list formatting to an item: {0}", ex.Message));
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("Could not apply list formatting: {0}", ex.Message));
            }
        }

        /// <summary>
        /// Every block a list's items contain, in document order. Paragraphs become the numbered
        /// or bulleted lines; anything else is rendered as a standalone block beneath its item.
        /// Rendering needs a live Word instance, so this traversal is separated out to keep the
        /// "a list item never swallows its contents" guarantee testable on its own.
        /// </summary>
        internal static List<Block> PlanListBlocks(ListBlock list)
        {
            var planned = new List<Block>();
            if (list == null) return planned;

            foreach (var item in list)
            {
                var listItem = item as ListItemBlock;
                if (listItem == null) continue;

                foreach (var child in listItem)
                {
                    planned.Add(child);
                }
            }
            return planned;
        }

        /// <summary>
        /// Takes list numbering/bullets and the list indent off the insertion point, so a block
        /// started here is not absorbed into the surrounding list's formatting.
        /// </summary>
        private static void ClearListFormat(Word.Range range)
        {
            try
            {
                range.ListFormat.RemoveNumbers(Word.Enums.WdNumberType.wdNumberParagraph);
            }
            catch { }
            try
            {
                range.ParagraphFormat.LeftIndent = 0.0f;
                range.ParagraphFormat.FirstLineIndent = 0.0f;
            }
            catch { }
        }

        private static bool RenderTable(Word.Application app, Word.Range range, Table table, string docFont, float docSize)
        {
            var rows = new List<TableRow>();
            int maxCols = 0;
            foreach (var child in table)
            {
                var row = child as TableRow;
                if (row != null)
                {
                    rows.Add(row);
                    if (row.Count > maxCols) maxCols = row.Count;
                }
            }

            if (rows.Count == 0 || maxCols == 0) return false;

            var wordDoc = app.ActiveDocument;
            if (wordDoc == null) return false;

            // The insertion point may still carry the numbering/indent of a list the table is
            // nested inside; Word would build a numbered, indented table from it.
            ClearListFormat(range);

            Word.Table wordTable = null;
            try
            {
                wordTable = wordDoc.Tables.Add(range, rows.Count, maxCols);
                wordTable.Borders.Enable = 1;
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("Failed to add Word table: {0}", ex.Message));
                return false;
            }

            for (int r = 0; r < rows.Count; r++)
            {
                var tableRow = rows[r];
                bool isHeader = tableRow.IsHeader || r == 0;

                if (isHeader)
                {
                    try
                    {
                        wordTable.Rows[r + 1].HeadingFormat = -1;
                        wordTable.Rows[r + 1].Range.Font.Bold = 1;
                    }
                    catch { }
                }

                for (int c = 0; c < tableRow.Count && c < maxCols; c++)
                {
                    var cell = tableRow[c] as TableCell;
                    if (cell != null)
                    {
                        try
                        {
                            var cellRange = wordTable.Cell(r + 1, c + 1).Range;
                            // A cell range includes Word's end-of-cell marker.  Do not render
                            // into that marker or it can result in an extra paragraph/cell.
                            if (cellRange.End > cellRange.Start)
                                cellRange.End = cellRange.End - 1;
                            cellRange.Text = string.Empty;
                            cellRange.Collapse(Word.Enums.WdCollapseDirection.wdCollapseStart);
                            cellRange.Font.Name = docFont;
                            cellRange.Font.Size = docSize;
                            cellRange.Font.Bold = isHeader ? 1 : 0;
                            ApplyColumnAlignment(cellRange, table, c);

                            foreach (var cellBlock in cell)
                            {
                                var cellPara = cellBlock as ParagraphBlock;
                                if (cellPara != null && cellPara.Inline != null)
                                {
                                    RenderInlines(cellRange, cellPara.Inline, docFont, docSize, isHeader, false, false, docFont, docSize);
                                }
                            }
                        }
                        catch { }
                    }
                }
            }

            try
            {
                // Tables.Add inserts the table ahead of the paragraph the range sat in, so that
                // paragraph normally survives just past the table's end and is where the next
                // block belongs. Only add one when it did not -- unconditionally inserting left a
                // blank paragraph under every table.
                range.SetRange(wordTable.Range.End, wordTable.Range.End);
                if (IsInsideTable(range))
                {
                    range.InsertParagraphAfter();
                    range.Collapse(Word.Enums.WdCollapseDirection.wdCollapseEnd);
                }
                // The trailing paragraph inherits the table's formatting; reset it so following
                // text is not bold or indented.
                try
                {
                    range.Font.Bold = 0;
                    range.ParagraphFormat.LeftIndent = 0.0f;
                    range.ParagraphFormat.SpaceBefore = 0.0f;
                }
                catch { }
            }
            catch { }
            return true;
        }

        /// <summary>
        /// Mirrors the markdown delimiter row's alignment (:---, :---:, ---:) onto the Word cell.
        /// Markdig appends a trailing column definition beyond the real columns, so an index that
        /// runs past the table's actual width is ignored rather than trusted.
        /// </summary>
        private static void ApplyColumnAlignment(Word.Range cellRange, Table table, int columnIndex)
        {
            try
            {
                if (table.ColumnDefinitions == null || columnIndex < 0 || columnIndex >= table.ColumnDefinitions.Count)
                    return;

                var alignment = table.ColumnDefinitions[columnIndex].Alignment;
                if (!alignment.HasValue) return;

                switch (alignment.Value)
                {
                    case TableColumnAlign.Center:
                        cellRange.ParagraphFormat.Alignment = Word.Enums.WdParagraphAlignment.wdAlignParagraphCenter;
                        break;
                    case TableColumnAlign.Right:
                        cellRange.ParagraphFormat.Alignment = Word.Enums.WdParagraphAlignment.wdAlignParagraphRight;
                        break;
                    default:
                        cellRange.ParagraphFormat.Alignment = Word.Enums.WdParagraphAlignment.wdAlignParagraphLeft;
                        break;
                }
            }
            catch { }
        }

        private static bool IsInsideTable(Word.Range range)
        {
            try
            {
                object inTable = range.Information(Word.Enums.WdInformation.wdWithInTable);
                return inTable != null && Convert.ToBoolean(inTable);
            }
            catch
            {
                // Cannot tell -- assume we are still in the table so a paragraph is added and the
                // next block never lands inside a cell.
                return true;
            }
        }

        /// <summary>
        /// Last-resort rendering for a table Word would not create: one paragraph per row with
        /// tab-separated cells. Ugly, but the data reaches the document.
        /// </summary>
        private static void RenderTableAsText(Word.Range range, Table table, string docFont, float docSize)
        {
            foreach (var child in table)
            {
                var row = child as TableRow;
                if (row == null) continue;

                for (int c = 0; c < row.Count; c++)
                {
                    if (c > 0)
                    {
                        try
                        {
                            range.Text = "\t";
                            range.Collapse(Word.Enums.WdCollapseDirection.wdCollapseEnd);
                        }
                        catch { }
                    }

                    var cell = row[c] as TableCell;
                    if (cell == null) continue;

                    foreach (var cellBlock in cell)
                    {
                        var cellPara = cellBlock as ParagraphBlock;
                        if (cellPara != null && cellPara.Inline != null)
                        {
                            RenderInlines(range, cellPara.Inline, docFont, docSize, row.IsHeader, false, false, docFont, docSize);
                        }
                    }
                }

                range.InsertParagraphAfter();
                range.Collapse(Word.Enums.WdCollapseDirection.wdCollapseEnd);
            }
        }

        private static void RenderCodeBlock(Word.Application app, Word.Range range, CodeBlock codeBlock, float docSize)
        {
            string codeText = codeBlock.Lines.ToString();
            if (string.IsNullOrEmpty(codeText)) return;

            try
            {
                range.Font.Name = "Consolas";
                range.Font.Size = Math.Max(8.5f, docSize - 1.0f);
                range.Font.Bold = 0;
                range.Font.Italic = 0;
                range.Font.StrikeThrough = 0;
                range.ParagraphFormat.LeftIndent = 14.0f;
                range.ParagraphFormat.SpaceBefore = 2.0f;
                range.ParagraphFormat.SpaceAfter = 2.0f;
                range.ParagraphFormat.LineSpacingRule = Word.Enums.WdLineSpacing.wdLineSpaceSingle;

                range.Text = codeText.TrimEnd('\r', '\n');
                range.Collapse(Word.Enums.WdCollapseDirection.wdCollapseEnd);
                range.InsertParagraphAfter();
                range.Collapse(Word.Enums.WdCollapseDirection.wdCollapseEnd);
                range.ParagraphFormat.LeftIndent = 0.0f;
            }
            catch { }
        }

        private static void RenderQuote(Word.Application app, Word.Range range, QuoteBlock quote, string docFont, float docSize)
        {
            foreach (var child in quote)
            {
                var para = child as ParagraphBlock;
                if (para != null)
                {
                    try
                    {
                        range.ParagraphFormat.LeftIndent = 18.0f;
                        range.ParagraphFormat.SpaceBefore = 2.0f;
                        range.ParagraphFormat.SpaceAfter = 2.0f;
                    }
                    catch { }

                    if (para.Inline != null)
                    {
                        RenderInlines(range, para.Inline, docFont, docSize, false, true, false, docFont, docSize);
                    }

                    range.InsertParagraphAfter();
                    range.Collapse(Word.Enums.WdCollapseDirection.wdCollapseEnd);
                }
            }

            try
            {
                range.ParagraphFormat.LeftIndent = 0.0f;
            }
            catch { }
        }

        private static void RenderThematicBreak(Word.Application app, Word.Range range)
        {
            try
            {
                range.InsertParagraphAfter();
                range.Borders[Word.Enums.WdBorderType.wdBorderBottom].LineStyle = Word.Enums.WdLineStyle.wdLineStyleSingle;
                range.Collapse(Word.Enums.WdCollapseDirection.wdCollapseEnd);
            }
            catch { }
        }

        private static void RenderInlines(Word.Range range, ContainerInline inlines, string currentFont, float currentSize,
            bool isBold, bool isItalic, bool isStrike, string docFont, float docBaseSize)
        {
            if (inlines == null) return;

            foreach (var inline in inlines)
            {
                var lit = inline as LiteralInline;
                if (lit != null)
                {
                    string text = lit.Content.ToString();
                    if (!string.IsNullOrEmpty(text))
                    {
                        range.Text = text;
                        range.Font.Name = currentFont;
                        range.Font.Size = currentSize;
                        range.Font.Bold = isBold ? 1 : 0;
                        range.Font.Italic = isItalic ? 1 : 0;
                        range.Font.StrikeThrough = isStrike ? 1 : 0;
                        range.Collapse(Word.Enums.WdCollapseDirection.wdCollapseEnd);
                    }
                    continue;
                }

                var emph = inline as EmphasisInline;
                if (emph != null)
                {
                    bool nodeBold = (emph.DelimiterCount == 2 || emph.DelimiterCount == 3);
                    bool nodeItalic = (emph.DelimiterCount == 1 || emph.DelimiterCount == 3);
                    RenderInlines(range, emph, currentFont, currentSize, isBold || nodeBold, isItalic || nodeItalic, isStrike, docFont, docBaseSize);
                    continue;
                }

                var code = inline as CodeInline;
                if (code != null)
                {
                    string text = code.Content;
                    if (!string.IsNullOrEmpty(text))
                    {
                        range.Text = text;
                        range.Font.Name = "Consolas";
                        range.Font.Size = Math.Max(9.0f, docBaseSize - 1.0f);
                        range.Font.Bold = 0;
                        range.Font.Italic = 0;
                        range.Font.StrikeThrough = 0;
                        range.Collapse(Word.Enums.WdCollapseDirection.wdCollapseEnd);
                    }
                    continue;
                }

                var lineBreak = inline as LineBreakInline;
                if (lineBreak != null)
                {
                    // Chr(11), Word's manual line break -- NOT "\n", which Word turns into a
                    // paragraph mark. Inside a list item that split the item in two and the
                    // continuation line picked up a number of its own, so a single point came
                    // out as "1. Background:" / "2. The Electrical Division is unable to...".
                    range.Text = "\v";
                    range.Collapse(Word.Enums.WdCollapseDirection.wdCollapseEnd);
                    continue;
                }

                var autoLink = inline as AutolinkInline;
                if (autoLink != null)
                {
                    string url = autoLink.Url;
                    if (!string.IsNullOrEmpty(url))
                    {
                        range.Text = url;
                        range.Font.Name = currentFont;
                        range.Font.Size = currentSize;
                        range.Font.Underline = Word.Enums.WdUnderline.wdUnderlineSingle;
                        range.Collapse(Word.Enums.WdCollapseDirection.wdCollapseEnd);
                    }
                    continue;
                }

                var link = inline as LinkInline;
                if (link != null)
                {
                    if (link.FirstChild != null)
                    {
                        RenderInlines(range, link, currentFont, currentSize, isBold, isItalic, isStrike, docFont, docBaseSize);
                    }
                    else if (!string.IsNullOrEmpty(link.Url))
                    {
                        range.Text = link.Url;
                        range.Font.Name = currentFont;
                        range.Font.Size = currentSize;
                        range.Font.Underline = Word.Enums.WdUnderline.wdUnderlineSingle;
                        range.Collapse(Word.Enums.WdCollapseDirection.wdCollapseEnd);
                    }
                    continue;
                }

                var container = inline as ContainerInline;
                if (container != null)
                {
                    RenderInlines(range, container, currentFont, currentSize, isBold, isItalic, isStrike, docFont, docBaseSize);
                }
            }
        }
    }
}
