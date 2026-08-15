using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Markdig;
using Markdig.Extensions.Tables;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;
using MistralOfficeAddin.Core;
using Word = NetOffice.WordApi;

namespace MistralOfficeAddin.Hosts
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
                // 1. Detect base font name and size from current selection or document Normal style
                string docFont = "Calibri";
                float docSize = 11.0f;

                try
                {
                    if (app.Selection != null && app.Selection.Font != null)
                    {
                        string selFont = app.Selection.Font.Name;
                        if (!string.IsNullOrWhiteSpace(selFont)) docFont = selFont;
                        float selSize = app.Selection.Font.Size;
                        if (selSize >= 4.0f && selSize <= 100.0f) docSize = selSize;
                    }
                    else if (app.ActiveDocument != null && app.ActiveDocument.Styles != null)
                    {
                        var normalStyle = app.ActiveDocument.Styles[Word.Enums.WdBuiltinStyle.wdStyleNormal];
                        if (normalStyle != null && normalStyle.Font != null)
                        {
                            if (!string.IsNullOrWhiteSpace(normalStyle.Font.Name)) docFont = normalStyle.Font.Name;
                            if (normalStyle.Font.Size >= 4.0f && normalStyle.Font.Size <= 100.0f) docSize = normalStyle.Font.Size;
                        }
                    }
                }
                catch (Exception ex)
                {
                    Logger.Warn(string.Format("Could not detect active font properties: {0}", ex.Message));
                }

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
                RenderTable(app, range, table, docFont, docSize);
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
            foreach (var item in list)
            {
                var listItem = item as ListItemBlock;
                if (listItem != null)
                {
                    foreach (var child in listItem)
                    {
                        var para = child as ParagraphBlock;
                        if (para != null)
                        {
                            try
                            {
                                if (list.IsOrdered)
                                    range.ListFormat.ApplyNumberDefault();
                                else
                                    range.ListFormat.ApplyBulletDefault();

                                range.ParagraphFormat.SpaceBefore = 0.0f;
                                range.ParagraphFormat.SpaceAfter = 2.0f;
                            }
                            catch { }

                            if (para.Inline != null)
                            {
                                RenderInlines(range, para.Inline, docFont, docSize, false, false, false, docFont, docSize);
                            }

                            range.InsertParagraphAfter();
                            range.Collapse(Word.Enums.WdCollapseDirection.wdCollapseEnd);
                        }
                        else
                        {
                            var nestedList = child as ListBlock;
                            if (nestedList != null)
                            {
                                RenderList(app, range, nestedList, docFont, docSize);
                            }
                        }
                    }
                }
            }

            try
            {
                range.ListFormat.RemoveNumbers(Word.Enums.WdNumberType.wdNumberParagraph);
            }
            catch { }
        }

        private static void RenderTable(Word.Application app, Word.Range range, Table table, string docFont, float docSize)
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

            if (rows.Count == 0 || maxCols == 0) return;

            var wordDoc = app.ActiveDocument;
            if (wordDoc == null) return;

            Word.Table wordTable = null;
            try
            {
                wordTable = wordDoc.Tables.Add(range, rows.Count, maxCols);
                wordTable.Borders.Enable = 1;
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("Failed to add Word table: {0}", ex.Message));
                return;
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
                            cellRange.Font.Name = docFont;
                            cellRange.Font.Size = docSize;
                            cellRange.Font.Bold = isHeader ? 1 : 0;

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
                range.SetRange(wordTable.Range.End, wordTable.Range.End);
                range.InsertParagraphAfter();
                range.Collapse(Word.Enums.WdCollapseDirection.wdCollapseEnd);
            }
            catch { }
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
                    range.Text = "\n";
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
