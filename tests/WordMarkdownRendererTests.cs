using System;
using System.Collections.Generic;
using Markdig;
using Markdig.Extensions.Tables;
using Markdig.Syntax;
using MSOfficeAIAssistant.Hosts;

namespace MSOfficeAIAssistant.Tests
{
    /// <summary>
    /// Covers the AST traversal WordMarkdownRenderer performs. The rendering itself needs a live
    /// Word instance, but the traversal is where content goes missing -- a table nested under a
    /// numbered point was dropped on the floor and never reached the document.
    /// </summary>
    public static class WordMarkdownRendererTests
    {
        // The shape a model actually produced for an official letter: a table nested under item 2
        // of an ordered list. This is the regression that motivated the fix.
        private const string LetterWithNestedTable =
            "Sir,\n" +
            "\n" +
            "1. **Background:**\n" +
            "   The Electrical Division is unable to conduct dewatering operations.\n" +
            "\n" +
            "2. **Request:**\n" +
            "   The details of the requirement are as follows:\n" +
            "\n" +
            "   | Parameter | Specification |\n" +
            "   |-----------|---------------|\n" +
            "   | Capacity  | 5 HP          |\n" +
            "   | Quantity  | 1 No.         |\n" +
            "\n" +
            "3. **Assurance:**\n" +
            "   The equipment will be returned in its original condition.\n";

        public static void RunAll()
        {
            TestNestedTableIsPlannedForRendering();
            TestNestedCodeAndQuoteArePlannedForRendering();
            TestListParagraphsAreStillPlanned();
            TestNestedListIsPlannedForRendering();
            TestPlanListBlocksHandlesNullAndEmpty();
            TestTopLevelTableParsesWithExpectedShape();
        }

        private static MarkdownDocument Parse(string markdown)
        {
            var pipeline = new MarkdownPipelineBuilder().UseAdvancedExtensions().Build();
            return Markdown.Parse(markdown, pipeline);
        }

        private static ListBlock FirstList(MarkdownDocument document)
        {
            foreach (var block in document)
            {
                var list = block as ListBlock;
                if (list != null) return list;
            }
            throw new Exception("Expected the markdown to contain a list block");
        }

        /// <summary>
        /// The bug: Markdig nests the table inside the ListItemBlock, and the renderer's list walk
        /// only handled paragraphs and nested lists, so the table never reached Word.
        /// </summary>
        private static void TestNestedTableIsPlannedForRendering()
        {
            var document = Parse(LetterWithNestedTable);
            var list = FirstList(document);

            // Guard the premise: the table really is nested inside a list item, not a sibling.
            bool tableIsTopLevel = false;
            foreach (var block in document)
            {
                if (block is Table) tableIsTopLevel = true;
            }
            Assert(!tableIsTopLevel, "Premise broken: the table should be nested inside the list, not top level");

            var planned = WordMarkdownRenderer.PlanListBlocks(list);
            Assert(CountOf<Table>(planned) == 1, "The table nested in the list item must be planned for rendering");
        }

        private static void TestNestedCodeAndQuoteArePlannedForRendering()
        {
            const string markdown =
                "- Item with extras\n" +
                "\n" +
                "  ```\n" +
                "  code line\n" +
                "  ```\n" +
                "\n" +
                "  > quoted line\n";

            var planned = WordMarkdownRenderer.PlanListBlocks(FirstList(Parse(markdown)));

            Assert(CountOf<FencedCodeBlock>(planned) == 1, "A fenced code block inside a list item must be planned for rendering");
            Assert(CountOf<QuoteBlock>(planned) == 1, "A quote inside a list item must be planned for rendering");
        }

        private static void TestListParagraphsAreStillPlanned()
        {
            var planned = WordMarkdownRenderer.PlanListBlocks(FirstList(Parse(LetterWithNestedTable)));

            // One paragraph per numbered point; the table is planned alongside them.
            Assert(CountOf<ParagraphBlock>(planned) == 3, "Every list item paragraph must still be planned for rendering");
            Assert(planned.Count == 4, "Planned blocks should be the three paragraphs plus the table");
        }

        /// <summary>
        /// A nested list has to come back out of the traversal as a ListBlock so the renderer can
        /// recurse into it within the same numbering pass. Handing it to the generic block
        /// dispatch instead would start a separate pass, and the outer list's next item then
        /// continued the inner list's sequence rather than its own.
        /// </summary>
        private static void TestNestedListIsPlannedForRendering()
        {
            const string markdown =
                "1. one\n" +
                "   1. inner a\n" +
                "   2. inner b\n" +
                "2. two\n";

            var outer = FirstList(Parse(markdown));
            var planned = WordMarkdownRenderer.PlanListBlocks(outer);

            Assert(CountOf<ListBlock>(planned) == 1, "A nested list must be planned for rendering");
            Assert(CountOf<ParagraphBlock>(planned) == 2, "Both outer list item paragraphs must be planned");

            ListBlock inner = null;
            foreach (var block in planned)
            {
                inner = block as ListBlock;
                if (inner != null) break;
            }
            var innerPlanned = WordMarkdownRenderer.PlanListBlocks(inner);
            Assert(CountOf<ParagraphBlock>(innerPlanned) == 2, "Both nested list item paragraphs must be planned");
        }

        private static void TestPlanListBlocksHandlesNullAndEmpty()
        {
            var planned = WordMarkdownRenderer.PlanListBlocks(null);
            Assert(planned != null && planned.Count == 0, "PlanListBlocks(null) must return an empty list, not null");
        }

        /// <summary>
        /// Markdig appends a column definition beyond the table's real width, so the renderer must
        /// size the Word table from the widest row instead. Pins that behaviour against upgrades.
        /// </summary>
        private static void TestTopLevelTableParsesWithExpectedShape()
        {
            const string markdown =
                "| Left | Center | Right |\n" +
                "|:-----|:------:|------:|\n" +
                "| a    | b      | c     |\n";

            Table table = null;
            foreach (var block in Parse(markdown))
            {
                table = block as Table;
                if (table != null) break;
            }
            Assert(table != null, "Expected a top-level table");

            int widestRow = 0;
            int rows = 0;
            foreach (var child in table)
            {
                var row = child as TableRow;
                if (row == null) continue;
                rows++;
                if (row.Count > widestRow) widestRow = row.Count;
            }

            Assert(rows == 2, "Expected a header row and one body row");
            Assert(widestRow == 3, "Expected three columns from the widest row");
            Assert(table.ColumnDefinitions.Count > widestRow,
                "Premise: Markdig reports more column definitions than real columns, so row width must drive table size");
            Assert(table.ColumnDefinitions[1].Alignment == TableColumnAlign.Center, "Centre alignment must survive parsing");
            Assert(table.ColumnDefinitions[2].Alignment == TableColumnAlign.Right, "Right alignment must survive parsing");
        }

        private static int CountOf<T>(IEnumerable<Block> blocks) where T : Block
        {
            int count = 0;
            foreach (var block in blocks)
            {
                if (block is T) count++;
            }
            return count;
        }

        private static void Assert(bool condition, string message)
        {
            if (!condition) throw new Exception("ASSERTION FAILED: " + message);
            Console.WriteLine("  [PASS] " + message);
        }
    }
}
