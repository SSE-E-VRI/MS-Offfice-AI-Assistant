using System;
using System.Collections.Generic;
using MSOfficeAIAssistant.Core;

namespace MSOfficeAIAssistant.Tests
{
    public static class SpreadsheetActionParserTests
    {
        public static void RunAll()
        {
            TestIsSafeRangeBounds();
            TestTryParseRangeExtent();
            TestIndexToColumnLetter();
            TestBuildRangeAddress();
            TestExtractActions();
            TestDtdRejection();
            TestMaxActionsCap();
            TestRejectedTargets();
            TestUndoableFlags();
        }

        private static void TestIsSafeRangeBounds()
        {
            Assert(SpreadsheetActionParser.IsSafeRangeBounds(1, 1, 10, 5), "Valid range 1,1 to 10,5");
            Assert(SpreadsheetActionParser.IsSafeRangeBounds(1, 1, 1048576, 16384) == false, "Full sheet exceeds MaxCellExtent");
            Assert(SpreadsheetActionParser.IsSafeRangeBounds(10, 5, 1, 1) == false, "Inverted range is invalid");
            Assert(SpreadsheetActionParser.IsSafeRangeBounds(0, 1, 10, 5) == false, "Zero column is invalid");
            Assert(SpreadsheetActionParser.IsSafeRangeBounds(-1, 1, 10, 5) == false, "Negative column is invalid");
            Assert(SpreadsheetActionParser.IsSafeRangeBounds(1, 1, 500, 200) == true, "500x200 is 100,000 cells (exact max allowed)");
            Assert(SpreadsheetActionParser.IsSafeRangeBounds(1, 1, 501, 200) == false, "501x200 exceeds 100,000 max cells");
        }

        private static void TestTryParseRangeExtent()
        {
            int startCol, startRow, endCol, endRow;
            Assert(SpreadsheetActionParser.TryParseRangeExtent("A1", out startCol, out startRow, out endCol, out endRow) &&
                   startCol == 1 && startRow == 1 && endCol == 1 && endRow == 1, "A1 extent");

            Assert(SpreadsheetActionParser.TryParseRangeExtent("B2:D10", out startCol, out startRow, out endCol, out endRow) &&
                   startCol == 2 && startRow == 2 && endCol == 4 && endRow == 10, "B2:D10 extent");

            Assert(SpreadsheetActionParser.TryParseRangeExtent("A1:A100", out startCol, out startRow, out endCol, out endRow) &&
                   startCol == 1 && startRow == 1 && endCol == 1 && endRow == 100, "A1:A100 extent");

            Assert(SpreadsheetActionParser.TryParseRangeExtent("A:A", out startCol, out startRow, out endCol, out endRow) == false, "Unbounded column range rejected");
            Assert(SpreadsheetActionParser.TryParseRangeExtent("1:1", out startCol, out startRow, out endCol, out endRow) == false, "Unbounded row range rejected");
            Assert(SpreadsheetActionParser.TryParseRangeExtent("InvalidRange!", out startCol, out startRow, out endCol, out endRow) == false, "Malformed range rejected");
        }

        private static void TestIndexToColumnLetter()
        {
            Assert(SpreadsheetActionParser.IndexToColumnLetter(1) == "A", "Col 1 is A");
            Assert(SpreadsheetActionParser.IndexToColumnLetter(26) == "Z", "Col 26 is Z");
            Assert(SpreadsheetActionParser.IndexToColumnLetter(27) == "AA", "Col 27 is AA");
            Assert(SpreadsheetActionParser.IndexToColumnLetter(28) == "AB", "Col 28 is AB");
            Assert(SpreadsheetActionParser.IndexToColumnLetter(702) == "ZZ", "Col 702 is ZZ");
            Assert(SpreadsheetActionParser.IndexToColumnLetter(703) == "AAA", "Col 703 is AAA");
        }

        private static void TestBuildRangeAddress()
        {
            Assert(SpreadsheetActionParser.BuildRangeAddress(1, 1, 1, 1) == "A1", "1,1 to 1,1 is A1");
            Assert(SpreadsheetActionParser.BuildRangeAddress(2, 2, 4, 10) == "B2:D10", "2,2 to 4,10 is B2:D10");
        }

        private static void TestExtractActions()
        {
            string aiResponse = "Here are the suggested updates:\n\n" +
                                "<excel_actions>\n" +
                                "  <excel_action target=\"B2:B10\" type=\"formula\" formula=\"=SUM(A2:A10)\" description=\"Compute sum\" />\n" +
                                "  <excel_action target=\"C2:C100\" type=\"remove_duplicates\" value=\"columns:1,2\" description=\"Remove duplicate rows\" />\n" +
                                "</excel_actions>\n\n" +
                                "Let me know if you want further changes.";

            string cleanContent;
            List<SpreadsheetAction> actions = SpreadsheetActionParser.ExtractActions(aiResponse, out cleanContent);

            Assert(actions != null && actions.Count == 2, "Parsed 2 actions");
            Assert(actions[0].Target == "B2:B10", "Action 0 target B2:B10");
            Assert(actions[0].Type == SpreadsheetActionType.Formula, "Action 0 type Formula");
            Assert(actions[0].Content == "=SUM(A2:A10)", "Action 0 content");
            Assert(actions[0].IsUndoable == true, "Formula action is undoable");

            Assert(actions[1].Target == "C2:C100", "Action 1 target C2:C100");
            Assert(actions[1].Type == SpreadsheetActionType.RemoveDuplicates, "Action 1 type RemoveDuplicates");
            Assert(actions[1].IsUndoable == false, "Remove duplicates is NOT undoable");

            Assert(cleanContent.IndexOf("<excel_actions>", StringComparison.Ordinal) < 0, "XML block stripped from clean content");
            Assert(cleanContent.IndexOf("Here are the suggested updates", StringComparison.Ordinal) >= 0, "Conversational text preserved");
        }

        private static void TestDtdRejection()
        {
            string dtdAttack = "<excel_actions>\n" +
                               "  <!DOCTYPE test [ <!ENTITY xxe SYSTEM \"file:///c:/windows/win.ini\"> ]>\n" +
                               "  <excel_action target=\"A1\" type=\"value\" value=\"&xxe;\" />\n" +
                               "</excel_actions>";
            string clean;
            List<SpreadsheetAction> actions = SpreadsheetActionParser.ExtractActions(dtdAttack, out clean);
            Assert(actions != null && actions.Count == 0, "DTD entity expansion rejected with 0 parsed actions");
        }

        private static void TestMaxActionsCap()
        {
            var sb = new System.Text.StringBuilder();
            sb.AppendLine("<excel_actions>");
            for (int i = 1; i <= 35; i++)
            {
                sb.AppendFormat("  <excel_action target=\"A{0}\" type=\"value\" value=\"{0}\" />\n", i);
            }
            sb.AppendLine("</excel_actions>");

            string clean;
            List<SpreadsheetAction> actions = SpreadsheetActionParser.ExtractActions(sb.ToString(), out clean);
            Assert(actions != null && actions.Count == 25, "Actions capped strictly at 25 actions");
        }

        private static void TestRejectedTargets()
        {
            // Sheet-qualified targets
            Assert(!SpreadsheetActionParser.IsSafeTarget("'Sheet2'!A1"), "'Sheet2'!A1 is rejected");
            Assert(!SpreadsheetActionParser.IsSafeTarget("Sheet1!B2:C10"), "Sheet1!B2:C10 is rejected");
            // Multi-area targets
            Assert(!SpreadsheetActionParser.IsSafeTarget("A1,C1"), "A1,C1 multi-area is rejected");
            Assert(!SpreadsheetActionParser.IsSafeTarget("A1:B2,D1:E2"), "A1:B2,D1:E2 is rejected");
            // Whole column / row targets
            Assert(!SpreadsheetActionParser.IsSafeTarget("A:A"), "A:A whole column is rejected");
            Assert(!SpreadsheetActionParser.IsSafeTarget("1:1"), "1:1 whole row is rejected");
            // Special / injected addresses
            Assert(!SpreadsheetActionParser.IsSafeTarget("CMD|'/C calc'!A1"), "DDE injection address is rejected");
        }

        private static void TestUndoableFlags()
        {
            SpreadsheetAction actFormula = new SpreadsheetAction { Type = SpreadsheetActionType.Formula };
            SpreadsheetAction actValue = new SpreadsheetAction { Type = SpreadsheetActionType.Value };
            SpreadsheetAction actFilldown = new SpreadsheetAction { Type = SpreadsheetActionType.FillDown };
            SpreadsheetAction actDuplicates = new SpreadsheetAction { Type = SpreadsheetActionType.RemoveDuplicates };
            SpreadsheetAction actTable = new SpreadsheetAction { Type = SpreadsheetActionType.CreateTable };
            SpreadsheetAction actChart = new SpreadsheetAction { Type = SpreadsheetActionType.Chart };
            SpreadsheetAction actPivot = new SpreadsheetAction { Type = SpreadsheetActionType.PivotTable };
            SpreadsheetAction actNamedRange = new SpreadsheetAction { Type = SpreadsheetActionType.NamedRange };

            Assert(actFormula.IsUndoable, "Formula is undoable");
            Assert(actValue.IsUndoable, "Value is undoable");
            Assert(actFilldown.IsUndoable, "Filldown is undoable");
            Assert(!actDuplicates.IsUndoable, "remove_duplicates is not undoable");
            Assert(!actTable.IsUndoable, "create_table is not undoable");
            Assert(!actChart.IsUndoable, "chart is not undoable");
            Assert(!actPivot.IsUndoable, "pivot_table is not undoable");
            Assert(!actNamedRange.IsUndoable, "named_range is not undoable");
        }

        private static void Assert(bool condition, string message)
        {
            if (!condition)
            {
                throw new Exception("Assertion Failed: " + message);
            }
        }
    }
}
