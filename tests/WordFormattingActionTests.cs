using System;
using System.Collections.Generic;
using MSOfficeAIAssistant.Core;
using MSOfficeAIAssistant.Core.Actions;

namespace MSOfficeAIAssistant.Tests
{
    public static class WordFormattingActionTests
    {
        public static void RunAll()
        {
            TestWordToolRegistryEntries();
            TestWordFormatActionTypesList();
            TestExtractorFindReplace();
            TestExtractorApplyStyle();
            TestExtractorSetCase();
            TestExtractorReorganize();
            TestExtractorNormalize();
            TestWordHostGuard();
        }

        private static void TestWordToolRegistryEntries()
        {
            var find = ToolRegistry.GetTool("word.find_replace", "Word");
            Assert(find != null, "word.find_replace should be registered");
            Assert(find.RiskLevel == 2, "find_replace risk 2");
            Assert(find.RequiresApproval == true, "find_replace requires approval");
            Assert(find.IsUndoable == true, "find_replace undoable");

            var style = ToolRegistry.GetTool("word.apply_style", "Word");
            Assert(style != null, "word.apply_style registered");
            Assert(style.RiskLevel == 1, "apply_style risk 1");

            var scase = ToolRegistry.GetTool("word.set_case", "Word");
            Assert(scase != null, "word.set_case registered");

            var reorder = ToolRegistry.GetTool("word.reorganize_paragraphs", "Word");
            Assert(reorder != null, "word.reorganize_paragraphs registered");
            Assert(reorder.RiskLevel == 2, "reorganize risk 2");

            var norm = ToolRegistry.GetTool("word.normalize_whitespace", "Word");
            Assert(norm != null, "word.normalize_whitespace registered");
        }

        private static void TestWordFormatActionTypesList()
        {
            string list = ToolRegistry.FormatActionTypesList("Word");
            Assert(list.IndexOf("find_replace", StringComparison.OrdinalIgnoreCase) >= 0, "list should contain find_replace");
            Assert(list.IndexOf("apply_style", StringComparison.OrdinalIgnoreCase) >= 0, "list should contain apply_style");
            Assert(list.IndexOf("set_case", StringComparison.OrdinalIgnoreCase) >= 0, "list should contain set_case");
            Assert(list.IndexOf("reorganize_paragraphs", StringComparison.OrdinalIgnoreCase) >= 0, "list should contain reorganize_paragraphs");
            Assert(list.IndexOf("normalize_whitespace", StringComparison.OrdinalIgnoreCase) >= 0, "list should contain normalize_whitespace");
            Assert(list.IndexOf("add_comment", StringComparison.OrdinalIgnoreCase) >= 0, "list should still contain add_comment");
        }

        private static void TestExtractorFindReplace()
        {
            string text = "<office_actions>[{\"host\":\"Word\",\"operation\":\"word.find_replace\",\"target\":{},\"input\":{\"find\":\"hello\",\"replace\":\"hi\"},\"risk_level\":2}]</office_actions> leftover";
            var res = ActionExtractor.Extract(text, "Word");
            Assert(res.HasActions, "Should extract find_replace");
            Assert(res.Actions.Count == 1, "One action");
            Assert(res.Actions[0].Operation == "word.find_replace", "operation name");
            Assert(res.Actions[0].GetParameterString("find") == "hello", "find param");
            Assert(res.Actions[0].GetParameterString("replace") == "hi", "replace param");
        }

        private static void TestExtractorApplyStyle()
        {
            string text = "<office_actions>[{\"host\":\"Word\",\"operation\":\"word.apply_style\",\"target\":{},\"input\":{\"style\":\"Heading 1\",\"paragraph\":2},\"risk_level\":1}]</office_actions>";
            var res = ActionExtractor.Extract(text, "Word");
            Assert(res.HasActions, "apply_style extract");
            Assert(res.Actions[0].GetParameterString("style") == "Heading 1", "style param");
            Assert(res.Actions[0].GetParameterInt("paragraph") == 2, "paragraph param");
        }

        private static void TestExtractorSetCase()
        {
            string text = "<office_actions>[{\"host\":\"Word\",\"operation\":\"word.set_case\",\"target\":{},\"input\":{\"case_type\":\"upper\"},\"risk_level\":1}]</office_actions>";
            var res = ActionExtractor.Extract(text, "Word");
            Assert(res.HasActions, "set_case extract");
            Assert(res.Actions[0].GetParameterString("case_type") == "upper", "case_type");
        }

        private static void TestExtractorReorganize()
        {
            string text = "<office_actions>[{\"host\":\"Word\",\"operation\":\"word.reorganize_paragraphs\",\"target\":{},\"input\":{\"order\":\"3,1,2\"},\"risk_level\":2}]</office_actions>";
            var res = ActionExtractor.Extract(text, "Word");
            Assert(res.HasActions, "reorganize extract");
            Assert(res.Actions[0].GetParameterString("order") == "3,1,2", "order param");
        }

        private static void TestExtractorNormalize()
        {
            string text = "<office_actions>[{\"host\":\"Word\",\"operation\":\"word.normalize_whitespace\",\"target\":{},\"input\":{},\"risk_level\":1}]</office_actions>";
            var res = ActionExtractor.Extract(text, "Word");
            Assert(res.HasActions, "normalize extract");
            Assert(res.Actions[0].Operation == "word.normalize_whitespace", "op normalize");
        }

        private static void TestWordHostGuard()
        {
            // Ensure non-Word controller cannot execute Word tool
            var action = new OfficeAction { Host = "Word", Operation = "word.find_replace", Target = new ActionTarget() };
            action.Parameters = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
            action.Parameters["find"] = "a";
            action.Parameters["replace"] = "b";
            var mock = new MockGenericHost();
            var res = ToolRegistry.Execute(mock, action);
            Assert(!res.Success, "Word tool on non-Word controller should fail");
            Assert(res.ErrorMessage.IndexOf("Controller type mismatch", StringComparison.Ordinal) >= 0, "Should be controller mismatch");
        }

        private class MockGenericHost
        {
            // Intentionally not WordController, name does not contain Word/Excel/PowerPoint
        }

        private static void Assert(bool condition, string message)
        {
            if (!condition) throw new Exception("Assertion failed: " + message);
        }
    }
}
