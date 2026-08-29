using System;
using System.Collections.Generic;
using System.Linq;
using MSOfficeAIAssistant.Core;
using MSOfficeAIAssistant.Core.Actions;

namespace MSOfficeAIAssistant.Tests
{
    public static class ToolRegistryTests
    {
        public static void RunAll()
        {
            TestBuiltInToolsRegistered();
            TestExcelToolsDefinitions();
            TestWordToolsDefinitions();
            TestPowerPointToolsDefinitions();
            TestPromptAllowListMatchesSSOT();
            TestOpenAiFunctionSchemaGeneration();
            TestNegativeMutationValidation();
        }

        private static void TestBuiltInToolsRegistered()
        {
            var excelTools = ToolRegistry.GetToolsForHost("Excel");
            Assert(excelTools.Count >= 10, "Expected at least 10 Excel tools");

            var pptTools = ToolRegistry.GetToolsForHost("PowerPoint");
            Assert(pptTools.Count >= 4, "Expected at least 4 PowerPoint tools");

            var wordTools = ToolRegistry.GetToolsForHost("Word");
            Assert(wordTools.Count >= 2, "Expected at least 2 Word tools");
        }

        private static void TestExcelToolsDefinitions()
        {
            var formulaTool = ToolRegistry.GetTool("excel.write_formula");
            Assert(formulaTool != null, "excel.write_formula must be registered");
            Assert(formulaTool.RiskLevel == 2, "Formula RiskLevel should be 2");
            Assert(formulaTool.IsUndoable == true, "Formula must be undoable");
            Assert(formulaTool.Parameters.Any(p => p.Name == "target" && p.IsRequired), "Formula requires target");
            Assert(formulaTool.Parameters.Any(p => p.Name == "formula" && p.IsRequired), "Formula requires formula parameter");

            // Alias lookup
            Assert(ToolRegistry.GetTool("formula") == formulaTool, "Alias 'formula' should resolve to excel.write_formula");
            Assert(ToolRegistry.GetTool("excel_write_formula") == formulaTool, "Alias 'excel_write_formula' should resolve to excel.write_formula");

            var dedupeTool = ToolRegistry.GetTool("excel.remove_duplicates");
            Assert(dedupeTool != null, "excel.remove_duplicates must be registered");
            Assert(dedupeTool.RiskLevel == 3, "remove_duplicates RiskLevel should be 3 (destructive)");
            Assert(dedupeTool.IsUndoable == false, "remove_duplicates is not undoable in Excel");
            // Host-scoped lookups
            var excelTable = ToolRegistry.GetTool("table", "Excel");
            Assert(excelTable != null && excelTable.Name == "excel.table", "Host-scoped 'table' in Excel must resolve to excel.table");
        }

        private static void TestWordToolsDefinitions()
        {
            var commentTool = ToolRegistry.GetTool("word.add_comment");
            Assert(commentTool != null, "word.add_comment must be registered");
            Assert(commentTool.Host == "Word", "Host should be Word");
            Assert(commentTool.RiskLevel == 1, "Comment RiskLevel should be 1");
            Assert(commentTool.IsUndoable == true, "Comment must be undoable");
            Assert(ToolRegistry.GetTool("add_comment") == commentTool, "Alias 'add_comment' should resolve to word.add_comment");

            var tableTool = ToolRegistry.GetTool("word.insert_table");
            Assert(tableTool != null, "word.insert_table must be registered");
            Assert(tableTool.Host == "Word", "Host should be Word");
            Assert(tableTool.RiskLevel == 2, "Table RiskLevel should be 2");
            Assert(tableTool.IsUndoable == true, "Table must be undoable");
            Assert(tableTool.Parameters.Any(p => p.Name == "rows" && p.IsRequired), "insert_table requires rows");
            Assert(tableTool.Parameters.Any(p => p.Name == "cols" && p.IsRequired), "insert_table requires cols");

            var wordTable = ToolRegistry.GetTool("table", "Word");
            Assert(wordTable != null && wordTable.Name == "word.insert_table", "Host-scoped 'table' in Word must resolve to word.insert_table");
        }

        private static void TestPowerPointToolsDefinitions()
        {
            var moveTool = ToolRegistry.GetTool("powerpoint.move_slide");
            Assert(moveTool != null, "powerpoint.move_slide must be registered");
            Assert(moveTool.RiskLevel == 2, "move_slide RiskLevel should be 2");
            Assert(moveTool.IsUndoable == true, "move_slide must be undoable");
            Assert(ToolRegistry.GetTool("move_slide") == moveTool, "Alias 'move_slide' should resolve");

            var sectionTool = ToolRegistry.GetTool("powerpoint.create_section");
            Assert(sectionTool != null, "powerpoint.create_section must be registered");
            Assert(ToolRegistry.GetTool("section+") == sectionTool, "Alias 'section+' should resolve");

            var notesTool = ToolRegistry.GetTool("powerpoint.set_notes");
            Assert(notesTool != null, "powerpoint.set_notes must be registered");
            Assert(notesTool.RiskLevel == 1, "set_notes RiskLevel should be 1");
        }

        private static void TestPromptAllowListMatchesSSOT()
        {
            string excelAllowList = ToolRegistry.FormatActionTypesList("Excel");
            Assert(excelAllowList.Contains("formula") && excelAllowList.Contains("remove_duplicates") && excelAllowList.Contains("find_replace") && excelAllowList.Contains("text_to_columns") && excelAllowList.Contains("format_cells") && excelAllowList.Contains("add_worksheet"),
                "Excel allow-list should contain key Excel action types");

            string pptAllowList = ToolRegistry.FormatActionTypesList("PowerPoint");
            Assert(pptAllowList.Contains("move_slide") && pptAllowList.Contains("create_section") && pptAllowList.Contains("rename_section") && pptAllowList.Contains("set_notes") && pptAllowList.Contains("create_slide") && pptAllowList.Contains("insert_image") && pptAllowList.Contains("delete_slide") && pptAllowList.Contains("duplicate_slide"),
                "PowerPoint allow-list should contain all PowerPoint action types");
        }

        private static void TestOpenAiFunctionSchemaGeneration()
        {
            var schemas = ToolRegistry.GetOpenAiToolsForHost("Excel");
            Assert(schemas.Count >= 10, "Expected at least 10 schemas for Excel");

            var formulaSchema = schemas.FirstOrDefault(s =>
            {
                var fn = s["function"] as Dictionary<string, object>;
                return fn != null && (string)fn["name"] == "excel_write_formula";
            });
            Assert(formulaSchema != null, "Expected excel_write_formula OpenAI schema");

            var fnDict = (Dictionary<string, object>)formulaSchema["function"];
            Assert((string)fnDict["name"] == "excel_write_formula", "Function name mismatch");
            Assert(fnDict.ContainsKey("parameters"), "Schema must have parameters dict");

            var paramsDict = (Dictionary<string, object>)fnDict["parameters"];
            Assert((string)paramsDict["type"] == "object", "Parameters type must be object");
            Assert(paramsDict.ContainsKey("properties"), "Parameters must contain properties");

            var required = (List<string>)paramsDict["required"];
            Assert(required.Contains("target"), "Target must be required");
            Assert(required.Contains("formula"), "Formula must be required");
        }

        private static void TestNegativeMutationValidation()
        {
            // Unregistered tool must return null
            Assert(ToolRegistry.GetTool("nonexistent.tool") == null, "Nonexistent tool must return null");
            Assert(ToolRegistry.GetTool("") == null, "Empty tool name must return null");
            Assert(ToolRegistry.GetTool(null) == null, "Null tool name must return null");

            ToolDefinition dummy;
            Assert(ToolRegistry.TryGetTool("invalid_alias_123", out dummy) == false, "TryGetTool on invalid name must return false");
            Assert(dummy == null, "Output tool must be null on failure");
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
