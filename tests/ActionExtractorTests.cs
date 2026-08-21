using System;
using System.Collections.Generic;
using MSOfficeAIAssistant.Core;
using MSOfficeAIAssistant.Core.Actions;

namespace MSOfficeAIAssistant.Tests
{
    public static class ActionExtractorTests
    {
        public static void RunAll()
        {
            TestJsonOfficeActionsExtraction();
            TestJsonCodeBlockExtraction();
            TestLegacyExcelActionsXmlExtraction();
            TestLegacyPowerPointActionsXmlExtraction();
            TestNativeToolCallExtraction();
            TestMalformedJsonYieldsExtractionFailure();
            TestOfficeActionSpreadsheetAdapterRoundTrip();
            TestOfficeActionPowerPointAdapterRoundTrip();
            TestStatusMappingRoundTrips();
            TestUnknownOperationDoesNotDefaultToMutation();
            TestBadgesAndDescriptions();
        }

        private static void TestStatusMappingRoundTrips()
        {
            // Pending & Approved -> Pending
            Assert(OfficeAction.ToSpreadsheetActionStatus(OfficeActionStatus.Pending) == SpreadsheetActionStatus.Pending, "Pending -> Pending");
            Assert(OfficeAction.ToSpreadsheetActionStatus(OfficeActionStatus.Approved) == SpreadsheetActionStatus.Pending, "Approved -> Pending (not yet executed)");
            Assert(OfficeAction.ToPowerPointActionStatus(OfficeActionStatus.Pending) == PowerPointActionStatus.Pending, "Pending -> Pending PPT");
            Assert(OfficeAction.ToPowerPointActionStatus(OfficeActionStatus.Approved) == PowerPointActionStatus.Pending, "Approved -> Pending PPT");

            // Applying -> Applying
            Assert(OfficeAction.ToSpreadsheetActionStatus(OfficeActionStatus.Applying) == SpreadsheetActionStatus.Applying, "Applying -> Applying");
            Assert(OfficeAction.ToPowerPointActionStatus(OfficeActionStatus.Applying) == PowerPointActionStatus.Applying, "Applying -> Applying PPT");

            // Applied -> Applied (NOT Error!)
            Assert(OfficeAction.ToSpreadsheetActionStatus(OfficeActionStatus.Applied) == SpreadsheetActionStatus.Applied, "Applied -> Applied");
            Assert(OfficeAction.ToPowerPointActionStatus(OfficeActionStatus.Applied) == PowerPointActionStatus.Applied, "Applied -> Applied PPT");

            // Failed & Rejected -> Error
            Assert(OfficeAction.ToSpreadsheetActionStatus(OfficeActionStatus.Failed) == SpreadsheetActionStatus.Error, "Failed -> Error");
            Assert(OfficeAction.ToSpreadsheetActionStatus(OfficeActionStatus.Rejected) == SpreadsheetActionStatus.Error, "Rejected -> Error");
            Assert(OfficeAction.ToPowerPointActionStatus(OfficeActionStatus.Failed) == PowerPointActionStatus.Error, "Failed -> Error PPT");
            Assert(OfficeAction.ToPowerPointActionStatus(OfficeActionStatus.Rejected) == PowerPointActionStatus.Error, "Rejected -> Error PPT");

            // Legacy enum -> OfficeActionStatus
            Assert(OfficeAction.FromSpreadsheetActionStatus(SpreadsheetActionStatus.Pending) == OfficeActionStatus.Pending, "Pending -> Pending OA");
            Assert(OfficeAction.FromSpreadsheetActionStatus(SpreadsheetActionStatus.Applying) == OfficeActionStatus.Applying, "Applying -> Applying OA");
            Assert(OfficeAction.FromSpreadsheetActionStatus(SpreadsheetActionStatus.Applied) == OfficeActionStatus.Applied, "Applied -> Applied OA");
            Assert(OfficeAction.FromSpreadsheetActionStatus(SpreadsheetActionStatus.Error) == OfficeActionStatus.Failed, "Error -> Failed OA");
        }

        private static void TestUnknownOperationDoesNotDefaultToMutation()
        {
            var unknownPpt = new OfficeAction
            {
                Host = "PowerPoint",
                Operation = "powerpoint.nonexistent_feature",
                Parameters = new Dictionary<string, object> { { "source", 0 }, { "target", 0 } }
            };
            Assert(unknownPpt.ToPowerPointAction() == null, "Unknown PowerPoint operation must return null and not default to mutating move_slide");

            var unknownExcel = new OfficeAction
            {
                Host = "Excel",
                Operation = "excel.unsupported_action"
            };
            Assert(unknownExcel.ToSpreadsheetAction() == null, "Unknown Excel operation must return null");
        }

        private static void TestJsonOfficeActionsExtraction()
        {
            string raw =
                "Here is your railway failure analysis.\n\n" +
                "<office_actions>\n" +
                "[\n" +
                "  {\n" +
                "    \"action_id\": \"act-1\",\n" +
                "    \"host\": \"Excel\",\n" +
                "    \"operation\": \"excel.write_formula\",\n" +
                "    \"target\": { \"sheet\": \"Summary\", \"range\": \"D10\" },\n" +
                "    \"input\": { \"formula\": \"=SUM(D2:D9)\" },\n" +
                "    \"expected_result\": \"Total failures in D10\",\n" +
                "    \"risk_level\": 2,\n" +
                "    \"requires_approval\": true,\n" +
                "    \"source_reason\": \"Summing column D totals\",\n" +
                "    \"evidence\": [\n" +
                "      { \"location\": \"Summary!D2:D9\", \"extracted_value\": \"14\", \"evidence_level\": \"DIRECTLY_OBSERVED\" }\n" +
                "    ]\n" +
                "  }\n" +
                "]\n" +
                "</office_actions>\n\n" +
                "Let me know if you need additional formatting.";

            var result = ActionExtractor.Extract(raw, "Excel");
            Assert(result.HasActions, "Expected actions to be extracted");
            Assert(result.Actions.Count == 1, "Expected exactly 1 action");
            Assert(!result.HasFailure, "Expected no extraction failure");

            var act = result.Actions[0];
            Assert(act.ActionId == "act-1", "ActionId mismatch");
            Assert(act.Host == "Excel", "Host mismatch");
            Assert(act.Operation == "excel.write_formula", "Operation mismatch");
            Assert(act.Target.Sheet == "Summary", "Target.Sheet mismatch");
            Assert(act.Target.Range == "D10", "Target.Range mismatch");
            Assert(act.RiskLevel == 2, "RiskLevel mismatch");
            Assert(act.RequiresApproval == true, "RequiresApproval mismatch");
            Assert(act.Evidence.Count == 1, "Expected 1 evidence claim");
            Assert(act.Evidence[0].EvidenceLevel == "DIRECTLY_OBSERVED", "EvidenceLevel mismatch");

            Assert(!result.CleanText.Contains("<office_actions>"), "Clean text should not contain action tags");
            Assert(result.CleanText.Contains("Here is your railway failure analysis."), "Clean text should retain prose");
        }

        private static void TestJsonCodeBlockExtraction()
        {
            string raw =
                "I will create a table for you:\n\n" +
                "```json\n" +
                "[\n" +
                "  {\n" +
                "    \"operation\": \"word.insert_table\",\n" +
                "    \"target\": { \"range\": \"Section 2\" },\n" +
                "    \"input\": { \"rows\": 4, \"cols\": 3 },\n" +
                "    \"expected_result\": \"Table with 4 rows and 3 columns\"\n" +
                "  }\n" +
                "]\n" +
                "```";

            var result = ActionExtractor.Extract(raw, "Word");
            Assert(result.HasActions, "Expected actions from JSON codeblock");
            Assert(result.Actions.Count == 1, "Expected 1 action");
            Assert(result.Actions[0].Operation == "word.insert_table", "Operation mismatch");
            Assert(result.Actions[0].Host == "Word", "Host should default to current host");
        }

        private static void TestLegacyExcelActionsXmlExtraction()
        {
            string raw =
                "Here are the spreadsheet changes:\n\n" +
                "<excel_actions>\n" +
                "  <excel_action target=\"A1\" type=\"value\" value=\"Revenue Summary\" description=\"Header cell\" />\n" +
                "  <excel_action target=\"B2:B10\" type=\"fill_down\" formula=\"=A2*1.15\" description=\"Fill formula\" />\n" +
                "</excel_actions>\n\n" +
                "Please review.";

            var result = ActionExtractor.Extract(raw, "Excel");
            Assert(result.HasActions, "Expected legacy XML actions to be parsed");
            Assert(result.Actions.Count == 2, "Expected 2 parsed actions");

            var a1 = result.Actions[0];
            Assert(a1.Host == "Excel", "Host mismatch");
            Assert(a1.Target.Range == "A1", "Target mismatch");
            Assert(a1.ActionBadge == "val", "Badge mismatch");

            var a2 = result.Actions[1];
            Assert(a2.Target.Range == "B2:B10", "Target mismatch");
            Assert(a2.ActionBadge == "fill", "Badge mismatch");

            Assert(!result.CleanText.Contains("<excel_actions>"), "Clean text should strip XML tags");
        }

        private static void TestLegacyPowerPointActionsXmlExtraction()
        {
            string raw =
                "Reorganizing deck:\n\n" +
                "<powerpoint_actions>\n" +
                "  <action type=\"move_slide\" source=\"3\" target=\"1\" />\n" +
                "  <action type=\"create_section\" name=\"Financials\" slide=\"2\" />\n" +
                "</powerpoint_actions>";

            var result = ActionExtractor.Extract(raw, "PowerPoint");
            Assert(result.HasActions, "Expected legacy PPT XML actions");
            Assert(result.Actions.Count == 2, "Expected 2 actions");

            var act1 = result.Actions[0];
            Assert(act1.Operation == "powerpoint.move_slide", "Operation mismatch");
            Assert(act1.GetParameterInt("source") == 3, "Source parameter mismatch");
            Assert(act1.GetParameterInt("target") == 1, "Target parameter mismatch");

            var act2 = result.Actions[1];
            Assert(act2.Operation == "powerpoint.create_section", "Operation mismatch");
            Assert(act2.GetParameterString("name") == "Financials", "Name parameter mismatch");
        }

        private static void TestNativeToolCallExtraction()
        {
            var nativeCalls = new List<MSOfficeAIAssistant.Providers.ToolCallDto>
            {
                new MSOfficeAIAssistant.Providers.ToolCallDto
                {
                    Id = "call_abc123",
                    Name = "excel_write_formula",
                    Arguments = "{\"target\": \"Sheet1!C5\", \"formula\": \"=AVERAGE(C1:C4)\", \"expected_result\": \"Average value in C5\"}"
                }
            };

            var result = ActionExtractor.Extract(null, "Excel", nativeCalls);
            Assert(result.HasActions, "Expected native tool call to produce action");
            Assert(result.Actions.Count == 1, "Expected 1 action");
            Assert(result.Actions[0].ActionId == "call_abc123", "ActionId should be preserved from ToolCallDto");
            Assert(result.Actions[0].Operation == "excel.write_formula", "Normalized operation name mismatch");
            Assert(result.Actions[0].Target.Sheet == "Sheet1", "Sheet mismatch");
            Assert(result.Actions[0].Target.Range == "C5", "Range mismatch");
        }

        private static void TestMalformedJsonYieldsExtractionFailure()
        {
            string malformed =
                "<office_actions>\n" +
                "[ { \"operation\": \"excel.write_formula\", \"target\": { \"range\": \"A1\" } UNTERMINATED JSON\n" +
                "</office_actions>";

            var result = ActionExtractor.Extract(malformed, "Excel");
            Assert(!result.HasActions, "Malformed JSON should not produce valid actions");
            Assert(result.HasFailure, "Expected extraction failure object");
            Assert(result.Failure.FailureType == "MalformedJsonOfficeActions", "FailureType mismatch");
            Assert(!string.IsNullOrEmpty(result.Failure.ErrorMessage), "Expected error message");
            Assert(!string.IsNullOrEmpty(result.Failure.RawSnippet), "Expected preserved raw snippet for plan routing");
        }

        private static void TestOfficeActionSpreadsheetAdapterRoundTrip()
        {
            var sa = new SpreadsheetAction
            {
                Type = SpreadsheetActionType.Formula,
                Target = "Sheet2!B5",
                Content = "=VLOOKUP(A5, D1:E10, 2, FALSE)",
                Description = "Lookup product rate"
            };

            var oa = OfficeAction.FromSpreadsheetAction(sa);
            Assert(oa.Host == "Excel", "Host mismatch");
            Assert(oa.Target.Sheet == "Sheet2", "Sheet mismatch");
            Assert(oa.Target.Range == "B5", "Range mismatch");
            Assert(oa.RiskLevel == 2, "Formula risk level should be 2");

            var projected = oa.ToSpreadsheetAction();
            Assert(projected.Type == SpreadsheetActionType.Formula, "Type mismatch");
            Assert(projected.Target == "B5" || projected.Target == "Sheet2!B5", "Target mismatch");
            Assert(projected.Content == "=VLOOKUP(A5, D1:E10, 2, FALSE)", "Content mismatch");
        }

        private static void TestOfficeActionPowerPointAdapterRoundTrip()
        {
            var pa = new PowerPointAction
            {
                Type = "set_notes",
                Slide = 4,
                Notes = "Emphasize supply chain resilience."
            };

            var oa = OfficeAction.FromPowerPointAction(pa);
            Assert(oa.Host == "PowerPoint", "Host mismatch");
            Assert(oa.Target.Slide == 4, "Slide mismatch");
            Assert(oa.RiskLevel == 1, "set_notes risk level should be 1");

            var projected = oa.ToPowerPointAction();
            Assert(projected.Type == "set_notes", "Type mismatch");
            Assert(projected.Slide == 4, "Slide mismatch");
            Assert(projected.Notes == "Emphasize supply chain resilience.", "Notes mismatch");
        }

        private static void TestBadgesAndDescriptions()
        {
            var a1 = new OfficeAction { Operation = "excel.write_formula" };
            Assert(a1.ActionBadge == "fx", "Badge mismatch for formula");

            var a2 = new OfficeAction { Operation = "excel.write_value" };
            Assert(a2.ActionBadge == "val", "Badge mismatch for value");

            var a3 = new OfficeAction { Operation = "word.add_comment" };
            Assert(a3.ActionBadge == "comm", "Badge mismatch for comment");

            var a4 = new OfficeAction { Operation = "powerpoint.create_slide" };
            Assert(a4.ActionBadge == "sld", "Badge mismatch for slide");
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
