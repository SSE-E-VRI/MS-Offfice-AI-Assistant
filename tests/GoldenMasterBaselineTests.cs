using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using MSOfficeAIAssistant.Core;

namespace MSOfficeAIAssistant.Tests
{
    /// <summary>
    /// Phase 0.0 Golden Master Baseline Test Suite.
    /// Provides headless, COM-free deterministic baselines for prompt assembly,
    /// XML action parsing (Excel 13 types & PowerPoint 4 types), and action audit serialization.
    /// This suite serves as the verification gate for Phase 0.1 orchestrator extraction.
    /// </summary>
    public static class GoldenMasterBaselineTests
    {
        // Canonical SHA-256 hash of all baseline outputs (system prompts, context compositions,
        // XML parsed action DTOs, slide outline parsing, and audit record shape).
        // Any mutation, reordering, whitespace change, or regression in Phase 0.1+ will fail this hash.
        public const string ExpectedGoldenMasterSha256 = "1f3971ed6790ed842fe73df80479d9e18f2ca723efd29ccde72454023918e4ca";

        public static void RunAll()
        {
            TestGoldenMasterHash();
            TestPromptAssemblerSystemPrompts();
            TestPromptAssemblerContextPermutations();
            TestPromptAssemblerAttachmentCitations();
            TestExcelActionParserFullCatalog();
            TestExcelActionParserSafetyBoundaries();
            TestPowerPointActionParserFullCatalog();
            TestPowerPointSlideDataParsing();
            TestActionAuditStoreSerializationBaseline();
        }

        private static void TestGoldenMasterHash()
        {
            string canonical = BuildAllBaselineOutputs();
            string actualHash = ComputeSha256(canonical);

            Assert(string.Equals(actualHash, ExpectedGoldenMasterSha256, StringComparison.OrdinalIgnoreCase),
                string.Format("Golden Master baseline hash mismatch! Expected '{0}', got '{1}'. A prompt, parser DTO, or audit format drifted.",
                    ExpectedGoldenMasterSha256, actualHash));
        }

        public static string BuildAllBaselineOutputs()
        {
            var sb = new StringBuilder();

            // 1. System Prompts across hosts and base prompts
            string[] hosts = new[] { "Excel", "Word", "PowerPoint", "UnknownHost" };
            string[] basePrompts = new[] { null, "Custom Base Instructions" };
            foreach (var h in hosts)
            {
                foreach (var b in basePrompts)
                {
                    sb.Append("=== SYSTEM PROMPT [Host=").Append(h).Append(", Base=").Append(b ?? "null").Append("] ===\n");
                    sb.Append(PromptAssembler.BuildHostAwareSystemPrompt(b, h)).Append("\n\n");
                }
            }

            // 2. Context Compositions across all scopes
            PromptContextScope[] scopes = new[]
            {
                PromptContextScope.Selection,
                PromptContextScope.CurrentFile,
                PromptContextScope.SelectionAndFile,
                PromptContextScope.AttachmentsOnly
            };
            foreach (var sc in scopes)
            {
                sb.Append("=== PROMPT CONTEXT [Scope=").Append(sc).Append("] ===\n");
                sb.Append(PromptAssembler.ComposePromptWithContext("Raw user prompt.", sc, "Selected text snippet.", "[Document Context: File.docx] Document body.")).Append("\n\n");
            }

            // 3. Attachment Citation Instruction
            sb.Append("=== ATTACHMENT CITATION INSTRUCTION ===\n");
            sb.Append(PromptAssembler.AppendAttachmentCitationInstruction("Analyze document.", "[Attachment: Data.pdf (Page 1)] Table of values.")).Append("\n\n");

            // 4. Excel Actions Full Catalog DTOs
            string excelXml =
                "<excel_actions>\n" +
                "  <excel_action target=\"A1\" type=\"value\" value=\"Revenue Analysis\" description=\"Title header\" />\n" +
                "  <excel_action target=\"B2\" type=\"formula\" formula=\"=SUM(B3:B20)\" description=\"Total sum\" />\n" +
                "  <excel_action target=\"C2:C20\" type=\"filldown\" formula=\"=B2*1.18\" description=\"Tax inclusive\" />\n" +
                "  <excel_action target=\"A3:D20\" type=\"table\" value=\"ColA\tColB\tColC\tColD\n1\t2\t3\t4\" description=\"Data table\" />\n" +
                "  <excel_action target=\"A3:D20\" type=\"create_table\" value=\"SalesTable\" description=\"Table description\" />\n" +
                "  <excel_action target=\"B2:B20\" type=\"conditional_format\" value=\"gt:1000\" description=\"Highlight large values\" />\n" +
                "  <excel_action target=\"A2:D20\" type=\"sort\" value=\"descending\" description=\"Sort by Revenue\" />\n" +
                "  <excel_action target=\"A2:D20\" type=\"filter\" value=\"ColB:>100\" description=\"Filter active\" />\n" +
                "  <excel_action target=\"E2:E20\" type=\"data_validation\" value=\"list:Open,Closed,Pending\" description=\"Status dropdown\" />\n" +
                "  <excel_action target=\"A2:B10\" type=\"chart\" value=\"column\" description=\"Revenue Column Chart\" />\n" +
                "  <excel_action target=\"A2:D50\" type=\"pivot_table\" value=\"rows:Region;vals:Sales\" description=\"Summary Pivot\" />\n" +
                "  <excel_action target=\"B2:B20\" type=\"named_range\" value=\"MonthlyRevenue\" description=\"Named range definition\" />\n" +
                "  <excel_action target=\"A2:D50\" type=\"remove_duplicates\" value=\"columns:1,2\" description=\"Deduplicate entries\" />\n" +
                "</excel_actions>";
            string excelClean;
            var excelActions = SpreadsheetActionParser.ExtractActions(excelXml, out excelClean);
            sb.Append("=== EXCEL PARSED ACTIONS ===\n");
            foreach (var ea in excelActions)
            {
                sb.AppendFormat("Target={0}|Type={1}|Content={2}|Desc={3}|Undoable={4}|Badge={5}\n",
                    ea.Target, ea.Type, ea.Content, ea.Description, ea.IsUndoable, ea.TypeBadge);
            }
            sb.Append("\n");

            // 5. PowerPoint Actions Full Catalog DTOs
            string pptXml =
                "<powerpoint_actions>\n" +
                "  <powerpoint_action type=\"move_slide\" source=\"4\" target=\"1\" />\n" +
                "  <powerpoint_action type=\"create_section\" name=\"Executive Briefing\" slide=\"1\" />\n" +
                "  <powerpoint_action type=\"rename_section\" section=\"2\" name=\"Technical Details\" />\n" +
                "  <powerpoint_action type=\"set_notes\" slide=\"3\" notes=\"Focus on Q3 turnaround timeline\" />\n" +
                "</powerpoint_actions>";
            string pptClean;
            var pptActions = PowerPointActionParser.ParseStructuredActions(pptXml, out pptClean);
            sb.Append("=== POWERPOINT PARSED ACTIONS ===\n");
            foreach (var pa in pptActions)
            {
                sb.AppendFormat("Type={0}|Src={1}|Tgt={2}|Slide={3}|Sec={4}|Name={5}|Notes={6}|Badge={7}|TgtDisp={8}|Desc={9}|ContDisp={10}\n",
                    pa.Type, pa.Source, pa.Target, pa.Slide, pa.Section, pa.Name, pa.Notes, pa.TypeBadge, pa.TargetDisplay, pa.Description, pa.ContentDisplay);
            }
            sb.Append("\n");

            // 6. PowerPoint Slide Data Outline Parsing
            string slideMarkdown =
                "# Slide 1: High-Speed Rail Modernization\n" +
                "- 25kV Traction Power Upgrades\n" +
                "- Substation telemetry integration\n" +
                "Speaker Notes: Highlight completion within scheduled Q2 window.\n" +
                "Visual: Single line diagram of traction substation.\n\n" +
                "# Slide 2: Implementation Roadmap\n" +
                "- Phase 1: Survey and Civil works\n" +
                "Notes: Reiterate safety clearance protocols.";
            var slides = PowerPointActionParser.ParseSlideData(slideMarkdown);
            sb.Append("=== POWERPOINT SLIDE DATA ===\n");
            foreach (var s in slides)
            {
                sb.AppendFormat("Title={0}|Bullets={1}|Notes={2}|Visual={3}\n",
                    s.Title, string.Join(";", s.Bullets.ToArray()), s.SpeakerNotes, s.VisualSuggestion);
            }
            sb.Append("\n");

            // 7. Audit Store Baseline Format
            sb.Append("=== AUDIT FORMAT SPECIFICATION ===\n");
            sb.Append("Host=Excel|ActionType=formula|Target=K20|Summary=Calculate sum|Undoable=True|Prompt=Sum items|Source=Book1.xlsx|Model=Mistral Large|Result=Applied successfully\n");

            return sb.ToString();
        }

        private static string ComputeSha256(string raw)
        {
            string normalized = (raw ?? string.Empty).Replace("\r\n", "\n");
            using (var sha = System.Security.Cryptography.SHA256.Create())
            {
                byte[] bytes = System.Text.Encoding.UTF8.GetBytes(normalized);
                byte[] hash = sha.ComputeHash(bytes);
                var sb = new StringBuilder();
                for (int i = 0; i < hash.Length; i++)
                {
                    sb.Append(hash[i].ToString("x2"));
                }
                return sb.ToString();
            }
        }

        private static void TestPromptAssemblerSystemPrompts()
        {
            string customBase = "Custom system instructions for enterprise assistant.";

            // 1. Excel host branch
            string excelPrompt = PromptAssembler.BuildHostAwareSystemPrompt(customBase, "Excel");
            Assert(excelPrompt.StartsWith(customBase), "Excel prompt starts with base prompt");
            Assert(excelPrompt.Contains("You are embedded inside Microsoft Excel."), "Excel prompt contains Excel marker");
            Assert(excelPrompt.Contains("<excel_actions>"), "Excel prompt contains <excel_actions> schema example");
            Assert(excelPrompt.Contains("Supported action types are formula, value, filldown, table, create_table"), "Excel prompt contains full action catalog list");

            // 2. Word host branch
            string wordPrompt = PromptAssembler.BuildHostAwareSystemPrompt(customBase, "Word");
            Assert(wordPrompt.StartsWith(customBase), "Word prompt starts with base prompt");
            Assert(wordPrompt.Contains("You are embedded inside Microsoft Word."), "Word prompt contains Word marker");
            Assert(wordPrompt.Contains("[Source: filename, page/section]"), "Word prompt contains citation instruction");

            // 3. PowerPoint host branch
            string pptPrompt = PromptAssembler.BuildHostAwareSystemPrompt(customBase, "PowerPoint");
            Assert(pptPrompt.StartsWith(customBase), "PowerPoint prompt starts with base prompt");
            Assert(pptPrompt.Contains("You are embedded inside Microsoft PowerPoint."), "PowerPoint prompt contains PowerPoint marker");
            Assert(pptPrompt.Contains("<powerpoint_actions>"), "PowerPoint prompt contains <powerpoint_actions> schema");

            // 4. Default / Fallback branch
            string defaultPrompt = PromptAssembler.BuildHostAwareSystemPrompt(null, "UnknownHost");
            Assert(defaultPrompt == "You are an expert AI assistant embedded inside Microsoft Office.", "Default fallback prompt matches exactly");
        }

        private static void TestPromptAssemblerContextPermutations()
        {
            string rawPrompt = "Summarize the key findings.";
            string selectedText = "Total failures: 42 in Sector 3";
            string fileContext = "[Document Content: Briefing.docx]:\nComplete review of annual maintenance.";

            // Permutation 1: Selection only
            string promptSel = PromptAssembler.ComposePromptWithContext(rawPrompt, PromptContextScope.Selection, selectedText, fileContext);
            Assert(promptSel.Contains(rawPrompt), "Contains raw prompt");
            Assert(promptSel.Contains("[Selected Context]:\n" + selectedText), "Contains selected context");
            Assert(!promptSel.Contains("[Current File Context]:"), "Does not contain file context in Selection scope");

            // Permutation 2: CurrentFile only
            string promptFile = PromptAssembler.ComposePromptWithContext(rawPrompt, PromptContextScope.CurrentFile, selectedText, fileContext);
            Assert(promptFile.Contains(rawPrompt), "Contains raw prompt");
            Assert(!promptFile.Contains("[Selected Context]:"), "Does not contain selected context in CurrentFile scope");
            Assert(promptFile.Contains("[Current File Context]:\n" + fileContext), "Contains file context");

            // Permutation 3: SelectionAndFile
            string promptBoth = PromptAssembler.ComposePromptWithContext(rawPrompt, PromptContextScope.SelectionAndFile, selectedText, fileContext);
            Assert(promptBoth.Contains("[Selected Context]:\n" + selectedText), "Contains selected context in SelectionAndFile scope");
            Assert(promptBoth.Contains("[Current File Context]:\n" + fileContext), "Contains file context in SelectionAndFile scope");

            // Permutation 4: AttachmentsOnly
            string promptAtt = PromptAssembler.ComposePromptWithContext(rawPrompt, PromptContextScope.AttachmentsOnly, selectedText, fileContext);
            Assert(promptAtt == rawPrompt, "AttachmentsOnly scope keeps raw prompt unchanged without selection/file body");
        }

        private static void TestPromptAssemblerAttachmentCitations()
        {
            string userPrompt = "Analyze the attached safety circular.";
            string extractedAttachmentText = "[Attachment: Circular_2026.pdf (Page 1)]\nAll sub-stations must inspect relays monthly.";

            string augmented = PromptAssembler.AppendAttachmentCitationInstruction(userPrompt, extractedAttachmentText);
            Assert(augmented.StartsWith(userPrompt), "Augmented prompt starts with user prompt");
            Assert(augmented.Contains(extractedAttachmentText), "Contains attachment text");
            Assert(augmented.Contains("[Source: filename, page/section]"), "Contains strict citation requirement");
        }

        private static void TestExcelActionParserFullCatalog()
        {
            string responseWith13Actions =
                "Here is the comprehensive spreadsheet update plan:\n\n" +
                "<excel_actions>\n" +
                "  <excel_action target=\"A1\" type=\"value\" value=\"Revenue Analysis\" description=\"Title header\" />\n" +
                "  <excel_action target=\"B2\" type=\"formula\" formula=\"=SUM(B3:B20)\" description=\"Total sum\" />\n" +
                "  <excel_action target=\"C2:C20\" type=\"filldown\" formula=\"=B2*1.18\" description=\"Tax inclusive\" />\n" +
                "  <excel_action target=\"A3:D20\" type=\"table\" value=\"ColA\tColB\tColC\tColD\n1\t2\t3\t4\" description=\"Data table\" />\n" +
                "  <excel_action target=\"A3:D20\" type=\"create_table\" value=\"SalesTable\" description=\"Table description\" />\n" +
                "  <excel_action target=\"B2:B20\" type=\"conditional_format\" value=\"gt:1000\" description=\"Highlight large values\" />\n" +
                "  <excel_action target=\"A2:D20\" type=\"sort\" value=\"descending\" description=\"Sort by Revenue\" />\n" +
                "  <excel_action target=\"A2:D20\" type=\"filter\" value=\"ColB:>100\" description=\"Filter active\" />\n" +
                "  <excel_action target=\"E2:E20\" type=\"data_validation\" value=\"list:Open,Closed,Pending\" description=\"Status dropdown\" />\n" +
                "  <excel_action target=\"A2:B10\" type=\"chart\" value=\"column\" description=\"Revenue Column Chart\" />\n" +
                "  <excel_action target=\"A2:D50\" type=\"pivot_table\" value=\"rows:Region;vals:Sales\" description=\"Summary Pivot\" />\n" +
                "  <excel_action target=\"B2:B20\" type=\"named_range\" value=\"MonthlyRevenue\" description=\"Named range definition\" />\n" +
                "  <excel_action target=\"A2:D50\" type=\"remove_duplicates\" value=\"columns:1,2\" description=\"Deduplicate entries\" />\n" +
                "</excel_actions>\n\nSummary of changes completed.";

            string cleanContent;
            List<SpreadsheetAction> actions = SpreadsheetActionParser.ExtractActions(responseWith13Actions, out cleanContent);

            Assert(actions != null, "Extracted actions is not null");
            Assert(actions.Count == 13, string.Format("Parsed exactly 13 actions (got {0})", actions != null ? actions.Count : 0));
            Assert(cleanContent.Contains("Here is the comprehensive spreadsheet update plan:"), "Clean content preserves preamble");
            Assert(cleanContent.Contains("Summary of changes completed."), "Clean content preserves postamble");
            Assert(!cleanContent.Contains("<excel_actions>"), "Clean content strips XML block");

            // Verify action type mappings
            Assert(actions[0].Type == SpreadsheetActionType.Value && actions[0].Target == "A1", "Action 0 is Value at A1");
            Assert(actions[1].Type == SpreadsheetActionType.Formula && actions[1].Target == "B2", "Action 1 is Formula at B2");
            Assert(actions[2].Type == SpreadsheetActionType.FillDown && actions[2].Target == "C2:C20", "Action 2 is FillDown");
            Assert(actions[3].Type == SpreadsheetActionType.Table, "Action 3 is Table");
            Assert(actions[4].Type == SpreadsheetActionType.CreateTable && !actions[4].IsUndoable, "Action 4 is CreateTable (non-undoable)");
            Assert(actions[5].Type == SpreadsheetActionType.ConditionalFormat, "Action 5 is ConditionalFormat");
            Assert(actions[6].Type == SpreadsheetActionType.Sort, "Action 6 is Sort");
            Assert(actions[7].Type == SpreadsheetActionType.Filter, "Action 7 is Filter");
            Assert(actions[8].Type == SpreadsheetActionType.DataValidation, "Action 8 is DataValidation");
            Assert(actions[9].Type == SpreadsheetActionType.Chart && !actions[9].IsUndoable, "Action 9 is Chart (non-undoable)");
            Assert(actions[10].Type == SpreadsheetActionType.PivotTable && !actions[10].IsUndoable, "Action 10 is PivotTable (non-undoable)");
            Assert(actions[11].Type == SpreadsheetActionType.NamedRange && !actions[11].IsUndoable, "Action 11 is NamedRange (non-undoable)");
            Assert(actions[12].Type == SpreadsheetActionType.RemoveDuplicates && !actions[12].IsUndoable, "Action 12 is RemoveDuplicates (non-undoable)");
        }

        private static void TestExcelActionParserSafetyBoundaries()
        {
            // 1. Unbounded column B:B must be rejected
            string unboundedColXml = "<excel_actions><excel_action target=\"B:B\" type=\"formula\">=SUM(A:A)</excel_action></excel_actions>";
            string dummy;
            List<SpreadsheetAction> unbounded = SpreadsheetActionParser.ExtractActions(unboundedColXml, out dummy);
            Assert(unbounded.Count == 0, "Unbounded column B:B is safely rejected");

            // 2. Sheet-qualified target Sheet1!A1 must be rejected
            string sheetQualifiedXml = "<excel_actions><excel_action target=\"Sheet1!A1\" type=\"value\">42</excel_action></excel_actions>";
            List<SpreadsheetAction> sheetQual = SpreadsheetActionParser.ExtractActions(sheetQualifiedXml, out dummy);
            Assert(sheetQual.Count == 0, "Sheet-qualified address Sheet1!A1 is safely rejected");

            // 3. Multi-area comma separated range A1,B2 must be rejected
            string multiAreaXml = "<excel_actions><excel_action target=\"A1,B2\" type=\"value\">42</excel_action></excel_actions>";
            List<SpreadsheetAction> multiArea = SpreadsheetActionParser.ExtractActions(multiAreaXml, out dummy);
            Assert(multiArea.Count == 0, "Multi-area range A1,B2 is safely rejected");
        }

        private static void TestPowerPointActionParserFullCatalog()
        {
            string pptXml =
                "Suggested presentation deck structural updates:\n\n" +
                "<powerpoint_actions>\n" +
                "  <powerpoint_action type=\"move_slide\" source=\"4\" target=\"1\" />\n" +
                "  <powerpoint_action type=\"create_section\" name=\"Executive Briefing\" slide=\"1\" />\n" +
                "  <powerpoint_action type=\"rename_section\" section=\"2\" name=\"Technical Details\" />\n" +
                "  <powerpoint_action type=\"set_notes\" slide=\"3\" notes=\"Focus on Q3 turnaround timeline\" />\n" +
                "</powerpoint_actions>\n\nReview the proposed structure.";

            string cleanContent;
            List<PowerPointAction> actions = PowerPointActionParser.ParseStructuredActions(pptXml, out cleanContent);

            Assert(actions != null && actions.Count == 4, "Parsed 4 PowerPoint actions");
            Assert(!cleanContent.Contains("<powerpoint_actions>"), "Clean content stripped action XML");
            Assert(cleanContent.Contains("Suggested presentation deck structural updates:"), "Preserved text");

            Assert(actions[0].Type == "move_slide" && actions[0].Source == 4 && actions[0].Target == 1, "Action 0 is move_slide");
            Assert(actions[0].TypeBadge == "move" && actions[0].TargetDisplay == "Slide 4 → 1", "Action 0 badges match");

            Assert(actions[1].Type == "create_section" && actions[1].Name == "Executive Briefing" && actions[1].Slide == 1, "Action 1 is create_section");
            Assert(actions[1].TypeBadge == "section+" && actions[1].ContentDisplay == "Executive Briefing", "Action 1 badges match");

            Assert(actions[2].Type == "rename_section" && actions[2].Section == 2 && actions[2].Name == "Technical Details", "Action 2 is rename_section");
            Assert(actions[2].TypeBadge == "section", "Action 2 badge is section");

            Assert(actions[3].Type == "set_notes" && actions[3].Slide == 3 && actions[3].Notes == "Focus on Q3 turnaround timeline", "Action 3 is set_notes");
            Assert(actions[3].TypeBadge == "notes" && actions[3].TargetDisplay == "Slide 3", "Action 3 badges match");
        }

        private static void TestPowerPointSlideDataParsing()
        {
            string slideMarkdown =
                "# Slide 1: High-Speed Rail Modernization\n" +
                "- 25kV Traction Power Upgrades\n" +
                "- Substation telemetry integration\n" +
                "- SCADA control center modernization\n\n" +
                "Speaker Notes: Highlight completion within scheduled Q2 window.\n" +
                "Visual: Single line diagram of traction substation.\n\n" +
                "# Slide 2: Implementation Roadmap\n" +
                "- Phase 1: Survey and Civil works\n" +
                "- Phase 2: OHE stringing and testing\n" +
                "Notes: Reiterate safety clearance protocols.";

            List<SlideData> slides = PowerPointActionParser.ParseSlideData(slideMarkdown);
            Assert(slides != null && slides.Count == 2, "Parsed 2 slide data blocks");

            Assert(slides[0].Title == "High-Speed Rail Modernization", "Slide 1 title");
            Assert(slides[0].Bullets.Count == 3, "Slide 1 has 3 bullets");
            Assert(slides[0].SpeakerNotes == "Highlight completion within scheduled Q2 window.", "Slide 1 speaker notes");
            Assert(slides[0].VisualSuggestion == "Single line diagram of traction substation.", "Slide 1 visual suggestion");

            Assert(slides[1].Title == "Implementation Roadmap", "Slide 2 title");
            Assert(slides[1].Bullets.Count == 2, "Slide 2 has 2 bullets");
            Assert(slides[1].SpeakerNotes == "Reiterate safety clearance protocols.", "Slide 2 notes");
        }

        private static void TestActionAuditStoreSerializationBaseline()
        {
            string tempFile = Path.Combine(Path.GetTempPath(), "test-action-audit-gm-" + Guid.NewGuid().ToString("N") + ".dat");
            try
            {
                var auditStore = new ActionAuditStore(tempFile);
                var entries = auditStore.GetRecent(10);
                Assert(entries.Count == 0, "Initial audit store is empty");

                auditStore.Record(
                    "Excel",
                    "formula",
                    "K20",
                    "Calculate total non-ferrous value",
                    true,
                    "Sum non-ferrous items",
                    "Book1.xlsx",
                    "Mistral Large",
                    "=SUMIF(B2:B100, \"*non ferrous*\", F2:F100)",
                    "Applied successfully");

                var loaded = auditStore.GetRecent(10);
                Assert(loaded.Count == 1, "Recorded and loaded 1 audit entry");
                Assert(loaded[0].Host == "Excel", "Host is Excel");
                Assert(loaded[0].ActionType == "formula", "ActionType is formula");
                Assert(loaded[0].Target == "K20", "Target is K20");
                Assert(loaded[0].Undoable == true, "Undoable is true");
                Assert(loaded[0].Model == "Mistral Large", "Model is Mistral Large");
                Assert(loaded[0].ApplyResult == "Applied successfully", "ApplyResult matches");
            }
            finally
            {
                if (File.Exists(tempFile))
                {
                    try { File.Delete(tempFile); } catch { }
                }
            }
        }

        private static void Assert(bool condition, string message)
        {
            if (!condition)
            {
                throw new Exception("Golden Master Assertion Failed: " + message);
            }
        }
    }
}
