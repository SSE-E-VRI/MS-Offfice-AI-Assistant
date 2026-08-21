using System;
using System.Collections.Generic;
using MSOfficeAIAssistant.API.Models;
using MSOfficeAIAssistant.Core;
using MSOfficeAIAssistant.Core.Actions;
using Newtonsoft.Json;

namespace MSOfficeAIAssistant.Tests
{
    public static class ChatMessageSerializationTests
    {
        public static void RunAll()
        {
            TestChatMessageActionsSerialization();
            TestConversationStoreHistoryPersistence();
            TestLegacySpreadsheetActionsHydrateOfficeActions();
            TestLegacyPowerPointActionsHydrateOfficeActions();
            TestOfficeActionsDirectSerialization();
        }

        private static void TestLegacySpreadsheetActionsHydrateOfficeActions()
        {
            // Simulates reading legacy JSON from disk (no officeActions property, only legacy actions)
            string legacyJson =
                "{\n" +
                "  \"role\": \"assistant\",\n" +
                "  \"content\": \"Here are the requested formulas:\",\n" +
                "  \"actions\": [\n" +
                "    {\n" +
                "      \"target\": \"B2:B20\",\n" +
                "      \"type\": 0,\n" + // Formula
                "      \"content\": \"=SUM(A2:A20)\",\n" +
                "      \"description\": \"Total sales sum\",\n" +
                "      \"status\": 0\n" + // Pending
                "    },\n" +
                "    {\n" +
                "      \"target\": \"A1:D50\",\n" +
                "      \"type\": 12,\n" + // RemoveDuplicates
                "      \"content\": \"columns:1,2\",\n" +
                "      \"description\": \"Deduplicate entries\",\n" +
                "      \"status\": 2\n" + // Applied
                "    }\n" +
                "  ]\n" +
                "}";

            ChatMessage deserialized = JsonConvert.DeserializeObject<ChatMessage>(legacyJson);
            Assert(deserialized != null, "Deserialized message is not null");
            Assert(deserialized.HasOfficeActions == true, "HasOfficeActions is true after legacy hydration");
            Assert(deserialized.OfficeActions != null && deserialized.OfficeActions.Count == 2, "OfficeActions hydrated 2 actions");

            var a0 = deserialized.OfficeActions[0];
            Assert(a0.TargetDisplay == "B2:B20", "Action 0 target display matches");
            Assert(a0.Operation == "excel.write_formula", "Action 0 operation mapped to excel.write_formula");
            Assert(a0.ContentDisplay == "=SUM(A2:A20)", "Action 0 formula preserved in ContentDisplay");
            Assert(a0.Status == OfficeActionStatus.Pending, "Action 0 status is Pending");
            Assert(a0.IsUndoable == true, "Action 0 is undoable");

            var a1 = deserialized.OfficeActions[1];
            Assert(a1.TargetDisplay == "A1:D50", "Action 1 target display matches");
            Assert(a1.Operation == "excel.remove_duplicates", "Action 1 operation mapped to excel.remove_duplicates");
            Assert(a1.Status == OfficeActionStatus.Applied, "Action 1 status is Applied");
            Assert(a1.IsUndoable == false, "Action 1 is not undoable");
        }

        private static void TestLegacyPowerPointActionsHydrateOfficeActions()
        {
            // Simulates reading legacy PowerPoint JSON from disk
            string legacyPptJson =
                "{\n" +
                "  \"role\": \"assistant\",\n" +
                "  \"content\": \"Reorganized deck slides:\",\n" +
                "  \"powerPointActions\": [\n" +
                "    {\n" +
                "      \"type\": \"move_slide\",\n" +
                "      \"source\": 4,\n" +
                "      \"target\": 1,\n" +
                "      \"description\": \"Move appendix slide to intro\",\n" +
                "      \"status\": 0\n" +
                "    },\n" +
                "    {\n" +
                "      \"type\": \"set_notes\",\n" +
                "      \"slide\": 2,\n" +
                "      \"notes\": \"Executive briefing remarks\",\n" +
                "      \"description\": \"Add notes to slide 2\",\n" +
                "      \"status\": 2\n" +
                "    }\n" +
                "  ]\n" +
                "}";

            ChatMessage deserialized = JsonConvert.DeserializeObject<ChatMessage>(legacyPptJson);
            Assert(deserialized != null, "Deserialized message is not null");
            Assert(deserialized.HasOfficeActions == true, "HasOfficeActions is true after PPT hydration");
            Assert(deserialized.OfficeActions != null && deserialized.OfficeActions.Count == 2, "OfficeActions hydrated 2 PPT actions");

            var p0 = deserialized.OfficeActions[0];
            Assert(p0.Host == "PowerPoint", "PPT action 0 host is PowerPoint");
            Assert(p0.Operation == "powerpoint.move_slide", "PPT action 0 operation is powerpoint.move_slide");
            Assert(p0.Status == OfficeActionStatus.Pending, "PPT action 0 status is Pending");
            Assert(p0.Parameters.ContainsKey("source") && Convert.ToString(p0.Parameters["source"]) == "4", "PPT action 0 source parameter preserved");
            Assert(p0.Parameters.ContainsKey("target") && Convert.ToString(p0.Parameters["target"]) == "1", "PPT action 0 target parameter preserved");

            var p1 = deserialized.OfficeActions[1];
            Assert(p1.Operation == "powerpoint.set_notes", "PPT action 1 operation is powerpoint.set_notes");
            Assert(p1.Status == OfficeActionStatus.Applied, "PPT action 1 status is Applied");
            Assert(p1.ContentDisplay == "Executive briefing remarks", "PPT action 1 notes preserved in ContentDisplay");
        }

        private static void TestOfficeActionsDirectSerialization()
        {
            ChatMessage msg = new ChatMessage("assistant", "Native OfficeAction block:");
            var act = new OfficeAction
            {
                Host = "Word",
                Operation = "word.add_comment",
                Status = OfficeActionStatus.Pending,
                IsUndoable = true,
                SourceReason = "Review comment on introduction"
            };
            act.Target.Range = "Paragraph 3";
            act.Parameters["comment_text"] = "Verify Q3 metrics citation.";
            msg.OfficeActions.Add(act);

            string json = JsonConvert.SerializeObject(msg, Formatting.Indented);
            Assert(json.IndexOf("\"officeActions\"", StringComparison.Ordinal) >= 0, "JSON contains officeActions property");

            ChatMessage roundTripped = JsonConvert.DeserializeObject<ChatMessage>(json);
            Assert(roundTripped != null, "Roundtripped message is not null");
            Assert(roundTripped.HasOfficeActions == true, "HasOfficeActions is true");
            Assert(roundTripped.OfficeActions.Count == 1, "OfficeActions count is 1");
            Assert(roundTripped.OfficeActions[0].Host == "Word", "Host is Word");
            Assert(roundTripped.OfficeActions[0].Operation == "word.add_comment", "Operation is word.add_comment");
            Assert(roundTripped.OfficeActions[0].ContentDisplay == "Verify Q3 metrics citation.", "ContentDisplay matches");
        }

        private static void TestChatMessageActionsSerialization()
        {
            ChatMessage msg = new ChatMessage("assistant", "Here is the spreadsheet formula:");
            msg.Actions.Add(new SpreadsheetAction
            {
                Target = "K20",
                Type = SpreadsheetActionType.Formula,
                Content = "=SUM(B2:B100)",
                Description = "Total sum",
                Status = SpreadsheetActionStatus.Pending
            });
            msg.Actions.Add(new SpreadsheetAction
            {
                Target = "K21",
                Type = SpreadsheetActionType.RemoveDuplicates,
                Content = "columns:1",
                Description = "Remove duplicates",
                Status = SpreadsheetActionStatus.Applied
            });

            string json = JsonConvert.SerializeObject(msg, Formatting.Indented);
            Assert(json.IndexOf("\"actions\"", StringComparison.Ordinal) >= 0, "JSON contains actions property");
            Assert(json.IndexOf("=SUM(B2:B100)", StringComparison.Ordinal) >= 0, "JSON contains action formula");

            ChatMessage deserialized = JsonConvert.DeserializeObject<ChatMessage>(json);
            Assert(deserialized != null, "Deserialized message is not null");
            Assert(deserialized.Role == "assistant", "Role is assistant");
            Assert(deserialized.Content == "Here is the spreadsheet formula:", "Content matches");
            Assert(deserialized.Actions != null && deserialized.Actions.Count == 2, "Deserialized actions count is 2");
            Assert(deserialized.Actions[0].Target == "K20", "Action 0 target K20");
            Assert(deserialized.Actions[0].Content == "=SUM(B2:B100)", "Action 0 content");
            Assert(deserialized.Actions[0].Status == SpreadsheetActionStatus.Pending, "Action 0 status pending");
            Assert(deserialized.Actions[0].IsUndoable == true, "Action 0 is undoable");
            Assert(deserialized.Actions[1].Status == SpreadsheetActionStatus.Applied, "Action 1 status applied");
            Assert(deserialized.Actions[1].IsUndoable == false, "Action 1 is not undoable");
        }

        private static void TestConversationStoreHistoryPersistence()
        {
            string tempDir = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "test-conv-" + Guid.NewGuid().ToString("N"));
            try
            {
                var store = new ConversationStore(tempDir);
                string docKey = "UnitTestDoc_" + Guid.NewGuid().ToString("N");
                List<ChatMessage> history = new List<ChatMessage>();

                ChatMessage userMsg = new ChatMessage("user", "Please calculate totals.");
                ChatMessage assistantMsg = new ChatMessage("assistant", "Here is the total:");
                assistantMsg.Actions.Add(new SpreadsheetAction
                {
                    Target = "C10",
                    Type = SpreadsheetActionType.Formula,
                    Content = "=SUM(C2:C9)",
                    Description = "Calculate column C sum"
                });

                history.Add(userMsg);
                history.Add(assistantMsg);

                store.SaveHistory(docKey, history);

                // Clear memory cache so GetHistory MUST read and decrypt from disk via DPAPI
                store.ClearMemoryCache();

                List<ChatMessage> loaded = store.GetHistory(docKey);
                Assert(loaded != null && loaded.Count == 2, "Loaded history count is 2 from disk");
                Assert(loaded[0].Content == "Please calculate totals.", "User message loaded");
                Assert(loaded[1].Actions != null && loaded[1].Actions.Count == 1, "Assistant action loaded");
                Assert(loaded[1].Actions[0].Target == "C10", "Action target preserved across reload");
                Assert(loaded[1].Actions[0].Content == "=SUM(C2:C9)", "Action formula preserved across reload");

                // Also test fresh store instance reading the exact same disk storage
                var freshStore = new ConversationStore(tempDir);
                List<ChatMessage> fromFreshStore = freshStore.GetHistory(docKey);
                Assert(fromFreshStore != null && fromFreshStore.Count == 2, "Fresh store instance loaded from disk");
                Assert(fromFreshStore[1].Actions[0].Target == "C10", "Fresh store action target matches");

                store.ClearHistory(docKey);
                List<ChatMessage> cleared = store.GetHistory(docKey);
                Assert(cleared == null || cleared.Count == 0, "History cleared");
            }
            finally
            {
                try { if (System.IO.Directory.Exists(tempDir)) System.IO.Directory.Delete(tempDir, true); } catch { }
            }
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
