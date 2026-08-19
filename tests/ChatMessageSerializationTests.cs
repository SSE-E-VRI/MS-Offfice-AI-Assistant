using System;
using System.Collections.Generic;
using MistralOfficeAddin.API.Models;
using MistralOfficeAddin.Core;
using Newtonsoft.Json;

namespace MSOfficeAIAssistant.Tests
{
    public static class ChatMessageSerializationTests
    {
        public static void RunAll()
        {
            TestChatMessageActionsSerialization();
            TestConversationStoreHistoryPersistence();
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
