using System;
using System.Collections.Generic;
using MistralOfficeAddin.Core;

namespace MSOfficeAIAssistant.Tests
{
    public static class ActionAuditStoreTests
    {
        public static void RunAll()
        {
            TestRecordAndRetrieve();
            TestTruncation();
        }

        private static void TestRecordAndRetrieve()
        {
            string tempFile = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "test-action-audit-" + Guid.NewGuid().ToString("N") + ".dat");
            try
            {
                var store = new ActionAuditStore(tempFile);
                string host = "Excel";
                string actionType = "formula";
                string target = "B2:B10";
                string summary = "Calculate sum for column B";
                bool undoable = true;
                string prompt = "Calculate the totals in column B";
                string sourceContext = "Sheet1 with headers in Row 1";
                string model = "mistral-large-latest";
                string proposed = "<excel_actions><excel_action target=\"B2:B10\" type=\"formula\" formula=\"=SUM(A2:A10)\"/></excel_actions>";
                string result = "Success (9 cells updated)";

                store.Record(
                    host,
                    actionType,
                    target,
                    summary,
                    undoable,
                    prompt,
                    sourceContext,
                    model,
                    proposed,
                    result);

                List<ActionAuditEntry> recent = store.GetRecent(10);
                Assert(recent != null && recent.Count == 1, "Exactly 1 audit entry retrieved");

                ActionAuditEntry latest = recent[0];
                Assert(latest.Host == host, "Host matches");
                Assert(latest.ActionType == actionType, "ActionType matches");
                Assert(latest.Target == target, "Target matches");
                Assert(latest.Summary == summary, "Summary matches");
                Assert(latest.Undoable == undoable, "Undoable matches");
                Assert(latest.Prompt == prompt, "Prompt matches");
                Assert(latest.SourceContext == sourceContext, "SourceContext matches");
                Assert(latest.Model == model, "Model matches");
                Assert(latest.FullProposedAction == proposed, "FullProposedAction matches");
                Assert(latest.ApplyResult == result, "ApplyResult matches");
            }
            finally
            {
                try { if (System.IO.File.Exists(tempFile)) System.IO.File.Delete(tempFile); } catch { }
            }
        }

        private static void TestTruncation()
        {
            string tempFile = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "test-action-audit-" + Guid.NewGuid().ToString("N") + ".dat");
            try
            {
                var store = new ActionAuditStore(tempFile);
                string longText = new string('A', 3000);
                store.Record("Word", "Insert", "Doc", longText, true);

                List<ActionAuditEntry> recent = store.GetRecent(5);
                Assert(recent != null && recent.Count == 1, "Recent entry found");
                Assert(recent[0].Summary.Length <= 2020, "Summary was truncated to max length");
                Assert(recent[0].Summary.IndexOf("[truncated]", StringComparison.Ordinal) >= 0, "Summary has truncation marker");
            }
            finally
            {
                try { if (System.IO.File.Exists(tempFile)) System.IO.File.Delete(tempFile); } catch { }
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
