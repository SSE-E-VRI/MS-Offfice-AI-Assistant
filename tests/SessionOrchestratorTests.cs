using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using MSOfficeAIAssistant.API;
using MSOfficeAIAssistant.API.Models;
using MSOfficeAIAssistant.Core;
using MSOfficeAIAssistant.Core.Session;

namespace MSOfficeAIAssistant.Tests
{
    public static class SessionOrchestratorTests
    {
        public static void RunAll()
        {
            TestStreamCoordinatorAccumulationAndThrottling();
            TestStreamCoordinatorStreamFinishedGuard();
            TestStreamCoordinatorCancellation();
            TestAssistantSessionPreparePayload();
            TestAssistantSessionProcessExcelResponse();
            TestAssistantSessionProcessPowerPointResponse();
            TestAssistantSessionHistoryRoundtrip();
        }

        private static void TestStreamCoordinatorAccumulationAndThrottling()
        {
            var coordinator = new StreamCoordinator();
            var cts = coordinator.BeginStream();
            Assert(coordinator.IsSending, "Coordinator should be sending after BeginStream");
            Assert(!coordinator.IsStreamFinished, "Coordinator stream should not be finished");

            var throttledSnapshots = new List<string>();
            Action<string> onUpdate = snapshot => throttledSnapshots.Add(snapshot);

            // Send 12 deltas with throttle step 5 -> updates at 5 and 10
            for (int i = 1; i <= 12; i++)
            {
                coordinator.AccumulateDelta("t" + i + " ", onUpdate, throttleStep: 5);
            }

            Assert(throttledSnapshots.Count == 2, "Expected 2 throttled snapshots at delta 5 and 10, got " + throttledSnapshots.Count);
            Assert(throttledSnapshots[0] == "t1 t2 t3 t4 t5 ", "Snapshot 1 mismatch: " + throttledSnapshots[0]);
            Assert(throttledSnapshots[1] == "t1 t2 t3 t4 t5 t6 t7 t8 t9 t10 ", "Snapshot 2 mismatch: " + throttledSnapshots[1]);

            string full = coordinator.FinishStream();
            Assert(coordinator.IsStreamFinished, "IsStreamFinished should be true after FinishStream");
            Assert(full == "t1 t2 t3 t4 t5 t6 t7 t8 t9 t10 t11 t12 ", "Final text mismatch: " + full);

            coordinator.EndSending();
            Assert(!coordinator.IsSending, "IsSending should be false after EndSending");
        }

        private static void TestStreamCoordinatorStreamFinishedGuard()
        {
            var coordinator = new StreamCoordinator();
            coordinator.BeginStream();

            int updateCount = 0;
            Action<string> onUpdate = s => updateCount++;

            coordinator.AccumulateDelta("Part 1", onUpdate, 1);
            Assert(updateCount == 1, "Expected 1 update");

            string final = coordinator.FinishStream();
            Assert(final == "Part 1", "Final mismatch");

            // Late delta after stream finished should be ignored
            coordinator.AccumulateDelta("Late arrival", onUpdate, 1);
            Assert(updateCount == 1, "Late delta should not trigger throttled updates");
            Assert(coordinator.FinishStream() == "Part 1", "Late delta should not alter final accumulator");

            coordinator.EndSending();
        }

        private static void TestStreamCoordinatorCancellation()
        {
            var coordinator = new StreamCoordinator();
            var cts = coordinator.BeginStream();
            Assert(!cts.IsCancellationRequested, "CTS should not be cancelled initially");

            coordinator.Cancel();
            Assert(cts.IsCancellationRequested, "CTS should be cancelled after coordinator.Cancel()");

            coordinator.EndSending();
        }

        private static void TestAssistantSessionPreparePayload()
        {
            var messages = new ObservableCollection<ChatMessage>();
            messages.Add(new ChatMessage("user", "Summarize Q3 earnings") { FullContent = "Summarize Q3 earnings with details" });
            var assistantMsg = new ChatMessage("assistant", "") { IsStreaming = true };
            messages.Add(assistantMsg);

            var session = new AssistantSession(null, messages);
            session.HostType = "Word";

            var task = session.PreparePayloadAsync("gpt-4o", null, assistantMsg);
            task.Wait();
            var prepared = task.Result;

            Assert(prepared.Request != null, "Prepared request should not be null");
            Assert(prepared.Request.Model == "gpt-4o", "Model mismatch");
            Assert(prepared.Request.Messages.Count == 2, "Expected system + user message in payload, got " + prepared.Request.Messages.Count);
            Assert(prepared.Request.Messages[0].IsSystem, "First message should be system message");
            Assert(prepared.Request.Messages[1].Content == "Summarize Q3 earnings with details", "User full content should be preserved in payload");
            Assert(prepared.EffectiveSystemPrompt.Contains("Microsoft Word"), "System prompt should contain host rules for Word");
        }

        private static void TestAssistantSessionProcessExcelResponse()
        {
            var messages = new ObservableCollection<ChatMessage>();
            var session = new AssistantSession(null, messages);
            session.HostType = "Excel";

            var assistantMsg = new ChatMessage("assistant", "") { IsStreaming = true };
            string rawResponse = "Here are the formulas:\n<excel_actions>\n  <excel_action target=\"B2\" type=\"formula\" formula=\"=SUM(B3:B10)\" description=\"Total sum\" />\n</excel_actions>\nDone.";

            session.ProcessAssistantResponse(rawResponse, assistantMsg);

            Assert(!assistantMsg.IsStreaming, "IsStreaming should be false");
            Assert(!assistantMsg.Content.Contains("<excel_actions>"), "Prose should be cleaned of XML tags");
            Assert(assistantMsg.Content.Contains("Here are the formulas:"), "Prose should contain introductory text");
            Assert(assistantMsg.HasOfficeActions, "HasOfficeActions should be true");
            Assert(assistantMsg.OfficeActions.Count == 1, "Expected 1 parsed Excel action, got " + assistantMsg.OfficeActions.Count);
            Assert(assistantMsg.OfficeActions[0].TargetDisplay == "B2", "Action target mismatch");
            Assert(assistantMsg.OfficeActions[0].Operation == "excel.write_formula", "Action operation mismatch");
            Assert(assistantMsg.OfficeActions[0].ContentDisplay == "=SUM(B3:B10)", "Action content mismatch");
        }

        private static void TestAssistantSessionProcessPowerPointResponse()
        {
            var messages = new ObservableCollection<ChatMessage>();
            var session = new AssistantSession(null, messages);
            session.HostType = "PowerPoint";

            var assistantMsg = new ChatMessage("assistant", "") { IsStreaming = true };
            string rawResponse = "Slide deck reorganization:\n<powerpoint_actions>\n  <powerpoint_action type=\"move_slide\" source=\"3\" target=\"1\" />\n</powerpoint_actions>\nEnjoy.";

            session.ProcessAssistantResponse(rawResponse, assistantMsg);

            Assert(!assistantMsg.IsStreaming, "IsStreaming should be false");
            Assert(!assistantMsg.Content.Contains("<powerpoint_actions>"), "Prose should be cleaned of XML tags");
            Assert(assistantMsg.Content.Contains("Slide deck reorganization:"), "Prose should contain introductory text");
            Assert(assistantMsg.HasOfficeActions, "HasOfficeActions should be true");
            Assert(assistantMsg.OfficeActions.Count == 1, "Expected 1 parsed PowerPoint action, got " + assistantMsg.OfficeActions.Count);
            Assert(assistantMsg.OfficeActions[0].Operation == "powerpoint.move_slide", "Action operation mismatch");
            Assert(Convert.ToString(assistantMsg.OfficeActions[0].Parameters["source"]) == "3", "Source slide mismatch");
            Assert(Convert.ToString(assistantMsg.OfficeActions[0].Parameters["target"]) == "1", "Target position mismatch");
        }

        private static void TestAssistantSessionHistoryRoundtrip()
        {
            string testKey = "TestDocSession_" + Guid.NewGuid().ToString("N");
            var messages = new ObservableCollection<ChatMessage>();
            messages.Add(new ChatMessage("user", "Hello Assistant"));
            messages.Add(new ChatMessage("assistant", "Hello User"));

            var session = new AssistantSession(null, messages);
            session.CurrentDocumentKey = testKey;
            session.SaveHistory();

            var loadedSession = new AssistantSession();
            loadedSession.LoadHistory(testKey);

            Assert(loadedSession.Messages.Count == 2, "Expected 2 loaded messages, got " + loadedSession.Messages.Count);
            Assert(loadedSession.Messages[0].Content == "Hello Assistant", "User content mismatch");
            Assert(loadedSession.Messages[1].Content == "Hello User", "Assistant content mismatch");
        }

        private static void Assert(bool condition, string message)
        {
            if (!condition)
            {
                throw new Exception(message);
            }
        }
    }
}
