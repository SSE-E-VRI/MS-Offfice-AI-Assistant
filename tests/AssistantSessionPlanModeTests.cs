using System;
using System.Collections.Generic;
using MSOfficeAIAssistant.API.Models;
using MSOfficeAIAssistant.Core;
using MSOfficeAIAssistant.Core.Actions;
using MSOfficeAIAssistant.Core.Session;

namespace MSOfficeAIAssistant.Tests
{
    public static class AssistantSessionPlanModeTests
    {
        public static void RunAll()
        {
            TestPlanModeNoActions();
            TestPlanModeWithActions();
            TestEditModeWithActions();
            TestChatModeWithActions();
        }

        private static void TestPlanModeNoActions()
        {
            // Test: Mode == Plan with plain text (no action markup)
            var session = new AssistantSession();
            session.Mode = SessionMode.Plan;
            session.HostType = "Excel";

            var assistantMsg = new ChatMessage("assistant", "");

            string plainText = "This is just a plain response with no actions.";
            session.ProcessAssistantResponse(plainText, assistantMsg);

            Assert(assistantMsg.Plan == null, "Plan should be null for plain text in Plan mode");
            Assert(assistantMsg.HasPlan == false, "HasPlan should be false for plain text in Plan mode");
            Assert(assistantMsg.Content == plainText, "Content should be set to the plain text");
            Assert(assistantMsg.HasOfficeActions == false, "OfficeActions should not be populated in Plan mode");
        }

        private static void TestPlanModeWithActions()
        {
            // Test: Mode == Plan with action XML
            var session = new AssistantSession();
            session.Mode = SessionMode.Plan;
            session.HostType = "Excel";

            var assistantMsg = new ChatMessage("assistant", "");

            // Sample action XML from ActionExtractorTests
            string withActions =
                "Here are the spreadsheet changes:\n\n" +
                "<excel_actions>\n" +
                "  <excel_action target=\"A1\" type=\"value\" value=\"Revenue Summary\" description=\"Header cell\" />\n" +
                "  <excel_action target=\"B2:B10\" type=\"fill_down\" formula=\"=A2*1.15\" description=\"Fill formula\" />\n" +
                "</excel_actions>\n\n" +
                "Please review.";

            session.ProcessAssistantResponse(withActions, assistantMsg);

            Assert(assistantMsg.Plan != null, "Plan should be created in Plan mode with actions");
            Assert(assistantMsg.HasPlan == true, "HasPlan should be true");
            Assert(assistantMsg.Plan.Steps.Count == 2, "Plan should have 2 steps matching the 2 actions");
            Assert(!assistantMsg.Content.Contains("<excel_actions>"), "Content should be cleaned text without action tags");
            Assert(assistantMsg.Content.Contains("Please review"), "Content should retain the prose text");
            Assert(assistantMsg.HasOfficeActions == false, "OfficeActions collection should remain empty in Plan mode");
        }

        private static void TestEditModeWithActions()
        {
            // Test: Mode == Edit with action XML (existing behavior should be unchanged)
            var session = new AssistantSession();
            session.Mode = SessionMode.Edit;
            session.HostType = "Excel";

            var assistantMsg = new ChatMessage("assistant", "");

            string withActions =
                "Here are the spreadsheet changes:\n\n" +
                "<excel_actions>\n" +
                "  <excel_action target=\"A1\" type=\"value\" value=\"Revenue Summary\" description=\"Header cell\" />\n" +
                "  <excel_action target=\"B2:B10\" type=\"fill_down\" formula=\"=A2*1.15\" description=\"Fill formula\" />\n" +
                "</excel_actions>\n\n" +
                "Please review.";

            session.ProcessAssistantResponse(withActions, assistantMsg);

            Assert(assistantMsg.Plan == null, "Plan should be null in Edit mode");
            Assert(assistantMsg.HasPlan == false, "HasPlan should be false in Edit mode");
            Assert(assistantMsg.HasOfficeActions == true, "OfficeActions should be populated in Edit mode");
            Assert(assistantMsg.OfficeActions.Count == 2, "Should have 2 actions attached");
            Assert(!assistantMsg.Content.Contains("<excel_actions>"), "Content should be cleaned text");
        }

        private static void TestChatModeWithActions()
        {
            // Test: Mode == Chat with action XML (existing behavior should be unchanged)
            var session = new AssistantSession();
            session.Mode = SessionMode.Chat;
            session.HostType = "Excel";

            var assistantMsg = new ChatMessage("assistant", "");

            string withActions =
                "Here are the spreadsheet changes:\n\n" +
                "<excel_actions>\n" +
                "  <excel_action target=\"A1\" type=\"value\" value=\"Revenue Summary\" description=\"Header cell\" />\n" +
                "</excel_actions>\n\n" +
                "Please review.";

            session.ProcessAssistantResponse(withActions, assistantMsg);

            Assert(assistantMsg.Plan == null, "Plan should be null in Chat mode");
            Assert(assistantMsg.HasPlan == false, "HasPlan should be false in Chat mode");
            Assert(assistantMsg.HasOfficeActions == true, "OfficeActions should be populated in Chat mode (same as Edit)");
            Assert(assistantMsg.OfficeActions.Count == 1, "Should have 1 action attached");
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
