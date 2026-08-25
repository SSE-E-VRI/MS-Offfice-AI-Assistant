using System;
using System.Collections.ObjectModel;
using MSOfficeAIAssistant.Core.Actions;
using MSOfficeAIAssistant.Core.Session;

namespace MSOfficeAIAssistant.Tests
{
    public static class AssistantSessionModeTests
    {
        public static void RunAll()
        {
            TestDefaultModeIsEdit();
            TestIsActionAllowedNullAction();
            TestChatModeReadOnlyAllowed();
            TestChatModeRiskLevel1Blocked();
            TestChatModeRiskLevel3Blocked();
            TestPlanModeHighRiskAllowed();
            TestEditModeHighRiskAllowed();
        }

        private static void TestDefaultModeIsEdit()
        {
            var session = new AssistantSession();
            Assert(session.Mode == SessionMode.Edit, "Default mode should be Edit");
        }

        private static void TestIsActionAllowedNullAction()
        {
            var session = new AssistantSession();
            session.Mode = SessionMode.Chat;
            var result = session.IsActionAllowed(null);
            Assert(result.Allowed == true, "Null action should be allowed");
            Assert(result.Reason == null, "Null action should have no reason");
        }

        private static void TestChatModeReadOnlyAllowed()
        {
            var session = new AssistantSession();
            session.Mode = SessionMode.Chat;
            var action = new OfficeAction { Host = "Excel", Operation = "excel.read_data", RiskLevel = 0 };
            var result = session.IsActionAllowed(action);
            Assert(result.Allowed == true, "Chat mode should allow RiskLevel 0 (read-only)");
            Assert(result.Reason == null, "Allowed action should have no reason");
        }

        private static void TestChatModeRiskLevel1Blocked()
        {
            var session = new AssistantSession();
            session.Mode = SessionMode.Chat;
            var action = new OfficeAction { Host = "Excel", Operation = "excel.write_value", RiskLevel = 1 };
            var result = session.IsActionAllowed(action);
            Assert(result.Allowed == false, "Chat mode should block RiskLevel 1");
            Assert(!string.IsNullOrEmpty(result.Reason), "Blocked action should have a reason");
            Assert(result.Reason.Contains("Chat mode is read-only"), "Reason should mention Chat mode is read-only");
        }

        private static void TestChatModeRiskLevel3Blocked()
        {
            var session = new AssistantSession();
            session.Mode = SessionMode.Chat;
            var action = new OfficeAction { Host = "Excel", Operation = "excel.remove_duplicates", RiskLevel = 3 };
            var result = session.IsActionAllowed(action);
            Assert(result.Allowed == false, "Chat mode should block RiskLevel 3");
            Assert(!string.IsNullOrEmpty(result.Reason), "Blocked action should have a reason");
        }

        private static void TestPlanModeHighRiskAllowed()
        {
            var session = new AssistantSession();
            session.Mode = SessionMode.Plan;
            var action = new OfficeAction { Host = "Excel", Operation = "excel.remove_duplicates", RiskLevel = 3 };
            var result = session.IsActionAllowed(action);
            Assert(result.Allowed == true, "Plan mode should allow high-risk actions");
            Assert(result.Reason == null, "Allowed action should have no reason");
        }

        private static void TestEditModeHighRiskAllowed()
        {
            var session = new AssistantSession();
            session.Mode = SessionMode.Edit;
            var action = new OfficeAction { Host = "Excel", Operation = "excel.remove_duplicates", RiskLevel = 3 };
            var result = session.IsActionAllowed(action);
            Assert(result.Allowed == true, "Edit mode should allow high-risk actions");
            Assert(result.Reason == null, "Allowed action should have no reason");
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
