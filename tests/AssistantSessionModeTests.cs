using System;
using System.Collections.Generic;
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
            TestChatModeBlocksAfterPreVerifySyncsRiskLevel();
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

        /// <summary>
        /// Regression test for the ExecuteOfficeAction ordering fix (§2.7): the Chat-mode gate must
        /// read the ToolRegistry-synced RiskLevel, not whatever a locally-forged/stale action carried.
        /// "excel.write_formula" is registered at RiskLevel 2; a caller constructing the OfficeAction
        /// with RiskLevel 0 (e.g. a buggy extractor) must still be blocked in Chat mode once PreVerify
        /// has run — exactly the order ExecuteOfficeAction now uses (PreVerify before IsActionAllowed).
        /// </summary>
        private static void TestChatModeBlocksAfterPreVerifySyncsRiskLevel()
        {
            var session = new AssistantSession();
            session.Mode = SessionMode.Chat;
            var action = new OfficeAction
            {
                Host = "Excel",
                Operation = "excel.write_formula",
                RiskLevel = 0,
                Target = new ActionTarget { Range = "B2" }
            };
            action.Parameters = new Dictionary<string, object> { { "formula", "=SUM(A1:A10)" } };

            // Simulate ExecuteOfficeAction's fixed order: PreVerify syncs RiskLevel from ToolRegistry
            // (2, not the locally-set 0) before the mode gate is consulted.
            var pre = ActionVerifier.PreVerify(action, "Excel");
            Assert(pre.Tool != null, "excel.write_formula should be a recognized tool");
            Assert(action.RiskLevel == 2, "PreVerify should sync RiskLevel to the registry value (2), even when other validation is still pending");

            var result = session.IsActionAllowed(action);
            Assert(result.Allowed == false, "Chat mode should block the action once RiskLevel is synced to 2");
            Assert(!string.IsNullOrEmpty(result.Reason), "Blocked action should have a reason");
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
