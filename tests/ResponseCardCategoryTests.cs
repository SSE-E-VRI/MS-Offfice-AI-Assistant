using System;
using System.Collections.ObjectModel;
using MSOfficeAIAssistant.API.Models;
using MSOfficeAIAssistant.Core.Actions;
using MSOfficeAIAssistant.Core.Planning;
using MSOfficeAIAssistant.UI.Cards;

namespace MSOfficeAIAssistant.Tests
{
    public static class ResponseCardCategoryTests
    {
        public static void RunAll()
        {
            TestClassifyNull();
            TestClassifyWithOfficeActions();
            TestClassifyWithPlan();
            TestClassifyWithBothPlanAndOfficeActions();
            TestClassifyWarning();
            TestClassifyFinding();
            TestClassifyRecommendation();
            TestClassifySummary();
            TestClassifyPlainText();
            TestClassifyWithLeadingWhitespace();
            TestClassifyBothOfficeActionsAndWarningMarker();
        }

        private static void TestClassifyNull()
        {
            var result = ResponseCardCategoryClassifier.Classify(null);
            Assert(result == ResponseCardCategory.Text, "Classify(null) should return Text");
        }

        private static void TestClassifyWithOfficeActions()
        {
            var message = new ChatMessage("assistant", "Here is what I recommend");
            var action = new OfficeAction();
            action.ActionId = Guid.NewGuid().ToString("N");
            action.Host = "Word";
            action.Operation = "write_value";
            message.OfficeActions.Add(action);

            var result = ResponseCardCategoryClassifier.Classify(message);
            Assert(result == ResponseCardCategory.ActionPreview,
                "Message with HasOfficeActions=true should be ActionPreview");
        }

        private static void TestClassifyWithPlan()
        {
            var message = new ChatMessage("assistant", "Here is your plan");
            var plan = new Plan();
            plan.Title = "Test Plan";
            var step = new PlanStep
            {
                Order = 1,
                Description = "Step 1",
                Action = null,
                TargetHost = "Excel",
                Status = PlanStepStatus.Pending
            };
            plan.Steps.Add(step);
            message.Plan = plan;

            var result = ResponseCardCategoryClassifier.Classify(message);
            Assert(result == ResponseCardCategory.Plan,
                "Message with HasPlan=true should be Plan");
        }

        private static void TestClassifyWithBothPlanAndOfficeActions()
        {
            // Deliberately construct a message with both Plan and OfficeActions (even though
            // ProcessAssistantResponse doesn't produce this combination) to verify priority order
            var message = new ChatMessage("assistant", "Here is your plan with actions");

            var plan = new Plan();
            plan.Title = "Test Plan";
            plan.Steps.Add(new PlanStep
            {
                Order = 1,
                Description = "Step 1",
                Action = null,
                TargetHost = "Excel",
                Status = PlanStepStatus.Pending
            });
            message.Plan = plan;

            var action = new OfficeAction();
            action.ActionId = Guid.NewGuid().ToString("N");
            action.Host = "Word";
            action.Operation = "write_value";
            message.OfficeActions.Add(action);

            var result = ResponseCardCategoryClassifier.Classify(message);
            Assert(result == ResponseCardCategory.Plan,
                "Message with both Plan and OfficeActions should prioritize Plan");
        }

        private static void TestClassifyWarning()
        {
            var message1 = new ChatMessage("assistant", "**Warning:** This is a warning");
            var result1 = ResponseCardCategoryClassifier.Classify(message1);
            Assert(result1 == ResponseCardCategory.Warning,
                "Content starting with '**Warning:**' should be Warning");

            var message2 = new ChatMessage("assistant", "Warning: Another warning");
            var result2 = ResponseCardCategoryClassifier.Classify(message2);
            Assert(result2 == ResponseCardCategory.Warning,
                "Content starting with 'Warning:' should be Warning");

            var message3 = new ChatMessage("assistant", "⚠ Warning symbol");
            var result3 = ResponseCardCategoryClassifier.Classify(message3);
            Assert(result3 == ResponseCardCategory.Warning,
                "Content starting with '⚠' should be Warning");
        }

        private static void TestClassifyFinding()
        {
            var message1 = new ChatMessage("assistant", "**Finding:** X was observed");
            var result1 = ResponseCardCategoryClassifier.Classify(message1);
            Assert(result1 == ResponseCardCategory.Finding,
                "Content starting with '**Finding:**' should be Finding");

            var message2 = new ChatMessage("assistant", "Finding: Another observation");
            var result2 = ResponseCardCategoryClassifier.Classify(message2);
            Assert(result2 == ResponseCardCategory.Finding,
                "Content starting with 'Finding:' should be Finding");
        }

        private static void TestClassifyRecommendation()
        {
            var message1 = new ChatMessage("assistant", "**Recommendation:** Do Y");
            var result1 = ResponseCardCategoryClassifier.Classify(message1);
            Assert(result1 == ResponseCardCategory.Recommendation,
                "Content starting with '**Recommendation:**' should be Recommendation");

            var message2 = new ChatMessage("assistant", "Recommendation: Another suggestion");
            var result2 = ResponseCardCategoryClassifier.Classify(message2);
            Assert(result2 == ResponseCardCategory.Recommendation,
                "Content starting with 'Recommendation:' should be Recommendation");
        }

        private static void TestClassifySummary()
        {
            var message1 = new ChatMessage("assistant", "**Summary:** Overview here");
            var result1 = ResponseCardCategoryClassifier.Classify(message1);
            Assert(result1 == ResponseCardCategory.Summary,
                "Content starting with '**Summary:**' should be Summary");

            var message2 = new ChatMessage("assistant", "Summary: Quick overview");
            var result2 = ResponseCardCategoryClassifier.Classify(message2);
            Assert(result2 == ResponseCardCategory.Summary,
                "Content starting with 'Summary:' should be Summary");
        }

        private static void TestClassifyPlainText()
        {
            var message1 = new ChatMessage("assistant", "Just a normal reply");
            var result1 = ResponseCardCategoryClassifier.Classify(message1);
            Assert(result1 == ResponseCardCategory.Text,
                "Regular content should be Text");

            var message2 = new ChatMessage("assistant", "");
            var result2 = ResponseCardCategoryClassifier.Classify(message2);
            Assert(result2 == ResponseCardCategory.Text,
                "Empty content should be Text");

            var message3 = new ChatMessage("assistant", null);
            var result3 = ResponseCardCategoryClassifier.Classify(message3);
            Assert(result3 == ResponseCardCategory.Text,
                "Null content should be Text");
        }

        private static void TestClassifyWithLeadingWhitespace()
        {
            var message = new ChatMessage("assistant", "   Warning: Spaced warning");
            var result = ResponseCardCategoryClassifier.Classify(message);
            Assert(result == ResponseCardCategory.Warning,
                "Content with leading whitespace before marker should still classify marker");
        }

        private static void TestClassifyBothOfficeActionsAndWarningMarker()
        {
            // Priority test: ActionPreview should win over Warning marker
            var message = new ChatMessage("assistant", "**Warning:** I made changes");
            var action = new OfficeAction();
            action.ActionId = Guid.NewGuid().ToString("N");
            action.Host = "Word";
            action.Operation = "write_value";
            message.OfficeActions.Add(action);

            var result = ResponseCardCategoryClassifier.Classify(message);
            Assert(result == ResponseCardCategory.ActionPreview,
                "ActionPreview (HasOfficeActions) should have priority over content markers");
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
