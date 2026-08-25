using System;
using System.Collections.Generic;
using System.Linq;
using MSOfficeAIAssistant.Core;
using MSOfficeAIAssistant.Core.Actions;
using MSOfficeAIAssistant.Core.Planning;

namespace MSOfficeAIAssistant.Tests
{
    public static class PlannerTests
    {
        public static void RunAll()
        {
            TestBuildPlanFromActionsPreservesOrder();
            TestMoveStepUpSwapsOrder();
            TestMoveStepDownSwapsOrder();
            TestRenumberStepsAfterEdits();
            TestRemoveStepAndRenumber();
            TestInsertReasoningStepAtPosition();
            TestIsFullyAutoRunnableTrueForRiskLevel0();
            TestIsFullyAutoRunnableFalseForRiskLevel1();
            TestValidateEmptyListForValidPlan();
            TestValidateDetectsOrderGap();
            TestValidateDetectsUnresolvedOperation();
            TestDistinctHostsOrderPreserved();
        }

        private static void TestBuildPlanFromActionsPreservesOrder()
        {
            var actions = new List<OfficeAction>
            {
                CreateTestAction("excel.write_formula", "Excel", "Write formula A1"),
                CreateTestAction("excel.create_table", "Excel", "Create table"),
                CreateTestAction("word.add_comment", "Word", "Add comment")
            };

            var plan = Planner.BuildPlanFromActions("Test Plan", "User request", actions);

            Assert(plan != null, "Plan should not be null");
            Assert(plan.Steps.Count == 3, "Plan should have 3 steps");
            Assert(plan.Steps[0].Action == actions[0], "First step should match first action");
            Assert(plan.Steps[1].Action == actions[1], "Second step should match second action");
            Assert(plan.Steps[2].Action == actions[2], "Third step should match third action");
            Assert(plan.Steps[0].Order == 1, "First step Order should be 1");
            Assert(plan.Steps[1].Order == 2, "Second step Order should be 2");
            Assert(plan.Steps[2].Order == 3, "Third step Order should be 3");
            Assert(plan.Title == "Test Plan", "Title should match");
            Assert(plan.SourceRequest == "User request", "SourceRequest should match");
        }

        private static void TestMoveStepUpSwapsOrder()
        {
            var actions = new List<OfficeAction>
            {
                CreateTestAction("excel.write_formula", "Excel", "First"),
                CreateTestAction("excel.create_table", "Excel", "Second"),
                CreateTestAction("word.add_comment", "Word", "Third")
            };

            var plan = Planner.BuildPlanFromActions("Test", "Request", actions);
            string secondStepId = plan.Steps[1].StepId;

            plan.MoveStepUp(secondStepId);

            Assert(plan.Steps[0].Action.ExpectedResult == "Second", "Second should now be first");
            Assert(plan.Steps[1].Action.ExpectedResult == "First", "First should now be second");
            Assert(plan.Steps[0].Order == 1, "New first step should have Order 1");
            Assert(plan.Steps[1].Order == 2, "New second step should have Order 2");
        }

        private static void TestMoveStepDownSwapsOrder()
        {
            var actions = new List<OfficeAction>
            {
                CreateTestAction("excel.write_formula", "Excel", "First"),
                CreateTestAction("excel.create_table", "Excel", "Second"),
                CreateTestAction("word.add_comment", "Word", "Third")
            };

            var plan = Planner.BuildPlanFromActions("Test", "Request", actions);
            string firstStepId = plan.Steps[0].StepId;

            plan.MoveStepDown(firstStepId);

            Assert(plan.Steps[0].Action.ExpectedResult == "Second", "Second should now be first");
            Assert(plan.Steps[1].Action.ExpectedResult == "First", "First should now be second");
            Assert(plan.Steps[0].Order == 1, "New first step should have Order 1");
            Assert(plan.Steps[1].Order == 2, "New second step should have Order 2");
        }

        private static void TestRenumberStepsAfterEdits()
        {
            var actions = new List<OfficeAction>
            {
                CreateTestAction("excel.write_formula", "Excel", "A"),
                CreateTestAction("excel.create_table", "Excel", "B"),
                CreateTestAction("word.add_comment", "Word", "C"),
                CreateTestAction("powerpoint.move_slide", "PowerPoint", "D")
            };

            var plan = Planner.BuildPlanFromActions("Test", "Request", actions);

            // Do several edits
            plan.MoveStepDown(plan.Steps[0].StepId); // A moves after B
            plan.MoveStepUp(plan.Steps[3].StepId);   // D moves before C
            plan.RemoveStep(plan.Steps[1].StepId);   // Remove one

            // After all edits, Order should still be a clean 1..N sequence
            var orders = plan.Steps.Select(s => s.Order).OrderBy(o => o).ToList();
            for (int i = 0; i < orders.Count; i++)
            {
                Assert(orders[i] == i + 1, string.Format("Order should be {0}, got {1}", i + 1, orders[i]));
            }
        }

        private static void TestRemoveStepAndRenumber()
        {
            var actions = new List<OfficeAction>
            {
                CreateTestAction("excel.write_formula", "Excel", "A"),
                CreateTestAction("excel.create_table", "Excel", "B"),
                CreateTestAction("word.add_comment", "Word", "C")
            };

            var plan = Planner.BuildPlanFromActions("Test", "Request", actions);
            string stepToRemoveId = plan.Steps[1].StepId;

            plan.RemoveStep(stepToRemoveId);

            Assert(plan.Steps.Count == 2, "Plan should have 2 steps after removal");
            Assert(plan.Steps[0].Order == 1, "First step Order should be 1");
            Assert(plan.Steps[1].Order == 2, "Second step Order should be 2");
            Assert(plan.Steps[0].Action.ExpectedResult == "A", "First should still be A");
            Assert(plan.Steps[1].Action.ExpectedResult == "C", "Second should be C");
        }

        private static void TestInsertReasoningStepAtPosition()
        {
            var actions = new List<OfficeAction>
            {
                CreateTestAction("excel.write_formula", "Excel", "Formula 1"),
                CreateTestAction("excel.create_table", "Excel", "Create table")
            };

            var plan = Planner.BuildPlanFromActions("Test", "Request", actions);

            Planner.InsertReasoningStep(plan, 2, "Analyze the data", "Excel");

            Assert(plan.Steps.Count == 3, "Plan should have 3 steps");
            Assert(plan.Steps[1].IsReasoningOnly == true, "Middle step should be reasoning-only");
            Assert(plan.Steps[1].Action == null, "Reasoning step should have no Action");
            Assert(plan.Steps[1].Description == "Analyze the data", "Description should match");
            Assert(plan.Steps[1].TargetHost == "Excel", "Host should be Excel");
            Assert(plan.Steps[0].Order == 1, "First step Order should be 1");
            Assert(plan.Steps[1].Order == 2, "Reasoning step Order should be 2");
            Assert(plan.Steps[2].Order == 3, "Last step Order should be 3");
        }

        private static void TestIsFullyAutoRunnableTrueForRiskLevel0()
        {
            var actions = new List<OfficeAction>();

            // Add RiskLevel 0 actions
            var action1 = CreateTestAction("excel.write_formula", "Excel", "Low risk");
            action1.RiskLevel = 0;
            actions.Add(action1);

            var action2 = CreateTestAction("word.add_comment", "Word", "Low risk");
            action2.RiskLevel = 0;
            actions.Add(action2);

            var plan = Planner.BuildPlanFromActions("Test", "Request", actions);

            // Add a reasoning step
            Planner.InsertReasoningStep(plan, 1, "Prepare", "Excel");

            Assert(Planner.IsFullyAutoRunnable(plan) == true, "Plan with only RiskLevel 0 actions and reasoning steps should be auto-runnable");
        }

        private static void TestIsFullyAutoRunnableFalseForRiskLevel1()
        {
            var actions = new List<OfficeAction>();

            // Add a RiskLevel 1 action
            var action1 = CreateTestAction("excel.write_formula", "Excel", "Medium risk");
            action1.RiskLevel = 1;
            actions.Add(action1);

            var plan = Planner.BuildPlanFromActions("Test", "Request", actions);

            Assert(Planner.IsFullyAutoRunnable(plan) == false, "Plan with RiskLevel 1 pending action should NOT be auto-runnable");
        }

        private static void TestValidateEmptyListForValidPlan()
        {
            var actions = new List<OfficeAction>
            {
                CreateTestAction("excel.write_formula", "Excel", "Valid formula"),
                CreateTestAction("word.add_comment", "Word", "Valid comment")
            };

            var plan = Planner.BuildPlanFromActions("Valid Plan", "Request", actions);

            var errors = Planner.Validate(plan);

            Assert(errors.Count == 0, "Valid plan should have no validation errors");
        }

        private static void TestValidateDetectsOrderGap()
        {
            var plan = new Plan();
            plan.Steps.Add(new PlanStep { Order = 1 });
            plan.Steps.Add(new PlanStep { Order = 3 }); // Gap: missing Order 2

            var errors = Planner.Validate(plan);

            Assert(errors.Count > 0, "Validation should detect Order gap");
            Assert(errors.Any(e => e.Contains("gap")), "Error message should mention gap");
        }

        private static void TestValidateDetectsUnresolvedOperation()
        {
            var action = new OfficeAction
            {
                Operation = "nonexistent.operation",
                Host = "Excel",
                ExpectedResult = "Invalid operation"
            };

            var plan = Planner.BuildPlanFromActions("Test", "Request", new[] { action });

            var errors = Planner.Validate(plan);

            Assert(errors.Count > 0, "Validation should detect unresolved operation");
            Assert(errors.Any(e => e.Contains("not found")), "Error message should mention operation not found");
        }

        private static void TestDistinctHostsOrderPreserved()
        {
            var actions = new List<OfficeAction>
            {
                CreateTestAction("excel.write_formula", "Excel", "Excel 1"),
                CreateTestAction("word.add_comment", "Word", "Word 1"),
                CreateTestAction("excel.create_table", "Excel", "Excel 2"),
                CreateTestAction("powerpoint.move_slide", "PowerPoint", "PowerPoint 1")
            };

            var plan = Planner.BuildPlanFromActions("Test", "Request", actions);

            var hosts = plan.DistinctHosts;

            Assert(hosts.Count == 3, "Should have 3 distinct hosts");
            Assert(hosts[0] == "Excel", "First host should be Excel");
            Assert(hosts[1] == "Word", "Second host should be Word");
            Assert(hosts[2] == "PowerPoint", "Third host should be PowerPoint");
        }

        #region Test Helpers

        private static OfficeAction CreateTestAction(string operation, string host, string expectedResult)
        {
            var tool = ToolRegistry.GetTool(operation, host);
            var action = new OfficeAction
            {
                ActionId = Guid.NewGuid().ToString(),
                Operation = operation,
                Host = host,
                ExpectedResult = expectedResult,
                SourceReason = "Test",
                Status = OfficeActionStatus.Pending,
                RiskLevel = tool != null ? tool.RiskLevel : 1,
                IsUndoable = tool != null ? tool.IsUndoable : true,
                RequiresApproval = tool != null ? tool.RequiresApproval : true
            };
            return action;
        }

        private static void Assert(bool condition, string message)
        {
            if (!condition)
            {
                throw new Exception("Assertion failed: " + message);
            }
        }

        #endregion
    }
}
