using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using MSOfficeAIAssistant.Core;
using MSOfficeAIAssistant.Core.Actions;
using MSOfficeAIAssistant.Core.Planning;
using MSOfficeAIAssistant.Hosts;

namespace MSOfficeAIAssistant.Tests
{
    public static class CrossHostPlanCoordinatorTests
    {
        public static void RunAll()
        {
            Console.WriteLine("=== CrossHostPlanCoordinator Multi-Host Workflow Tests === START");

            TestIsMultiHostReturnsFalseForSingleHost();
            TestIsMultiHostReturnsTrueForMultipleHosts();
            TestExecuteForCurrentHostSkipsOutOfHostStepsAndContinues();
            TestGetNextPendingHostAfterMultiHostExecution();
            TestPausedForDifferentHostResult();
            TestSingleHostPlanExecutesCompletely();
            TestComputeStatusForMultiHostAwaitingDifferentHost();
            TestComputeStatusForSingleHostCompleted();
            TestStepGatedForApprovalStopsExecutionOnCurrentHost();

            Console.WriteLine("All CrossHostPlanCoordinator tests passed!");
        }

        private static void Assert(bool condition, string message)
        {
            if (!condition)
            {
                throw new Exception("Assertion Failed: " + message);
            }
        }

        private static void TestIsMultiHostReturnsFalseForSingleHost()
        {
            // A plan with all steps targeting the same host should return false
            var plan = new Plan();

            var step1 = new PlanStep
            {
                StepId = "step-1",
                Order = 1,
                Description = "Excel step 1",
                TargetHost = "Excel",
                Status = PlanStepStatus.Pending
            };
            plan.Steps.Add(step1);

            var step2 = new PlanStep
            {
                StepId = "step-2",
                Order = 2,
                Description = "Excel step 2",
                TargetHost = "Excel",
                Status = PlanStepStatus.Pending
            };
            plan.Steps.Add(step2);

            bool isMulti = CrossHostPlanCoordinator.IsMultiHost(plan);

            Assert(!isMulti, "Plan with all steps on same host should return false");
            Console.WriteLine("  [PASS] IsMultiHost returns false for single-host plan");
        }

        private static void TestIsMultiHostReturnsTrueForMultipleHosts()
        {
            // A plan with steps targeting different hosts should return true
            var plan = new Plan();

            var step1 = new PlanStep
            {
                StepId = "step-1",
                Order = 1,
                Description = "Excel step",
                TargetHost = "Excel",
                Status = PlanStepStatus.Pending
            };
            plan.Steps.Add(step1);

            var step2 = new PlanStep
            {
                StepId = "step-2",
                Order = 2,
                Description = "Word step",
                TargetHost = "Word",
                Status = PlanStepStatus.Pending
            };
            plan.Steps.Add(step2);

            bool isMulti = CrossHostPlanCoordinator.IsMultiHost(plan);

            Assert(isMulti, "Plan with steps on different hosts should return true");
            Console.WriteLine("  [PASS] IsMultiHost returns true for multi-host plan");
        }

        private static void TestExecuteForCurrentHostSkipsOutOfHostStepsAndContinues()
        {
            // 3-step plan: step 1 Excel (approved), step 2 Word, step 3 Excel (approved)
            // When executing on Excel, should execute step 1, SKIP step 2, CONTINUE and execute step 3
            var plan = new Plan();

            var step1 = new PlanStep
            {
                StepId = "step-1",
                Order = 1,
                Description = "Excel step 1",
                Action = new OfficeAction
                {
                    ActionId = "act-1",
                    Host = "Excel",
                    Operation = "excel.write_formula",
                    Target = new ActionTarget { Range = "B2" },
                    Parameters = new Dictionary<string, object> { { "formula", "=1+1" } },
                    RiskLevel = 2,
                    RequiresApproval = false
                },
                TargetHost = "Excel",
                Status = PlanStepStatus.Approved
            };
            plan.Steps.Add(step1);

            var step2 = new PlanStep
            {
                StepId = "step-2",
                Order = 2,
                Description = "Word step (should be skipped)",
                TargetHost = "Word",
                Status = PlanStepStatus.Pending
            };
            plan.Steps.Add(step2);

            var step3 = new PlanStep
            {
                StepId = "step-3",
                Order = 3,
                Description = "Excel step 2",
                TargetHost = "Excel",
                Status = PlanStepStatus.Pending
            };
            plan.Steps.Add(step3);

            var executor = new PlanExecutor(plan, new ExcelController(null));

            var result = CrossHostPlanCoordinator.ExecuteForCurrentHost(plan, executor, "Excel");

            // Step 1 should be Applied or Failed (attempted on headless controller)
            Assert(step1.Status == PlanStepStatus.Applied || step1.Status == PlanStepStatus.Failed,
                "Step 1 should be attempted and reach terminal state");

            // Step 2 should remain untouched - NOT gated for approval, just skipped
            Assert(step2.Status == PlanStepStatus.Pending,
                "Step 2 (Word) should be skipped, not attempted");

            // Step 3 should be Applied (reasoning-only step)
            Assert(step3.Status == PlanStepStatus.Applied,
                "Step 3 (Excel reasoning-only) should be executed after step 2 was skipped");

            // At least 2 steps should have been executed (step 1 and step 3)
            Assert(result.StepsExecutedOnThisHost >= 2,
                "Should have executed at least 2 steps on Excel (step 1 and step 3)");

            Console.WriteLine("  [PASS] ExecuteForCurrentHost skips out-of-host steps and continues to same-host steps later in sequence");
        }

        private static void TestGetNextPendingHostAfterMultiHostExecution()
        {
            // After executing the above scenario, the next pending step should be step 2 (Word)
            var plan = new Plan();

            var step1 = new PlanStep
            {
                StepId = "step-1",
                Order = 1,
                Description = "Excel step",
                Action = new OfficeAction
                {
                    ActionId = "act-1",
                    Host = "Excel",
                    Operation = "excel.write_formula",
                    Target = new ActionTarget { Range = "B2" },
                    Parameters = new Dictionary<string, object> { { "formula", "=1+1" } },
                    RiskLevel = 2,
                    RequiresApproval = false
                },
                TargetHost = "Excel",
                Status = PlanStepStatus.Approved
            };
            plan.Steps.Add(step1);

            var step2 = new PlanStep
            {
                StepId = "step-2",
                Order = 2,
                Description = "Word step",
                TargetHost = "Word",
                Status = PlanStepStatus.Pending
            };
            plan.Steps.Add(step2);

            var step3 = new PlanStep
            {
                StepId = "step-3",
                Order = 3,
                Description = "Excel step 2",
                TargetHost = "Excel",
                Status = PlanStepStatus.Pending
            };
            plan.Steps.Add(step3);

            var executor = new PlanExecutor(plan, new ExcelController(null));
            CrossHostPlanCoordinator.ExecuteForCurrentHost(plan, executor, "Excel");

            // After execution, step 1 should be in terminal state, step 2 is first pending
            string nextHost = CrossHostPlanCoordinator.GetNextPendingHost(plan);

            Assert(nextHost == "Word",
                string.Format("Next pending host should be Word, got {0}", nextHost));

            Console.WriteLine("  [PASS] GetNextPendingHost returns Word after Excel execution");
        }

        private static void TestPausedForDifferentHostResult()
        {
            // Setup: 3-step plan (Excel, Word, Excel), execute on Excel
            var plan = new Plan();

            var step1 = new PlanStep
            {
                StepId = "step-1",
                Order = 1,
                Description = "Excel step",
                Action = new OfficeAction
                {
                    ActionId = "act-1",
                    Host = "Excel",
                    Operation = "excel.write_formula",
                    Target = new ActionTarget { Range = "B2" },
                    Parameters = new Dictionary<string, object> { { "formula", "=1+1" } },
                    RiskLevel = 2,
                    RequiresApproval = false
                },
                TargetHost = "Excel",
                Status = PlanStepStatus.Approved
            };
            plan.Steps.Add(step1);

            var step2 = new PlanStep
            {
                StepId = "step-2",
                Order = 2,
                Description = "Word step",
                TargetHost = "Word",
                Status = PlanStepStatus.Pending
            };
            plan.Steps.Add(step2);

            var step3 = new PlanStep
            {
                StepId = "step-3",
                Order = 3,
                Description = "Excel step 2",
                TargetHost = "Excel",
                Status = PlanStepStatus.Pending
            };
            plan.Steps.Add(step3);

            var executor = new PlanExecutor(plan, new ExcelController(null));
            var result = CrossHostPlanCoordinator.ExecuteForCurrentHost(plan, executor, "Excel");

            Assert(result.PausedForDifferentHost,
                "Result should indicate paused for different host");
            Assert(result.NextHost == "Word",
                string.Format("NextHost should be Word, got {0}", result.NextHost));
            Assert(!result.PlanFullyComplete,
                "Plan should not be fully complete");

            Console.WriteLine("  [PASS] PausedForDifferentHost and NextHost are correctly set");
        }

        private static void TestSingleHostPlanExecutesCompletely()
        {
            // Single-host plan: all steps on Excel, all should execute and complete
            var plan = new Plan();

            var step1 = new PlanStep
            {
                StepId = "step-1",
                Order = 1,
                Description = "Reasoning step",
                TargetHost = "Excel",
                Status = PlanStepStatus.Pending
            };
            plan.Steps.Add(step1);

            var step2 = new PlanStep
            {
                StepId = "step-2",
                Order = 2,
                Description = "Another reasoning step",
                TargetHost = "Excel",
                Status = PlanStepStatus.Pending
            };
            plan.Steps.Add(step2);

            var executor = new PlanExecutor(plan, new ExcelController(null));
            var result = CrossHostPlanCoordinator.ExecuteForCurrentHost(plan, executor, "Excel");

            Assert(result.PlanFullyComplete,
                "Single-host plan should be fully complete after execution");
            Assert(!result.PausedForDifferentHost,
                "Single-host plan should not pause for different host");
            Assert(result.NextHost == null,
                "NextHost should be null for completed single-host plan");
            Assert(step1.Status == PlanStepStatus.Applied,
                "Step 1 should be applied");
            Assert(step2.Status == PlanStepStatus.Applied,
                "Step 2 should be applied");

            Console.WriteLine("  [PASS] Single-host plan executes completely");
        }

        private static void TestComputeStatusForMultiHostAwaitingDifferentHost()
        {
            // Setup: 2-step plan (Excel, Word), execute on Excel
            var plan = new Plan();

            var step1 = new PlanStep
            {
                StepId = "step-1",
                Order = 1,
                Description = "Excel step",
                TargetHost = "Excel",
                Status = PlanStepStatus.Applied
            };
            plan.Steps.Add(step1);

            var step2 = new PlanStep
            {
                StepId = "step-2",
                Order = 2,
                Description = "Word step",
                TargetHost = "Word",
                Status = PlanStepStatus.Pending
            };
            plan.Steps.Add(step2);

            var executor = new PlanExecutor(plan, new ExcelController(null));

            string status = CrossHostPlanCoordinator.ComputeStatus(plan, executor, "Excel");

            Assert(status.Contains("Awaiting host:") && status.Contains("Word"),
                string.Format("Status should indicate awaiting Word host, got: {0}", status));

            Console.WriteLine("  [PASS] ComputeStatus returns 'Awaiting host: Word' for multi-host scenario");
        }

        private static void TestComputeStatusForSingleHostCompleted()
        {
            // Setup: 2-step plan (Excel, Excel), both applied
            var plan = new Plan();

            var step1 = new PlanStep
            {
                StepId = "step-1",
                Order = 1,
                Description = "Excel step 1",
                TargetHost = "Excel",
                Status = PlanStepStatus.Applied
            };
            plan.Steps.Add(step1);

            var step2 = new PlanStep
            {
                StepId = "step-2",
                Order = 2,
                Description = "Excel step 2",
                TargetHost = "Excel",
                Status = PlanStepStatus.Applied
            };
            plan.Steps.Add(step2);

            var executor = new PlanExecutor(plan, new ExcelController(null));

            string status = CrossHostPlanCoordinator.ComputeStatus(plan, executor, "Excel");

            Assert(status == "Completed",
                string.Format("Status should be 'Completed', got: {0}", status));

            Console.WriteLine("  [PASS] ComputeStatus returns 'Completed' for fully-applied single-host plan");
        }

        private static void TestStepGatedForApprovalStopsExecutionOnCurrentHost()
        {
            // Setup: 3-step plan on same host (PowerPoint), step 2 gated for approval
            // Execution should stop at step 2 and not attempt step 3
            var plan = new Plan();

            var step1 = new PlanStep
            {
                StepId = "step-1",
                Order = 1,
                Description = "Reasoning step (executes automatically)",
                TargetHost = "PowerPoint",
                Status = PlanStepStatus.Pending
            };
            plan.Steps.Add(step1);

            var step2 = new PlanStep
            {
                StepId = "step-2",
                Order = 2,
                Description = "Gated step (requires approval)",
                Action = new OfficeAction
                {
                    ActionId = "act-2",
                    Host = "PowerPoint",
                    Operation = "powerpoint.set_notes",
                    Target = new ActionTarget { Slide = 1 },
                    Parameters = new Dictionary<string, object> { { "slide", 1 }, { "notes", "Test notes" } },
                    RiskLevel = 1,
                    RequiresApproval = true
                },
                TargetHost = "PowerPoint",
                Status = PlanStepStatus.Pending
            };
            plan.Steps.Add(step2);

            var step3 = new PlanStep
            {
                StepId = "step-3",
                Order = 3,
                Description = "Should not execute (stops at approval gate)",
                TargetHost = "PowerPoint",
                Status = PlanStepStatus.Pending
            };
            plan.Steps.Add(step3);

            var executor = new PlanExecutor(plan, new PowerPointController(null));
            var result = CrossHostPlanCoordinator.ExecuteForCurrentHost(plan, executor, "PowerPoint");

            // Step 1 should be applied (reasoning-only, same host)
            Assert(step1.Status == PlanStepStatus.Applied,
                "Reasoning-only step 1 should be applied");

            // Step 2 should remain Pending (gated for approval, not executed)
            Assert(step2.Status == PlanStepStatus.Pending,
                "Step 2 should remain Pending (gated for approval)");

            // Step 3 should remain Pending (not attempted after approval gate)
            Assert(step3.Status == PlanStepStatus.Pending,
                "Step 3 should remain Pending (not attempted after approval gate on step 2)");

            // Executor should be in AwaitingApproval state
            Assert(executor.State == PlanExecutionState.AwaitingApproval,
                "Executor state should be AwaitingApproval");

            string status = CrossHostPlanCoordinator.ComputeStatus(plan, executor, "PowerPoint");
            Assert(status == "Awaiting approval",
                string.Format("ComputeStatus should return 'Awaiting approval', got: {0}", status));

            Console.WriteLine("  [PASS] Execution stops at approval gate on current host step");
        }
    }
}
