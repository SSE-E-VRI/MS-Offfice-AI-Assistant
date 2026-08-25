using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using MSOfficeAIAssistant.Core;
using MSOfficeAIAssistant.Core.Actions;
using MSOfficeAIAssistant.Core.Planning;
using MSOfficeAIAssistant.Hosts;

namespace MSOfficeAIAssistant.Tests
{
    public static class PlanExecutorTests
    {
        public static void RunAll()
        {
            Console.WriteLine("=== PlanExecutor State Machine Tests === START");

            TestRiskLevel0StepExecutesAutomatically();
            TestRiskLevel1StepWithPendingStatusAwaitingApproval();
            TestApprovedGatedStepExecutesOnNextCall();
            TestPreVerifyValidationErrorMarksStepFailed();
            TestExecuteAllStopsAtAwaitingApprovalStep();
            TestExecuteAllStopsAtFailedStep();
            TestRollbackAllDelegatesAndSetsStateRolledBack();
            TestRollbackAllHealsActionStatusAfterRoundTripAndUpdatesPlanStep();
            TestBusyRetryableStateRetriesUpTo3Times();
            TestPostVerifyNonRetryableFailureSetsExecutorFailed();

            Console.WriteLine("All PlanExecutor State Machine tests passed!");
        }

        private static void Assert(bool condition, string message)
        {
            if (!condition)
            {
                throw new Exception("Assertion Failed: " + message);
            }
        }

        private static void TestRiskLevel0StepExecutesAutomatically()
        {
            // NOTE: No RiskLevel-0 mutating tool is registered in ToolRegistry today - every
            // registered operation (excel.write_formula, powerpoint.set_notes, etc.) requires
            // approval by design (Phase C2/C4). ActionVerifier.PreVerify re-synchronizes
            // RiskLevel/RequiresApproval from the ToolDefinition on every call (the D-5 single
            // source of truth), so a test cannot override RiskLevel on a real registered tool -
            // it will simply be resynced back before ExecuteStep's gate check runs. And an
            // OfficeAction with an unresolvable Operation fails PreVerify.IsValid instead of
            // reaching the gate at all.
            //
            // The one category of step that genuinely always executes without approval gating,
            // regardless of risk, is a reasoning-only step (Action == null) - that is what this
            // test actually verifies.
            var plan = new Plan();
            var step = new PlanStep
            {
                StepId = "step-1",
                Order = 1,
                Description = "Reasoning-only step, no Action",
                TargetHost = "Excel",
                Status = PlanStepStatus.Pending
                // Action left null -> IsReasoningOnly == true
            };
            plan.Steps.Add(step);

            var executor = new PlanExecutor(plan, new ExcelController(null));

            var result = executor.ExecuteStep(1);

            Assert(step.Status == PlanStepStatus.Applied, "Reasoning-only step should be marked Applied automatically");
            Assert(executor.State != PlanExecutionState.AwaitingApproval, "Reasoning-only step must not gate for approval");
            Assert(result.Success, "ExecuteStep should report success for a reasoning-only step");
            Console.WriteLine("  [PASS] Reasoning-only step executes automatically without requiring Approved status");
        }

        private static void TestRiskLevel1StepWithPendingStatusAwaitingApproval()
        {
            var plan = new Plan();
            var step = new PlanStep
            {
                StepId = "step-2",
                Order = 1,
                Description = "Test RiskLevel 1 gating",
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
            plan.Steps.Add(step);

            var executor = new PlanExecutor(plan, new PowerPointController(null));

            // Execute the step - should NOT execute, should set state to AwaitingApproval
            var result = executor.ExecuteStep(1);

            Assert(executor.State == PlanExecutionState.AwaitingApproval, "Executor state should be AwaitingApproval for gated step");
            Assert(step.Status == PlanStepStatus.Pending, "Step status should remain Pending (not executed)");
            Assert(!result.Success, "Execution result should indicate failure (step not executed)");
            Assert(result.ErrorMessage.Contains("requires approval"), "Error message should mention approval requirement");

            Console.WriteLine("  [PASS] RiskLevel 1 step with Pending status sets state to AwaitingApproval and does not execute");
        }

        private static void TestApprovedGatedStepExecutesOnNextCall()
        {
            var plan = new Plan();
            var step = new PlanStep
            {
                StepId = "step-3",
                Order = 1,
                Description = "Test approved gated step execution",
                Action = new OfficeAction
                {
                    ActionId = "act-3",
                    Host = "PowerPoint",
                    Operation = "powerpoint.set_notes",
                    Target = new ActionTarget { Slide = 1 },
                    Parameters = new Dictionary<string, object> { { "slide", 1 }, { "notes", "Updated notes" } },
                    RiskLevel = 1,
                    RequiresApproval = true
                },
                TargetHost = "PowerPoint",
                Status = PlanStepStatus.Pending
            };
            plan.Steps.Add(step);

            var executor = new PlanExecutor(plan, new PowerPointController(null));

            // First call should gate the step
            var firstResult = executor.ExecuteStep(1);
            Assert(step.Status == PlanStepStatus.Pending, "First call should not execute");

            // Manually approve the step
            step.Status = PlanStepStatus.Approved;

            // Second call should execute it
            var secondResult = executor.ExecuteStep(1);

            // With headless PowerPoint controller, execution will fail, but status should change from Approved
            Assert(step.Status != PlanStepStatus.Approved, "Step status should change from Approved after ExecuteStep");
            // Status should be either Applied or Failed depending on execution outcome
            Assert(step.Status == PlanStepStatus.Applied || step.Status == PlanStepStatus.Failed, "Step should be in terminal state after execution");

            Console.WriteLine("  [PASS] Approved gated step executes on subsequent ExecuteStep/ContinueFromStep call");
        }

        private static void TestPreVerifyValidationErrorMarksStepFailed()
        {
            var plan = new Plan();
            var step = new PlanStep
            {
                StepId = "step-4",
                Order = 1,
                Description = "Test validation error",
                Action = new OfficeAction
                {
                    ActionId = "act-4",
                    Host = "Excel",
                    Operation = "excel.write_formula",
                    Target = new ActionTarget { Range = "INVALID_RANGE!!!$" },
                    Parameters = new Dictionary<string, object> { { "formula", "=SUM(A1:A10)" } },
                    // Missing target parameter, or using unsafe range
                    RiskLevel = 2,
                    RequiresApproval = false
                },
                TargetHost = "Excel",
                // Approved (not Pending): excel.write_formula is a registered RiskLevel-2 tool
                // that always requires approval (PreVerify resyncs RequiresApproval from the
                // registry regardless of the object initializer above), so a Pending step here
                // would gate for approval before ever reaching PreVerify's validation check. Set
                // Approved to bypass the gate and actually exercise the validation-error path
                // this test targets.
                Status = PlanStepStatus.Approved
            };
            plan.Steps.Add(step);

            var executor = new PlanExecutor(plan, new ExcelController(null));

            // Execute the step
            var result = executor.ExecuteStep(1);

            // Step should be marked Failed due to validation error
            Assert(step.Status == PlanStepStatus.Failed, "Step with validation error should be marked Failed");
            Assert(!string.IsNullOrEmpty(step.ErrorMessage), "ErrorMessage should be populated");
            // The error should mention unsafe or invalid address
            Assert(step.ErrorMessage.Contains("Invalid") || step.ErrorMessage.Contains("unsafe") || step.ErrorMessage.Contains("parameter"),
                "Error message should describe the validation failure");

            Console.WriteLine("  [PASS] Step with PreVerify validation error is marked Failed without calling ToolRegistry.Execute");
        }

        private static void TestExecuteAllStopsAtAwaitingApprovalStep()
        {
            var plan = new Plan();

            // Step 1: RiskLevel 0, should execute
            var step1 = new PlanStep
            {
                StepId = "step-1",
                Order = 1,
                Description = "Low risk step",
                Status = PlanStepStatus.Pending
            };
            plan.Steps.Add(step1);

            // Step 2: RiskLevel 1 Pending, should gate and stop
            var step2 = new PlanStep
            {
                StepId = "step-2",
                Order = 2,
                Description = "Gated step",
                Action = new OfficeAction
                {
                    ActionId = "act-2",
                    Host = "PowerPoint",
                    Operation = "powerpoint.set_notes",
                    Target = new ActionTarget { Slide = 1 },
                    Parameters = new Dictionary<string, object> { { "slide", 1 }, { "notes", "Test" } },
                    RiskLevel = 1,
                    RequiresApproval = true
                },
                TargetHost = "PowerPoint",
                Status = PlanStepStatus.Pending
            };
            plan.Steps.Add(step2);

            // Step 3: Should NOT execute because ExecuteAll should stop at step 2
            var step3 = new PlanStep
            {
                StepId = "step-3",
                Order = 3,
                Description = "Should not execute",
                Status = PlanStepStatus.Pending
            };
            plan.Steps.Add(step3);

            var executor = new PlanExecutor(plan, new PowerPointController(null));
            var progressCalls = new List<PlanExecutionProgress>();

            executor.ExecuteAll(p => progressCalls.Add(p));

            // Step 1 should be Applied (reasoning-only)
            Assert(step1.Status == PlanStepStatus.Applied, "Step 1 should be Applied");

            // Step 2 should remain Pending (gated, awaiting approval)
            Assert(step2.Status == PlanStepStatus.Pending, "Step 2 should remain Pending (gated)");

            // Step 3 should remain Pending (not attempted)
            Assert(step3.Status == PlanStepStatus.Pending, "Step 3 should remain Pending (not attempted after gating)");

            // Executor state should be AwaitingApproval
            Assert(executor.State == PlanExecutionState.AwaitingApproval, "Executor state should be AwaitingApproval");

            Console.WriteLine("  [PASS] ExecuteAll stops at first AwaitingApproval step without attempting subsequent steps");
        }

        private static void TestExecuteAllStopsAtFailedStep()
        {
            var plan = new Plan();

            // Step 1: RiskLevel 0, should execute
            var step1 = new PlanStep
            {
                StepId = "step-1",
                Order = 1,
                Description = "Low risk reasoning step",
                Status = PlanStepStatus.Pending
            };
            plan.Steps.Add(step1);

            // Step 2: Invalid parameters, should fail validation
            var step2 = new PlanStep
            {
                StepId = "step-2",
                Order = 2,
                Description = "Will fail validation",
                Action = new OfficeAction
                {
                    ActionId = "act-2",
                    Host = "Excel",
                    Operation = "excel.write_formula",
                    Target = new ActionTarget { Range = "UNSAFE_RANGE" },
                    Parameters = new Dictionary<string, object>(),
                    RiskLevel = 2,
                    RequiresApproval = false
                },
                TargetHost = "Excel",
                // Approved: excel.write_formula always requires approval per the registry, so a
                // Pending step would gate rather than reach validation - see the comment in
                // TestPreVerifyValidationErrorMarksStepFailed for the full explanation.
                Status = PlanStepStatus.Approved
            };
            plan.Steps.Add(step2);

            // Step 3: Should NOT execute because ExecuteAll should stop at step 2
            var step3 = new PlanStep
            {
                StepId = "step-3",
                Order = 3,
                Description = "Should not execute",
                Status = PlanStepStatus.Pending
            };
            plan.Steps.Add(step3);

            var executor = new PlanExecutor(plan, new ExcelController(null));
            executor.ExecuteAll(null);

            // Step 1 should be Applied
            Assert(step1.Status == PlanStepStatus.Applied, "Step 1 should be Applied");

            // Step 2 should be Failed
            Assert(step2.Status == PlanStepStatus.Failed, "Step 2 should be Failed");

            // Step 3 should remain Pending
            Assert(step3.Status == PlanStepStatus.Pending, "Step 3 should remain Pending (not attempted after failure)");

            // Executor state should be Failed
            Assert(executor.State == PlanExecutionState.Failed, "Executor state should be Failed");

            Console.WriteLine("  [PASS] ExecuteAll stops at first Failed step without attempting subsequent steps");
        }

        private static void TestRollbackAllDelegatesAndSetsStateRolledBack()
        {
            var plan = new Plan();

            // Step 1: Create a step with Applied status and rollback info
            var step1 = new PlanStep
            {
                StepId = "step-1",
                Order = 1,
                Description = "Applied step for rollback",
                Action = new OfficeAction
                {
                    ActionId = "act-1",
                    Host = "Excel",
                    Operation = "excel.write_formula",
                    Target = new ActionTarget { Range = "B2" },
                    Parameters = new Dictionary<string, object> { { "formula", "=1+1" } },
                    RiskLevel = 2,
                    RequiresApproval = false,
                    Status = OfficeActionStatus.Applied,
                    Rollback = new RollbackInfo("mock_success")
                    {
                        IsRollbackPossible = true,
                        CapturedAt = DateTime.UtcNow
                    }
                },
                TargetHost = "Excel",
                Status = PlanStepStatus.Applied
            };
            plan.Steps.Add(step1);

            var executor = new PlanExecutor(plan, new ExcelController(null));

            // Call RollbackAll
            var result = executor.RollbackAll();

            // Result should indicate success (we're using mock_success strategy)
            Assert(result.Success, "RollbackAll should succeed with mock_success strategy");

            // Executor state should be RolledBack
            Assert(executor.State == PlanExecutionState.RolledBack, "Executor state should be RolledBack after successful RollbackAll");

            // PlanStep status must track the action-level rollback
            Assert(step1.Status == PlanStepStatus.RolledBack,
                "PlanStep.Status should be RolledBack after successful RollbackAll");

            Console.WriteLine("  [PASS] RollbackAll delegates to RollbackExecutor.RollbackBatch and sets state to RolledBack");
        }

        private static void TestRollbackAllHealsActionStatusAfterRoundTripAndUpdatesPlanStep()
        {
            // OfficeAction.Status was historically [JsonIgnore]; even with serialization fixed,
            // PlanStep.Status is the plan-level authority. Simulate a stale Action.Status=Pending
            // while PlanStep is Applied (the post-round-trip footgun) and confirm RollbackAll
            // still rolls back rather than silently succeeding with "No applied actions".
            var plan = new Plan();
            var action = new OfficeAction
            {
                ActionId = "act-heal",
                Host = "Excel",
                Operation = "excel.write_formula",
                Target = new ActionTarget { Range = "B2" },
                Status = OfficeActionStatus.Pending, // stale — as if JsonIgnore wiped Applied
                Rollback = new RollbackInfo("mock_success")
                {
                    IsRollbackPossible = true,
                    CapturedAt = DateTime.UtcNow
                }
            };
            var step = new PlanStep
            {
                StepId = "step-heal",
                Order = 1,
                Description = "Applied step with stale Action.Status",
                Action = action,
                TargetHost = "Excel",
                Status = PlanStepStatus.Applied
            };
            plan.Steps.Add(step);

            var executor = new PlanExecutor(plan, new ExcelController(null));
            var result = executor.RollbackAll();

            Assert(result.Success, "RollbackAll should succeed after healing Action.Status from PlanStep");
            Assert(action.Status == OfficeActionStatus.RolledBack,
                "Action.Status should be RolledBack after mock rollback");
            Assert(step.Status == PlanStepStatus.RolledBack,
                "PlanStep.Status should be RolledBack after successful RollbackAll");
            Assert(executor.State == PlanExecutionState.RolledBack,
                "Executor state should be RolledBack");
            Assert(result.ErrorMessage == null || !result.ErrorMessage.Contains("No applied actions"),
                "Must not silently no-op with 'No applied actions' when PlanStep is Applied");

            Console.WriteLine("  [PASS] RollbackAll heals stale Action.Status from PlanStep and updates PlanStep to RolledBack");
        }

        private static void TestBusyRetryableStateRetriesUpTo3Times()
        {
            // powerpoint.set_notes is RiskLevel 1, so ExecuteStep skips the RiskLevel>=2
            // BeforeState capture gate that blocks excel.write_formula under a null COM app.
            // SetSpeakerNotesForSlide is virtual — a test subclass can throw COMException
            // 0x800AC472 so PostVerify returns HostBusyRetryable and the retry loop runs.
            var controller = new BusyRetryPowerPointController();
            var plan = new Plan();
            var step = new PlanStep
            {
                StepId = "step-busy",
                Order = 1,
                Description = "Force HostBusyRetryable via subclassed PowerPointController",
                Action = new OfficeAction
                {
                    ActionId = "act-busy",
                    Host = "PowerPoint",
                    Operation = "powerpoint.set_notes",
                    Target = new ActionTarget { Slide = 1 },
                    Parameters = new Dictionary<string, object> { { "slide", 1 }, { "notes", "retry me" } }
                },
                TargetHost = "PowerPoint",
                Status = PlanStepStatus.Approved
            };
            plan.Steps.Add(step);

            var executor = new PlanExecutor(plan, controller);
            var result = executor.ExecuteStep(1);

            Assert(controller.CallCount == 3,
                string.Format("ExecuteStep must retry exactly 3 times on HostBusyRetryable, got {0}", controller.CallCount));
            Assert(step.Status == PlanStepStatus.Failed,
                "Step should be Failed after exhausting HostBusyRetryable retries");
            Assert(executor.State == PlanExecutionState.Failed,
                "Executor state should be Failed after retry exhaustion");
            Assert(!result.Success, "ExecuteStep should report failure after retry exhaustion");
            Assert(step.ErrorMessage != null && step.ErrorMessage.Contains("retries"),
                "ErrorMessage should mention retries");

            Console.WriteLine("  [PASS] HostBusyRetryable path retries exactly 3 times then marks Failed");
        }

        private static void TestPostVerifyNonRetryableFailureSetsExecutorFailed()
        {
            // Covers the PostVerify non-retryable failure branch (distinct from HostBusyRetryable
            // retry exhaustion). SetSpeakerNotesForSlide returning false yields ErrorCode 0,
            // which PostVerify classifies as ExecutionError — not HostBusyRetryable.
            var controller = new NonRetryFailPowerPointController();
            var plan = new Plan();
            var step = new PlanStep
            {
                StepId = "step-postverify",
                Order = 1,
                Description = "Non-retryable PostVerify failure",
                Action = new OfficeAction
                {
                    ActionId = "act-pv",
                    Host = "PowerPoint",
                    Operation = "powerpoint.set_notes",
                    Target = new ActionTarget { Slide = 1 },
                    Parameters = new Dictionary<string, object> { { "slide", 1 }, { "notes", "x" } }
                },
                TargetHost = "PowerPoint",
                Status = PlanStepStatus.Approved
            };
            plan.Steps.Add(step);

            // Step 2 must not run if ExecuteAll halts on _state=Failed from step 1
            var step2 = new PlanStep
            {
                StepId = "step-after",
                Order = 2,
                Description = "Must remain Pending",
                TargetHost = "PowerPoint",
                Status = PlanStepStatus.Pending
            };
            plan.Steps.Add(step2);

            var executor = new PlanExecutor(plan, controller);
            executor.ExecuteAll(null);

            Assert(controller.CallCount == 1, "Non-retryable failure must not retry");
            Assert(step.Status == PlanStepStatus.Failed, "Step 1 should be Failed");
            Assert(executor.State == PlanExecutionState.Failed,
                "Executor _state must be Failed so ExecuteAll halts (PostVerify non-retryable branch)");
            Assert(step2.Status == PlanStepStatus.Pending,
                "ExecuteAll must not continue past PostVerify non-retryable failure");

            Console.WriteLine("  [PASS] PostVerify non-retryable failure sets executor Failed and halts ExecuteAll");
        }

        /// <summary>
        /// Test double: SetSpeakerNotesForSlide is virtual on PowerPointController.
        /// Throws 0x800AC472 so ActionVerifier.PostVerify classifies HostBusyRetryable.
        /// </summary>
        private class BusyRetryPowerPointController : PowerPointController
        {
            public int CallCount;

            public BusyRetryPowerPointController()
                : base(null)
            {
            }

            public override bool SetSpeakerNotesForSlide(int slideNumber, string notes)
            {
                CallCount++;
                throw new System.Runtime.InteropServices.COMException(
                    "Excel busy",
                    unchecked((int)0x800AC472));
            }
        }

        private class NonRetryFailPowerPointController : PowerPointController
        {
            public int CallCount;

            public NonRetryFailPowerPointController()
                : base(null)
            {
            }

            public override bool SetSpeakerNotesForSlide(int slideNumber, string notes)
            {
                CallCount++;
                return false;
            }
        }
    }
}
