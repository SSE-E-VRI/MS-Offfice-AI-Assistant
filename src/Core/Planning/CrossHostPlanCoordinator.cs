using System;
using System.Collections.Generic;

namespace MSOfficeAIAssistant.Core.Planning
{
    /// <summary>
    /// Coordinates execution of multi-host plans, allowing a Plan to be executed sequentially
    /// across different Office applications. When execution reaches a step targeting a DIFFERENT
    /// host than the current one, execution pauses and persists via WorkSession for later resumption
    /// in that host's context.
    ///
    /// Execution strategy (ExecuteForCurrentHost):
    /// - Execute ONLY steps whose TargetHost matches the current host, plus any reasoning-only steps
    /// - Walk steps in Order sequence, skipping (not attempting) any steps for a DIFFERENT host
    /// - CONTINUE past a skipped different-host step to execute any same-host steps later in the
    ///   sequence (do NOT stop at the first different-host step)
    /// - STOP at the first in-scope step that leaves executor.State as AwaitingApproval or Failed
    ///   (matches PlanExecutor.ExecuteAll's own stop condition - a failed step is never silently
    ///   skipped past, since a later same-host step may depend on state it never produced)
    /// - Detect when all steps for the current host are complete and the plan is waiting for a
    ///   different host to continue
    /// </summary>
    public static class CrossHostPlanCoordinator
    {
        /// <summary>
        /// Determines if a plan is multi-host: has at least one step targeting a DIFFERENT host
        /// than the others, among all incomplete (non-Applied, non-RolledBack, non-Skipped) steps.
        /// Reasoning-only steps' TargetHost values are considered; a plan with reasoning-only steps
        /// on different hosts is considered multi-host.
        /// </summary>
        public static bool IsMultiHost(Plan plan)
        {
            if (plan == null || plan.Steps == null || plan.Steps.Count == 0)
            {
                return false;
            }

            // Collect distinct hosts from all INCOMPLETE steps (not terminal)
            string firstHost = null;
            foreach (var step in plan.Steps)
            {
                // Skip terminal statuses
                if (step.Status == PlanStepStatus.Applied || step.Status == PlanStepStatus.RolledBack || step.Status == PlanStepStatus.Skipped)
                {
                    continue;
                }

                // Skip steps with no host affinity
                if (string.IsNullOrEmpty(step.TargetHost))
                {
                    continue;
                }

                // First incomplete step sets the baseline host
                if (firstHost == null)
                {
                    firstHost = step.TargetHost;
                    continue;
                }

                // If any step targets a different host, plan is multi-host
                if (step.TargetHost != firstHost)
                {
                    return true;
                }
            }

            return false;
        }

        /// <summary>
        /// Executes only the steps in the plan whose TargetHost matches currentHost.
        /// Walks steps in Order sequence, skipping (not attempting) steps targeting a DIFFERENT host,
        /// and CONTINUING past those to execute any further same-host steps.
        ///
        /// Stops early only when an in-scope step leaves executor.State as AwaitingApproval
        /// (user approval required) or Failed (matches PlanExecutor.ExecuteAll's own stop
        /// condition - resume via PlanExecutor.ContinueFromStep after review).
        ///
        /// Returns a CrossHostExecutionResult describing the outcome: whether the plan is fully
        /// complete, paused for a different host, and the number of steps executed on this host.
        /// </summary>
        public static CrossHostExecutionResult ExecuteForCurrentHost(Plan plan, PlanExecutor executor, string currentHost)
        {
            if (plan == null || plan.Steps == null || executor == null)
            {
                return new CrossHostExecutionResult
                {
                    PlanFullyComplete = false,
                    PausedForDifferentHost = false,
                    NextHost = null,
                    StepsExecutedOnThisHost = 0,
                    StatusMessage = "Invalid plan or executor provided."
                };
            }

            int stepsExecuted = 0;

            // Execute in Order sequence, skipping out-of-scope steps, stopping at approval/failure
            foreach (var step in plan.Steps)
            {
                // Skip already-terminal steps
                if (step.Status == PlanStepStatus.Applied || step.Status == PlanStepStatus.RolledBack)
                {
                    continue;
                }

                // Determine if this step targets the current host
                // A step targets the current host if its TargetHost matches currentHost,
                // or if TargetHost is empty/null (no specific host affinity).
                bool isForCurrentHost = string.IsNullOrEmpty(step.TargetHost) || step.TargetHost == currentHost;

                // Skip steps for a DIFFERENT host without attempting them
                if (!isForCurrentHost)
                {
                    continue;
                }

                // This step targets the current host - execute it
                var result = executor.ExecuteStep(step.Order);
                stepsExecuted++;

                // Stop at AwaitingApproval (user action required) or Failed (matches
                // PlanExecutor.ExecuteAll's own stop condition from Phase D2 - a failed step
                // must not be silently skipped past, since a later same-host step may depend on
                // state the failed step never produced). The user resumes via
                // PlanExecutor.ContinueFromStep after reviewing/fixing the failure.
                if (executor.State == PlanExecutionState.AwaitingApproval ||
                    executor.State == PlanExecutionState.Failed)
                {
                    break;
                }
            }

            // Determine the plan's next state
            string nextHost = GetNextPendingHost(plan);
            bool planFullyComplete = (nextHost == null);
            bool pausedForDifferentHost = false;

            if (!planFullyComplete && !string.IsNullOrEmpty(nextHost) && nextHost != currentHost)
            {
                pausedForDifferentHost = true;
            }

            // Build status message
            string statusMessage = ComputeStatus(plan, executor, currentHost);

            return new CrossHostExecutionResult
            {
                PlanFullyComplete = planFullyComplete,
                PausedForDifferentHost = pausedForDifferentHost,
                NextHost = pausedForDifferentHost ? nextHost : null,
                StepsExecutedOnThisHost = stepsExecuted,
                StatusMessage = statusMessage
            };
        }

        /// <summary>
        /// Returns the TargetHost of the NEXT step to be executed in the plan (by Order, walking
        /// forward from step 1). Returns null if there are no remaining steps awaiting execution
        /// or if the plan/steps are null/empty.
        ///
        /// A step is considered "awaiting execution" if its Status is Pending or Approved.
        /// Failed steps are not considered (they remain on the current host for retry/resolution).
        /// </summary>
        public static string GetNextPendingHost(Plan plan)
        {
            if (plan == null || plan.Steps == null || plan.Steps.Count == 0)
            {
                return null;
            }

            // Walk steps in Order sequence
            var stepList = new List<PlanStep>(plan.Steps);
            stepList.Sort(delegate(PlanStep a, PlanStep b) { return a.Order.CompareTo(b.Order); });

            foreach (var step in stepList)
            {
                // Find the first step that is Pending or Approved (awaiting execution)
                if (step.Status == PlanStepStatus.Pending || step.Status == PlanStepStatus.Approved)
                {
                    return step.TargetHost;
                }
            }

            return null;
        }

        /// <summary>
        /// Builds a WorkSession.Status vocabulary string describing the current state of the plan
        /// given its execution progress. Used to communicate multi-host workflow state to the UI.
        ///
        /// Rules (in order):
        /// 1. If GetNextPendingHost returns null and no step is Failed: "Completed"
        /// 2. Else if any step Status == Failed: "Failed"
        /// 3. Else if GetNextPendingHost returns a non-null host DIFFERENT from currentHost:
        ///    "Awaiting host: {host}"
        /// 4. Else if executor.State == AwaitingApproval: "Awaiting approval"
        /// 5. Else: "In progress"
        /// </summary>
        public static string ComputeStatus(Plan plan, PlanExecutor executor, string currentHost)
        {
            if (plan == null || plan.Steps == null)
            {
                return "Unknown";
            }

            // Check if plan is fully complete (all steps terminal)
            string nextPendingHost = GetNextPendingHost(plan);
            bool hasPendingSteps = (nextPendingHost != null);
            bool hasFailedStep = false;

            foreach (var step in plan.Steps)
            {
                if (step.Status == PlanStepStatus.Failed)
                {
                    hasFailedStep = true;
                    break;
                }
            }

            // Rule 1: No pending steps and no failures = Completed
            if (!hasPendingSteps && !hasFailedStep)
            {
                return "Completed";
            }

            // Rule 2: Any failed step = Failed
            if (hasFailedStep)
            {
                return "Failed";
            }

            // Rule 3: Pending steps exist for a DIFFERENT host
            if (hasPendingSteps && nextPendingHost != currentHost)
            {
                return string.Format("Awaiting host: {0}", nextPendingHost);
            }

            // Rule 4: Executor is awaiting approval
            if (executor != null && executor.State == PlanExecutionState.AwaitingApproval)
            {
                return "Awaiting approval";
            }

            // Rule 5: Else in progress
            return "In progress";
        }
    }

    /// <summary>
    /// Result of ExecuteForCurrentHost: describes the outcome of executing a plan's
    /// host-specific subset of steps.
    /// </summary>
    public class CrossHostExecutionResult
    {
        /// <summary>
        /// True if every step in the plan is now in a terminal state (Applied/Skipped/RolledBack)
        /// and no step is Failed. The plan is fully done.
        /// </summary>
        public bool PlanFullyComplete { get; set; }

        /// <summary>
        /// True if there are remaining incomplete steps targeting a DIFFERENT host than currentHost.
        /// When true, the caller should persist the session and notify the user to switch to the
        /// host indicated by NextHost.
        /// </summary>
        public bool PausedForDifferentHost { get; set; }

        /// <summary>
        /// The host that the next incomplete step targets (e.g. "Word", "Excel"), or null if
        /// PausedForDifferentHost is false or the plan is fully complete.
        /// </summary>
        public string NextHost { get; set; }

        /// <summary>
        /// Number of steps actually executed (attempted via ExecuteStep) on this host during
        /// this call. Does not include skipped steps for other hosts.
        /// </summary>
        public int StepsExecutedOnThisHost { get; set; }

        /// <summary>
        /// Human-readable status message describing the outcome, suitable for WorkSession.Status.
        /// Examples: "Completed", "Awaiting host: Word", "Awaiting approval", "Failed".
        /// </summary>
        public string StatusMessage { get; set; }
    }
}
