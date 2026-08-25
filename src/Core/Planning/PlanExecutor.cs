using System;
using System.Collections.Generic;
using System.Threading;
using MSOfficeAIAssistant.Core.Actions;
using MSOfficeAIAssistant.Hosts;

namespace MSOfficeAIAssistant.Core.Planning
{
    /// <summary>
    /// Execution state of the plan executor.
    /// </summary>
    public enum PlanExecutionState
    {
        Queued,
        Running,
        AwaitingApproval,
        Paused,
        Completed,
        Failed,
        Cancelled,
        RolledBack
    }

    /// <summary>
    /// A snapshot of execution progress handed to progress callbacks after each step.
    /// Not persisted, used for live UI updates.
    /// </summary>
    public class PlanExecutionProgress
    {
        public string PlanId { get; set; }
        public int CurrentStepOrder { get; set; }
        public int TotalSteps { get; set; }
        public PlanExecutionState State { get; set; }
        public string LastMessage { get; set; }
    }

    /// <summary>
    /// Executes a Plan step-by-step with risk gating, verification, retry logic, and rollback support.
    /// Synchronous, single-threaded execution. Delegates to ToolRegistry.Execute and RollbackExecutor.
    /// </summary>
    public class PlanExecutor
    {
        private Plan _plan;
        private object _controller;
        private PlanExecutionState _state;

        /// <summary>
        /// Constructs a new PlanExecutor with the Plan to execute and the live host controller object.
        /// </summary>
        public PlanExecutor(Plan plan, object controller)
        {
            _plan = plan;
            _controller = controller;
            _state = PlanExecutionState.Queued;
        }

        /// <summary>
        /// Current execution state of this executor.
        /// </summary>
        public PlanExecutionState State
        {
            get { return _state; }
        }

        /// <summary>
        /// Executes a single step by Order, respecting risk gating and verification.
        /// If the step is gated (RiskLevel >= 1 and RequiresApproval is true) and its Status is still Pending,
        /// sets state to AwaitingApproval and returns without mutating anything.
        /// Otherwise executes the step with PreVerify, ToolRegistry.Execute, PostVerify, and retry logic.
        /// Returns the execution result.
        /// </summary>
        public HostOperationResult ExecuteStep(int order)
        {
            if (_plan == null || _plan.Steps == null || _plan.Steps.Count == 0)
            {
                return HostOperationResult.Failed("No plan or steps available for execution.");
            }

            PlanStep step = null;
            foreach (var s in _plan.Steps)
            {
                if (s.Order == order)
                {
                    step = s;
                    break;
                }
            }

            if (step == null)
            {
                return HostOperationResult.Failed(string.Format("Step with Order {0} not found.", order));
            }

            // Risk gating: if step requires approval and is still Pending, await approval
            if (!step.IsReasoningOnly && step.Action != null)
            {
                // Pre-verify to synchronize risk properties
                var preVerifyRes = ActionVerifier.PreVerify(step.Action);
                if (preVerifyRes.RiskLevel >= 1 && preVerifyRes.RequiresApproval && step.Status == PlanStepStatus.Pending)
                {
                    _state = PlanExecutionState.AwaitingApproval;
                    return HostOperationResult.Failed(
                        string.Format("Step {0} requires approval before execution (RiskLevel {1}).", order, preVerifyRes.RiskLevel),
                        0, step.Action.TargetDisplay);
                }
            }

            // Reasoning-only steps: mark as Applied immediately
            if (step.IsReasoningOnly)
            {
                step.Status = PlanStepStatus.Applied;
                step.ResultText = "Reasoning step completed.";
                return HostOperationResult.Ok("Reasoning step applied.", string.Empty);
            }

            if (step.Action == null)
            {
                return HostOperationResult.Failed("Step has no action to execute.", 0, string.Empty);
            }

            _state = PlanExecutionState.Running;
            step.Status = PlanStepStatus.Applying;

            // Pre-verify the action
            var preVerifyResult = ActionVerifier.PreVerify(step.Action);
            if (!preVerifyResult.IsValid)
            {
                step.Status = PlanStepStatus.Failed;
                _state = PlanExecutionState.Failed;
                step.ErrorMessage = string.Join("; ", preVerifyResult.ValidationErrors);
                return HostOperationResult.Failed(step.ErrorMessage, 0, step.Action.TargetDisplay);
            }

            // Before high-risk undoable steps, capture BeforeState
            if (step.Action.RiskLevel >= 2 && step.Action.IsUndoable)
            {
                var captureRes = RollbackExecutor.CaptureBeforeState(_controller, step.Action);
                if (!captureRes.Success && step.Action.RiskLevel >= 2)
                {
                    step.Status = PlanStepStatus.Failed;
                    _state = PlanExecutionState.Failed;
                    step.ErrorMessage = "Failed to capture BeforeState for rollback: " + captureRes.ErrorMessage;
                    return HostOperationResult.Failed(step.ErrorMessage, captureRes.ErrorCode, step.Action.TargetDisplay);
                }
            }

            // Execute the action with retry logic for HostBusyRetryable
            HostOperationResult execResult = null;
            int retryCount = 0;
            const int maxRetries = 3;
            const int retryDelayMs = 500;

            while (retryCount < maxRetries)
            {
                execResult = ToolRegistry.Execute(_controller, step.Action);
                var postVerifyResult = ActionVerifier.PostVerify(step.Action, execResult);

                if (postVerifyResult.Outcome == VerificationOutcome.HostBusyRetryable)
                {
                    retryCount++;
                    if (retryCount < maxRetries)
                    {
                        Thread.Sleep(retryDelayMs);
                        continue;
                    }
                    else
                    {
                        // Exceeded retries
                        step.Status = PlanStepStatus.Failed;
                        _state = PlanExecutionState.Failed;
                        step.ErrorMessage = string.Format("Host remained busy after {0} retries: {1}", maxRetries, postVerifyResult.DiagnosticMessage);
                        return HostOperationResult.Failed(step.ErrorMessage, execResult.ErrorCode, step.Action.TargetDisplay);
                    }
                }
                else
                {
                    // Success or non-retryable failure
                    if (postVerifyResult.Verified)
                    {
                        step.Status = PlanStepStatus.Applied;
                        step.ResultText = postVerifyResult.ObservedValue ?? postVerifyResult.DiagnosticMessage;
                        return HostOperationResult.Ok(step.ResultText, step.Action.TargetDisplay);
                    }
                    else
                    {
                        step.Status = PlanStepStatus.Failed;
                        _state = PlanExecutionState.Failed;
                        step.ErrorMessage = postVerifyResult.DiagnosticMessage;
                        return HostOperationResult.Failed(step.ErrorMessage, execResult.ErrorCode, step.Action.TargetDisplay);
                    }
                }
            }

            step.Status = PlanStepStatus.Failed;
            _state = PlanExecutionState.Failed;
            step.ErrorMessage = "Unexpected execution failure.";
            return HostOperationResult.Failed(step.ErrorMessage, 0, step.Action.TargetDisplay);
        }

        /// <summary>
        /// Executes all remaining Pending/Approved steps in Order sequence.
        /// Stops as soon as a step becomes AwaitingApproval or Failed.
        /// Invokes the optional progress callback after every step attempt.
        /// </summary>
        public void ExecuteAll(Action<PlanExecutionProgress> onProgress)
        {
            if (_plan == null || _plan.Steps == null)
            {
                return;
            }

            _state = PlanExecutionState.Running;

            foreach (var step in _plan.Steps)
            {
                if (step.Status == PlanStepStatus.Applied || step.Status == PlanStepStatus.RolledBack)
                {
                    continue; // Skip already-applied steps
                }

                var result = ExecuteStep(step.Order);

                if (onProgress != null)
                {
                    onProgress(new PlanExecutionProgress
                    {
                        PlanId = _plan.PlanId,
                        CurrentStepOrder = step.Order,
                        TotalSteps = _plan.TotalStepCount,
                        State = _state,
                        LastMessage = result.ErrorMessage ?? (result.Value != null ? Convert.ToString(result.Value) : "Step executed.")
                    });
                }

                if (_state == PlanExecutionState.AwaitingApproval || _state == PlanExecutionState.Failed)
                {
                    break;
                }
            }

            if (_state == PlanExecutionState.Running)
            {
                _state = PlanExecutionState.Completed;
            }
        }

        /// <summary>
        /// Resumes execution starting at a specific step Order (e.g. after approving a gated step or retrying a failed step).
        /// Same stop conditions as ExecuteAll.
        /// </summary>
        public void ContinueFromStep(int order, Action<PlanExecutionProgress> onProgress)
        {
            if (_plan == null || _plan.Steps == null)
            {
                return;
            }

            _state = PlanExecutionState.Running;

            bool found = false;
            foreach (var step in _plan.Steps)
            {
                if (step.Order < order)
                {
                    continue;
                }

                found = true;

                if (step.Status == PlanStepStatus.Applied || step.Status == PlanStepStatus.RolledBack)
                {
                    continue; // Skip already-applied steps
                }

                var result = ExecuteStep(step.Order);

                if (onProgress != null)
                {
                    onProgress(new PlanExecutionProgress
                    {
                        PlanId = _plan.PlanId,
                        CurrentStepOrder = step.Order,
                        TotalSteps = _plan.TotalStepCount,
                        State = _state,
                        LastMessage = result.ErrorMessage ?? (result.Value != null ? Convert.ToString(result.Value) : "Step executed.")
                    });
                }

                if (_state == PlanExecutionState.AwaitingApproval || _state == PlanExecutionState.Failed)
                {
                    break;
                }
            }

            if (!found)
            {
                _state = PlanExecutionState.Failed;
            }
            else if (_state == PlanExecutionState.Running)
            {
                _state = PlanExecutionState.Completed;
            }
        }

        /// <summary>
        /// Sets State to Cancelled. Does not roll back already-Applied steps.
        /// Simply stops further execution. Any step that was Applying at the moment of cancellation
        /// is left in whatever terminal state ExecuteStep already resolved it to.
        /// </summary>
        public void Cancel()
        {
            _state = PlanExecutionState.Cancelled;
        }

        /// <summary>
        /// Rolls back every Applied step in the plan, strict LIFO.
        /// Delegates to RollbackExecutor.RollbackBatch.
        /// On success sets State to RolledBack.
        /// Returns the structured diagnostic from RollbackBatch.
        /// </summary>
        public HostOperationResult RollbackAll()
        {
            if (_plan == null || _plan.Steps == null)
            {
                return HostOperationResult.Failed("No plan or steps available for rollback.");
            }

            var appliedSteps = new List<PlanStep>();
            var appliedActions = new List<OfficeAction>();
            foreach (var s in _plan.Steps)
            {
                if (s.Status == PlanStepStatus.Applied && s.Action != null)
                {
                    // PlanStep.Status is the plan-level authority (and is what WorkSession
                    // persists). Heal Action.Status here so a stale Pending value (e.g. from
                    // an older session file written when Status was [JsonIgnore]) cannot make
                    // RollbackBatch's OfficeActionStatus.Applied filter silently no-op.
                    if (s.Action.Status != OfficeActionStatus.Applied)
                    {
                        s.Action.Status = OfficeActionStatus.Applied;
                    }
                    appliedSteps.Add(s);
                    appliedActions.Add(s.Action);
                }
            }

            var result = RollbackExecutor.RollbackBatch(_controller, appliedActions);

            if (result.Success)
            {
                // Keep PlanStep status in sync with action-level RolledBack (RollbackBatch
                // updates OfficeAction.Status but historically left PlanStep.Status Applied).
                foreach (var s in appliedSteps)
                {
                    if (s.Action != null && s.Action.Status == OfficeActionStatus.RolledBack)
                    {
                        s.Status = PlanStepStatus.RolledBack;
                    }
                }
                _state = PlanExecutionState.RolledBack;
            }

            return result;
        }
    }
}
