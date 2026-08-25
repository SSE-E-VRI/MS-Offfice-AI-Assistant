using System;
using System.Collections.Generic;
using System.Linq;
using MSOfficeAIAssistant.Core.Actions;

namespace MSOfficeAIAssistant.Core.Planning
{
    /// <summary>
    /// Static orchestration class for deterministic Plan construction and editing.
    /// Wraps extracted actions into ordered, editable plans without touching prompt construction.
    /// </summary>
    public static class Planner
    {
        /// <summary>
        /// Builds a Plan from a sequence of already-extracted OfficeActions.
        /// Does not call any AI provider or modify prompt text - pure transformation over existing data.
        /// </summary>
        /// <param name="title">Human-readable plan title</param>
        /// <param name="sourceRequest">Original user prompt that produced the actions</param>
        /// <param name="actions">Extracted actions to convert into plan steps</param>
        /// <returns>A new Plan with steps in the same order as input actions, Order values 1..N</returns>
        public static Plan BuildPlanFromActions(string title, string sourceRequest, IEnumerable<OfficeAction> actions)
        {
            var plan = new Plan();
            plan.Title = title;
            plan.SourceRequest = sourceRequest;

            if (actions == null)
            {
                return plan;
            }

            int order = 1;
            foreach (var action in actions)
            {
                if (action == null) continue;

                var step = new PlanStep
                {
                    Order = order,
                    Action = action,
                    TargetHost = action.Host,
                    Description = action.ExpectedResult ?? action.Operation ?? "Unnamed action",
                    Status = PlanStepStatus.Pending
                };

                plan.Steps.Add(step);
                order++;
            }

            return plan;
        }

        /// <summary>
        /// Inserts a reasoning-only step (no Action) at a given position.
        /// Used to add narrative context between action steps.
        /// </summary>
        /// <param name="plan">The plan to modify</param>
        /// <param name="atOrder">The desired Order position for the new step (1-based)</param>
        /// <param name="description">Human-readable narrative text</param>
        /// <param name="targetHost">Host context, e.g. "Excel" for analysis steps</param>
        public static void InsertReasoningStep(Plan plan, int atOrder, string description, string targetHost)
        {
            if (plan == null || plan.Steps == null) return;

            // Create the reasoning step with no Action
            var reasoningStep = new PlanStep
            {
                Order = atOrder,
                Action = null,
                TargetHost = targetHost,
                Description = description,
                Status = PlanStepStatus.Pending
            };

            // Insert at the appropriate position
            if (atOrder < 1)
            {
                // Insert at beginning
                plan.Steps.Insert(0, reasoningStep);
            }
            else if (atOrder > plan.Steps.Count)
            {
                // Insert at end
                plan.Steps.Add(reasoningStep);
            }
            else
            {
                // Insert in middle, shifting existing steps
                plan.Steps.Insert(atOrder - 1, reasoningStep);
            }

            // Renumber all steps to maintain 1..N sequence with no gaps
            plan.RenumberSteps();
            plan.UpdatedUtc = DateTime.UtcNow;
        }

        /// <summary>
        /// Determines if a plan can be executed without any approval gates.
        /// Returns true only if every Pending step has Action == null OR Action.RiskLevel == 0.
        /// </summary>
        /// <param name="plan">The plan to check</param>
        /// <returns>True if the entire plan is safe to auto-run; false if any risk >= 1 action is pending</returns>
        public static bool IsFullyAutoRunnable(Plan plan)
        {
            if (plan == null || plan.Steps == null) return true;

            foreach (var step in plan.Steps)
            {
                // Only check Pending steps - Applied/Failed steps don't gate execution
                if (step.Status != PlanStepStatus.Pending) continue;

                // Reasoning-only steps are safe
                if (step.IsReasoningOnly) continue;

                // Action steps with RiskLevel 0 are safe
                if (step.Action != null && step.Action.RiskLevel == 0) continue;

                // Any other case is unsafe (RiskLevel >= 1)
                return false;
            }

            return true;
        }

        /// <summary>
        /// Validates that a plan is internally consistent.
        /// Checks Order sequence (1..N with no gaps/duplicates after RenumberSteps),
        /// and verifies every action Operation is resolvable via ToolRegistry for its host.
        /// </summary>
        /// <param name="plan">The plan to validate</param>
        /// <returns>List of human-readable validation error strings; empty list = valid</returns>
        public static List<string> Validate(Plan plan)
        {
            var errors = new List<string>();

            if (plan == null)
            {
                errors.Add("Plan cannot be null");
                return errors;
            }

            if (plan.Steps == null)
            {
                errors.Add("Plan.Steps cannot be null");
                return errors;
            }

            // Check Order sequence: should be 1..N with no gaps or duplicates
            if (plan.Steps.Count > 0)
            {
                var orders = plan.Steps.Select(s => s.Order).ToList();
                orders.Sort();

                // Check for duplicates
                var uniqueOrders = new HashSet<int>(orders);
                if (uniqueOrders.Count != orders.Count)
                {
                    errors.Add("Step Order values contain duplicates");
                }

                // Check for gaps
                for (int i = 0; i < orders.Count; i++)
                {
                    if (orders[i] != i + 1)
                    {
                        errors.Add(string.Format("Step Order sequence has a gap: expected {0}, found {1}", i + 1, orders[i]));
                        break;
                    }
                }
            }

            // Validate each step's action
            foreach (var step in plan.Steps)
            {
                if (step.Action == null) continue; // Reasoning steps need no validation

                string operation = step.Action.Operation;
                string host = step.Action.Host;

                if (string.IsNullOrEmpty(operation))
                {
                    errors.Add(string.Format("Step {0} ({1}): Operation cannot be empty", step.Order, step.Description));
                    continue;
                }

                if (string.IsNullOrEmpty(host))
                {
                    errors.Add(string.Format("Step {0} ({1}): Host cannot be empty", step.Order, step.Description));
                    continue;
                }

                // Verify the operation is registered in ToolRegistry
                var tool = ToolRegistry.GetTool(operation, host);
                if (tool == null)
                {
                    errors.Add(string.Format("Step {0} ({1}): Operation '{2}' not found in ToolRegistry for host '{3}'",
                        step.Order, step.Description, operation, host));
                }
            }

            return errors;
        }
    }
}
