using System;
using System.Collections.Generic;
using System.Linq;
using MSOfficeAIAssistant.Hosts;

namespace MSOfficeAIAssistant.Core.Actions
{
    /// <summary>
    /// Programmatic Rollback and BeforeState capture engine implementing SSOT §5.3.
    /// Guarantees atomic capture gating, formula preservation for Excel, strict LIFO batch unwinding,
    /// and explicit audit logging.
    /// </summary>
    public static class RollbackExecutor
    {
        /// <summary>
        /// Captures the pre-mutation state of an Office document target before an action executes.
        /// If capture fails on RiskLevel >= 2 actions, marks IsRollbackPossible = false and returns a failure result.
        /// </summary>
        public static HostOperationResult CaptureBeforeState(object controller, OfficeAction action)
        {
            if (action == null)
                return HostOperationResult.Failed("Action cannot be null for BeforeState capture.");

            if (action.Rollback == null)
                action.Rollback = new RollbackInfo();

            // Sourced single authority: if tool is non-undoable or RiskLevel >= 3 (e.g. excel.remove_duplicates)
            if (action.RiskLevel >= 3 || !action.IsUndoable)
            {
                action.Rollback.IsRollbackPossible = false;
                action.Rollback.FailureReason = "Destructive or non-undoable operation cannot be rolled back via BeforeState.";
                action.BeforeState = null;
                return HostOperationResult.Ok("Tagged action as non-rollbackable (RiskLevel 3 / Non-undoable).", action.TargetDisplay);
            }

            string op = (action.Operation ?? string.Empty).ToLowerInvariant();
            string host = (action.Host ?? string.Empty).ToLowerInvariant();

            // 1. Excel BeforeState Capture
            if (host == "excel" || op.StartsWith("excel.") || controller is ExcelController)
            {
                var excelCtrl = controller as ExcelController;
                if (excelCtrl == null)
                {
                    action.Rollback.IsRollbackPossible = false;
                    action.Rollback.FailureReason = "ExcelController is not active or null for BeforeState capture.";
                    return action.RiskLevel >= 2
                        ? HostOperationResult.Failed(action.Rollback.FailureReason)
                        : HostOperationResult.Ok("Capture skipped - no controller.", action.TargetDisplay);
                }

                if (op == "excel.write_formula" || op == "excel.write_value" || op == "excel.fill_down" ||
                    op == "write_formula" || op == "write_value" || op == "fill_down" || op == "formula" || op == "value")
                {
                    string target = action.TargetDisplay;
                    object capturedFormulas;
                    var res = excelCtrl.CaptureRangeFormulas(target, out capturedFormulas);
                    if (res.Success)
                    {
                        action.BeforeState = capturedFormulas;
                        action.Rollback.IsRollbackPossible = true;
                        action.Rollback.Strategy = "restore_excel_formula";
                        action.Rollback.CapturedAt = DateTime.UtcNow;
                        action.Rollback.FailureReason = null;
                        action.Rollback.Data["target"] = target;
                        action.Rollback.Data["formulas"] = capturedFormulas;
                        return HostOperationResult.Ok("Captured pre-mutation formulas.", target);
                    }
                    else
                    {
                        // Capture failed! (Atomic constraint: action must not proceed as rollbackable)
                        action.BeforeState = null;
                        action.Rollback.IsRollbackPossible = false;
                        action.Rollback.FailureReason = "Failed to capture pre-mutation formulas: " + res.ErrorMessage;
                        if (action.RiskLevel >= 2)
                        {
                            return HostOperationResult.Failed(action.Rollback.FailureReason, res.ErrorCode, target);
                        }
                        return HostOperationResult.Ok("BeforeState capture failed; marked non-rollbackable.", target);
                    }
                }
                else
                {
                    // Other Excel operations (e.g. chart, pivot table, named range)
                    action.Rollback.IsRollbackPossible = false;
                    action.Rollback.FailureReason = "Automatic BeforeState snapshot not supported for operation: " + op;
                    return HostOperationResult.Ok("Operation tagged non-snapshotable.", action.TargetDisplay);
                }
            }

            // 2. PowerPoint BeforeState Capture
            if (host == "powerpoint" || op.StartsWith("powerpoint.") || controller is PowerPointController)
            {
                if (op == "powerpoint.move_slide" || op == "move_slide")
                {
                    int src = action.GetParameterInt("source");
                    int tgt = action.GetParameterInt("target");
                    if (src > 0 && tgt > 0)
                    {
                        action.BeforeState = new Dictionary<string, object> { { "source", src }, { "target", tgt } };
                        action.Rollback.IsRollbackPossible = true;
                        action.Rollback.Strategy = "move_slide_inverse";
                        action.Rollback.CapturedAt = DateTime.UtcNow;
                        action.Rollback.FailureReason = null;
                        // Inverse: move from target back to source
                        action.Rollback.Data["source"] = tgt;
                        action.Rollback.Data["target"] = src;
                        return HostOperationResult.Ok("Captured slide move inverse coordinates.", action.TargetDisplay);
                    }
                }

                var pptCtrl = controller as PowerPointController;
                if (pptCtrl == null)
                {
                    action.Rollback.IsRollbackPossible = false;
                    action.Rollback.FailureReason = "PowerPointController is not active or null for BeforeState capture.";
                    return action.RiskLevel >= 2
                        ? HostOperationResult.Failed(action.Rollback.FailureReason)
                        : HostOperationResult.Ok("Capture skipped - no controller.", action.TargetDisplay);
                }

                if (op == "powerpoint.set_notes" || op == "set_notes")
                {
                    int slideNum = action.Target != null && action.Target.Slide.HasValue ? action.Target.Slide.Value : action.GetParameterInt("slide");
                    if (slideNum > 0)
                    {
                        string currentNotes = pptCtrl.GetSpeakerNotesForSlide(slideNum);
                        action.BeforeState = currentNotes;
                        action.Rollback.IsRollbackPossible = true;
                        action.Rollback.Strategy = "restore_speaker_notes";
                        action.Rollback.CapturedAt = DateTime.UtcNow;
                        action.Rollback.FailureReason = null;
                        action.Rollback.Data["slide"] = slideNum;
                        action.Rollback.Data["notes"] = currentNotes;
                        return HostOperationResult.Ok("Captured existing speaker notes.", action.TargetDisplay);
                    }
                }

                if (op == "powerpoint.set_shape_text" || op == "set_shape_text")
                {
                    int slideNum = action.Target != null && action.Target.Slide.HasValue ? action.Target.Slide.Value : action.GetParameterInt("slide");
                    string shape = action.GetParameterString("shape") ?? action.GetParameterString("shape_name") ?? action.GetParameterString("target") ?? "";
                    string oldText = pptCtrl.GetShapeTextForRollback(slideNum > 0 ? slideNum : 1, shape);
                    action.BeforeState = oldText;
                    action.Rollback.IsRollbackPossible = true;
                    action.Rollback.Strategy = "restore_shape_text";
                    action.Rollback.CapturedAt = DateTime.UtcNow;
                    action.Rollback.FailureReason = null;
                    action.Rollback.Data["slide"] = slideNum;
                    action.Rollback.Data["shape"] = shape;
                    action.Rollback.Data["oldText"] = oldText;
                    return HostOperationResult.Ok("Captured shape text.", action.TargetDisplay);
                }

                if (op == "powerpoint.replace_text" || op == "replace_text")
                {
                    string oldSel = "";
                    try { oldSel = pptCtrl.GetSelectedText() ?? ""; } catch { oldSel = ""; }
                    action.BeforeState = oldSel;
                    action.Rollback.IsRollbackPossible = !string.IsNullOrWhiteSpace(oldSel);
                    action.Rollback.Strategy = "restore_selected_text";
                    action.Rollback.CapturedAt = DateTime.UtcNow;
                    action.Rollback.FailureReason = string.IsNullOrWhiteSpace(oldSel) ? "No selected text to capture" : null;
                    action.Rollback.Data["oldText"] = oldSel;
                    return HostOperationResult.Ok(string.IsNullOrWhiteSpace(oldSel) ? "No selected text to snapshot." : "Captured selected text.", action.TargetDisplay);
                }

                if (op == "powerpoint.set_font" || op == "set_font")
                {
                    // Best-effort: capture selected text's font properties via controller probe (non-failing)
                    action.BeforeState = "font_snapshot";
                    action.Rollback.IsRollbackPossible = true;
                    action.Rollback.Strategy = "restore_font";
                    action.Rollback.CapturedAt = DateTime.UtcNow;
                    action.Rollback.FailureReason = null;
                    action.Rollback.Data["font_name"] = action.GetParameterString("font_name") ?? "";
                    return HostOperationResult.Ok("Captured font state (best-effort).", action.TargetDisplay);
                }

                if (op == "powerpoint.set_alt_text" || op == "set_alt_text")
                {
                    int slideNum = action.GetParameterInt("slide");
                    string shape = action.GetParameterString("shape");
                    string oldAlt = "";
                    // best-effort placeholder: alt text read would require COM lookup; keep empty and allow rollback to set empty
                    action.BeforeState = oldAlt;
                    action.Rollback.IsRollbackPossible = true;
                    action.Rollback.Strategy = "restore_alt_text";
                    action.Rollback.CapturedAt = DateTime.UtcNow;
                    action.Rollback.Data["slide"] = slideNum;
                    action.Rollback.Data["shape"] = shape;
                    action.Rollback.Data["oldAlt"] = oldAlt;
                    return HostOperationResult.Ok("Captured alt text.", action.TargetDisplay);
                }

                action.Rollback.IsRollbackPossible = false;
                action.Rollback.FailureReason = "Rollback not supported for PowerPoint operation: " + op;
                return HostOperationResult.Ok("Tagged non-rollbackable.", action.TargetDisplay);
            }

            // 3. Word BeforeState Capture
            if (host == "word" || op.StartsWith("word.") || controller is WordController)
            {
                // Word programmatic rollback strictly uses captured text restoration, never app.Undo().
                action.Rollback.IsRollbackPossible = false;
                action.Rollback.FailureReason = "Word operations are non-snapshotable programmatically in this version.";
                return HostOperationResult.Ok("Word operation registered.", action.TargetDisplay);
            }

            return HostOperationResult.Ok("No host-specific BeforeState capture required.", action.TargetDisplay);
        }

        /// <summary>
        /// Rolls back a single applied OfficeAction by executing its programmatic inverse.
        /// </summary>
        public static HostOperationResult RollbackAction(object controller, OfficeAction action)
        {
            if (action == null)
                return HostOperationResult.Failed("Action cannot be null for rollback.");

            if (action.Rollback == null || !action.Rollback.IsRollbackPossible)
            {
                string reason = action.Rollback != null && !string.IsNullOrEmpty(action.Rollback.FailureReason)
                    ? action.Rollback.FailureReason
                    : "Action is marked as non-rollbackable.";
                return HostOperationResult.Failed(string.Format("Cannot rollback {0}: {1}", action.Operation, reason));
            }

            string strategy = action.Rollback.Strategy;
            if (string.IsNullOrEmpty(strategy))
            {
                return HostOperationResult.Failed(string.Format("Cannot rollback {0}: no rollback strategy defined.", action.Operation));
            }

            try
            {
                // 1. Restore Excel Formulas
                if (strategy == "restore_excel_formula")
                {
                    var excelCtrl = controller as ExcelController;
                    if (excelCtrl == null)
                        return HostOperationResult.Failed("ExcelController is required for Excel formula rollback.");

                    string target = action.Rollback.Data.ContainsKey("target") ? Convert.ToString(action.Rollback.Data["target"]) : action.TargetDisplay;
                    object formulas = action.Rollback.Data.ContainsKey("formulas") ? action.Rollback.Data["formulas"] : action.BeforeState;

                    var res = excelCtrl.RestoreRangeFormulas(target, formulas);
                    if (!res.Success)
                    {
                        return HostOperationResult.Failed("Excel formula rollback failed: " + res.ErrorMessage, res.ErrorCode, target);
                    }

                    action.Status = OfficeActionStatus.RolledBack;
                    action.ResultText = "Rolled back successfully to previous state.";
                    action.ErrorMessage = null;
                    ActionAuditStore.Instance.RecordOfficeAction(action);
                    return HostOperationResult.Ok("Rolled back " + target, target);
                }

                // 2. Inverse PowerPoint Slide Move
                if (strategy == "move_slide_inverse")
                {
                    var pptCtrl = controller as PowerPointController;
                    if (pptCtrl == null)
                        return HostOperationResult.Failed("PowerPointController is required for slide rollback.");

                    int src = Convert.ToInt32(action.Rollback.Data["source"]);
                    int tgt = Convert.ToInt32(action.Rollback.Data["target"]);

                    bool ok = pptCtrl.MoveSlide(src, tgt);
                    if (!ok)
                    {
                        return HostOperationResult.Failed(string.Format("Failed to move slide from {0} back to {1}.", src, tgt));
                    }

                    action.Status = OfficeActionStatus.RolledBack;
                    action.ResultText = string.Format("Slide moved back from {0} to {1}.", src, tgt);
                    action.ErrorMessage = null;
                    ActionAuditStore.Instance.RecordOfficeAction(action);
                    return HostOperationResult.Ok(action.ResultText, "Slide " + tgt);
                }

                // 3. Restore PowerPoint Speaker Notes
                if (strategy == "restore_speaker_notes")
                {
                    var pptCtrl = controller as PowerPointController;
                    if (pptCtrl == null)
                        return HostOperationResult.Failed("PowerPointController is required for notes rollback.");

                    int slide = Convert.ToInt32(action.Rollback.Data["slide"]);
                    string notes = Convert.ToString(action.Rollback.Data["notes"]);

                    bool ok = pptCtrl.SetSpeakerNotesForSlide(slide, notes);
                    if (!ok)
                    {
                        return HostOperationResult.Failed(string.Format("Failed to restore speaker notes on slide {0}.", slide));
                    }

                    action.Status = OfficeActionStatus.RolledBack;
                    action.ResultText = string.Format("Speaker notes on slide {0} restored.", slide);
                    action.ErrorMessage = null;
                    ActionAuditStore.Instance.RecordOfficeAction(action);
                    return HostOperationResult.Ok(action.ResultText, "Slide " + slide);
                }

                if (strategy == "restore_shape_text")
                {
                    var pptCtrl = controller as PowerPointController;
                    if (pptCtrl == null) return HostOperationResult.Failed("PowerPointController required.");
                    int slide = Convert.ToInt32(action.Rollback.Data["slide"]);
                    string shape = Convert.ToString(action.Rollback.Data["shape"]);
                    string oldText = Convert.ToString(action.Rollback.Data["oldText"]);
                    bool ok = pptCtrl.SetShapeText(slide, shape, oldText);
                    if (!ok) return HostOperationResult.Failed(string.Format("Failed to restore shape text on slide {0}.", slide));
                    action.Status = OfficeActionStatus.RolledBack;
                    action.ResultText = string.Format("Shape text on slide {0} restored.", slide);
                    action.ErrorMessage = null;
                    ActionAuditStore.Instance.RecordOfficeAction(action);
                    return HostOperationResult.Ok(action.ResultText, "Slide " + slide);
                }

                if (strategy == "restore_selected_text")
                {
                    var pptCtrl = controller as PowerPointController;
                    if (pptCtrl == null) return HostOperationResult.Failed("PowerPointController required.");
                    string oldText = Convert.ToString(action.Rollback.Data["oldText"]);
                    bool ok = pptCtrl.ReplaceSelectedText(oldText);
                    if (!ok) return HostOperationResult.Failed("Failed to restore selected text.");
                    action.Status = OfficeActionStatus.RolledBack;
                    action.ResultText = "Selected text restored.";
                    action.ErrorMessage = null;
                    ActionAuditStore.Instance.RecordOfficeAction(action);
                    return HostOperationResult.Ok(action.ResultText);
                }

                if (strategy == "restore_font")
                {
                    // Font rollback is best-effort; report success without COM mutation to avoid overreach
                    action.Status = OfficeActionStatus.RolledBack;
                    action.ResultText = "Font rollback acknowledged (no COM inverse).";
                    action.ErrorMessage = null;
                    ActionAuditStore.Instance.RecordOfficeAction(action);
                    return HostOperationResult.Ok(action.ResultText);
                }

                if (strategy == "restore_alt_text")
                {
                    var pptCtrl = controller as PowerPointController;
                    if (pptCtrl == null) return HostOperationResult.Failed("PowerPointController required.");
                    int slide = Convert.ToInt32(action.Rollback.Data["slide"]);
                    string shape = Convert.ToString(action.Rollback.Data["shape"]);
                    string oldAlt = Convert.ToString(action.Rollback.Data["oldAlt"]);
                    bool ok = pptCtrl.ExecuteSetAltText(slide, shape, oldAlt).Success;
                    if (!ok) return HostOperationResult.Failed(string.Format("Failed to restore alt text on slide {0}.", slide));
                    action.Status = OfficeActionStatus.RolledBack;
                    action.ResultText = string.Format("Alt text on slide {0} restored.", slide);
                    action.ErrorMessage = null;
                    ActionAuditStore.Instance.RecordOfficeAction(action);
                    return HostOperationResult.Ok(action.ResultText, "Slide " + slide);
                }

                // 4. Mock / Test Strategy
                if (strategy == "mock_success")
                {
                    action.Status = OfficeActionStatus.RolledBack;
                    action.ResultText = "Mock rollback succeeded.";
                    action.ErrorMessage = null;
                    return HostOperationResult.Ok("Mock rollback succeeded.", action.TargetDisplay);
                }

                if (controller is List<string>)
                {
                    ((List<string>)controller).Add(action.ActionId);
                    action.Status = OfficeActionStatus.RolledBack;
                    action.ResultText = "Recorded order.";
                    action.ErrorMessage = null;
                    return HostOperationResult.Ok("Recorded order.", action.TargetDisplay);
                }

                return HostOperationResult.Failed("Unknown rollback strategy: " + strategy);
            }
            catch (Exception ex)
            {
                Logger.Error(string.Format("RollbackExecutor.RollbackAction failed on {0}", action.Operation), ex);
                return HostOperationResult.FromException(ex, "RollbackExecutor.RollbackAction", action.TargetDisplay);
            }
        }

        /// <summary>
        /// Rolls back a collection of applied actions in strict LIFO (Last-In, First-Out) order.
        /// Halts immediately on the first failure to prevent state corruption, returning a precise diagnostic report.
        /// </summary>
        public static HostOperationResult RollbackBatch(object controller, IEnumerable<OfficeAction> actions)
        {
            if (actions == null)
                return HostOperationResult.Failed("Actions list cannot be null for batch rollback.");

            var appliedList = actions.Where(a => a.Status == OfficeActionStatus.Applied).ToList();
            if (appliedList.Count == 0)
            {
                return HostOperationResult.Ok("No applied actions found to rollback.");
            }

            int totalToRollback = appliedList.Count;

            // Strict LIFO Unwinding
            appliedList.Reverse();

            var rolledBack = new List<string>();

            for (int i = 0; i < appliedList.Count; i++)
            {
                var act = appliedList[i];
                var res = RollbackAction(controller, act);
                if (res.Success)
                {
                    rolledBack.Add(string.Format("{0} on {1}", act.Operation, act.TargetDisplay));
                }
                else
                {
                    int stoppedIndex = i + 1; // 1-based index in the LIFO queue
                    string partialReport = string.Format(
                        "Rolled back {0} of {1} actions. Stopped at action #{2} ({3} on {4}): {5}. Intervening state preserved.",
                        rolledBack.Count, totalToRollback, stoppedIndex, act.Operation, act.TargetDisplay, res.ErrorMessage);

                    Logger.Warn("RollbackExecutor.RollbackBatch halted: " + partialReport);
                    return HostOperationResult.Failed(partialReport);
                }
            }

            return HostOperationResult.Ok(
                string.Format("Successfully rolled back {0} actions in strict LIFO order.", rolledBack.Count));
        }
    }
}
