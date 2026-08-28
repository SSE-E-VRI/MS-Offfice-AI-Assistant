using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using MSOfficeAIAssistant.Hosts;

namespace MSOfficeAIAssistant.Core.Actions
{
    public enum VerificationOutcome
    {
        Success,
        ValidationError,
        HostBusyRetryable,
        EvaluatedToError,
        ExecutionError
    }

    public class PreVerificationResult
    {
        public bool IsValid { get; set; }
        public VerificationOutcome Outcome { get; set; }
        public ToolDefinition Tool { get; set; }
        public int RiskLevel { get; set; }
        public bool RequiresApproval { get; set; }
        public bool IsUndoable { get; set; }
        public List<string> ValidationErrors { get; set; }
        public string ConfirmationPrompt { get; set; }

        public PreVerificationResult()
        {
            ValidationErrors = new List<string>();
            IsValid = true;
            Outcome = VerificationOutcome.Success;
            RiskLevel = 1;
            RequiresApproval = true;
            IsUndoable = true;
        }
    }

    public class PostVerificationResult
    {
        public bool Verified { get; set; }
        public VerificationOutcome Outcome { get; set; }
        public int ErrorCode { get; set; }
        public string DiagnosticMessage { get; set; }
        public string ObservedValue { get; set; }
        public string ExpectedValue { get; set; }
        public bool IsRetryable { get; set; }
        public bool WasApplied { get; set; }

        public PostVerificationResult()
        {
            Verified = false;
            Outcome = VerificationOutcome.ExecutionError;
            WasApplied = false;
            IsRetryable = false;
        }
    }

    /// <summary>
    /// Verification and Risk-Gating Engine per SSOT §5.3.
    /// Provides pre-execution parameter and syntax validation, risk gating from ToolDefinition,
    /// and post-execution verification inspecting COM HRESULTs, error codes, and spreadsheet error literals (#REF!, #VALUE!, etc.).
    /// </summary>
    public static class ActionVerifier
    {
        public const int HRESULT_VBA_E_IGNORE = unchecked((int)0x800AC472); // -2146777998: Excel in-cell edit or modal busy
        public const int HRESULT_NAME_NOT_FOUND = unchecked((int)0x800A03EC); // -2146827284: Invalid formula/range syntax
        public const int HRESULT_RPC_E_RETRY = unchecked((int)0x8001010A); // -2147417846: RPC busy

        private static readonly string[] ExcelErrorLiterals = new string[]
        {
            "#REF!",
            "#VALUE!",
            "#DIV/0!",
            "#N/A",
            "#NAME?",
            "#NULL!",
            "#NUM!",
            "#SPILL!",
            "#CALC!"
        };

        /// <summary>
        /// Pre-verifies an action against the ToolRegistry, checks parameter requirements and target safety,
        /// and extracts authoritative RiskLevel and IsUndoable from ToolDefinition (SSOT single source).
        /// </summary>
        public static PreVerificationResult PreVerify(OfficeAction action, string host = null)
        {
            var res = new PreVerificationResult();
            if (action == null)
            {
                res.IsValid = false;
                res.Outcome = VerificationOutcome.ValidationError;
                res.ValidationErrors.Add("Action cannot be null.");
                return res;
            }

            string effectiveHost = !string.IsNullOrEmpty(host) ? host : (!string.IsNullOrEmpty(action.Host) ? action.Host : null);
            ToolDefinition tool = ToolRegistry.GetTool(action.Operation, effectiveHost);

            if (tool == null)
            {
                res.IsValid = false;
                res.Outcome = VerificationOutcome.ValidationError;
                res.ValidationErrors.Add(string.Format("Unrecognized operation '{0}' for host '{1}'.", action.Operation, effectiveHost ?? "unspecified"));
                return res;
            }

            res.Tool = tool;
            res.RiskLevel = tool.RiskLevel;
            res.RequiresApproval = tool.RequiresApproval;
            res.IsUndoable = tool.IsUndoable;

            // Synchronize action's risk properties from single source of truth
            action.RiskLevel = tool.RiskLevel;
            action.RequiresApproval = tool.RequiresApproval;
            action.IsUndoable = tool.IsUndoable;
            if (string.IsNullOrEmpty(action.Host))
            {
                action.Host = tool.Host;
            }

            // Check required parameters
            if (tool.Parameters != null)
            {
                foreach (var param in tool.Parameters)
                {
                    if (param.IsRequired)
                    {
                        bool present = false;
                        if (action.Parameters != null && action.Parameters.ContainsKey(param.Name))
                        {
                            object val = action.Parameters[param.Name];
                            if (val != null && !string.IsNullOrWhiteSpace(Convert.ToString(val)))
                            {
                                present = true;
                            }
                        }

                        // Also check standard properties mapped to parameters
                        if (!present)
                        {
                            if (string.Equals(param.Name, "target", StringComparison.OrdinalIgnoreCase) && !string.IsNullOrEmpty(action.TargetDisplay)) present = true;
                            else if (string.Equals(param.Name, "formula", StringComparison.OrdinalIgnoreCase) && !string.IsNullOrEmpty(action.ContentDisplay)) present = true;
                            else if (string.Equals(param.Name, "value", StringComparison.OrdinalIgnoreCase) && !string.IsNullOrEmpty(action.ContentDisplay)) present = true;
                            else if (string.Equals(param.Name, "comment", StringComparison.OrdinalIgnoreCase) && !string.IsNullOrEmpty(action.ContentDisplay)) present = true;
                            else if (string.Equals(param.Name, "slide", StringComparison.OrdinalIgnoreCase) && action.Target != null && action.Target.Slide.HasValue) present = true;
                            // Word tool handlers accept these parameters through alternate channels
                            // (see ToolRegistry's word.find_replace/apply_style/reorganize_paragraphs
                            // handlers) -- the required-parameter check must recognize the same
                            // fallbacks, or a validly-populated action gets rejected here before it
                            // ever reaches the handler that would have accepted it.
                            else if (string.Equals(param.Name, "find", StringComparison.OrdinalIgnoreCase) &&
                                !string.IsNullOrWhiteSpace(action.GetParameterString("target_text"))) present = true;
                            else if (string.Equals(param.Name, "style", StringComparison.OrdinalIgnoreCase) && !string.IsNullOrEmpty(action.ContentDisplay)) present = true;
                            else if (string.Equals(param.Name, "order", StringComparison.OrdinalIgnoreCase) && !string.IsNullOrEmpty(action.ContentDisplay)) present = true;
                        }

                        if (!present)
                        {
                            res.IsValid = false;
                            res.Outcome = VerificationOutcome.ValidationError;
                            res.ValidationErrors.Add(string.Format("Missing required parameter '{0}' for tool '{1}'.", param.Name, tool.Name));
                        }
                    }
                }
            }

            // Excel specific target syntax safety check
            if (string.Equals(tool.Host, "Excel", StringComparison.OrdinalIgnoreCase))
            {
                string target = action.TargetDisplay;
                if (!string.IsNullOrEmpty(target) && !SpreadsheetActionParser.IsSafeTarget(target))
                {
                    res.IsValid = false;
                    res.Outcome = VerificationOutcome.ValidationError;
                    res.ValidationErrors.Add(string.Format("Invalid or unsafe cell/range target address: '{0}'.", target));
                }
            }

            // Generate structured confirmation prompt
            var sb = new StringBuilder();
            sb.AppendFormat("{0} will execute {1} on {2}.\n", tool.Host, tool.Name, action.TargetDisplay);
            if (!string.IsNullOrWhiteSpace(action.PreviewDescription))
            {
                sb.AppendFormat("Description: {0}\n", action.PreviewDescription);
            }
            if (!string.IsNullOrWhiteSpace(action.ContentDisplay))
            {
                sb.AppendFormat("Proposed change: {0}\n", action.ContentDisplay);
            }
            if (res.RiskLevel >= 3 || !res.IsUndoable)
            {
                sb.AppendLine("\n⚠ WARNING: This action cannot be reliably undone by Office Undo.");
            }

            res.ConfirmationPrompt = sb.ToString().TrimEnd();
            return res;
        }

        /// <summary>
        /// Post-verifies the execution result by inspecting COM HRESULTs, error codes, and checking
        /// spreadsheet return values for error literals (#REF!, #VALUE!, etc.).
        /// </summary>
        public static PostVerificationResult PostVerify(OfficeAction action, HostOperationResult result)
        {
            var pv = new PostVerificationResult();
            if (result == null)
            {
                pv.Verified = false;
                pv.Outcome = VerificationOutcome.ExecutionError;
                pv.WasApplied = false;
                pv.DiagnosticMessage = "No execution result was returned by host controller.";
                return pv;
            }

            pv.ErrorCode = result.ErrorCode;

            // 1. Check for 0x800AC472 (Excel busy / in-cell edit mode) -> NOT APPLIED, retryable
            if (result.ErrorCode == HRESULT_VBA_E_IGNORE ||
                (!string.IsNullOrEmpty(result.ErrorMessage) && (result.ErrorMessage.Contains("0x800AC472") || result.ErrorMessage.IndexOf("0x800AC472", StringComparison.OrdinalIgnoreCase) >= 0)))
            {
                pv.Verified = false;
                pv.Outcome = VerificationOutcome.HostBusyRetryable;
                pv.WasApplied = false;
                pv.IsRetryable = true;
                pv.DiagnosticMessage = "Excel is busy (in-cell edit mode or open dialog). Action was not applied. Finish editing in Excel and try again.";
                return pv;
            }

            // 2. Check for 0x800A03EC (Invalid formula or range syntax)
            if (result.ErrorCode == HRESULT_NAME_NOT_FOUND ||
                (!string.IsNullOrEmpty(result.ErrorMessage) && result.ErrorMessage.Contains("0x800A03EC")))
            {
                pv.Verified = false;
                pv.Outcome = VerificationOutcome.ExecutionError;
                pv.WasApplied = false;
                pv.IsRetryable = false;
                pv.DiagnosticMessage = "Excel error 0x800A03EC: Invalid formula syntax, unrecognized function name, or invalid range target.";
                return pv;
            }

            // 3. General failure
            if (!result.Success)
            {
                pv.Verified = false;
                pv.Outcome = VerificationOutcome.ExecutionError;
                pv.WasApplied = false;
                pv.DiagnosticMessage = !string.IsNullOrEmpty(result.ErrorMessage) ? result.ErrorMessage : "Host operation failed.";
                return pv;
            }

            // 4. Success path: inspect observed value for spreadsheet calculation error literals (#REF!, #VALUE!, etc.)
            pv.WasApplied = true;
            string observed = result.Value != null ? Convert.ToString(result.Value) : string.Empty;
            pv.ObservedValue = observed;
            if (action != null)
            {
                pv.ExpectedValue = action.ExpectedResult;
            }

            string detectedError = DetectSpreadsheetErrorLiteral(observed);
            if (detectedError != null)
            {
                pv.Verified = false;
                pv.Outcome = VerificationOutcome.EvaluatedToError;
                pv.DiagnosticMessage = string.Format("Formula applied to {0} but evaluated to calculation error '{1}' in Excel.", action != null ? action.TargetDisplay : "cell", detectedError);
                return pv;
            }

            pv.Verified = true;
            pv.Outcome = VerificationOutcome.Success;
            pv.DiagnosticMessage = !string.IsNullOrEmpty(observed) ? observed : "Applied successfully";
            return pv;
        }

        public static string DetectSpreadsheetErrorLiteral(string text)
        {
            if (string.IsNullOrEmpty(text)) return null;
            foreach (var err in ExcelErrorLiterals)
            {
                if (text.IndexOf(err, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return err;
                }
            }
            return null;
        }
    }
}
