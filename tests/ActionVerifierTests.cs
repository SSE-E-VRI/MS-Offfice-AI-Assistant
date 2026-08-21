using System;
using System.Collections.Generic;
using MSOfficeAIAssistant.Core;
using MSOfficeAIAssistant.Core.Actions;
using MSOfficeAIAssistant.Hosts;

namespace MSOfficeAIAssistant.Tests
{
    public class ActionVerifierTests
    {
        public static void RunAll()
        {
            Console.WriteLine("=== ActionVerifier & Risk Gating Tests ===");

            TestRiskLevelSingleSourcedFromToolDefinition();
            TestPreVerifyMissingRequiredParameters();
            TestPreVerifyUnsafeTargetAddress();
            TestPreVerifyValidActionGeneratesPrompt();
            TestPostVerifyHostBusyRetryable();
            TestPostVerifyInvalidSyntaxHResult();
            TestPostVerifySpreadsheetErrorLiterals();
            TestPostVerifyAppliedClean();
            TestHostGuardRejectsHostMismatch();
            TestControllerTypeGuardRejectsWrongType();

            Console.WriteLine("All ActionVerifier & Risk Gating tests passed!");
        }

        private static void Assert(bool condition, string message)
        {
            if (!condition)
            {
                throw new Exception("Assertion Failed: " + message);
            }
        }

        private static void TestRiskLevelSingleSourcedFromToolDefinition()
        {
            var removeDupTool = ToolRegistry.GetTool("excel.remove_duplicates", "Excel");
            Assert(removeDupTool != null, "Tool excel.remove_duplicates must exist");
            Assert(removeDupTool.RiskLevel == 3, "excel.remove_duplicates risk level must be 3");
            Assert(!removeDupTool.IsUndoable, "excel.remove_duplicates must not be undoable");

            // Verify OfficeAction inherits from ToolDefinition
            var action = new OfficeAction
            {
                Host = "Excel",
                Operation = "excel.remove_duplicates",
                Target = new ActionTarget { Range = "A1:C10" }
            };

            var pre = ActionVerifier.PreVerify(action, "Excel");
            Assert(pre.IsValid, "PreVerify should be valid for remove_duplicates");
            Assert(pre.RiskLevel == 3, "PreVerify must report RiskLevel 3 from ToolDefinition");
            Assert(!pre.IsUndoable, "PreVerify must report IsUndoable == false from ToolDefinition");
            Assert(action.RiskLevel == 3, "OfficeAction.RiskLevel must be synchronized to 3");
            Assert(!action.IsUndoable, "OfficeAction.IsUndoable must be synchronized to false");
            Assert(pre.ConfirmationPrompt.Contains("WARNING"), "Confirmation prompt must contain warning for high risk / non-undoable tool");

            Console.WriteLine("  [PASS] RiskLevel and IsUndoable single-sourced from ToolDefinition");
        }

        private static void TestPreVerifyMissingRequiredParameters()
        {
            var action = new OfficeAction
            {
                Host = "Excel",
                Operation = "excel.write_formula"
                // Missing both target and formula
            };

            var pre = ActionVerifier.PreVerify(action, "Excel");
            Assert(!pre.IsValid, "PreVerify must fail when required parameters are missing");
            Assert(pre.Outcome == VerificationOutcome.ValidationError, "Outcome must be ValidationError");
            Assert(pre.ValidationErrors.Count >= 2, "Must report errors for both missing target and formula");

            Console.WriteLine("  [PASS] PreVerify detects missing required parameters");
        }

        private static void TestPreVerifyUnsafeTargetAddress()
        {
            var action = new OfficeAction
            {
                Host = "Excel",
                Operation = "excel.write_value",
                Target = new ActionTarget { Range = "INVALID_ADDR$$$!!!" },
                Parameters = new Dictionary<string, object>
                {
                    { "target", "INVALID_ADDR$$$!!!" },
                    { "value", "123" }
                }
            };

            var pre = ActionVerifier.PreVerify(action, "Excel");
            Assert(!pre.IsValid, "PreVerify must fail on unsafe/malformed target range");
            Assert(pre.ValidationErrors.Exists(e => e.Contains("Invalid or unsafe")), "Must report unsafe address error");

            Console.WriteLine("  [PASS] PreVerify validates target syntax safety");
        }

        private static void TestPreVerifyValidActionGeneratesPrompt()
        {
            var action = new OfficeAction
            {
                Host = "Excel",
                Operation = "excel.write_formula",
                Target = new ActionTarget { Range = "K20" },
                Parameters = new Dictionary<string, object>
                {
                    { "target", "K20" },
                    { "formula", "=SUM(A1:A10)" },
                    { "description", "Calculate total" }
                }
            };

            var pre = ActionVerifier.PreVerify(action, "Excel");
            Assert(pre.IsValid, "PreVerify must pass for valid action");
            Assert(!string.IsNullOrWhiteSpace(pre.ConfirmationPrompt), "ConfirmationPrompt must be populated");
            Assert(pre.ConfirmationPrompt.Contains("K20"), "Prompt must reference target cell K20");

            Console.WriteLine("  [PASS] PreVerify valid action generates structured prompt");
        }

        private static void TestPostVerifyHostBusyRetryable()
        {
            var action = new OfficeAction
            {
                Host = "Excel",
                Operation = "excel.write_formula",
                Target = new ActionTarget { Range = "B2" }
            };

            var busyResult = HostOperationResult.Failed("Excel COM call returned 0x800AC472 (VBA_E_IGNORE)", ActionVerifier.HRESULT_VBA_E_IGNORE, "B2");
            var pv = ActionVerifier.PostVerify(action, busyResult);

            Assert(!pv.Verified, "Host busy result must not be verified");
            Assert(pv.Outcome == VerificationOutcome.HostBusyRetryable, "Outcome must be HostBusyRetryable");
            Assert(!pv.WasApplied, "WasApplied must be FALSE for 0x800AC472");
            Assert(pv.IsRetryable, "IsRetryable must be TRUE for 0x800AC472");
            Assert(pv.DiagnosticMessage.Contains("Excel is busy"), "DiagnosticMessage must explain in-cell edit / busy state");

            Console.WriteLine("  [PASS] PostVerify classifies 0x800AC472 as HostBusyRetryable (not applied)");
        }

        private static void TestPostVerifyInvalidSyntaxHResult()
        {
            var action = new OfficeAction
            {
                Host = "Excel",
                Operation = "excel.write_formula",
                Target = new ActionTarget { Range = "B2" }
            };

            var syntaxErrResult = HostOperationResult.Failed("Name not found 0x800A03EC", ActionVerifier.HRESULT_NAME_NOT_FOUND, "B2");
            var pv = ActionVerifier.PostVerify(action, syntaxErrResult);

            Assert(!pv.Verified, "Syntax error result must not be verified");
            Assert(pv.Outcome == VerificationOutcome.ExecutionError, "Outcome must be ExecutionError");
            Assert(!pv.WasApplied, "WasApplied must be FALSE for syntax error");
            Assert(!pv.IsRetryable, "IsRetryable must be FALSE for syntax error");

            Console.WriteLine("  [PASS] PostVerify identifies 0x800A03EC syntax error");
        }

        private static void TestPostVerifySpreadsheetErrorLiterals()
        {
            string[] testLiterals = new string[] { "#REF!", "#VALUE!", "#DIV/0!", "#N/A", "#NAME?", "#NULL!", "#NUM!", "#SPILL!", "#CALC!" };

            foreach (var errLiteral in testLiterals)
            {
                var action = new OfficeAction
                {
                    Host = "Excel",
                    Operation = "excel.write_formula",
                    Target = new ActionTarget { Range = "C5" }
                };

                // The COM call succeeded (S_OK), but the formula cell evaluated to an error literal
                var successResultWithErrorValue = HostOperationResult.Ok(errLiteral, "C5");
                var pv = ActionVerifier.PostVerify(action, successResultWithErrorValue);

                Assert(!pv.Verified, string.Format("PostVerify must detect {0} evaluation error", errLiteral));
                Assert(pv.Outcome == VerificationOutcome.EvaluatedToError, string.Format("Outcome must be EvaluatedToError for {0}", errLiteral));
                Assert(pv.WasApplied, "WasApplied must be TRUE because the cell was modified");
                Assert(pv.ObservedValue == errLiteral, "ObservedValue must match returned error literal");
                Assert(pv.DiagnosticMessage.Contains(errLiteral), "Diagnostic message must report the error literal");
            }

            Console.WriteLine("  [PASS] PostVerify detects all standard Excel calculation error literals (#REF!, #VALUE!, #DIV/0!, etc.)");
        }

        private static void TestPostVerifyAppliedClean()
        {
            var action = new OfficeAction
            {
                Host = "Excel",
                Operation = "excel.write_formula",
                Target = new ActionTarget { Range = "D10" }
            };

            var cleanResult = HostOperationResult.Ok("1500.00", "D10");
            var pv = ActionVerifier.PostVerify(action, cleanResult);

            Assert(pv.Verified, "Clean calculation result must be verified");
            Assert(pv.Outcome == VerificationOutcome.Success, "Outcome must be Success");
            Assert(pv.WasApplied, "WasApplied must be TRUE");
            Assert(pv.ObservedValue == "1500.00", "ObservedValue must be 1500.00");

            Console.WriteLine("  [PASS] PostVerify verifies clean execution result");
        }

        private static void TestHostGuardRejectsHostMismatch()
        {
            var action = new OfficeAction
            {
                Host = "Word",
                Operation = "excel.write_formula",
                Target = new ActionTarget { Range = "A1" }
            };

            var res = ToolRegistry.Execute("DummyController", action);
            Assert(!res.Success, "Execute must fail when action specifies Word but tool is Excel");
            Assert(res.ErrorMessage.Contains("Host mismatch"), "Error message must report Host mismatch");

            Console.WriteLine("  [PASS] ToolRegistry.Execute enforces Host Guard against host mismatch");
        }

        private static void TestControllerTypeGuardRejectsWrongType()
        {
            var action = new OfficeAction
            {
                Host = "Excel",
                Operation = "excel.write_formula",
                Target = new ActionTarget { Range = "A1" },
                Parameters = new Dictionary<string, object>
                {
                    { "target", "A1" },
                    { "formula", "=1+1" }
                }
            };

            // Pass a string or wrong object instead of ExcelController
            var res = ToolRegistry.Execute("InvalidStringControllerObject", action);
            Assert(!res.Success, "Execute must fail when controller is of the wrong type");
            Assert(res.ErrorMessage.Contains("Controller type mismatch"), "Error message must report Controller type mismatch rather than throwing");

            Console.WriteLine("  [PASS] ToolRegistry.Execute enforces Controller Type Guard without throwing InvalidCastException");
        }
    }
}
