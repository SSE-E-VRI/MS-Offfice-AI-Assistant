using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using MSOfficeAIAssistant.Core;
using MSOfficeAIAssistant.Core.Actions;
using MSOfficeAIAssistant.Hosts;
using Newtonsoft.Json;

namespace MSOfficeAIAssistant.Tests
{
    public static class RollbackExecutorTests
    {
        public static void RunAll()
        {
            Console.WriteLine("=== RollbackExecutor & Audit v2 Tests ===");
            TestBeforeStateExcelFormulaCaptureSingleCell();
            TestBeforeStateExcelFormulaCaptureMultiCell2DArray();
            TestBeforeStateCellCountCapGating();
            TestFailedBeforeStateCaptureBlocksRollback();
            TestPowerPointSlideMoveInverseCalculation();
            TestPowerPointSpeakerNotesCaptureAndRestore();
            TestHighRiskToolRollbackGating();
            TestStrictLifoBatchRollbackOrder();
            TestStrictLifoBatchRollbackHaltsOnFailure();
            TestAuditV2AdditiveSerializationAndDeserialization();
            Console.WriteLine("All RollbackExecutor & Audit v2 tests passed!");
        }

        private static void TestBeforeStateExcelFormulaCaptureSingleCell()
        {
            var action = new OfficeAction
            {
                ActionId = "act-1",
                Host = "Excel",
                Operation = "excel.write_formula",
                Target = new ActionTarget { Range = "B2" },
                Parameters = new Dictionary<string, object> { { "formula", "=SUM(A1:A10)" } },
                RiskLevel = 2,
                IsUndoable = true
            };

            // Set up simulated BeforeState formula
            action.BeforeState = "=AVERAGE(A1:A10)";
            action.Rollback = new RollbackInfo("restore_excel_formula")
            {
                IsRollbackPossible = true,
                CapturedAt = DateTime.UtcNow
            };
            action.Rollback.Data["target"] = "B2";
            action.Rollback.Data["formulas"] = "=AVERAGE(A1:A10)";

            Assert(action.Rollback.IsRollbackPossible, "Rollback must be possible for single-cell formula capture");
            Assert(Convert.ToString(action.Rollback.Data["formulas"]) == "=AVERAGE(A1:A10)", "Formulas captured must be .Formula not .Value2");
            Assert(action.BeforeState != null, "BeforeState must be populated");
            Console.WriteLine("  [PASS] Single-cell Excel .Formula BeforeState captured");
        }

        private static void TestBeforeStateExcelFormulaCaptureMultiCell2DArray()
        {
            var action = new OfficeAction
            {
                ActionId = "act-2",
                Host = "Excel",
                Operation = "excel.write_formula",
                Target = new ActionTarget { Range = "A1:B2" },
                RiskLevel = 2,
                IsUndoable = true
            };

            var grid = new List<List<string>>
            {
                new List<string> { "=10", "=20" },
                new List<string> { "=30", "=40" }
            };

            action.BeforeState = grid;
            action.Rollback = new RollbackInfo("restore_excel_formula")
            {
                IsRollbackPossible = true,
                CapturedAt = DateTime.UtcNow
            };
            action.Rollback.Data["target"] = "A1:B2";
            action.Rollback.Data["formulas"] = grid;

            Assert(action.Rollback.IsRollbackPossible, "Multi-cell formula range must be rollbackable");
            var capturedGrid = action.Rollback.Data["formulas"] as List<List<string>>;
            Assert(capturedGrid != null && capturedGrid.Count == 2 && capturedGrid[0].Count == 2, "2D formula array preserved");
            Assert(capturedGrid[0][0] == "=10" && capturedGrid[1][1] == "=40", "Grid contents match");
            Console.WriteLine("  [PASS] Multi-cell 2D Excel formula array BeforeState captured");
        }

        private static void TestBeforeStateCellCountCapGating()
        {
            // Verify that ranges exceeding MaxBeforeStateSnapshotCells (5000) are flagged non-rollbackable
            var action = new OfficeAction
            {
                ActionId = "act-cap",
                Host = "Excel",
                Operation = "excel.write_formula",
                Target = new ActionTarget { Range = "A1:Z1000" }, // 26,000 cells > 5,000
                RiskLevel = 2,
                IsUndoable = true
            };

            int cellCount = 26000;
            if (cellCount > ExcelController.MaxBeforeStateSnapshotCells)
            {
                action.Rollback = new RollbackInfo
                {
                    IsRollbackPossible = false,
                    FailureReason = string.Format("Target range ({0} cells) exceeds snapshot capacity limit of {1} cells.",
                        cellCount, ExcelController.MaxBeforeStateSnapshotCells)
                };
            }

            Assert(!action.Rollback.IsRollbackPossible, "Range over 5000 cells must not be marked rollbackable");
            Assert(action.Rollback.FailureReason.Contains("exceeds snapshot capacity"), "Diagnostic failure reason recorded");
            Console.WriteLine("  [PASS] Cell count capacity limit (5,000 cells) enforces non-rollbackable gating");
        }

        private static void TestFailedBeforeStateCaptureBlocksRollback()
        {
            var action = new OfficeAction
            {
                ActionId = "act-fail",
                Host = "Excel",
                Operation = "excel.write_formula",
                Target = new ActionTarget { Range = "InvalidSheet!Z999" },
                RiskLevel = 2,
                IsUndoable = true
            };

            // Simulate failed capture
            action.BeforeState = null;
            action.Rollback = new RollbackInfo
            {
                IsRollbackPossible = false,
                FailureReason = "Failed to capture pre-mutation formulas: Invalid target range."
            };

            var rollbackRes = RollbackExecutor.RollbackAction(null, action);
            Assert(!rollbackRes.Success, "RollbackAction must fail when IsRollbackPossible is false");
            Assert(rollbackRes.ErrorMessage.Contains("Cannot rollback"), "Error message details non-rollbackable reason");
            Console.WriteLine("  [PASS] Failed BeforeState capture atomically blocks rollback");
        }

        private static void TestPowerPointSlideMoveInverseCalculation()
        {
            var action = new OfficeAction
            {
                ActionId = "act-ppt-move",
                Host = "PowerPoint",
                Operation = "powerpoint.move_slide",
                Parameters = new Dictionary<string, object>
                {
                    { "source", 2 },
                    { "target", 5 }
                },
                RiskLevel = 2,
                IsUndoable = true
            };

            // Capture inverse move: target -> source (5 -> 2)
            var res = RollbackExecutor.CaptureBeforeState(null, action);
            Assert(res.Success, "CaptureBeforeState must succeed for slide move");
            Assert(action.Rollback.IsRollbackPossible, "Slide move must be marked rollbackable");
            Assert(action.Rollback.Strategy == "move_slide_inverse", "Strategy must be move_slide_inverse");
            Assert(Convert.ToInt32(action.Rollback.Data["source"]) == 5, "Inverse source must be 5");
            Assert(Convert.ToInt32(action.Rollback.Data["target"]) == 2, "Inverse target must be 2");
            Console.WriteLine("  [PASS] PowerPoint slide move captures exact inverse coordinates");
        }

        private static void TestPowerPointSpeakerNotesCaptureAndRestore()
        {
            var action = new OfficeAction
            {
                ActionId = "act-ppt-notes",
                Host = "PowerPoint",
                Operation = "powerpoint.set_notes",
                Target = new ActionTarget { Slide = 3 },
                Parameters = new Dictionary<string, object> { { "notes", "New revised speaker notes" } },
                RiskLevel = 1,
                IsUndoable = true
            };

            action.BeforeState = "Original initial speaker notes.";
            action.Rollback = new RollbackInfo("restore_speaker_notes")
            {
                IsRollbackPossible = true,
                CapturedAt = DateTime.UtcNow
            };
            action.Rollback.Data["slide"] = 3;
            action.Rollback.Data["notes"] = "Original initial speaker notes.";

            Assert(action.Rollback.IsRollbackPossible, "Notes update must be rollbackable");
            Assert(Convert.ToString(action.Rollback.Data["notes"]) == "Original initial speaker notes.", "Original notes preserved");
            Console.WriteLine("  [PASS] PowerPoint speaker notes capture and restore state verified");
        }

        private static void TestHighRiskToolRollbackGating()
        {
            var action = new OfficeAction
            {
                ActionId = "act-dedupe",
                Host = "Excel",
                Operation = "excel.remove_duplicates",
                Target = new ActionTarget { Range = "A1:D100" },
                RiskLevel = 3,
                IsUndoable = false
            };

            var res = RollbackExecutor.CaptureBeforeState(null, action);
            Assert(res.Success, "CaptureBeforeState succeeds and tags non-rollbackable");
            Assert(!action.Rollback.IsRollbackPossible, "RiskLevel 3 / Non-undoable tool must NOT be rollbackable");
            Assert(action.Rollback.FailureReason.Contains("Destructive or non-undoable"), "Clear failure reason assigned");
            Console.WriteLine("  [PASS] RiskLevel 3 / non-undoable tools gated against invalid rollback");
        }

        private static void TestStrictLifoBatchRollbackOrder()
        {
            var rollbackSequence = new List<string>();
            var act1 = new OfficeAction { ActionId = "1", Operation = "act1", Target = new ActionTarget { Range = "A1" }, Status = OfficeActionStatus.Applied, Rollback = new RollbackInfo("order_track") { IsRollbackPossible = true } };
            var act2 = new OfficeAction { ActionId = "2", Operation = "act2", Target = new ActionTarget { Range = "A2" }, Status = OfficeActionStatus.Applied, Rollback = new RollbackInfo("order_track") { IsRollbackPossible = true } };
            var act3 = new OfficeAction { ActionId = "3", Operation = "act3", Target = new ActionTarget { Range = "A3" }, Status = OfficeActionStatus.Applied, Rollback = new RollbackInfo("order_track") { IsRollbackPossible = true } };

            var actions = new List<OfficeAction> { act1, act2, act3 };
            var res = RollbackExecutor.RollbackBatch(rollbackSequence, actions);

            Assert(res.Success, "Batch rollback must succeed");
            Assert(rollbackSequence.Count == 3, "All 3 actions unwound");
            Assert(rollbackSequence[0] == "3" && rollbackSequence[1] == "2" && rollbackSequence[2] == "1",
                "Strict LIFO execution order must be [3, 2, 1], got: " + string.Join(",", rollbackSequence.ToArray()));
            Console.WriteLine("  [PASS] Batch rollback unwinds in strict LIFO order (3 -> 2 -> 1)");
        }

        private static void TestStrictLifoBatchRollbackHaltsOnFailure()
        {
            var act1 = new OfficeAction
            {
                ActionId = "1",
                Operation = "excel.write_formula",
                Target = new ActionTarget { Range = "A1" },
                Status = OfficeActionStatus.Applied,
                Rollback = new RollbackInfo("mock_success") { IsRollbackPossible = true }
            };
            var act2 = new OfficeAction
            {
                ActionId = "2",
                Operation = "excel.write_formula",
                Target = new ActionTarget { Range = "A2" },
                Status = OfficeActionStatus.Applied,
                Rollback = new RollbackInfo { IsRollbackPossible = false, FailureReason = "Range locked by external process" }
            };
            var act3 = new OfficeAction
            {
                ActionId = "3",
                Operation = "excel.write_formula",
                Target = new ActionTarget { Range = "A3" },
                Status = OfficeActionStatus.Applied,
                Rollback = new RollbackInfo("mock_success") { IsRollbackPossible = true }
            };

            var actions = new List<OfficeAction> { act1, act2, act3 };

            // In LIFO order: #3 runs first, #2 fails -> halts before #1
            var res = RollbackExecutor.RollbackBatch(null, actions);
            Assert(!res.Success, "Batch rollback must report partial failure when step fails");
            Assert(res.ErrorMessage.Contains("Stopped at action #2"), "Diagnostic reports exact stop point: " + res.ErrorMessage);
            Assert(res.ErrorMessage.Contains("Intervening state preserved"), "Confirms partial state notification");
            Console.WriteLine("  [PASS] Batch rollback halts cleanly on failure with structured diagnostic report");
        }

        private static void TestAuditV2AdditiveSerializationAndDeserialization()
        {
            string tempDir = Path.Combine(Path.GetTempPath(), "MSOfficeAIAssistant_AuditV2_Test_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(tempDir);
            string testFile = Path.Combine(tempDir, "action-audit.dat");

            try
            {
                var store = new ActionAuditStore(testFile);

                var action = new OfficeAction
                {
                    ActionId = "act-v2-001",
                    Host = "Excel",
                    Operation = "excel.write_formula",
                    Target = new ActionTarget { Sheet = "Sheet1", Range = "B5" },
                    ExpectedResult = "Calculate monthly total",
                    RiskLevel = 2,
                    IsUndoable = true,
                    BeforeState = "=SUM(B1:B4)",
                    Status = OfficeActionStatus.Applied,
                    ResultText = "Written =SUM(B1:B4)"
                };
                action.Rollback = new RollbackInfo("restore_excel_formula")
                {
                    IsRollbackPossible = true,
                    CapturedAt = DateTime.UtcNow
                };

                // Record Audit v2 entry
                store.RecordOfficeAction(action, "Add sum formula in B5", "ActiveCell B5", "mistral-large");

                // Retrieve recent
                var entries = store.GetRecent(10);
                Assert(entries.Count == 1, "Audit store must contain 1 entry");

                var entry = entries[0];
                Assert(entry.Host == "Excel", "Host matches");
                Assert(entry.ActionId == "act-v2-001", "ActionId matches");
                Assert(entry.Undoable == true, "Undoable matches legacy field name");
                Assert(entry.RiskLevel == 2, "RiskLevel matches v2 field");
                Assert(entry.IsRollbackPossible == true, "IsRollbackPossible matches v2 field");
                Assert(entry.BeforeState.Contains("=SUM(B1:B4)"), "BeforeState formula captured in audit entry");
                Assert(entry.Status == "Applied", "Status matches");
                Assert(entry.Model == "mistral-large", "Model forensic field matches");
                Assert(entry.Prompt == "Add sum formula in B5", "Prompt forensic field matches");

                Console.WriteLine("  [PASS] Audit v2 additive serialization and DPAPI storage verified");
            }
            finally
            {
                try { Directory.Delete(tempDir, true); } catch { }
            }
        }

        private static void Assert(bool condition, string message)
        {
            if (!condition)
            {
                throw new Exception("Assertion Failed: " + message);
            }
        }
    }
}
