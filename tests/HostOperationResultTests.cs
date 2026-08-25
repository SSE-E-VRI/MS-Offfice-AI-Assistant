using System;
using System.Collections.Generic;
using MSOfficeAIAssistant.Core;
using MSOfficeAIAssistant.Hosts;

namespace MSOfficeAIAssistant.Tests
{
    public static class HostOperationResultTests
    {
        public static void RunAll()
        {
            TestResultOkCreation();
            TestResultFailedCreation();
            TestResultFromException();
            TestWordControllerHeadlessOperationResults();
            TestExcelControllerHeadlessOperationResults();
            TestPowerPointControllerHeadlessOperationResults();
        }

        private static void TestResultOkCreation()
        {
            var res = HostOperationResult.Ok("Formula written", "B2:B10");
            Assert(res.Success, "Expected Success to be true");
            Assert(res.ErrorCode == 0, "Expected ErrorCode to be 0");
            Assert(res.ErrorMessage == null, "Expected ErrorMessage to be null");
            Assert((string)res.Value == "Formula written", "Expected Value to match");
            Assert(res.TargetLocation == "B2:B10", "Expected TargetLocation to match");
        }

        private static void TestResultFailedCreation()
        {
            var res = HostOperationResult.Failed("Protected worksheet", unchecked((int)0x800AC472), "Sheet1!A1");
            Assert(!res.Success, "Expected Success to be false");
            Assert(res.ErrorCode == unchecked((int)0x800AC472), "Expected ErrorCode to be preserved");
            Assert(res.ErrorMessage == "Protected worksheet", "Expected ErrorMessage to match");
            Assert(res.TargetLocation == "Sheet1!A1", "Expected TargetLocation to match");
            Assert(res.Value == null, "Expected Value to be null");
        }

        private static void TestResultFromException()
        {
            var ex = new InvalidOperationException("Cell is locked and formula cannot be edited");
            var res = HostOperationResult.FromException(ex, "ExcelController.WriteFormula", "C4");
            Assert(!res.Success, "Expected Success to be false");
            Assert(res.ErrorMessage.Contains("ExcelController.WriteFormula"), "Expected operation name prefix in error message");
            Assert(res.ErrorMessage.Contains("Cell is locked"), "Expected exception message in error message");
            Assert(res.TargetLocation == "C4", "Expected TargetLocation to match");
        }

        private static void TestWordControllerHeadlessOperationResults()
        {
            var ctrl = new WordController(null);

            var insertRes = ctrl.ExecuteInsertText(null);
            Assert(!insertRes.Success, "ExecuteInsertText(null) should fail");
            Assert(insertRes.ErrorMessage.Contains("null or empty"), "Expected validation message on null text");

            var insertEmptyAppRes = ctrl.ExecuteInsertText("Hello World");
            Assert(!insertEmptyAppRes.Success, "ExecuteInsertText on null app should fail cleanly");
            Assert(insertEmptyAppRes.ErrorMessage.Contains("accessible") || insertEmptyAppRes.ErrorMessage.Contains("Word"),
                "Expected error message indicating Word application is inaccessible");

            var commentRes = ctrl.ExecuteAddComment(null, "Heading 1");
            Assert(!commentRes.Success, "ExecuteAddComment(null) should fail");
            Assert(commentRes.ErrorMessage.Contains("empty"), "Expected validation message on empty comment");

            var commentHeadlessRes = ctrl.ExecuteAddComment("Review this section", "Section 1");
            Assert(!commentHeadlessRes.Success, "ExecuteAddComment on null app should fail cleanly");

            var tableRes = ctrl.ExecuteInsertTable(0, 0);
            Assert(!tableRes.Success, "ExecuteInsertTable(0,0) should fail");

            var tableHeadlessRes = ctrl.ExecuteInsertTable(2, 3, new List<List<string>> { new List<string> { "A", "B", "C" } });
            Assert(!tableHeadlessRes.Success, "ExecuteInsertTable on null app should fail cleanly");

            var acceptRes = ctrl.ExecuteAcceptAllRevisions();
            Assert(!acceptRes.Success, "ExecuteAcceptAllRevisions on null app should fail cleanly");

            // D-13 Tier 2: new Execute* wrappers, headless (no Word app), should fail cleanly not throw.
            var acceptSelRes = ctrl.ExecuteAcceptRevisionsInSelection();
            Assert(!acceptSelRes.Success, "ExecuteAcceptRevisionsInSelection on null app should fail cleanly");

            var rejectSelRes = ctrl.ExecuteRejectRevisionsInSelection();
            Assert(!rejectSelRes.Success, "ExecuteRejectRevisionsInSelection on null app should fail cleanly");

            var undoRes = ctrl.ExecuteUndoLastChange();
            Assert(!undoRes.Success, "ExecuteUndoLastChange on null app should fail cleanly");
        }

        private static void TestExcelControllerHeadlessOperationResults()
        {
            var ctrl = new ExcelController(null);

            var actionNullRes = ctrl.ExecuteSpreadsheetAction(null);
            Assert(!actionNullRes.Success, "ExecuteSpreadsheetAction(null) should fail");

            var unsafeFormulaRes = ctrl.ExecuteWriteFormula("=SUM(A1:A10)", "INVALID_CELL_$$$");
            Assert(!unsafeFormulaRes.Success, "ExecuteWriteFormula on unsafe address should fail");
            Assert(unsafeFormulaRes.ErrorMessage.Contains("unsafe") || unsafeFormulaRes.ErrorMessage.Contains("Invalid"),
                "Expected unsafe cell error message");

            var formulaHeadlessRes = ctrl.ExecuteWriteFormula("=SUM(A1:A10)", "B2");
            Assert(!formulaHeadlessRes.Success, "ExecuteWriteFormula on null app should fail cleanly");

            var valueHeadlessRes = ctrl.ExecuteWriteValue("Test", "A1");
            Assert(!valueHeadlessRes.Success, "ExecuteWriteValue on null app should fail cleanly");

            // D-13 Tier 2: new Execute* wrappers, headless (no Excel app), should fail cleanly not throw.
            var undoRes = ctrl.ExecuteUndoLastAction();
            Assert(!undoRes.Success, "ExecuteUndoLastAction on null app should fail cleanly");

            var insertNullRes = ctrl.ExecuteInsertText(null);
            Assert(!insertNullRes.Success, "ExecuteInsertText(null) should fail");
            Assert(insertNullRes.ErrorMessage.Contains("empty"), "Expected validation message on null text");

            var insertHeadlessRes = ctrl.ExecuteInsertText("Hello World");
            Assert(!insertHeadlessRes.Success, "ExecuteInsertText on null app should fail cleanly");
        }

        private static void TestPowerPointControllerHeadlessOperationResults()
        {
            var ctrl = new PowerPointController(null);

            var outlineEmptyRes = ctrl.ExecuteCreateDeckFromOutline("");
            Assert(!outlineEmptyRes.Success, "ExecuteCreateDeckFromOutline(empty) should fail");

            var imageNotFoundRes = ctrl.ExecuteInsertImage("C:\\non_existent_file_path_12345.png");
            Assert(!imageNotFoundRes.Success, "ExecuteInsertImage on non-existent file should fail");
            Assert(imageNotFoundRes.ErrorMessage.Contains("not found"), "Expected file not found message");

            // D-13 Tier 1: Test new Execute* wrappers with validation failures and headless operation
            TestExecuteMoveSlideResults(ctrl);
            TestExecuteCreateSectionResults(ctrl);
            TestExecuteRenameSectionResults(ctrl);
            TestExecuteSetSpeakerNotesResults(ctrl);

            // D-13 Tier 2: new Execute* wrappers, headless (no PowerPoint app), should fail cleanly not throw.
            var undoRes = ctrl.ExecuteUndo();
            Assert(!undoRes.Success, "ExecuteUndo on null app should fail cleanly");

            var insertEmptyRes = ctrl.ExecuteInsertText("");
            Assert(!insertEmptyRes.Success, "ExecuteInsertText(empty) should fail validation");

            var insertHeadlessRes = ctrl.ExecuteInsertText("Hello World");
            Assert(!insertHeadlessRes.Success, "ExecuteInsertText on null app should fail cleanly");
        }

        private static void TestExecuteMoveSlideResults(PowerPointController ctrl)
        {
            // Validation: invalid source slide
            var invalidSrcRes = ctrl.ExecuteMoveSlide(0, 5);
            Assert(!invalidSrcRes.Success, "ExecuteMoveSlide(0, 5) should fail validation");
            Assert(invalidSrcRes.ErrorMessage.Contains("at least 1"), "Expected validation message for source slide");

            // Validation: invalid destination slide
            var invalidTgtRes = ctrl.ExecuteMoveSlide(5, -1);
            Assert(!invalidTgtRes.Success, "ExecuteMoveSlide(5, -1) should fail validation");
            Assert(invalidTgtRes.ErrorMessage.Contains("at least 1"), "Expected validation message for destination slide");

            // Headless: valid parameters but no presentation
            var headlessRes = ctrl.ExecuteMoveSlide(1, 2);
            Assert(!headlessRes.Success, "ExecuteMoveSlide(1, 2) on null app should fail cleanly");
            Assert(headlessRes.ErrorMessage.Contains("Failed"), "Expected failure message for headless operation");

            // Exception path: Mock controller that throws to force outer catch
            var throwingCtrl = new ThrowingMoveSlideController();
            var exceptionRes = throwingCtrl.ExecuteMoveSlide(1, 2);
            Assert(!exceptionRes.Success, "ExecuteMoveSlide should fail when MoveSlide throws");
            Assert(exceptionRes.ErrorCode != 0, "Expected non-zero ErrorCode (HRESULT) when exception caught");
            Assert(exceptionRes.ErrorMessage.Contains("ExecuteMoveSlide"), "Expected operation name in error message");
        }

        private static void TestExecuteCreateSectionResults(PowerPointController ctrl)
        {
            // Validation: empty section name
            var emptyNameRes = ctrl.ExecuteCreateSectionBeforeSlide("", 3);
            Assert(!emptyNameRes.Success, "ExecuteCreateSectionBeforeSlide(empty, 3) should fail");
            Assert(emptyNameRes.ErrorMessage.Contains("empty"), "Expected validation message for section name");

            // Validation: invalid slide number
            var invalidSlideRes = ctrl.ExecuteCreateSectionBeforeSlide("New Section", 0);
            Assert(!invalidSlideRes.Success, "ExecuteCreateSectionBeforeSlide('New Section', 0) should fail");
            Assert(invalidSlideRes.ErrorMessage.Contains("at least 1"), "Expected validation message for slide number");

            // Headless: valid parameters but no presentation
            var headlessRes = ctrl.ExecuteCreateSectionBeforeSlide("New Section", 3);
            Assert(!headlessRes.Success, "ExecuteCreateSectionBeforeSlide on null app should fail cleanly");

            // Exception path: Mock controller that throws to force outer catch
            var throwingCtrl = new ThrowingCreateSectionController();
            var exceptionRes = throwingCtrl.ExecuteCreateSectionBeforeSlide("New Section", 1);
            Assert(!exceptionRes.Success, "ExecuteCreateSectionBeforeSlide should fail when CreateSectionBeforeSlide throws");
            Assert(exceptionRes.ErrorCode != 0, "Expected non-zero ErrorCode (HRESULT) when exception caught");
            Assert(exceptionRes.ErrorMessage.Contains("ExecuteCreateSectionBeforeSlide"), "Expected operation name in error message");
        }

        private static void TestExecuteRenameSectionResults(PowerPointController ctrl)
        {
            // Validation: invalid section index
            var invalidIndexRes = ctrl.ExecuteRenameSectionInPlace(0, "Renamed");
            Assert(!invalidIndexRes.Success, "ExecuteRenameSectionInPlace(0, 'Renamed') should fail");
            Assert(invalidIndexRes.ErrorMessage.Contains("at least 1"), "Expected validation message for section index");

            // Validation: empty section name
            var emptyNameRes = ctrl.ExecuteRenameSectionInPlace(1, "");
            Assert(!emptyNameRes.Success, "ExecuteRenameSectionInPlace(1, empty) should fail");
            Assert(emptyNameRes.ErrorMessage.Contains("empty"), "Expected validation message for section name");

            // Headless: valid parameters but no presentation
            var headlessRes = ctrl.ExecuteRenameSectionInPlace(1, "Renamed");
            Assert(!headlessRes.Success, "ExecuteRenameSectionInPlace on null app should fail cleanly");

            // Exception path: Mock controller that throws to force outer catch
            var throwingCtrl = new ThrowingRenameSectionController();
            var exceptionRes = throwingCtrl.ExecuteRenameSectionInPlace(1, "Renamed");
            Assert(!exceptionRes.Success, "ExecuteRenameSectionInPlace should fail when RenameSection throws");
            Assert(exceptionRes.ErrorCode != 0, "Expected non-zero ErrorCode (HRESULT) when exception caught");
            Assert(exceptionRes.ErrorMessage.Contains("ExecuteRenameSectionInPlace"), "Expected operation name in error message");
        }

        private static void TestExecuteSetSpeakerNotesResults(PowerPointController ctrl)
        {
            // Validation: invalid slide number
            var invalidSlideRes = ctrl.ExecuteSetSpeakerNotesInPlace(-1, "Speaker notes");
            Assert(!invalidSlideRes.Success, "ExecuteSetSpeakerNotesInPlace(-1, notes) should fail");
            Assert(invalidSlideRes.ErrorMessage.Contains("at least 1"), "Expected validation message for slide number");

            // Validation: empty notes
            var emptyNotesRes = ctrl.ExecuteSetSpeakerNotesInPlace(1, "");
            Assert(!emptyNotesRes.Success, "ExecuteSetSpeakerNotesInPlace(1, empty) should fail");
            Assert(emptyNotesRes.ErrorMessage.Contains("empty"), "Expected validation message for notes");

            // Headless: valid parameters but no presentation
            var headlessRes = ctrl.ExecuteSetSpeakerNotesInPlace(1, "Speaker notes");
            Assert(!headlessRes.Success, "ExecuteSetSpeakerNotesInPlace on null app should fail cleanly");

            // Exception path: Mock controller that throws to force outer catch
            var throwingCtrl = new ThrowingSetSpeakerNotesController();
            var exceptionRes = throwingCtrl.ExecuteSetSpeakerNotesInPlace(1, "Speaker notes");
            Assert(!exceptionRes.Success, "ExecuteSetSpeakerNotesInPlace should fail when SetSpeakerNotesForSlide throws");
            Assert(exceptionRes.ErrorCode != 0, "Expected non-zero ErrorCode (HRESULT) when exception caught");
            Assert(exceptionRes.ErrorMessage.Contains("ExecuteSetSpeakerNotesInPlace"), "Expected operation name in error message");
        }

        private static void Assert(bool condition, string message)
        {
            if (!condition)
            {
                throw new Exception("Assertion failed: " + message);
            }
        }
    }

    // Mock controller subclasses that throw exceptions to force Execute* catch blocks (Rule 10 validation)
    internal class ThrowingMoveSlideController : PowerPointController
    {
        public ThrowingMoveSlideController() : base(new object()) { }

        public override bool MoveSlide(int sourceSlideNumber, int destinationSlideNumber)
        {
            throw new System.Runtime.InteropServices.COMException("Mock COM exception for MoveSlide", unchecked((int)0x800AC472));
        }
    }

    internal class ThrowingCreateSectionController : PowerPointController
    {
        public ThrowingCreateSectionController() : base(new object()) { }

        public override bool CreateSectionBeforeSlide(string sectionName, int slideNumber)
        {
            throw new System.Runtime.InteropServices.COMException("Mock COM exception for CreateSectionBeforeSlide", unchecked((int)0x800AC472));
        }
    }

    internal class ThrowingRenameSectionController : PowerPointController
    {
        public ThrowingRenameSectionController() : base(new object()) { }

        public override bool RenameSection(int sectionIndex, string sectionName)
        {
            throw new System.Runtime.InteropServices.COMException("Mock COM exception for RenameSection", unchecked((int)0x800AC472));
        }
    }

    internal class ThrowingSetSpeakerNotesController : PowerPointController
    {
        public ThrowingSetSpeakerNotesController() : base(new object()) { }

        public override bool SetSpeakerNotesForSlide(int slideNumber, string notes)
        {
            throw new System.Runtime.InteropServices.COMException("Mock COM exception for SetSpeakerNotesForSlide", unchecked((int)0x800AC472));
        }
    }
}
