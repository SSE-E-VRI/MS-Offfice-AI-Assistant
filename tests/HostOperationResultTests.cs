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
        }

        private static void TestPowerPointControllerHeadlessOperationResults()
        {
            var ctrl = new PowerPointController(null);

            var actionNullRes = ctrl.ExecutePowerPointAction(null);
            Assert(!actionNullRes.Success, "ExecutePowerPointAction(null) should fail");

            var outlineEmptyRes = ctrl.ExecuteCreateDeckFromOutline("");
            Assert(!outlineEmptyRes.Success, "ExecuteCreateDeckFromOutline(empty) should fail");

            var imageNotFoundRes = ctrl.ExecuteInsertImage("C:\\non_existent_file_path_12345.png");
            Assert(!imageNotFoundRes.Success, "ExecuteInsertImage on non-existent file should fail");
            Assert(imageNotFoundRes.ErrorMessage.Contains("not found"), "Expected file not found message");
        }

        private static void Assert(bool condition, string message)
        {
            if (!condition)
            {
                throw new Exception("Assertion failed: " + message);
            }
        }
    }
}
