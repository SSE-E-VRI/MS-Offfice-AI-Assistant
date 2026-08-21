using System;
using MSOfficeAIAssistant.Hosts;

namespace MSOfficeAIAssistant.Tests
{
    public static class OfficeHostControllerTests
    {
        public static void RunAll()
        {
            TestInterfaceImplementations();
            TestHeadlessSafeOperations();
            TestPolymorphicDispatch();
        }

        private static void TestInterfaceImplementations()
        {
            var word = new WordController(null);
            var excel = new ExcelController(null);
            var ppt = new PowerPointController(null);

            Assert(word is IOfficeHostController, "WordController must implement IOfficeHostController");
            Assert(excel is IOfficeHostController, "ExcelController must implement IOfficeHostController");
            Assert(ppt is IOfficeHostController, "PowerPointController must implement IOfficeHostController");

            Assert(((IOfficeHostController)word).HostType == "Word", "Word HostType mismatch");
            Assert(((IOfficeHostController)excel).HostType == "Excel", "Excel HostType mismatch");
            Assert(((IOfficeHostController)ppt).HostType == "PowerPoint", "PowerPoint HostType mismatch");
        }

        private static void TestHeadlessSafeOperations()
        {
            IOfficeHostController word = new WordController(null);
            IOfficeHostController excel = new ExcelController(null);
            IOfficeHostController ppt = new PowerPointController(null);

            // Document name fallbacks
            Assert(!string.IsNullOrEmpty(word.GetActiveDocumentName()), "Word document name fallback should not be empty");
            Assert(!string.IsNullOrEmpty(excel.GetActiveDocumentName()), "Excel document name fallback should not be empty");
            Assert(!string.IsNullOrEmpty(ppt.GetActiveDocumentName()), "PowerPoint document name fallback should not be empty");

            // Selected text safe defaults
            Assert(word.GetSelectedText() == string.Empty, "Word headless selected text must be empty string");
            Assert(excel.GetSelectedText() == string.Empty, "Excel headless selected text must be empty string");
            Assert(ppt.GetSelectedText() == string.Empty, "PowerPoint headless selected text must be empty string");

            // Context extraction safe defaults
            Assert(word.GetDocumentContext("test", 1000) == string.Empty, "Word headless document context must be empty string");
            Assert(excel.GetDocumentContext("test", 1000) == string.Empty, "Excel headless document context must be empty string");
            Assert(ppt.GetDocumentContext("test", 1000) == string.Empty, "PowerPoint headless document context must be empty string");

            // Mutation operations fail gracefully
            Assert(word.InsertText("sample") == false, "Word headless insert text must return false");
            Assert(excel.InsertText("sample") == false, "Excel headless insert text must return false");
            Assert(ppt.InsertText("sample") == false, "PowerPoint headless insert text must return false");

            // Undo operations fail gracefully
            Assert(word.Undo() == false, "Word headless undo must return false");
            Assert(excel.Undo() == false, "Excel headless undo must return false");
            Assert(ppt.Undo() == false, "PowerPoint headless undo must return false");
        }

        private static void TestPolymorphicDispatch()
        {
            IOfficeHostController[] controllers = new IOfficeHostController[]
            {
                new WordController(null),
                new ExcelController(null),
                new PowerPointController(null)
            };

            int count = 0;
            foreach (var ctrl in controllers)
            {
                Assert(ctrl.HostType != null, "HostType must not be null");
                Assert(ctrl.GetActiveDocumentName() != null, "GetActiveDocumentName must not be null");
                count++;
            }

            Assert(count == 3, "Expected 3 controllers in polymorphic test");
        }

        private static void Assert(bool condition, string message)
        {
            if (!condition)
            {
                throw new Exception(message);
            }
        }
    }
}
