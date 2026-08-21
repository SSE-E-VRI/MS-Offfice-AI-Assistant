using System;
using MSOfficeAIAssistant.Hosts;

namespace MSOfficeAIAssistant.Tests
{
    public static class WordDocumentContextBuilderTests
    {
        public static void RunAll()
        {
            TestInferHeadingLevel();
            TestTrimToLength();
            TestBuildDocumentOutlineMarkdown();
            TestBuildDocumentOutlineNumbered();
            TestBuildActionItemContext();
            TestBuildRelevantDocumentContextRanking();
        }

        private static void TestInferHeadingLevel()
        {
            Assert(WordDocumentContextBuilder.InferHeadingLevel("Heading 1") == 1, "InferHeadingLevel 'Heading 1' should be 1");
            Assert(WordDocumentContextBuilder.InferHeadingLevel("Heading 3") == 3, "InferHeadingLevel 'Heading 3' should be 3");
            Assert(WordDocumentContextBuilder.InferHeadingLevel("Custom 9") == 9, "InferHeadingLevel 'Custom 9' should be 9");
            Assert(WordDocumentContextBuilder.InferHeadingLevel("Normal") == 1, "InferHeadingLevel 'Normal' default to 1");
            Assert(WordDocumentContextBuilder.InferHeadingLevel(null) == 1, "InferHeadingLevel null default to 1");
        }

        private static void TestTrimToLength()
        {
            Assert(WordDocumentContextBuilder.TrimToLength("Hello World", 50) == "Hello World", "TrimToLength within budget");
            Assert(WordDocumentContextBuilder.TrimToLength("Hello World", 8) == "Hello...", "TrimToLength truncated with ellipsis");
            Assert(WordDocumentContextBuilder.TrimToLength(null, 10) == string.Empty, "TrimToLength null returns empty");
            Assert(WordDocumentContextBuilder.TrimToLength("Hello", 0) == string.Empty, "TrimToLength zero budget returns empty");
        }

        private static void TestBuildDocumentOutlineMarkdown()
        {
            string doc = "# Overview\nThis is an introduction.\n\n## Architecture\nDetails here.\n\n### Storage Layer\nDatabase info.";
            string outline = WordDocumentContextBuilder.BuildDocumentOutline(doc, 1000);

            Assert(!string.IsNullOrEmpty(outline), "Outline should not be empty");
            Assert(outline.IndexOf("Overview", StringComparison.Ordinal) >= 0, "Outline contains Overview");
            Assert(outline.IndexOf("Architecture", StringComparison.Ordinal) >= 0, "Outline contains Architecture");
            Assert(outline.IndexOf("Storage Layer", StringComparison.Ordinal) >= 0, "Outline contains Storage Layer");
        }

        private static void TestBuildDocumentOutlineNumbered()
        {
            string doc = "1. Introduction\nText here\n\n1.1 Scope\nScope text\n\n2. Implementation\nDetails";
            string outline = WordDocumentContextBuilder.BuildDocumentOutline(doc, 1000);

            Assert(!string.IsNullOrEmpty(outline), "Numbered outline should not be empty");
            Assert(outline.IndexOf("1. Introduction", StringComparison.Ordinal) >= 0, "Outline contains 1. Introduction");
            Assert(outline.IndexOf("1.1 Scope", StringComparison.Ordinal) >= 0, "Outline contains 1.1 Scope");
        }

        private static void TestBuildActionItemContext()
        {
            string doc = "Project status:\n- [ ] Todo: update dependencies by Friday\n- Follow-up with security team\nRegular paragraph with no action cues.";
            string actions = WordDocumentContextBuilder.BuildActionItemContext(doc, 1000);

            Assert(!string.IsNullOrEmpty(actions), "Actions should not be empty");
            Assert(actions.IndexOf("Todo", StringComparison.OrdinalIgnoreCase) >= 0, "Actions contains Todo");
            Assert(actions.IndexOf("Follow-up", StringComparison.OrdinalIgnoreCase) >= 0, "Actions contains Follow-up");
            Assert(actions.IndexOf("Regular paragraph", StringComparison.OrdinalIgnoreCase) < 0, "Actions excludes regular paragraphs");
        }

        private static void TestBuildRelevantDocumentContextRanking()
        {
            string doc = "Introduction section with basic background.\n\n" +
                         "Database Migration section discussing schema upgrade and data integrity.\n\n" +
                         "Conclusion and next steps for the deployment.";

            string context = WordDocumentContextBuilder.BuildRelevantDocumentContext(
                doc,
                "Explain the database migration and schema upgrade plan",
                500,
                null,
                null,
                null);

            Assert(!string.IsNullOrEmpty(context), "Relevant context should not be empty");
            Assert(context.IndexOf("Database Migration", StringComparison.Ordinal) >= 0, "Relevant context should include ranked database section");
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
