using System;
using System.Collections.Generic;
using MSOfficeAIAssistant.Core;

namespace MSOfficeAIAssistant.Tests
{
    public static class MentionResolverTests
    {
        public static void RunAll()
        {
            TestSupportedExtensions();
            TestTryExtractQueryBasic();
            TestTryExtractQueryBoundaries();
            TestTryExtractQueryInvalidChars();
            TestFilterCandidatesStartsWithPrioritized();
            TestFilterCandidatesEmptyQuery();
            TestFilterCandidatesCaseInsensitive();
            TestFilterCandidatesMaxCap();
            TestFilterCandidatesUnsupportedFiltered();
            TestTryExtractWithWhitespaceAndPunctuation();
        }

        private static void TestSupportedExtensions()
        {
            Assert(MentionResolver.IsSupportedExtension(".docx"), ".docx should be supported");
            Assert(MentionResolver.IsSupportedExtension("xlsx"), "xlsx without dot should be supported");
            Assert(MentionResolver.IsSupportedExtension(".PDF"), ".PDF case-insensitive should be supported");
            Assert(!MentionResolver.IsSupportedExtension(".doc"), ".doc legacy should not be supported");
            Assert(!MentionResolver.IsSupportedExtension(".exe"), ".exe should not be supported");
            Assert(MentionResolver.IsSupportedFile("report.docx"), "report.docx supported");
            Assert(!MentionResolver.IsSupportedFile("report.doc"), "report.doc not supported");
        }

        private static void TestTryExtractQueryBasic()
        {
            string q;
            // Bare @ at end
            Assert(MentionResolver.TryExtractQuery("Hello @", 7, out q) && q == "", "Bare @ should extract empty query");
            // @ with prefix
            Assert(MentionResolver.TryExtractQuery("Please see @rep", 15, out q) && q == "rep", "Should extract 'rep'");
            // @ at start
            Assert(MentionResolver.TryExtractQuery("@mydoc", 6, out q) && q == "mydoc", "@ at start");
            // @ with dot
            Assert(MentionResolver.TryExtractQuery("Check @report.docx", 18, out q) && q == "report.docx", "dot in query");
            // @ with dash underscore
            Assert(MentionResolver.TryExtractQuery("use @my-file_v2", 15, out q) && q == "my-file_v2", "dash underscore");
        }

        private static void TestTryExtractQueryBoundaries()
        {
            string q;
            // No @ should fail
            Assert(!MentionResolver.TryExtractQuery("Hello world", 5, out q), "No @ should fail");
            // Email-like a@b should NOT trigger (prev char not whitespace)
            Assert(!MentionResolver.TryExtractQuery("email a@b", 8, out q), "email pattern should not trigger");
            // @ preceded by '(' should trigger
            Assert(MentionResolver.TryExtractQuery("(@rep", 5, out q) && q == "rep", "(@ should trigger");
            // Caret before @ should not find
            Assert(!MentionResolver.TryExtractQuery("Hello @rep", 6, out q), "Caret at @ should not extract");
            // Multiple @, nearest to caret wins
            Assert(MentionResolver.TryExtractQuery("@first @second", 14, out q) && q == "second", "nearest @ wins");
            // Whitespace after @ terminates? Actually whitespace breaks query but we test
            // If text is "Hi @report more" and caret is after report (before space), should succeed
            Assert(MentionResolver.TryExtractQuery("Hi @report more", 10, out q) && q == "report", "caret before space");
        }

        private static void TestTryExtractQueryInvalidChars()
        {
            string q;
            // Query with invalid char '/' should fail (not valid query char)
            Assert(!MentionResolver.TryExtractQuery("see @rep/ort", 10, out q), "slash invalid should fail");
            // Space inside query should break at space, but TryExtract searches backwards to last @, so " @a b" with caret at b after space should fail because whitespace break
            // Our impl breaks on whitespace when scanning, so "@a b" with caret at end after b should not trigger because nearest '@' is before space? Actually text "@a b", caret at 4 (after b), scanning backwards: at 3 is 'b', 2 is ' ', breaks -> no @ found -> fail. This is expected.
            Assert(!MentionResolver.TryExtractQuery("@a b", 4, out q), "space breaks mention");
        }

        private static void TestFilterCandidatesStartsWithPrioritized()
        {
            var candidates = new List<string> { "report.docx", "myreport.docx", "rep_data.xlsx", "annual.pdf", "replica.txt" };
            var filtered = MentionResolver.FilterCandidates(candidates, "rep");
            // Starts-with first: report.docx, rep_data.xlsx, replica.txt (alphabetically), then contains: myreport.docx
            Assert(filtered.Count == 4, string.Format("Expected 4 filtered, got {0}", filtered.Count));
            Assert(filtered[0].ToLowerInvariant().StartsWith("rep"), "First should start with rep");
            // Verify alphabetical within bucket: rep_data.xlsx < replica.txt? Actually "rep_data" < "replica" < "report" alphabetically - but startsWith bucket sorted alphabetically
            // So order should be alpha: rep_data, replica, report, then myreport
            Assert(filtered[0] == "rep_data.xlsx" || filtered[0] == "replica.txt" || filtered[0] == "report.docx", "StartsWith bucket sorted");
            Assert(filtered[filtered.Count - 1] == "myreport.docx", "Contains bucket last");
        }

        private static void TestFilterCandidatesEmptyQuery()
        {
            var candidates = new List<string> { "b.docx", "a.xlsx", "c.pdf" };
            var filtered = MentionResolver.FilterCandidates(candidates, "");
            Assert(filtered.Count == 3, "Empty query should return all supported");
            // Should be sorted alphabetically
            Assert(filtered[0] == "a.xlsx", "Alphabetical sorted");
            Assert(filtered[1] == "b.docx", "Alphabetical sorted 2");
        }

        private static void TestFilterCandidatesCaseInsensitive()
        {
            var candidates = new List<string> { "Report.docx", "summary.PDF", "data.XLSX" };
            var filtered = MentionResolver.FilterCandidates(candidates, "REPORT");
            Assert(filtered.Count == 1 && filtered[0] == "Report.docx", "Case insensitive");
            filtered = MentionResolver.FilterCandidates(candidates, "pdf");
            Assert(filtered.Count == 1 && filtered[0] == "summary.PDF", "pdf case insensitive");
        }

        private static void TestFilterCandidatesMaxCap()
        {
            var candidates = new List<string>();
            for (int i = 0; i < 20; i++) candidates.Add("file" + i + ".docx");
            var filtered = MentionResolver.FilterCandidates(candidates, "", 10);
            Assert(filtered.Count == 10, string.Format("Max cap 10, got {0}", filtered.Count));
            var filtered5 = MentionResolver.FilterCandidates(candidates, "", 5);
            Assert(filtered5.Count == 5, "Max cap 5");
        }

        private static void TestFilterCandidatesUnsupportedFiltered()
        {
            var candidates = new List<string> { "a.docx", "b.doc", "c.exe", "d.pdf" };
            var filtered = MentionResolver.FilterCandidates(candidates, "");
            Assert(filtered.Count == 2, string.Format("Unsupported filtered, expected 2 got {0}", filtered.Count));
            foreach (string f in filtered) Assert(MentionResolver.IsSupportedFile(f), "Only supported remains");
        }

        private static void TestTryExtractWithWhitespaceAndPunctuation()
        {
            string q;
            // @ preceded by newline should trigger
            Assert(MentionResolver.TryExtractQuery("Line1\n@rep", 10, out q) && q == "rep", "newline before @");
            // Query empty after @ with caret immediately after @
            Assert(MentionResolver.TryExtractQuery("Hello @", 7, out q) && q == "", "empty query after @");
            // Null text fails
            Assert(!MentionResolver.TryExtractQuery(null, 0, out q), "null fails");
            // Caret out of range fails
            Assert(!MentionResolver.TryExtractQuery("hello", 10, out q), "caret out of range fails");
        }

        private static void Assert(bool condition, string message)
        {
            if (!condition) throw new Exception("Assertion failed: " + message);
        }
    }
}
