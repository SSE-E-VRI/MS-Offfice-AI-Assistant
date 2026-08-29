using System;
using System.Collections.Generic;
using System.Linq;
using MSOfficeAIAssistant.Core.QuickPrompts;

namespace MSOfficeAIAssistant.Tests
{
    public static class QuickPromptRegistryTests
    {
        public static void RunAll()
        {
            TestGetPromptsHostFiltering();
            TestDeckEntryHasCorrectLabelAndPromptText();
            TestHostFilterLogicWithSyntheticEntry();
            TestGetRibbonPromptsStructure();
            TestGetRibbonPromptsSummarizeEntry();
        }

        /// <summary>
        /// "Deck" carries HostFilter = "PowerPoint" (restoring the original XAML's
        /// isPowerPoint ? Visible : Collapsed rule on BtnQuickDeck) — only PowerPoint sees
        /// all 6; every other host (and an unknown/null host) sees the other 5.
        /// </summary>
        private static void TestGetPromptsHostFiltering()
        {
            var wordPrompts = QuickPromptRegistry.GetPrompts("Word");
            Assert(wordPrompts.Count == 5, string.Format("Expected 5 prompts for Word, got {0}", wordPrompts.Count));
            Assert(!wordPrompts.Any(p => p.Id == "Deck"), "Expected Deck to be excluded for Word");

            var excelPrompts = QuickPromptRegistry.GetPrompts("Excel");
            Assert(excelPrompts.Count == 7, string.Format("Expected 7 prompts for Excel, got {0}", excelPrompts.Count));
            Assert(!excelPrompts.Any(p => p.Id == "Deck"), "Expected Deck to be excluded for Excel");
            Assert(excelPrompts.Any(p => p.Id == "AnalyzeExcel"), "Expected AnalyzeExcel for Excel");
            Assert(excelPrompts.Any(p => p.Id == "FormulaVariants"), "Expected FormulaVariants for Excel");

            var pptPrompts = QuickPromptRegistry.GetPrompts("PowerPoint");
            Assert(pptPrompts.Count == 9, string.Format("Expected 9 prompts for PowerPoint, got {0}", pptPrompts.Count));
            Assert(pptPrompts.Any(p => p.Id == "Deck"), "Expected Deck to be included for PowerPoint");
            Assert(pptPrompts.Any(p => p.Id == "NotesForAll"), "Expected NotesForAll for PowerPoint");
            Assert(pptPrompts.Any(p => p.Id == "ExcelToSlide"), "Expected ExcelToSlide for PowerPoint");

            var unknownPrompts = QuickPromptRegistry.GetPrompts("UnknownHost");
            Assert(unknownPrompts.Count == 5, string.Format("Expected 5 prompts for unknown host, got {0}", unknownPrompts.Count));

            var nullPrompts = QuickPromptRegistry.GetPrompts(null);
            Assert(nullPrompts.Count == 5, string.Format("Expected 5 prompts for null host type, got {0}", nullPrompts.Count));
        }

        private static void TestDeckEntryHasCorrectLabelAndPromptText()
        {
            var prompts = QuickPromptRegistry.GetPrompts("PowerPoint");
            var deckPrompt = prompts.FirstOrDefault(p => p.Id == "Deck");
            Assert(deckPrompt != null, "Expected to find a Deck prompt");
            Assert(deckPrompt.Label == "Build deck", string.Format("Expected Label to be 'Build deck', got '{0}'", deckPrompt.Label));
            Assert(deckPrompt.PromptText.StartsWith("Create a concise, coherent slide deck"),
                string.Format("Expected PromptText to start with 'Create a concise, coherent slide deck', got '{0}'", deckPrompt.PromptText));
        }

        private static void TestHostFilterLogicWithSyntheticEntry()
        {
            // Create a synthetic list with a PowerPoint-specific entry
            var synthPrompts = new List<QuickPrompt>
            {
                new QuickPrompt { Id = "Test1", Label = "Test 1", PromptText = "Test prompt 1", HostFilter = null },
                new QuickPrompt { Id = "Test2", Label = "Test 2", PromptText = "Test prompt 2", HostFilter = "PowerPoint" },
                new QuickPrompt { Id = "Test3", Label = "Test 3", PromptText = "Test prompt 3", HostFilter = "Word" }
            };

            // Test MatchesHost with PowerPoint host type
            var pptMatches = synthPrompts.Where(p => QuickPromptRegistry.MatchesHost(p, "PowerPoint")).ToList();
            Assert(pptMatches.Count == 2, string.Format("Expected 2 prompts for PowerPoint, got {0}", pptMatches.Count));
            Assert(pptMatches.Any(p => p.Id == "Test1"), "Expected Test1 (no filter) to match PowerPoint");
            Assert(pptMatches.Any(p => p.Id == "Test2"), "Expected Test2 (PowerPoint filter) to match PowerPoint");
            Assert(!pptMatches.Any(p => p.Id == "Test3"), "Expected Test3 (Word filter) to NOT match PowerPoint");

            // Test MatchesHost with Word host type
            var wordMatches = synthPrompts.Where(p => QuickPromptRegistry.MatchesHost(p, "Word")).ToList();
            Assert(wordMatches.Count == 2, string.Format("Expected 2 prompts for Word, got {0}", wordMatches.Count));
            Assert(wordMatches.Any(p => p.Id == "Test1"), "Expected Test1 (no filter) to match Word");
            Assert(!wordMatches.Any(p => p.Id == "Test2"), "Expected Test2 (PowerPoint filter) to NOT match Word");
            Assert(wordMatches.Any(p => p.Id == "Test3"), "Expected Test3 (Word filter) to match Word");

            // Test case-insensitivity
            var lowerMatches = synthPrompts.Where(p => QuickPromptRegistry.MatchesHost(p, "powerpoint")).ToList();
            Assert(lowerMatches.Count == 2, string.Format("Expected case-insensitive match to work, got {0}", lowerMatches.Count));
            Assert(lowerMatches.Any(p => p.Id == "Test2"), "Expected Test2 (PowerPoint filter) to match 'powerpoint' (lowercase)");
        }

        private static void TestGetRibbonPromptsStructure()
        {
            var ribbonPrompts = QuickPromptRegistry.GetRibbonPrompts();

            // Should return exactly 10 ribbon prompts (one per simple ribbon method)
            Assert(ribbonPrompts.Count == 10, string.Format("Expected 10 ribbon prompts, got {0}", ribbonPrompts.Count));

            // Expected IDs (in order)
            string[] expectedIds = new[]
            {
                "Generate", "ContinueWriting", "Summarize", "Rewrite", "Expand",
                "Shorten", "Outline", "ActionItems", "ReviewContent", "BuildSlides"
            };

            for (int i = 0; i < expectedIds.Length; i++)
            {
                Assert(i < ribbonPrompts.Count && ribbonPrompts[i].Id == expectedIds[i],
                    string.Format("Ribbon prompt {0} should have Id '{1}'", i, expectedIds[i]));
            }

            // All entries should have non-empty Id, Label, PromptText
            foreach (var prompt in ribbonPrompts)
            {
                Assert(!string.IsNullOrWhiteSpace(prompt.Id), "Ribbon prompt must have non-empty Id");
                Assert(!string.IsNullOrWhiteSpace(prompt.Label), "Ribbon prompt must have non-empty Label");
                Assert(!string.IsNullOrWhiteSpace(prompt.PromptText), "Ribbon prompt must have non-empty PromptText");
                Assert(prompt.HostFilter == null, "Ribbon prompts should not have host filters");
            }
        }

        private static void TestGetRibbonPromptsSummarizeEntry()
        {
            // Spot-check: Verify the Summarize entry's exact PromptText matches the original hardcoded value
            var ribbonPrompts = QuickPromptRegistry.GetRibbonPrompts();
            var summarizeEntry = ribbonPrompts.Find(p => p.Id == "Summarize");

            Assert(summarizeEntry != null, "Should find Summarize entry in ribbon prompts");
            Assert(summarizeEntry.Label == "Summarize", "Summarize label should match");

            string expectedPromptText = "Provide a concise executive summary highlighting key takeaways and action items.";
            Assert(summarizeEntry.PromptText == expectedPromptText,
                string.Format("Summarize PromptText mismatch. Expected: '{0}', Got: '{1}'", expectedPromptText, summarizeEntry.PromptText));
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
