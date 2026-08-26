using System;
using System.Collections.Generic;
using System.Linq;
using MSOfficeAIAssistant.Core.QuickPrompts;
using MSOfficeAIAssistant.Core.Skills;

namespace MSOfficeAIAssistant.Tests
{
    public static class SkillPickerTests
    {
        public static void RunAll()
        {
            TestToQuickPromptNull();
            TestToQuickPromptMapsAllFields();
            TestIsFailureRelatedContextKeywordMatch();
            TestIsFailureRelatedContextUnrelated();
            TestIsFailureRelatedContextNullEmpty();
            TestSelectChipsNoFailureContextDoesNotPromote();
            TestSelectChipsWithFailureContextPromotesFailureAnalysis();
            TestSelectChipsGeneralPackNoFailureAnalysis();
            TestSelectChipsRespectsMaxChips();
            TestSelectChipsHostFiltering();
            TestSelectChipsNullArgumentsSafe();
        }

        private static void TestToQuickPromptNull()
        {
            QuickPrompt result = SkillPicker.ToQuickPrompt(null);
            Assert(result == null, "ToQuickPrompt(null) should return null");
        }

        private static void TestToQuickPromptMapsAllFields()
        {
            var skill = new Skill
            {
                Id = "test_skill_id",
                Name = "Test Skill Name",
                PromptTemplate = "Test prompt template text",
                PreferredHost = "Excel",
                Description = "Test description",
                RequiredContext = new List<string>(),
                OutputStructure = "JSON",
                DefaultMode = "Review",
                RiskCeiling = 1,
                DomainPack = "general"
            };

            QuickPrompt result = SkillPicker.ToQuickPrompt(skill);
            Assert(result != null, "ToQuickPrompt should not return null for a valid skill");
            Assert(result.Id == "test_skill_id", "Id should map correctly");
            Assert(result.Label == "Test Skill Name", "Label should map from Name");
            Assert(result.PromptText == "Test prompt template text", "PromptText should map from PromptTemplate");
            Assert(result.HostFilter == "Excel", "HostFilter should map from PreferredHost");
        }

        private static void TestIsFailureRelatedContextKeywordMatch()
        {
            bool result1 = SkillPicker.IsFailureRelatedContext("Columns: Col A [Failure Type] | Col B [Count]");
            Assert(result1, "Should detect 'failure' keyword");

            bool result2 = SkillPicker.IsFailureRelatedContext("Data contains breakdown analysis");
            Assert(result2, "Should detect 'breakdown' keyword");

            bool result3 = SkillPicker.IsFailureRelatedContext("FAULT LOG: System errors");
            Assert(result3, "Should detect 'FAULT' keyword (case-insensitive)");

            bool result4 = SkillPicker.IsFailureRelatedContext("Sheet with defect tracking");
            Assert(result4, "Should detect 'defect' keyword");

            bool result5 = SkillPicker.IsFailureRelatedContext("Equipment Breakdown Status");
            Assert(result5, "Should detect 'Breakdown' keyword with mixed case");
        }

        private static void TestIsFailureRelatedContextUnrelated()
        {
            bool result = SkillPicker.IsFailureRelatedContext("Columns: Col A [Name] | Col B [Date] | Col C [Value]");
            Assert(!result, "Should not detect failure context in unrelated text");
        }

        private static void TestIsFailureRelatedContextNullEmpty()
        {
            bool nullResult = SkillPicker.IsFailureRelatedContext(null);
            Assert(!nullResult, "IsFailureRelatedContext(null) should return false");

            bool emptyResult = SkillPicker.IsFailureRelatedContext("");
            Assert(!emptyResult, "IsFailureRelatedContext(\"\") should return false");

            bool whitespaceResult = SkillPicker.IsFailureRelatedContext("   ");
            Assert(!whitespaceResult, "IsFailureRelatedContext(whitespace) should return false");
        }

        private static void TestSelectChipsNoFailureContextDoesNotPromote()
        {
            // Load railway pack and check that failure_analysis_pareto is NOT forced to front when no failure context
            List<QuickPrompt> chips = SkillPicker.SelectChips("railway", "Excel", "", 3);

            // The failure_analysis_pareto skill may or may not be present in the first 3 entries,
            // but it should NOT be artificially moved to position 0 without a failure context signal
            if (chips.Count > 0 && chips[0].Id == "failure_analysis_pareto")
            {
                Assert(false, "failure_analysis_pareto should NOT be promoted to first position without failure context");
            }
        }

        private static void TestSelectChipsWithFailureContextPromotesFailureAnalysis()
        {
            // Load railway pack with failure-related context
            string contextWithFailure = "Columns: Col A [Failure Type] | Col B [Count]";
            List<QuickPrompt> chips = SkillPicker.SelectChips("railway", "Excel", contextWithFailure, 3);

            // failure_analysis_pareto should be the first entry if it exists in the pack
            bool found = false;
            foreach (var chip in chips)
            {
                if (chip.Id == "failure_analysis_pareto")
                {
                    found = true;
                    break;
                }
            }

            // Railway pack should contain failure_analysis_pareto, so it should be found
            Assert(found, "failure_analysis_pareto should be present in railway pack");

            // And if found, it should be first
            Assert(chips.Count > 0 && chips[0].Id == "failure_analysis_pareto",
                "failure_analysis_pareto should be the first entry when failure context is detected");
        }

        private static void TestSelectChipsGeneralPackNoFailureAnalysis()
        {
            // General pack does not have failure_analysis_pareto
            string contextWithFailure = "Columns: Col A [Failure Type] | Col B [Count]";
            List<QuickPrompt> chips = SkillPicker.SelectChips("general", "Excel", contextWithFailure, 3);

            // Should not error; should return a normal list
            Assert(chips != null, "SelectChips should not throw, should return a list");

            // Verify failure_analysis_pareto is NOT in the list (general pack doesn't have it)
            foreach (var chip in chips)
            {
                Assert(chip.Id != "failure_analysis_pareto",
                    "failure_analysis_pareto should not be in general pack");
            }
        }

        private static void TestSelectChipsRespectsMaxChips()
        {
            // maxChips = 1 should return at most 1 entry
            List<QuickPrompt> result1 = SkillPicker.SelectChips("railway", "Excel", "", 1);
            Assert(result1.Count <= 1, string.Format("SelectChips with maxChips=1 should return at most 1, got {0}", result1.Count));

            // maxChips = 3 should return at most 3 entries
            List<QuickPrompt> result3 = SkillPicker.SelectChips("railway", "Excel", "", 3);
            Assert(result3.Count <= 3, string.Format("SelectChips with maxChips=3 should return at most 3, got {0}", result3.Count));

            // maxChips = 0 should return empty
            List<QuickPrompt> result0 = SkillPicker.SelectChips("railway", "Excel", "", 0);
            Assert(result0.Count == 0, "SelectChips with maxChips=0 should return empty list");

            // maxChips = -1 should return empty
            List<QuickPrompt> resultNeg = SkillPicker.SelectChips("railway", "Excel", "", -1);
            Assert(resultNeg.Count == 0, "SelectChips with maxChips=-1 should return empty list");
        }

        private static void TestSelectChipsHostFiltering()
        {
            // Test that host filtering works: skills with non-matching PreferredHost are excluded
            // Railway pack has skills with various PreferredHost values
            List<QuickPrompt> excelChips = SkillPicker.SelectChips("railway", "Excel", "", 10);

            // All returned chips should either have null PreferredHost or "Excel" (case-insensitive)
            foreach (var chip in excelChips)
            {
                // We can't directly verify this without knowing the raw Skill objects,
                // but we can at least verify the method doesn't throw
                Assert(!string.IsNullOrEmpty(chip.Id), "All returned chips should have non-empty Id");
            }

            // Test that calling with a different host type returns results (no throw)
            List<QuickPrompt> wordChips = SkillPicker.SelectChips("railway", "Word", "", 10);
            Assert(wordChips != null, "SelectChips with Word host should not throw");
        }

        private static void TestSelectChipsNullArgumentsSafe()
        {
            // Null domainPack should return empty, not throw
            List<QuickPrompt> result1 = SkillPicker.SelectChips(null, "Excel", "", 3);
            Assert(result1 != null && result1.Count == 0, "SelectChips(null pack) should return empty list, not throw");

            // Null hostType should return empty, not throw
            List<QuickPrompt> result2 = SkillPicker.SelectChips("railway", null, "", 3);
            Assert(result2 != null && result2.Count == 0, "SelectChips(null hostType) should return empty list, not throw");

            // Null contextText should work fine (treated as empty)
            List<QuickPrompt> result3 = SkillPicker.SelectChips("railway", "Excel", null, 3);
            Assert(result3 != null, "SelectChips(null contextText) should not throw");

            // Empty pack should return empty
            List<QuickPrompt> result4 = SkillPicker.SelectChips("nonexistent", "Excel", "", 3);
            Assert(result4 != null && result4.Count == 0, "SelectChips(nonexistent pack) should return empty list, not throw");
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
