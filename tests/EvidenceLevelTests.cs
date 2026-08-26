using System;
using System.Collections.Generic;
using MSOfficeAIAssistant.UI.Cards;

namespace MSOfficeAIAssistant.Tests
{
    public static class EvidenceLevelTests
    {
        public static void RunAll()
        {
            TestCitationPatternWordParagraph();
            TestCitationPatternWordExcerpt();
            TestCitationPatternExcelSheetQualified();
            TestCitationPatternExcelBareCellTag();
            TestCitationPatternPowerPointSlide();
            TestCitationPatternPowerPointSlideTagBracketFormat();
            TestCitationPatternPowerPointSlideTagDashFormat();
            TestBracketedTagCalculated();
            TestBracketedTagStrongInference();
            TestBracketedTagPossibleInference();
            TestBracketedTagInsufficientEvidence();
            TestBracketedTagCaseInsensitive();
            TestCitationPatternWinsPriority();
            TestPlainContentWithoutSignals();
            TestNullAndEmptyContent();
            TestAllLevelsHaveNonEmptyLabels();
            TestAllLevelsHaveNonEmptyIcons();
            TestAllLevelIconsAreUnique();
            TestStripEvidenceTagRemovesLeadingTag();
            TestStripEvidenceTagPreservesContent();
        }

        private static void TestCitationPatternWordParagraph()
        {
            string content = "Finding: The document contains [¶12] some critical information";
            EvidenceLevel level = EvidenceLevelClassifier.Classify(content);
            Assert(level == EvidenceLevel.DirectlyObserved,
                "Finding content containing [¶12] pattern should be DirectlyObserved");
        }

        private static void TestCitationPatternWordExcerpt()
        {
            string content = "Finding: As mentioned in the excerpt ~Paragraph 5, the contract states...";
            EvidenceLevel level = EvidenceLevelClassifier.Classify(content);
            Assert(level == EvidenceLevel.DirectlyObserved,
                "Finding content containing ~Paragraph pattern should be DirectlyObserved");
        }

        private static void TestCitationPatternExcelSheetQualified()
        {
            string content = "Finding: The value in Sheet1!B7 shows a discrepancy";
            EvidenceLevel level = EvidenceLevelClassifier.Classify(content);
            Assert(level == EvidenceLevel.DirectlyObserved,
                "Finding content containing Sheet1!B7 pattern should be DirectlyObserved");

            string content2 = "Finding: Budget!C12 indicates the issue";
            level = EvidenceLevelClassifier.Classify(content2);
            Assert(level == EvidenceLevel.DirectlyObserved,
                "Finding content containing Budget!C12 pattern should be DirectlyObserved");
        }

        private static void TestCitationPatternExcelBareCellTag()
        {
            string content = "Finding: The cell B7=1234 shows abnormal activity";
            EvidenceLevel level = EvidenceLevelClassifier.Classify(content);
            Assert(level == EvidenceLevel.DirectlyObserved,
                "Finding content containing B7=1234 pattern should be DirectlyObserved");
        }

        private static void TestCitationPatternPowerPointSlide()
        {
            string content = "Finding: On Slide 3 of 12, the chart demonstrates this trend";
            EvidenceLevel level = EvidenceLevelClassifier.Classify(content);
            Assert(level == EvidenceLevel.DirectlyObserved,
                "Finding content containing Slide pattern should be DirectlyObserved");
        }

        /// <summary>
        /// GetSlideTextInternal's real emitted format — this is what actually reaches the model as
        /// content, unlike "Slide N of M" which is only the A2 context-readout UI label. Found
        /// missing during an adversarial review pass; fixed alongside this test.
        /// </summary>
        private static void TestCitationPatternPowerPointSlideTagBracketFormat()
        {
            string content = "Finding: [Slide #3: Overview] shows the trend clearly";
            EvidenceLevel level = EvidenceLevelClassifier.Classify(content);
            Assert(level == EvidenceLevel.DirectlyObserved,
                "Finding content containing [Slide #N: Title] pattern should be DirectlyObserved");
        }

        /// <summary>AttachmentExtractor's real emitted format for .pptx slide sections.</summary>
        private static void TestCitationPatternPowerPointSlideTagDashFormat()
        {
            string content = "Finding: --- Slide 5 --- shows the trend clearly";
            EvidenceLevel level = EvidenceLevelClassifier.Classify(content);
            Assert(level == EvidenceLevel.DirectlyObserved,
                "Finding content containing --- Slide N --- pattern should be DirectlyObserved");
        }

        private static void TestBracketedTagCalculated()
        {
            string content = "[Calculated] The total revenue is $1.5M based on these figures";
            EvidenceLevel level = EvidenceLevelClassifier.Classify(content);
            Assert(level == EvidenceLevel.Calculated,
                "Content starting with [Calculated] should classify as Calculated");
        }

        private static void TestBracketedTagStrongInference()
        {
            string content = "[Strong Inference] Given the trend, this pattern is likely to continue";
            EvidenceLevel level = EvidenceLevelClassifier.Classify(content);
            Assert(level == EvidenceLevel.StrongInference,
                "Content starting with [Strong Inference] should classify as StrongInference");
        }

        private static void TestBracketedTagPossibleInference()
        {
            string content = "[Possible Inference] The data might indicate this scenario";
            EvidenceLevel level = EvidenceLevelClassifier.Classify(content);
            Assert(level == EvidenceLevel.PossibleInference,
                "Content starting with [Possible Inference] should classify as PossibleInference");
        }

        private static void TestBracketedTagInsufficientEvidence()
        {
            string content = "[Insufficient Evidence] Without more data, I cannot confirm this";
            EvidenceLevel level = EvidenceLevelClassifier.Classify(content);
            Assert(level == EvidenceLevel.InsufficientEvidence,
                "Content starting with [Insufficient Evidence] should classify as InsufficientEvidence");
        }

        private static void TestBracketedTagCaseInsensitive()
        {
            string content1 = "[calculated] The amount is $500";
            EvidenceLevel level1 = EvidenceLevelClassifier.Classify(content1);
            Assert(level1 == EvidenceLevel.Calculated,
                "Bracketed tags should be case-insensitive (lowercase)");

            string content2 = "[STRONG INFERENCE] This trend suggests...";
            EvidenceLevel level2 = EvidenceLevelClassifier.Classify(content2);
            Assert(level2 == EvidenceLevel.StrongInference,
                "Bracketed tags should be case-insensitive (uppercase)");
        }

        private static void TestCitationPatternWinsPriority()
        {
            // Citation pattern should win over conflicting bracketed tag
            string content = "[Possible Inference] but see Sheet1!B7 for the data";
            EvidenceLevel level = EvidenceLevelClassifier.Classify(content);
            Assert(level == EvidenceLevel.DirectlyObserved,
                "Citation pattern (Sheet1!B7) should win priority over bracketed tag (Possible Inference)");
        }

        private static void TestPlainContentWithoutSignals()
        {
            string content = "Finding: This is a plain observation with no citations or tags";
            EvidenceLevel level = EvidenceLevelClassifier.Classify(content);
            Assert(level == EvidenceLevel.InsufficientEvidence,
                "Content with no citation pattern and no bracketed tag should be InsufficientEvidence");
        }

        private static void TestNullAndEmptyContent()
        {
            EvidenceLevel level1 = EvidenceLevelClassifier.Classify(null);
            Assert(level1 == EvidenceLevel.InsufficientEvidence,
                "Null content should return InsufficientEvidence");

            EvidenceLevel level2 = EvidenceLevelClassifier.Classify("");
            Assert(level2 == EvidenceLevel.InsufficientEvidence,
                "Empty content should return InsufficientEvidence");
        }

        private static void TestAllLevelsHaveNonEmptyLabels()
        {
            var values = (EvidenceLevel[])Enum.GetValues(typeof(EvidenceLevel));
            foreach (var level in values)
            {
                string label = EvidenceLevelClassifier.GetLabel(level);
                Assert(!string.IsNullOrEmpty(label),
                    string.Format("EvidenceLevel {0} has empty or null label", level));
            }
        }

        private static void TestAllLevelsHaveNonEmptyIcons()
        {
            var values = (EvidenceLevel[])Enum.GetValues(typeof(EvidenceLevel));
            foreach (var level in values)
            {
                string icon = EvidenceLevelClassifier.GetIcon(level);
                Assert(!string.IsNullOrEmpty(icon),
                    string.Format("EvidenceLevel {0} has empty or null icon", level));
            }
        }

        private static void TestAllLevelIconsAreUnique()
        {
            var values = (EvidenceLevel[])Enum.GetValues(typeof(EvidenceLevel));
            var seenIcons = new HashSet<string>();
            foreach (var level in values)
            {
                string icon = EvidenceLevelClassifier.GetIcon(level);
                Assert(!seenIcons.Contains(icon),
                    string.Format("Icon '{0}' is used by multiple levels (last was {1})", icon, level));
                seenIcons.Add(icon);
            }
        }

        private static void TestStripEvidenceTagRemovesLeadingTag()
        {
            string content = "[Calculated] The total is $1.5M";
            string stripped = EvidenceLevelClassifier.StripEvidenceTag(content);
            Assert(stripped.Equals("The total is $1.5M"),
                "StripEvidenceTag should remove [Calculated] prefix");

            string content2 = "[Strong Inference] This trend continues";
            string stripped2 = EvidenceLevelClassifier.StripEvidenceTag(content2);
            Assert(stripped2.Equals("This trend continues"),
                "StripEvidenceTag should remove [Strong Inference] prefix");
        }

        private static void TestStripEvidenceTagPreservesContent()
        {
            string content = "No tag here, just plain text";
            string stripped = EvidenceLevelClassifier.StripEvidenceTag(content);
            Assert(stripped.Equals(content),
                "StripEvidenceTag should return unchanged content when no tag present");

            string content2 = "This [mentions] something but no leading tag";
            string stripped2 = EvidenceLevelClassifier.StripEvidenceTag(content2);
            Assert(stripped2.Equals(content2),
                "StripEvidenceTag should only strip leading tags, not inline brackets");
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
