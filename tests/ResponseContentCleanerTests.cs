using System;
using MSOfficeAIAssistant.Core;

namespace MSOfficeAIAssistant.Tests
{
    public static class ResponseContentCleanerTests
    {
        public static void RunAll()
        {
            TestStripsLeadInAndCustomizationNotes();
            TestStripsKeyFeaturesTrailerButKeepsLetterRule();
            TestKeepsBodyKeyFeaturesWithoutTrailingRule();
            TestKeepsBodyHeadingsEndingInColon();
            TestKeepsNoteLineInsideLetterBody();
            TestStripsTrailingClosingRemark();
            TestLeavesCleanContentUntouched();
            TestNeverReturnsEmpty();
            TestHandlesNullAndEmpty();
            TestCurlyApostropheLeadIn();
            TestDetectsEditAnalysisReportWithComparisonTable();
            TestDetectsEditAnalysisReportByHeadingsAlone();
            TestDoesNotFlagALegitimateSpecTableInAFinishedLetter();
            TestDoesNotFlagASingleCoincidentalHeading();
        }

        /// <summary>
        /// The reported bug: the whole chat wrapper was being inserted into the letter --
        /// "Here's a polished draft..." at the top and the "Notes for Customization" list
        /// at the bottom, each separated from the body by a "---" rule.
        /// </summary>
        private static void TestStripsLeadInAndCustomizationNotes()
        {
            string response =
                "Here's a polished draft for your request letter:\n" +
                "\n---\n\n" +
                "**[Your Department's Letterhead]**\n" +
                "[Date]\n\n" +
                "Subject: Request for Temporary Allocation of 5 HP Submersible Pump\n\n" +
                "Respected Sir,\n\n" +
                "I kindly request the temporary allocation of one (1) unit.\n\n" +
                "Yours faithfully,\n" +
                "**[Your Name]**\n" +
                "\n---\n\n" +
                "**Notes for Customization:**\n" +
                "1. Replace placeholders with specific details.\n" +
                "2. Attach a supporting document if available.\n";

            string cleaned = ResponseContentCleaner.ExtractInsertableContent(response);

            Assert(!cleaned.Contains("polished draft"), "Lead-in sentence should be removed");
            Assert(!cleaned.Contains("Notes for Customization"), "Trailer heading should be removed");
            Assert(!cleaned.Contains("Replace placeholders"), "Trailer body should be removed");
            Assert(cleaned.StartsWith("**[Your Department's Letterhead]**"), "Body should start at the letterhead, got: " + First(cleaned));
            Assert(cleaned.Contains("Subject: Request for Temporary Allocation"), "Subject line must survive");
            Assert(cleaned.TrimEnd().EndsWith("**[Your Name]**"), "Body should end at the signature, got: " + Last(cleaned));
        }

        /// <summary>
        /// The official-letter case: the letter carries its own "---" rule before an "Encl."
        /// block, and the model appends a second rule plus a "Key Features" critique of its own
        /// draft. Only the LAST rule is the cut point, so the enclosure survives and the
        /// commentary does not.
        /// </summary>
        private static void TestStripsKeyFeaturesTrailerButKeepsLetterRule()
        {
            string response =
                "Here's a polished draft for your request letter:\n\n---\n\n" +
                "**SOUTHERN RAILWAY**\n**ELECTRICAL DEPARTMENT**\n\nSir,\n\n" +
                "1. **Reference:** Nil.\n2. **Context:**\n   2.1. It is submitted that the Division needs a pump.\n\n" +
                "**Yours faithfully,**\n\n[Your Full Name]\nSenior Section Engineer (Electrical)\n\n" +
                "---\n**Encl.:** [List any attachments]\n\n" +
                "---\n### **Key Features:**\n1. **Formal numbering** (e.g., \"2.1\").\n2. **Structure:** reference line for tracking.\n";

            string cleaned = ResponseContentCleaner.ExtractInsertableContent(response);

            Assert(!cleaned.Contains("polished draft"), "Lead-in should be removed");
            Assert(!cleaned.Contains("Key Features"), "Trailing critique should be removed");
            Assert(!cleaned.Contains("Formal numbering"), "Trailer body should be removed");
            Assert(cleaned.Contains("Encl.:"), "The letter's own rule + enclosure block must survive");
            Assert(cleaned.Contains("SOUTHERN RAILWAY"), "Letterhead must survive");
            Assert(cleaned.Contains("Senior Section Engineer (Electrical)"), "Signature block must survive");
        }

        /// <summary>
        /// "Key Features" is only a trailer cue when it follows a thematic break -- a product or
        /// spec document can carry that heading in its body, so the unguarded pass must ignore it.
        /// </summary>
        private static void TestKeepsBodyKeyFeaturesWithoutTrailingRule()
        {
            string response =
                "Product Overview\n\nThis release focuses on speed.\n\n" +
                "Key Features:\n- Faster indexing\n- Lower memory use\n\nAvailability: next month.";

            string cleaned = ResponseContentCleaner.ExtractInsertableContent(response);
            Assert(cleaned.Contains("Key Features:"), "A body 'Key Features' heading must survive: " + cleaned);
            Assert(cleaned.Contains("Availability: next month."), "Content after it must survive");
        }

        /// <summary>
        /// "Details of Required Pump:" ends in a colon like a lead-in does. Only the FIRST
        /// line is ever considered a preamble, so mid-body headings must be untouched.
        /// </summary>
        private static void TestKeepsBodyHeadingsEndingInColon()
        {
            string response =
                "Subject: Pump request\n\n" +
                "Respected Sir,\n\n" +
                "Details of Required Pump:\n" +
                "- Capacity: 5 HP\n" +
                "- Quantity: 1 No.\n\n" +
                "Thanking you in anticipation.";

            string cleaned = ResponseContentCleaner.ExtractInsertableContent(response);
            Assert(cleaned.Contains("Details of Required Pump:"), "Body heading must be preserved");
            Assert(cleaned.Contains("Capacity: 5 HP"), "List items must be preserved");
        }

        /// <summary>
        /// A real letter can carry its own "Note:" paragraph. Without a thematic break marking
        /// the end of the document, a bare note heading is body text, not a chat trailer.
        /// </summary>
        private static void TestKeepsNoteLineInsideLetterBody()
        {
            string response =
                "Respected Sir,\n\n" +
                "I request the allocation of one pump.\n\n" +
                "Note: The equipment will be returned in original condition.\n\n" +
                "Yours faithfully,\n[Your Name]";

            string cleaned = ResponseContentCleaner.ExtractInsertableContent(response);
            Assert(cleaned.Contains("Note: The equipment will be returned"), "Body note must survive");
            Assert(cleaned.TrimEnd().EndsWith("[Your Name]"), "Signature must survive");
        }

        private static void TestStripsTrailingClosingRemark()
        {
            string response =
                "Dear Sir,\n\n" +
                "Please approve the attached request at the earliest.\n\n" +
                "Yours faithfully,\n" +
                "[Your Name]\n\n" +
                "Let me know if you'd like a more formal tone.";

            string cleaned = ResponseContentCleaner.ExtractInsertableContent(response);
            Assert(!cleaned.Contains("Let me know"), "Closing chat remark should be removed");
            Assert(cleaned.TrimEnd().EndsWith("[Your Name]"), "Signature must survive, got: " + Last(cleaned));
        }

        private static void TestLeavesCleanContentUntouched()
        {
            string response =
                "# Quarterly Report\n\n" +
                "Revenue rose 12% this quarter.\n\n" +
                "## Outlook\n\n" +
                "Growth is expected to continue.";

            string cleaned = ResponseContentCleaner.ExtractInsertableContent(response);
            Assert(cleaned.Replace("\r\n", "\n") == response, "Clean markdown must pass through unchanged");
        }

        /// <summary>A response that is nothing but chatter must insert as-is, never as blank.</summary>
        private static void TestNeverReturnsEmpty()
        {
            string response = "Sure, here's the draft:";
            string cleaned = ResponseContentCleaner.ExtractInsertableContent(response);
            Assert(!string.IsNullOrWhiteSpace(cleaned), "Cleaner must never return empty content");
        }

        private static void TestHandlesNullAndEmpty()
        {
            Assert(ResponseContentCleaner.ExtractInsertableContent(null) == null, "null in, null out");
            Assert(ResponseContentCleaner.ExtractInsertableContent("") == "", "empty in, empty out");
        }

        /// <summary>Models emit U+2019, not the ASCII apostrophe, in "Here's".</summary>
        private static void TestCurlyApostropheLeadIn()
        {
            string response = "Here\u2019s the letter you asked for:\n\nRespected Sir,\n\nPlease approve.";
            string cleaned = ResponseContentCleaner.ExtractInsertableContent(response);
            Assert(cleaned.StartsWith("Respected Sir,"), "Curly-apostrophe lead-in should be removed, got: " + First(cleaned));
        }

        /// <summary>
        /// The reported failure: asking to grammar-check selected text got back a multi-section
        /// critique -- a revised draft, a comparison table, a "Key Improvements" rundown, and a
        /// second "recommended" draft -- and Insert applied the whole thing, table included, in
        /// place of the selection. This is the detector that lets the UI warn before that happens.
        /// </summary>
        private static void TestDetectsEditAnalysisReportWithComparisonTable()
        {
            string response =
                "### **Revised Version (Grammar-Checked & Polished)**\n" +
                "> **Subject:** Request for Temporary Allocation of a 5 HP Submersible Pump\n\n" +
                "---\n\n" +
                "### **Grammar & Clarity Corrections**\n" +
                "| **Original Text** | **Issue** | **Correction** | **Reason** |\n" +
                "|---|---|---|---|\n" +
                "| \"I, the undersigned,\" | Redundant | \"I, as the Senior...\" | Removes redundancy |\n\n" +
                "---\n\n" +
                "### **Key Improvements**\n" +
                "1. **Conciseness**: Removed \"the undersigned\".\n\n" +
                "---\n\n" +
                "### **Final Recommendation**\n" +
                "Use this version for formal letters.";

            Assert(ResponseContentCleaner.LooksLikeEditAnalysisReport(response),
                "A comparison-table response must be flagged as an analysis, not a single edit");
        }

        /// <summary>Two or more analysis headings alone (no table) is also enough to flag.</summary>
        private static void TestDetectsEditAnalysisReportByHeadingsAlone()
        {
            string response =
                "Revised Version:\nDear Sir, please approve the request.\n\n" +
                "Key Improvements:\n- Tightened the second paragraph.\n\n" +
                "Final Recommendation:\nUse the revised version above.";

            Assert(ResponseContentCleaner.LooksLikeEditAnalysisReport(response),
                "Two or more analysis-report headings must be flagged even without a table");
        }

        /// <summary>
        /// A finished letter can legitimately contain a real spec table ("Parameter |
        /// Specification") -- that must never be mistaken for a before/after comparison table.
        /// </summary>
        private static void TestDoesNotFlagALegitimateSpecTableInAFinishedLetter()
        {
            string response =
                "Respected Sir,\n\nI request one 5 HP submersible pump.\n\n" +
                "**Key Details of the Request:**\n" +
                "| **Parameter** | **Specification** |\n" +
                "|---|---|\n" +
                "| **Capacity** | 5 HP |\n" +
                "| **Quantity** | 1 No. |\n\n" +
                "Yours faithfully,\n[Your Name]";

            Assert(!ResponseContentCleaner.LooksLikeEditAnalysisReport(response),
                "A legitimate spec table inside a real letter must not be flagged");
        }

        /// <summary>One analysis-shaped heading by itself is not enough -- a real draft can be
        /// introduced with "Revised Version:" without being a critique report.</summary>
        private static void TestDoesNotFlagASingleCoincidentalHeading()
        {
            string response =
                "Revised Version:\n\nRespected Sir,\n\nPlease approve the attached request.\n\nYours faithfully,\n[Your Name]";

            Assert(!ResponseContentCleaner.LooksLikeEditAnalysisReport(response),
                "A single analysis-shaped heading alone must not be flagged: " + response);
        }

        private static string First(string s)
        {
            if (string.IsNullOrEmpty(s)) return "(empty)";
            string[] parts = s.Replace("\r\n", "\n").Split('\n');
            return parts[0];
        }

        private static string Last(string s)
        {
            if (string.IsNullOrEmpty(s)) return "(empty)";
            string[] parts = s.TrimEnd().Replace("\r\n", "\n").Split('\n');
            return parts[parts.Length - 1];
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
