using System;
using System.Collections.Generic;
using MSOfficeAIAssistant.Core;
using MSOfficeAIAssistant.Hosts;

namespace MSOfficeAIAssistant.Tests
{
    /// <summary>
    /// Test Suite #11: Doc-to-Deck (Briefing Deck Generator) Tests.
    /// Verifies prompt assembly, slide parsing, headless controller safety, and view state resilience.
    /// </summary>
    public static class DocToDeckTests
    {
        public static void RunAll()
        {
            TestBriefingDeckPromptAssembly();
            TestSlideOutlineParsingFromBriefingResponse();
            TestPowerPointControllerHeadlessDeckCreation();
            TestPowerPointViewStateResilience();
            TestEndToEndAttachmentToDeckSimulation();
        }

        private static void TestBriefingDeckPromptAssembly()
        {
            // 1. Topic and Attachment
            string promptWithBoth = PromptAssembler.BuildBriefingDeckPrompt(
                "Fleet Electrification Strategy",
                "[Attachment: Electrification_Brief_2026.docx]\nTotal fleet size is 1,200 locos; 450 scheduled for battery-electric conversion by 2028.",
                6);

            Assert(promptWithBoth.Contains("Create a concise, executive briefing deck of 6 slides focusing on: Fleet Electrification Strategy"),
                "Prompt should contain slide count and topic focus");
            Assert(promptWithBoth.Contains("[Source Document Excerpts]:"), "Prompt should contain document excerpts section");
            Assert(promptWithBoth.Contains("Total fleet size is 1,200 locos"), "Prompt should include excerpt content");
            Assert(promptWithBoth.Contains("Slide 1: [Executive Title]"), "Prompt should include slide structure template");
            Assert(promptWithBoth.Contains("Speaker Notes:"), "Prompt should mandate speaker notes");
            Assert(promptWithBoth.Contains("Visual suggestion:"), "Prompt should mandate visual suggestion");

            // 2. Attachment Only (No explicit topic)
            string promptAttachmentOnly = PromptAssembler.BuildBriefingDeckPrompt(
                null,
                "[Attachment: Q3_Financials.pdf]\nOperating ratio improved by 2.4% to 61.2%.",
                4);

            Assert(promptAttachmentOnly.Contains("Create a concise, executive briefing deck of 4 slides based on the attached document"),
                "Prompt should indicate attachment-based generation when topic is omitted");
            Assert(promptAttachmentOnly.Contains("Operating ratio improved by 2.4%"), "Prompt should contain excerpt");

            // 3. Fallback on invalid slide count
            string promptFallbackCount = PromptAssembler.BuildBriefingDeckPrompt("Safety Review", null, 0);
            Assert(promptFallbackCount.Contains("Create a concise, executive briefing deck of 5 slides"),
                "Prompt should default to 5 slides when target count <= 0");
        }

        private static void TestSlideOutlineParsingFromBriefingResponse()
        {
            string aiResponse =
                "Slide 1: High-Speed Rail Electrification\n" +
                "- 25kV 50Hz AC overhead catenary system deployment\n" +
                "- Phase 1 covers 450 route kilometers across Central Corridor\n" +
                "- Projected 28% reduction in traction energy costs\n" +
                "Visual suggestion: Map of Central Corridor highlighting Phase 1 sub-stations\n" +
                "Speaker Notes: Emphasize that civil engineering works are 94% complete as of Q2.\n\n" +
                "Slide 2: Traction Substation Telemetry & SCADA\n" +
                "- IEC 61850 protocol standard across all 12 substations\n" +
                "- Automated failover with redundant fiber optic backhaul\n" +
                "- Sub-second fault isolation and telemetry streaming\n" +
                "Visual suggestion: Architecture schematic of redundant substation optical ring\n" +
                "Speaker Notes: Address the cybersecurity compliance with EN 50126 and IEC 62443.\n\n" +
                "Slide 3: Project Milestones & Governance\n" +
                "- Commissioning trials start Q4 2026\n" +
                "- Independent safety assessor sign-off required prior to commercial service\n" +
                "Visual suggestion: Timeline milestone chart\n" +
                "Speaker Notes: Final board review is scheduled for November 15th.";

            var slides = PowerPointActionParser.ParseSlideData(aiResponse);
            Assert(slides.Count == 3, string.Format("Expected 3 parsed slides, got {0}", slides.Count));

            // Verify Slide 1
            Assert(slides[0].Title == "High-Speed Rail Electrification", "Slide 1 title mismatch");
            Assert(slides[0].Bullets.Count == 3, "Slide 1 should have 3 bullets");
            Assert(slides[0].Bullets[0].Contains("25kV 50Hz AC overhead"), "Slide 1 bullet 1 mismatch");
            Assert(slides[0].VisualSuggestion.Contains("Central Corridor"), "Slide 1 visual mismatch");
            Assert(slides[0].SpeakerNotes.Contains("civil engineering works"), "Slide 1 speaker notes mismatch");

            // Verify Slide 2
            Assert(slides[1].Title == "Traction Substation Telemetry & SCADA", "Slide 2 title mismatch");
            Assert(slides[1].Bullets.Count == 3, "Slide 2 should have 3 bullets");
            Assert(slides[1].VisualSuggestion.Contains("Architecture schematic"), "Slide 2 visual mismatch");
            Assert(slides[1].SpeakerNotes.Contains("IEC 62443"), "Slide 2 speaker notes mismatch");

            // Verify Slide 3
            Assert(slides[2].Title == "Project Milestones & Governance", "Slide 3 title mismatch");
            Assert(slides[2].Bullets.Count == 2, "Slide 3 should have 2 bullets");
            Assert(slides[2].SpeakerNotes.Contains("November 15th"), "Slide 3 speaker notes mismatch");
        }

        private static void TestPowerPointControllerHeadlessDeckCreation()
        {
            var pptCtrl = new PowerPointController(null);

            // Null app safety
            Assert(pptCtrl.InsertText(null) == false, "InsertText(null) should return false");
            Assert(pptCtrl.InsertText(string.Empty) == false, "InsertText(empty) should return false");
            Assert(pptCtrl.InsertText("Slide 1: Test\n- Bullet 1") == false, "Headless InsertText should return false without crashing");

            // Outline builder — routed through the HostOperationResult wrapper (adversarial-review
            // fix: the void CreateOrUpdateDeckFromOutline that used to sit here was deleted as dead
            // code once BtnInsertMessage_Click was rewired onto this same wrapper, see D-15/D-16).
            var deckResult = pptCtrl.ExecuteCreateDeckFromOutline("Slide 1: Heading\n- Item A\n- Item B");
            Assert(!deckResult.Success, "Headless ExecuteCreateDeckFromOutline should fail safely (no app) without crashing");

            // Context and name safety
            Assert(pptCtrl.GetActiveDocumentName() == "Presentation", "Document name fallback should be 'Presentation'");
            Assert(pptCtrl.GetSelectedText() == string.Empty, "Headless selected text should be empty string");
            Assert(pptCtrl.GetDocumentContext("review", 7000) == string.Empty, "Headless document context should be empty string");
            Assert(pptCtrl.Undo() == false, "Headless undo should return false without crashing");
        }

        private static void TestPowerPointViewStateResilience()
        {
            // Verify PowerPointController contracts when app is null or in non-standard views
            IOfficeHostController hostCtrl = new PowerPointController(null);
            Assert(hostCtrl.HostType == "PowerPoint", "HostType should be PowerPoint");
            Assert(hostCtrl.GetDocumentContext(null, 7000) == string.Empty, "GetDocumentContext should return empty string");
            Assert(hostCtrl.InsertText("Some text") == false, "InsertText should return false on null host");
            Assert(hostCtrl.Undo() == false, "Undo should return false on null host");
        }

        private static void TestEndToEndAttachmentToDeckSimulation()
        {
            // 1. Simulate document attachment text (e.g. from Word or PDF)
            string simulatedDocument =
                "[Attachment: Infrastructure_Audit.docx]\n" +
                "Audit Summary:\n" +
                "1. Signalling reliability index improved to 99.85%.\n" +
                "2. Axle counter upgrades reduced false track vacancy alarms by 72%.\n" +
                "3. Key remaining vulnerability is 13 legacy interlocking cabins slated for CBI migration.";

            // 2. Build prompt
            string deckPrompt = PromptAssembler.BuildBriefingDeckPrompt(
                "Signalling Modernization Status",
                simulatedDocument,
                3);

            Assert(deckPrompt.Contains("Create a concise, executive briefing deck of 3 slides focusing on: Signalling Modernization Status"),
                "Prompt should configure 3 slides");
            Assert(deckPrompt.Contains("Axle counter upgrades reduced false track"), "Prompt should contain audit excerpt");

            // 3. Simulate model completion response
            string modelOutput =
                "Slide 1: Signalling Reliability Performance\n" +
                "- Overall network reliability achieved 99.85% in Q3\n" +
                "- Axle counter upgrades drove 72% reduction in false alarms\n" +
                "Visual suggestion: Line graph comparing monthly reliability trend\n" +
                "Speaker Notes: Credit maintenance teams for rapid fault resolution.\n\n" +
                "Slide 2: Computer-Based Interlocking Migration\n" +
                "- 13 legacy relay interlocking cabins remain active\n" +
                "- Target completion for electronic CBI rollout is Q1 2027\n" +
                "Visual suggestion: Migration phase status heatmap\n" +
                "Speaker Notes: Funding for Phase 2 CBI contracts has been fully approved.";

            // 4. Parse slides from completion
            var parsedSlides = PowerPointActionParser.ParseSlideData(modelOutput);
            Assert(parsedSlides.Count == 2, "Expected 2 parsed slides from model output");
            Assert(parsedSlides[0].Title == "Signalling Reliability Performance", "Slide 1 title mismatch");
            Assert(parsedSlides[0].Bullets.Count == 2, "Slide 1 bullet count mismatch");
            Assert(parsedSlides[1].Title == "Computer-Based Interlocking Migration", "Slide 2 title mismatch");
            Assert(parsedSlides[1].Bullets.Count == 2, "Slide 2 bullet count mismatch");
        }

        private static void Assert(bool condition, string message)
        {
            if (!condition)
            {
                throw new InvalidOperationException("Assertion failed: " + message);
            }
        }
    }
}
