using System;
using System.Collections.Generic;
using MSOfficeAIAssistant.Core;

namespace MSOfficeAIAssistant.Tests
{
    public static class PowerPointActionParserTests
    {
        public static void RunAll()
        {
            TestParseStructuredActions();
            TestParseSlideData();
            TestCleanMarkdown();
        }

        private static void TestParseStructuredActions()
        {
            string aiResponse = "Here are the suggested structural changes:\n\n" +
                                "<powerpoint_actions>\n" +
                                "  <powerpoint_action type=\"create_section\" name=\"Financial Results\" slide=\"3\" />\n" +
                                "  <powerpoint_action type=\"move_slide\" source=\"5\" target=\"2\" />\n" +
                                "  <powerpoint_action type=\"set_notes\" slide=\"1\" notes=\"Welcome everyone to Q3 review\" />\n" +
                                "</powerpoint_actions>";

            List<PowerPointAction> actions = PowerPointActionParser.ParseStructuredActions(aiResponse);

            Assert(actions != null && actions.Count == 3, "Parsed 3 powerpoint actions");
            Assert(actions[0].Type == "create_section", "Action 0 type create_section");
            Assert(actions[0].Name == "Financial Results", "Action 0 name Financial Results");
            Assert(actions[0].Slide == 3, "Action 0 slide 3");

            Assert(actions[1].Type == "move_slide", "Action 1 type move_slide");
            Assert(actions[1].Source == 5, "Action 1 source 5");
            Assert(actions[1].Target == 2, "Action 1 target 2");

            Assert(actions[2].Type == "set_notes", "Action 2 type set_notes");
            Assert(actions[2].Slide == 1, "Action 2 slide 1");
            Assert(actions[2].Notes == "Welcome everyone to Q3 review", "Action 2 notes");
        }

        private static void TestParseSlideData()
        {
            string markdown = "# Slide 1: Introduction\n" +
                              "- Welcome to the session\n" +
                              "- Agenda overview\n\n" +
                              "Speaker Notes: Be brief on intro\n" +
                              "Visual: Photo of team\n\n" +
                              "# Slide 2: Next Steps\n" +
                              "- Rollout schedule\n" +
                              "- QA sign-off\n\n" +
                              "Notes: Emphasize timeline";

            List<SlideData> slides = PowerPointActionParser.ParseSlideData(markdown);

            Assert(slides != null && slides.Count == 2, "Parsed 2 slides");
            Assert(slides[0].Title == "Introduction", "Slide 0 title");
            Assert(slides[0].Bullets != null && slides[0].Bullets.Count == 2, "Slide 0 bullets count");
            Assert(slides[0].Bullets[0] == "Welcome to the session", "Slide 0 bullet 1");
            Assert(slides[0].SpeakerNotes == "Be brief on intro", "Slide 0 speaker notes");
            Assert(slides[0].VisualSuggestion == "Photo of team", "Slide 0 visual suggestion");

            Assert(slides[1].Title == "Next Steps", "Slide 1 title");
            Assert(slides[1].Bullets != null && slides[1].Bullets.Count == 2, "Slide 1 bullets count");
            Assert(slides[1].SpeakerNotes == "Emphasize timeline", "Slide 1 speaker notes");
        }

        private static void TestCleanMarkdown()
        {
            string raw = "```markdown\n# Title\nSome content\n<powerpoint_actions>\n<action/>\n</powerpoint_actions>\n```";
            string cleaned = PowerPointActionParser.CleanMarkdown(raw);

            Assert(!string.IsNullOrEmpty(cleaned), "Cleaned markdown should not be empty");
            Assert(cleaned.IndexOf("```", StringComparison.Ordinal) < 0, "Fenced backticks stripped");
            Assert(cleaned.IndexOf("<powerpoint_actions>", StringComparison.Ordinal) < 0, "PowerPoint action XML stripped");
            Assert(cleaned.IndexOf("Title", StringComparison.Ordinal) >= 0, "Slide title preserved");
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
