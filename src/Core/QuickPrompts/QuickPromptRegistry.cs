using System;
using System.Collections.Generic;
using System.Linq;

namespace MSOfficeAIAssistant.Core.QuickPrompts
{
    /// <summary>
    /// Represents a quick-prompt chip data model for the UI.
    /// </summary>
    public class QuickPrompt
    {
        public string Id { get; set; }
        public string Label { get; set; }
        public string PromptText { get; set; }
        public string HostFilter { get; set; }
    }

    /// <summary>
    /// Registry of quick prompts, populated with hardcoded entries for Phase A5.
    /// Future Phase B will replace this with skill-driven entries.
    /// </summary>
    public static class QuickPromptRegistry
    {
        /// <summary>
        /// Gets all quick prompts available for the specified host type.
        /// A null HostFilter matches all hosts; a non-null HostFilter only matches when
        /// it equals the hostType (case-insensitive).
        /// </summary>
        public static List<QuickPrompt> GetPrompts(string hostType)
        {
            var allPrompts = new List<QuickPrompt>
            {
                new QuickPrompt
                {
                    Id = "Summarize",
                    Label = "Summarize",
                    PromptText = "Summarize the selected text or document clearly.",
                    HostFilter = null
                },
                new QuickPrompt
                {
                    Id = "Rewrite",
                    Label = "Rewrite",
                    PromptText = "Rewrite the selected text for improved clarity and tone.",
                    HostFilter = null
                },
                new QuickPrompt
                {
                    Id = "Outline",
                    Label = "Outline",
                    PromptText = "Create a clear, hierarchical outline of the supplied content. Preserve key facts and show the recommended narrative flow.",
                    HostFilter = null
                },
                new QuickPrompt
                {
                    Id = "Actions",
                    Label = "Actions",
                    PromptText = "Extract decisions, action items, owners where stated, deadlines where stated, and risks. Do not invent owners or dates.",
                    HostFilter = null
                },
                new QuickPrompt
                {
                    Id = "Review",
                    Label = "Review",
                    PromptText = "Review this content for clarity, gaps, consistency, duplicated ideas, and the most valuable next edits. For presentations, also assess story flow and weak slides.",
                    HostFilter = null
                },
                new QuickPrompt
                {
                    Id = "Deck",
                    Label = "Build deck",
                    PromptText = "Create a concise, coherent slide deck from the supplied content. Return numbered slides with a title, 3-5 concise bullets, a Visual suggestion, and Speaker Notes for each slide.",
                    // Restores the original XAML's isPowerPoint ? Visible : Collapsed rule on
                    // BtnQuickDeck, which the data-driven ItemsControl replaced. Without this
                    // filter, "Build deck" would now show in Word and Excel too, which is a
                    // real behavior change from what shipped before this refactor.
                    HostFilter = "PowerPoint"
                }
            };

            return FilterByHost(allPrompts, hostType);
        }

        /// <summary>
        /// Internal filter logic, factored out for testability.
        /// </summary>
        public static bool MatchesHost(QuickPrompt prompt, string hostType)
        {
            if (prompt == null)
                return false;

            if (prompt.HostFilter == null)
                return true;

            return string.Equals(prompt.HostFilter, hostType, StringComparison.OrdinalIgnoreCase);
        }

        private static List<QuickPrompt> FilterByHost(List<QuickPrompt> prompts, string hostType)
        {
            return prompts.Where(p => MatchesHost(p, hostType)).ToList();
        }

        /// <summary>
        /// Gets all ribbon command prompts consolidated from RibbonCallback hardcoded prompts.
        /// Returns a list of QuickPrompt entries, one per simple (non-dynamic) ribbon method.
        /// All entries have HostFilter = null (ribbon commands are not currently host-conditional).
        /// </summary>
        public static List<QuickPrompt> GetRibbonPrompts()
        {
            return new List<QuickPrompt>
            {
                new QuickPrompt
                {
                    Id = "Generate",
                    Label = "Generate Draft",
                    PromptText = "Generate a comprehensive draft based on the topic or outline provided.",
                    HostFilter = null
                },
                new QuickPrompt
                {
                    Id = "ContinueWriting",
                    Label = "Continue Writing",
                    PromptText = "Continue writing seamlessly from the current point in the text.",
                    HostFilter = null
                },
                new QuickPrompt
                {
                    Id = "Summarize",
                    Label = "Summarize",
                    PromptText = "Provide a concise executive summary highlighting key takeaways and action items.",
                    HostFilter = null
                },
                new QuickPrompt
                {
                    Id = "Rewrite",
                    Label = "Rewrite",
                    PromptText = "Rewrite the selected text for maximum clarity, professional flow, and polished vocabulary.",
                    HostFilter = null
                },
                new QuickPrompt
                {
                    Id = "Expand",
                    Label = "Expand",
                    PromptText = "Elaborate on the selected text with supporting details, explanations, and context.",
                    HostFilter = null
                },
                new QuickPrompt
                {
                    Id = "Shorten",
                    Label = "Shorten",
                    PromptText = "Condense the selected text into a tight, impactful version without losing core meaning.",
                    HostFilter = null
                },
                new QuickPrompt
                {
                    Id = "Outline",
                    Label = "Outline",
                    PromptText = "Create a clear, hierarchical outline of the supplied content. Preserve key facts and show the recommended narrative flow.",
                    HostFilter = null
                },
                new QuickPrompt
                {
                    Id = "ActionItems",
                    Label = "Action Items",
                    PromptText = "Extract decisions, action items, owners where explicitly stated, deadlines where explicitly stated, and risks. Do not invent people or dates.",
                    HostFilter = null
                },
                new QuickPrompt
                {
                    Id = "ReviewContent",
                    Label = "Review Content",
                    PromptText = "Review the supplied content for clarity, gaps, consistency, duplicated ideas, and the most valuable next edits. For presentations, assess story flow and weak slides.",
                    HostFilter = null
                },
                new QuickPrompt
                {
                    Id = "BuildSlides",
                    Label = "Build Slides",
                    PromptText = "Create a concise, coherent slide deck from the supplied content. Return numbered slides with a title, 3-5 concise bullets, a Visual suggestion, and Speaker Notes for each slide.",
                    HostFilter = null
                }
            };
        }
    }
}
