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
    }
}
