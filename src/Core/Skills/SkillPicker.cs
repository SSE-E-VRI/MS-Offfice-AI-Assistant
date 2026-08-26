using System;
using System.Collections.Generic;
using System.Linq;
using MSOfficeAIAssistant.Core.QuickPrompts;

namespace MSOfficeAIAssistant.Core.Skills
{
    /// <summary>
    /// Converts skills to quick prompts and selects skill-derived chips with context-aware promotion.
    /// </summary>
    public static class SkillPicker
    {
        /// <summary>
        /// Converts a Skill to a QuickPrompt, mapping all relevant fields.
        /// Null input returns null; never throws.
        /// </summary>
        public static QuickPrompt ToQuickPrompt(Skill skill)
        {
            if (skill == null)
                return null;

            return new QuickPrompt
            {
                Id = skill.Id,
                Label = skill.Name,
                PromptText = skill.PromptTemplate,
                HostFilter = skill.PreferredHost
            };
        }

        /// <summary>
        /// Checks if the context text contains failure-related keywords.
        /// Returns true if contextText contains "failure", "breakdown", "fault", or "defect" (case-insensitive).
        /// Null/empty input returns false; never throws.
        /// </summary>
        public static bool IsFailureRelatedContext(string contextText)
        {
            if (string.IsNullOrEmpty(contextText))
                return false;

            string lower = contextText.ToLowerInvariant();
            return lower.Contains("failure") || lower.Contains("breakdown") || lower.Contains("fault") || lower.Contains("defect");
        }

        /// <summary>
        /// Selects skill-derived quick prompts for display, with context-aware promotion.
        ///
        /// Logic:
        /// 1. Loads the domain pack's skills
        /// 2. Filters to skills whose PreferredHost is null OR matches hostType (case-insensitive)
        /// 3. If failure context detected and pack contains failure_analysis_pareto skill, moves it to front
        /// 4. Converts to QuickPrompts, takes at most maxChips entries
        /// 5. Returns empty list on null arguments or empty pack; never throws
        /// </summary>
        public static List<QuickPrompt> SelectChips(string domainPack, string hostType, string contextText, int maxChips)
        {
            // Validate maxChips
            if (maxChips <= 0)
                return new List<QuickPrompt>();

            // Validate required arguments
            if (string.IsNullOrWhiteSpace(hostType))
                return new List<QuickPrompt>();

            // Load pack with null/empty safety
            List<Skill> skills = SkillRegistry.LoadPack(domainPack);
            if (skills == null || skills.Count == 0)
                return new List<QuickPrompt>();

            // Filter by host compatibility
            List<Skill> filtered = new List<Skill>();
            foreach (Skill skill in skills)
            {
                if (skill == null)
                    continue;

                // Null PreferredHost matches all hosts; otherwise match case-insensitively
                if (skill.PreferredHost == null || string.Equals(skill.PreferredHost, hostType, StringComparison.OrdinalIgnoreCase))
                {
                    filtered.Add(skill);
                }
            }

            if (filtered.Count == 0)
                return new List<QuickPrompt>();

            // Promotion: if failure context detected, move failure_analysis_pareto to front
            if (IsFailureRelatedContext(contextText))
            {
                Skill failureSkill = null;
                foreach (Skill skill in filtered)
                {
                    if (skill.Id == "failure_analysis_pareto")
                    {
                        failureSkill = skill;
                        break;
                    }
                }

                if (failureSkill != null)
                {
                    filtered.Remove(failureSkill);
                    filtered.Insert(0, failureSkill);
                }
            }

            // Convert to QuickPrompts and limit to maxChips
            List<QuickPrompt> result = new List<QuickPrompt>();
            for (int i = 0; i < filtered.Count && result.Count < maxChips; i++)
            {
                QuickPrompt qp = ToQuickPrompt(filtered[i]);
                if (qp != null)
                    result.Add(qp);
            }

            return result;
        }
    }
}
