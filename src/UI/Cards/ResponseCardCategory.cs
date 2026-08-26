using System;
using MSOfficeAIAssistant.API.Models;

namespace MSOfficeAIAssistant.UI.Cards
{
    /// <summary>
    /// Defines the display category of a chat response message.
    /// Used by the UI template selector to choose the appropriate visual presentation.
    /// </summary>
    public enum ResponseCardCategory
    {
        Text,
        ActionPreview,
        Plan,
        Warning,
        Finding,
        Recommendation,
        Summary
    }

    /// <summary>
    /// Pure, testable classifier that categorizes ChatMessages for response card rendering.
    /// No WPF dependencies — usable from COM-free test projects.
    /// </summary>
    public static class ResponseCardCategoryClassifier
    {
        /// <summary>
        /// Classifies a ChatMessage into one of the response card categories.
        /// Priority order: Plan (if HasPlan), then ActionPreview (if HasOfficeActions),
        /// then Warning/Finding/Recommendation/Summary content-prefix markers (in that order,
        /// case-insensitive, after trimming), else Text.
        /// </summary>
        public static ResponseCardCategory Classify(ChatMessage message)
        {
            // Null-safe: null message → Text (default)
            if (message == null)
            {
                return ResponseCardCategory.Text;
            }

            // Priority 1: Plan (checked first, defensively before ActionPreview)
            if (message.HasPlan)
            {
                return ResponseCardCategory.Plan;
            }

            // Priority 2: ActionPreview (exact today's behavior)
            if (message.HasOfficeActions)
            {
                return ResponseCardCategory.ActionPreview;
            }

            // Priority 3-6: Content-prefix markers (Warning → Finding → Recommendation → Summary)
            string content = message.Content;
            if (string.IsNullOrEmpty(content))
            {
                return ResponseCardCategory.Text;
            }

            string trimmed = content.Trim();

            // Warning: starts with "**Warning:**", "Warning:", or "⚠"
            if (StartsWithIgnoreCase(trimmed, "**Warning:**") ||
                StartsWithIgnoreCase(trimmed, "Warning:") ||
                trimmed.StartsWith("⚠", StringComparison.Ordinal))
            {
                return ResponseCardCategory.Warning;
            }

            // Finding: starts with "**Finding:**" or "Finding:"
            if (StartsWithIgnoreCase(trimmed, "**Finding:**") ||
                StartsWithIgnoreCase(trimmed, "Finding:"))
            {
                return ResponseCardCategory.Finding;
            }

            // Recommendation: starts with "**Recommendation:**" or "Recommendation:"
            if (StartsWithIgnoreCase(trimmed, "**Recommendation:**") ||
                StartsWithIgnoreCase(trimmed, "Recommendation:"))
            {
                return ResponseCardCategory.Recommendation;
            }

            // Summary: starts with "**Summary:**" or "Summary:"
            if (StartsWithIgnoreCase(trimmed, "**Summary:**") ||
                StartsWithIgnoreCase(trimmed, "Summary:"))
            {
                return ResponseCardCategory.Summary;
            }

            // Default: plain text
            return ResponseCardCategory.Text;
        }

        /// <summary>
        /// Case-insensitive string starts-with check.
        /// </summary>
        private static bool StartsWithIgnoreCase(string text, string prefix)
        {
            return text.StartsWith(prefix, StringComparison.OrdinalIgnoreCase);
        }
    }
}
