using System;
using System.Collections.Generic;

namespace MSOfficeAIAssistant.Core.Session
{
    /// <summary>
    /// Status states representing the lifecycle of an AI Assistant operation.
    /// These states reflect real transitions in SendMessageAsync, ExecuteOfficeAction, and ConfirmOfficeAction.
    /// </summary>
    public enum AssistantStatus
    {
        Ready,
        Thinking,
        ReadingAttachments,
        AwaitingApproval,
        Applying,
        Verifying,
        Done,
        Failed,
        Cancelled
    }

    /// <summary>
    /// Helper class providing display labels and icons for AssistantStatus states.
    /// Ensures every status has a unique, non-empty label and icon.
    /// </summary>
    public static class AssistantStatusDisplay
    {
        /// <summary>
        /// Gets a short, human-readable label for the given status.
        /// Never returns null or empty; suitable for screen reader announcements.
        /// </summary>
        public static string GetLabel(AssistantStatus status)
        {
            switch (status)
            {
                case AssistantStatus.Ready:
                    return "Ready";
                case AssistantStatus.Thinking:
                    return "Thinking…";
                case AssistantStatus.ReadingAttachments:
                    return "Reading attachments…";
                case AssistantStatus.AwaitingApproval:
                    return "Awaiting approval";
                case AssistantStatus.Applying:
                    return "Applying…";
                case AssistantStatus.Verifying:
                    return "Verifying…";
                case AssistantStatus.Done:
                    return "Done";
                case AssistantStatus.Failed:
                    return "Failed";
                case AssistantStatus.Cancelled:
                    return "Cancelled";
                default:
                    return "Unknown";
            }
        }

        /// <summary>
        /// Gets a visually distinct icon/glyph for the given status.
        /// Each status has a unique icon so status is never conveyed by color alone.
        /// </summary>
        public static string GetIcon(AssistantStatus status)
        {
            switch (status)
            {
                case AssistantStatus.Ready:
                    return "●";
                case AssistantStatus.Thinking:
                    return "💭";
                case AssistantStatus.ReadingAttachments:
                    return "📄";
                case AssistantStatus.AwaitingApproval:
                    return "⏸";
                case AssistantStatus.Applying:
                    return "⚙";
                case AssistantStatus.Verifying:
                    return "🔍";
                case AssistantStatus.Done:
                    return "✓";
                case AssistantStatus.Failed:
                    return "⚠";
                case AssistantStatus.Cancelled:
                    return "⏹";
                default:
                    return "?";
            }
        }
    }
}
