using System;
using System.Collections.Generic;
using MSOfficeAIAssistant.Core.Session;

namespace MSOfficeAIAssistant.Tests
{
    public static class AssistantStatusTests
    {
        public static void RunAll()
        {
            TestAllStatusesHaveNonEmptyLabels();
            TestAllStatusesHaveNonEmptyIcons();
            TestAllStatusIconsAreUnique();
        }

        private static void TestAllStatusesHaveNonEmptyLabels()
        {
            var values = (AssistantStatus[])Enum.GetValues(typeof(AssistantStatus));
            foreach (var status in values)
            {
                string label = AssistantStatusDisplay.GetLabel(status);
                Assert(!string.IsNullOrEmpty(label), string.Format("Status {0} has empty or null label", status));
            }
        }

        private static void TestAllStatusesHaveNonEmptyIcons()
        {
            var values = (AssistantStatus[])Enum.GetValues(typeof(AssistantStatus));
            foreach (var status in values)
            {
                string icon = AssistantStatusDisplay.GetIcon(status);
                Assert(!string.IsNullOrEmpty(icon), string.Format("Status {0} has empty or null icon", status));
            }
        }

        private static void TestAllStatusIconsAreUnique()
        {
            var values = (AssistantStatus[])Enum.GetValues(typeof(AssistantStatus));
            var seenIcons = new HashSet<string>();
            foreach (var status in values)
            {
                string icon = AssistantStatusDisplay.GetIcon(status);
                Assert(!seenIcons.Contains(icon), string.Format("Icon '{0}' is used by multiple statuses (last was {1})", icon, status));
                seenIcons.Add(icon);
            }
        }

        private static void Assert(bool condition, string message)
        {
            if (!condition)
            {
                throw new Exception(message);
            }
        }
    }
}
