namespace MSOfficeAIAssistant.Core.Session
{
    /// <summary>
    /// Represents the operational mode of the AssistantSession.
    /// Chat mode is read-only; Plan and Edit modes allow mutations.
    /// </summary>
    public enum SessionMode
    {
        Chat = 0,
        Plan = 1,
        Edit = 2
    }
}
