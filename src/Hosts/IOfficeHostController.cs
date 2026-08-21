using System;

namespace MSOfficeAIAssistant.Hosts
{
    /// <summary>
    /// Common host controller contract for Word, Excel, and PowerPoint.
    /// Exposes shared operations (document name, selected text, contextual snapshot, text insert, undo)
    /// while leaving host-specific operations (e.g. Word Track Changes, Excel SpreadsheetActions, PowerPoint deck actions)
    /// on typed controllers and the upcoming Tool Registry.
    /// </summary>
    public interface IOfficeHostController
    {
        string HostType { get; }
        string GetActiveDocumentName();
        string GetSelectedText();
        string GetDocumentContext(string prompt, int maxCharacters);
        bool InsertText(string text);
        bool Undo();
    }
}
