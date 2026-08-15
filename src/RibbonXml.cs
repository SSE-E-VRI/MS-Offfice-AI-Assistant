namespace MistralOfficeAddin
{
    /// <summary>
    /// Ribbon XML for all three hosts. The 2006/01 customUI schema loads
    /// identically in Office 2010, 2013, 2016, 2019, 2021, 2024 and 365.
    /// Callbacks are resolved by name via IDispatch on the Connect object.
    /// </summary>
    internal static class RibbonXml
    {
        public const string Ribbon =
            "<customUI xmlns=\"http://schemas.microsoft.com/office/2006/01/customUI\">" +
            "  <ribbon>" +
            "    <tabs>" +
            "      <tab id=\"tabMistralAI\" label=\"Mistral AI\">" +
            "        <group id=\"grpMistralAssistant\" label=\"AI Assistant\">" +
            "          <button id=\"btnMistralChat\" label=\"Chat Pane\" size=\"large\"" +
            "                  imageMso=\"HappyFace\" onAction=\"OnChatButtonClick\" />" +
            "          <button id=\"btnMistralSettings\" label=\"Settings\" size=\"large\"" +
            "                  imageMso=\"Info\" onAction=\"OnSettingsButtonClick\" />" +
            "        </group>" +
            "      </tab>" +
            "    </tabs>" +
            "  </ribbon>" +
            "</customUI>";
    }
}
