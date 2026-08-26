using System.Windows;
using System.Windows.Controls;
using MSOfficeAIAssistant.API.Models;

namespace MSOfficeAIAssistant.UI.Cards
{
    /// <summary>
    /// DataTemplateSelector that picks the appropriate template for a ChatMessage
    /// based on its ResponseCardCategory classification.
    /// </summary>
    public class ResponseCardTemplateSelector : DataTemplateSelector
    {
        /// <summary>
        /// Template for plain text responses (no special styling).
        /// </summary>
        public DataTemplate TextTemplate { get; set; }

        /// <summary>
        /// Template for responses with office actions (action preview cards).
        /// </summary>
        public DataTemplate ActionPreviewTemplate { get; set; }

        /// <summary>
        /// Template for warning responses (amber/warning accent border).
        /// </summary>
        public DataTemplate WarningTemplate { get; set; }

        /// <summary>
        /// Template for finding responses (blue/info accent border).
        /// </summary>
        public DataTemplate FindingTemplate { get; set; }

        /// <summary>
        /// Template for recommendation responses (green/success accent border).
        /// </summary>
        public DataTemplate RecommendationTemplate { get; set; }

        /// <summary>
        /// Template for summary responses (purple/neutral accent border).
        /// </summary>
        public DataTemplate SummaryTemplate { get; set; }

        /// <summary>
        /// Template for plan responses (multi-step execution preview).
        /// </summary>
        public DataTemplate PlanTemplate { get; set; }

        /// <summary>
        /// Selects the appropriate template for the given item (expected to be a ChatMessage).
        /// Falls back to TextTemplate if the classified category's template is not assigned.
        /// Never returns null.
        /// </summary>
        public override DataTemplate SelectTemplate(object item, DependencyObject container)
        {
            ChatMessage message = item as ChatMessage;
            if (message == null)
            {
                return TextTemplate ?? new DataTemplate();
            }

            ResponseCardCategory category = ResponseCardCategoryClassifier.Classify(message);

            switch (category)
            {
                case ResponseCardCategory.Plan:
                    return PlanTemplate ?? TextTemplate ?? new DataTemplate();
                case ResponseCardCategory.ActionPreview:
                    return ActionPreviewTemplate ?? TextTemplate ?? new DataTemplate();
                case ResponseCardCategory.Warning:
                    return WarningTemplate ?? TextTemplate ?? new DataTemplate();
                case ResponseCardCategory.Finding:
                    return FindingTemplate ?? TextTemplate ?? new DataTemplate();
                case ResponseCardCategory.Recommendation:
                    return RecommendationTemplate ?? TextTemplate ?? new DataTemplate();
                case ResponseCardCategory.Summary:
                    return SummaryTemplate ?? TextTemplate ?? new DataTemplate();
                default:
                    return TextTemplate ?? new DataTemplate();
            }
        }
    }
}
