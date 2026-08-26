using System;
using System.Globalization;
using System.Windows.Data;
using MSOfficeAIAssistant.UI.Cards;

namespace MSOfficeAIAssistant.UI.Converters
{
    /// <summary>
    /// Converts a Finding's Content text into evidence-level display properties:
    /// Label (e.g. "Directly Observed"), Icon (e.g. "✓"), or DisplayText (Content with tag stripped).
    ///
    /// ConverterParameter specifies what to output:
    ///   "Label" → EvidenceLevelClassifier.GetLabel(Classify(...))
    ///   "Icon" → EvidenceLevelClassifier.GetIcon(...)
    ///   "DisplayText" → EvidenceLevelClassifier.StripEvidenceTag(...) (for rendering in TextBlock)
    ///
    /// One-way converter (FindingTemplate never edits evidence level through the UI).
    /// </summary>
    public class EvidenceLevelConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            string content = value as string;
            if (content == null)
            {
                content = "";
            }

            string param = (parameter as string) ?? "Label";

            if (param.Equals("Label", StringComparison.OrdinalIgnoreCase))
            {
                EvidenceLevel level = EvidenceLevelClassifier.Classify(content);
                return EvidenceLevelClassifier.GetLabel(level);
            }
            else if (param.Equals("Icon", StringComparison.OrdinalIgnoreCase))
            {
                EvidenceLevel level = EvidenceLevelClassifier.Classify(content);
                return EvidenceLevelClassifier.GetIcon(level);
            }
            else if (param.Equals("DisplayText", StringComparison.OrdinalIgnoreCase))
            {
                return EvidenceLevelClassifier.StripEvidenceTag(content);
            }

            // Unknown parameter → return as-is
            return value;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotSupportedException("EvidenceLevelConverter is one-way.");
        }
    }
}
