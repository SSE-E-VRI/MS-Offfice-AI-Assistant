using System;
using System.Globalization;
using System.Windows.Data;

namespace MSOfficeAIAssistant.UI.Converters
{
    /// <summary>
    /// Converts a stored UTC DateTime (e.g. ActionAuditEntry.TimestampUtc,
    /// ConversationSessionSummary.LastUpdatedUtc) to the viewer's local time for display.
    /// One-way: history/audit timestamps are never edited back through the UI.
    /// </summary>
    public class UtcToLocalTimeConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is DateTime)
            {
                DateTime dt = (DateTime)value;
                DateTime utc = dt.Kind == DateTimeKind.Utc ? dt : DateTime.SpecifyKind(dt, DateTimeKind.Utc);
                return utc.ToLocalTime();
            }
            return value;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotSupportedException("UtcToLocalTimeConverter is one-way.");
        }
    }
}
