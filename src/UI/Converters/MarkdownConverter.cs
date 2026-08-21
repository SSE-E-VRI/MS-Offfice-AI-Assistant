using System;
using System.Globalization;
using System.Windows;
using System.Windows.Data;
using System.Windows.Documents;
using Markdig;
using Markdig.Wpf;

namespace MSOfficeAIAssistant.UI.Converters
{
    [ValueConversion(typeof(string), typeof(FlowDocument))]
    public class MarkdownToFlowDocumentConverter : IValueConverter
    {
        private static readonly MarkdownPipeline _pipeline = new MarkdownPipelineBuilder()
            .UseSupportedExtensions()
            .Build();

        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is string)
            {
                string markdown = (string)value;
                if (!string.IsNullOrWhiteSpace(markdown))
                {
                    try
                    {
                        var doc = Markdig.Wpf.Markdown.ToFlowDocument(markdown, _pipeline);
                        doc.PagePadding = new Thickness(0);
                        return doc;
                    }
                    catch
                    {
                        var fallbackDoc = new FlowDocument();
                        fallbackDoc.PagePadding = new Thickness(0);
                        fallbackDoc.Blocks.Add(new Paragraph(new Run(markdown)));
                        return fallbackDoc;
                    }
                }
            }

            var emptyDoc = new FlowDocument();
            emptyDoc.PagePadding = new Thickness(0);
            return emptyDoc;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotSupportedException();
        }
    }

    public class BooleanToVisibilityConverter : IValueConverter
    {
        public bool Invert { get; set; }

        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            bool boolVal = (value is bool) && (bool)value;
            if (Invert) boolVal = !boolVal;
            return boolVal ? Visibility.Visible : Visibility.Collapsed;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is Visibility)
            {
                Visibility v = (Visibility)value;
                bool res = (v == Visibility.Visible);
                return Invert ? !res : res;
            }
            return false;
        }
    }
}
